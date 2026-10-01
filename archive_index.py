"""archive_index.py — 범용 부품 5 `search_archive`의 1차(서버 불필요) 버전(PRD 16-5 ⑥).

과거 보고서 폴더를 훑어 문단 덩어리(chunk)로 쪼개 SQLite에 넣고, 질문과의
BM25 점수로 관련 덩어리를 찾는다. 임베딩 모델 없이도 "어느 파일 어느 자리에
그 얘기가 있나"를 찾는 데는 충분하고(기관명·사업명·연도 같은 고유 표현은
키워드 일치가 오히려 정확함), 나중에 서버의 임베딩 점수를 두 번째 점수로
합치면 된다(하이브리드 — PRD 16-4의 "카드 먼저 보고 표는 직접 읽기" 흐름의
검색 부분).

한국어 토큰화: 형태소 분석기 없이 **한글 연속 구간의 2글자 n-gram + 영문/숫자
토큰**을 쓴다("서류심사" → 서류, 류심, 심사). 조사·어미가 붙어도 앞부분
n-gram이 겹쳐 검색이 된다. 흔한 조사 n-gram이 점수를 흐리는 문제는 BM25의
IDF가 자연스럽게 눌러준다.

읽기는 부품 4(attachments.read_attachment)를 그대로 쓴다 — .hwp는 파일마다
새 한글 프로세스를 띄우므로(격리) 첫 색인은 파일당 수 초가 걸린다. 그래서
파일 수정시각(mtime)이 같으면 건너뛰는 증분 색인이다.

사용:
    python archive_index.py --add-folder "D:\\과거보고서"      # 폴더 등록
    python archive_index.py --reindex                        # 증분 색인
    python archive_index.py --search "2024년 서류심사 건수"  # 검색
설정(폴더 목록)과 색인 DB는 사용자 데이터라 git에 넣지 않는다(.gitignore)."""
import json
import math
import os
import re
import sqlite3
import sys
import time

CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "archive_folders.json")
DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "archive_index.sqlite")
CHUNK_CHARS = 700
CHUNK_OVERLAP = 120
SUPPORTED = (".hwp", ".hwpx", ".xlsx", ".xlsm", ".pdf", ".txt", ".md", ".csv")


# ---------------------------------------------------------------- 토큰화

_HANGUL_RUN = re.compile(r"[가-힣]+")
_ALNUM = re.compile(r"[A-Za-z0-9][A-Za-z0-9.,]*")


def tokenize(text: str) -> list:
    tokens = []
    for run in _HANGUL_RUN.findall(text):
        if len(run) == 1:
            tokens.append(run)
        else:
            tokens.extend(run[i:i + 2] for i in range(len(run) - 1))
    tokens.extend(t.strip(".,").lower() for t in _ALNUM.findall(text) if t.strip(".,"))
    return tokens


def guess_year(name: str):
    m = re.search(r"(20\d{2})", name)
    return int(m.group(1)) if m else None


def chunk_text(text: str, size: int = CHUNK_CHARS, overlap: int = CHUNK_OVERLAP) -> list:
    """문단 경계를 존중하며 size 안팎으로 자른다. 긴 문단은 overlap을 두고 자른다."""
    paras = [p.strip() for p in text.replace("\r\n", "\n").split("\n") if p.strip()]
    chunks, cur = [], ""
    for p in paras:
        if len(p) > size:
            if cur:
                chunks.append(cur); cur = ""
            start = 0
            while start < len(p):
                chunks.append(p[start:start + size])
                start += size - overlap
            continue
        if len(cur) + len(p) + 1 > size and cur:
            chunks.append(cur)
            cur = p
        else:
            cur = (cur + "\n" + p) if cur else p
    if cur:
        chunks.append(cur)
    return chunks


# ---------------------------------------------------------------- 설정·DB

def load_folders(path: str = CONFIG_PATH) -> list:
    if not os.path.exists(path):
        return []
    with open(path, encoding="utf-8") as f:
        return json.load(f).get("folders", [])


def save_folders(folders: list, path: str = CONFIG_PATH) -> None:
    with open(path, "w", encoding="utf-8") as f:
        json.dump({"folders": folders}, f, ensure_ascii=False, indent=2)


def _connect(db_path: str):
    conn = sqlite3.connect(db_path)
    conn.execute("""CREATE TABLE IF NOT EXISTS docs(
        id INTEGER PRIMARY KEY, path TEXT UNIQUE, name TEXT, mtime REAL, year INTEGER,
        error TEXT, indexed_at REAL)""")
    conn.execute("""CREATE TABLE IF NOT EXISTS chunks(
        id INTEGER PRIMARY KEY, doc_id INTEGER, label TEXT, text TEXT, tokens TEXT, length INTEGER,
        FOREIGN KEY(doc_id) REFERENCES docs(id))""")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_chunks_doc ON chunks(doc_id)")
    return conn


def _iter_files(folders: list):
    for folder in folders:
        for root, _dirs, files in os.walk(folder):
            for f in sorted(files):
                if f.startswith("~$") or f.startswith("_test"):
                    continue
                if os.path.splitext(f)[1].lower() in SUPPORTED:
                    yield os.path.join(root, f)


def reindex(folders: list = None, db_path: str = DB_PATH, progress=None, reader=None) -> dict:
    """증분 색인. mtime이 같은 파일은 건너뛴다. reader는 테스트용 주입(기본 attachments.read_attachment).
    반환: {"indexed": n, "skipped": n, "failed": [(path, error)], "removed": n}"""
    if reader is None:
        from attachments import read_attachment as reader
    folders = load_folders() if folders is None else folders
    conn = _connect(db_path)
    stats = {"indexed": 0, "skipped": 0, "failed": [], "removed": 0}
    seen = set()
    try:
        for path in _iter_files(folders):
            key = os.path.normcase(os.path.abspath(path))
            seen.add(key)
            mtime = os.path.getmtime(path)
            row = conn.execute("SELECT id, mtime, error FROM docs WHERE path=?", (key,)).fetchone()
            if row and abs(row[1] - mtime) < 1e-6 and not row[2]:
                stats["skipped"] += 1
                continue
            if progress:
                progress(f"색인 중: {os.path.basename(path)}")
            result = reader(path)
            name = os.path.basename(path)
            if row:
                conn.execute("DELETE FROM chunks WHERE doc_id=?", (row[0],))
                conn.execute("UPDATE docs SET mtime=?, error=?, indexed_at=?, year=? WHERE id=?",
                             (mtime, result.get("error"), time.time(), guess_year(name), row[0]))
                doc_id = row[0]
            else:
                cur = conn.execute("INSERT INTO docs(path, name, mtime, year, error, indexed_at) VALUES(?,?,?,?,?,?)",
                                   (key, name, mtime, guess_year(name), result.get("error"), time.time()))
                doc_id = cur.lastrowid
            if result.get("error"):
                stats["failed"].append((path, result["error"]))
                continue
            for part in result["parts"]:
                for chunk in chunk_text(part["text"]):
                    toks = tokenize(chunk)
                    conn.execute("INSERT INTO chunks(doc_id, label, text, tokens, length) VALUES(?,?,?,?,?)",
                                 (doc_id, part["label"], chunk, " ".join(toks), len(toks)))
            stats["indexed"] += 1
        # 사라진 파일 정리
        for (doc_id, path) in conn.execute("SELECT id, path FROM docs").fetchall():
            if path not in seen:
                conn.execute("DELETE FROM chunks WHERE doc_id=?", (doc_id,))
                conn.execute("DELETE FROM docs WHERE id=?", (doc_id,))
                stats["removed"] += 1
        conn.commit()
    finally:
        conn.close()
    return stats


# ---------------------------------------------------------------- 검색(BM25)

def search_archive(query: str, k: int = 5, year: int = None, db_path: str = DB_PATH,
                   name_contains: str = None) -> list:
    """BM25로 관련 덩어리를 찾는다. 반환: [{"file", "path", "label", "text", "score", "year"}]"""
    q_tokens = tokenize(query)
    if not q_tokens:
        return []
    conn = _connect(db_path)
    try:
        where, params = [], []
        if year is not None:
            where.append("d.year=?"); params.append(year)
        if name_contains:
            where.append("d.name LIKE ?"); params.append(f"%{name_contains}%")
        sql = ("SELECT c.id, c.tokens, c.length, c.text, c.label, d.name, d.path, d.year "
               "FROM chunks c JOIN docs d ON c.doc_id=d.id" + (" WHERE " + " AND ".join(where) if where else ""))
        rows = conn.execute(sql, params).fetchall()
    finally:
        conn.close()
    if not rows:
        return []
    n = len(rows)
    avgdl = sum(r[2] for r in rows) / n or 1.0
    # 문서 빈도
    df = {}
    docs_tokens = []
    for r in rows:
        toks = r[1].split(" ") if r[1] else []
        docs_tokens.append(toks)
        for t in set(toks):
            df[t] = df.get(t, 0) + 1
    k1, b = 1.5, 0.75
    scored = []
    q_set = set(q_tokens)
    for r, toks in zip(rows, docs_tokens):
        if not toks:
            continue
        tf = {}
        for t in toks:
            if t in q_set:
                tf[t] = tf.get(t, 0) + 1
        if not tf:
            continue
        dl = r[2]
        score = 0.0
        for t, f in tf.items():
            idf = math.log(1 + (n - df[t] + 0.5) / (df[t] + 0.5))
            score += idf * (f * (k1 + 1)) / (f + k1 * (1 - b + b * dl / avgdl))
        scored.append((score, r))
    scored.sort(key=lambda x: x[0], reverse=True)
    return [{"file": r[5], "path": r[6], "label": r[4], "text": r[3], "score": round(s, 3), "year": r[7]}
            for s, r in scored[:k]]


def format_results(results: list) -> str:
    return "\n\n".join(f"[{r['file']} · {r['label']}] (점수 {r['score']})\n{r['text']}" for r in results)


def index_stats(db_path: str = DB_PATH) -> dict:
    conn = _connect(db_path)
    try:
        docs = conn.execute("SELECT COUNT(*) FROM docs WHERE error IS NULL").fetchone()[0]
        failed = conn.execute("SELECT COUNT(*) FROM docs WHERE error IS NOT NULL").fetchone()[0]
        chunks = conn.execute("SELECT COUNT(*) FROM chunks").fetchone()[0]
    finally:
        conn.close()
    return {"docs": docs, "failed": failed, "chunks": chunks}


# ---------------------------------------------------------------- self-tests (서버·한글 불필요)

def _selftest_tokenize_and_chunk():
    assert tokenize("서류심사") == ["서류", "류심", "심사"], tokenize("서류심사")
    toks = tokenize("서류심사 건수 2024년 120건")
    assert "서류" in toks and "심사" in toks and "2024" in toks and "120" in toks, toks
    assert guess_year("2024년 서류심사 기본계획(안).hwp") == 2024 and guess_year("메모.txt") is None
    chunks = chunk_text("가\n" * 10 + ("나" * 1500) + "\n다", size=100, overlap=10)
    assert all(len(c) <= 100 for c in chunks) and chunks[-1] == "다", [len(c) for c in chunks]
    print("tokenize/guess_year/chunk_text 통과")


def _selftest_reindex_and_search_incremental():
    import tempfile, shutil
    d = tempfile.mkdtemp(prefix="_arch_")
    db = os.path.join(d, "idx.sqlite")
    try:
        p1 = os.path.join(d, "2024년 서류심사 결과보고.txt")
        p2 = os.path.join(d, "2025년 현장점검 계획.txt")
        open(p1, "w", encoding="utf-8").write("1. 추진 결과\n2024년 서류심사 건수는 총 120건이며 전년 대비 증가하였다.\n2. 향후 계획\n심사 기간 단축")
        open(p2, "w", encoding="utf-8").write("현장점검은 3월에 착수하여 상반기 7회 실시한다.\n점검 인원은 5명이다.")
        open(os.path.join(d, "x.docx"), "wb").close()  # 미지원 → 무시
        st = reindex([d], db_path=db)
        assert st["indexed"] == 2 and st["skipped"] == 0 and st["failed"] == [], st
        st = reindex([d], db_path=db)
        assert st["indexed"] == 0 and st["skipped"] == 2, st  # 증분

        r = search_archive("재작년 서류심사 건수 몇 건", db_path=db)
        assert r and r[0]["file"].startswith("2024년 서류심사") and "120건" in r[0]["text"] and r[0]["year"] == 2024, r[:1]
        r = search_archive("현장점검 몇 회", db_path=db, year=2025)
        assert r and "7회" in r[0]["text"], r
        assert search_archive("현장점검", db_path=db, year=2024) == []
        assert search_archive("", db_path=db) == []
        assert "[2024년 서류심사 결과보고.txt · 본문]" in format_results(search_archive("서류심사", db_path=db, k=1))

        # 수정 → 재색인, 삭제 → 정리
        time.sleep(0.01)
        open(p2, "w", encoding="utf-8").write("현장점검은 4월에 착수한다.")
        os.utime(p2, None)
        os.remove(p1)
        st = reindex([d], db_path=db)
        assert st["indexed"] == 1 and st["removed"] == 1, st
        assert search_archive("서류심사", db_path=db) == []
        assert "4월" in search_archive("현장점검 착수", db_path=db)[0]["text"]
        assert index_stats(db) == {"docs": 1, "failed": 0, "chunks": 1}, index_stats(db)

        # 읽기 실패 기록
        bad = os.path.join(d, "깨짐.pdf")
        open(bad, "wb").write(b"not a pdf")
        st = reindex([d], db_path=db)
        assert len(st["failed"]) == 1 and index_stats(db)["failed"] == 1, (st, index_stats(db))
        print("reindex/search_archive(증분·연도 필터·수정·삭제·실패 기록) 통과")
    finally:
        shutil.rmtree(d, ignore_errors=True)


def _main(argv):
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--add-folder"); ap.add_argument("--reindex", action="store_true")
    ap.add_argument("--search"); ap.add_argument("--year", type=int); ap.add_argument("--stats", action="store_true")
    a = ap.parse_args(argv)
    if a.add_folder:
        folders = load_folders()
        if a.add_folder not in folders:
            folders.append(a.add_folder); save_folders(folders)
        print("폴더:", folders)
    if a.reindex:
        print(reindex(progress=print))
    if a.search:
        print(format_results(search_archive(a.search, year=a.year)) or "(결과 없음)")
    if a.stats:
        print(index_stats())


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    if len(sys.argv) > 1:
        _main(sys.argv[1:])
    else:
        _selftest_tokenize_and_chunk()
        _selftest_reindex_and_search_incremental()
