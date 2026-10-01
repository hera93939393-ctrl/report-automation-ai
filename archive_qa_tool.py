"""archive_qa_tool.py — 과거 문서 아카이브 질의응답(PRD 16-4 "동료" 모드의 첫 조각).

archive_index.search_archive(BM25)로 관련 덩어리를 찾고, attachment_qa_tool의
answer_from_context로 "출처를 붙여, 없으면 없다고" 답한다. 사업 카드(요약)는
서버 LLM이 있어야 만들 수 있어 다음 단계로 두고, 여기서는 검색 → 답변만.

색인은 질문 때마다 증분으로 갱신한다(mtime 같으면 건너뛰므로 평소엔 즉시).
.hwp가 새로 추가되면 파일당 수 초가 걸리므로 progress 콜백으로 알린다."""
import os

from archive_index import (DB_PATH, format_results, index_stats, load_folders, reindex,
                           search_archive)
from attachment_qa_tool import answer_from_context


def ask_archive(question: str, client=None, folders: list = None, db_path: str = DB_PATH,
                year: int = None, k: int = 6, progress=None, do_reindex: bool = True) -> dict:
    """반환: {"ok": True, "answer", "results": [...], "index": stats}
        또는 {"ok": False, "error": "no_folders"|"no_results"|"server_unreachable"|"empty_response", "reason"}."""
    folders = load_folders() if folders is None else folders
    if not folders:
        return {"ok": False, "error": "no_folders",
                "reason": "과거 문서 폴더가 아직 등록되지 않았어요. 터미널에서 "
                          "`python archive_index.py --add-folder \"폴더경로\"` 로 등록한 뒤 다시 물어보세요."}
    if do_reindex:
        reindex(folders, db_path=db_path, progress=progress)
    results = search_archive(question, k=k, year=year, db_path=db_path)
    if not results:
        stats = index_stats(db_path)
        return {"ok": False, "error": "no_results",
                "reason": f"관련 내용을 찾지 못했어요(색인된 문서 {stats['docs']}개). 다른 표현으로 물어보시거나 폴더가 맞는지 확인해주세요."}
    context = format_results(results)
    r = answer_from_context(context, question, client=client, heading="과거 문서에서 찾은 부분")
    if not r["ok"]:
        return r
    return {"ok": True, "answer": r["answer"], "results": results, "index": index_stats(db_path)}


# ---------------------------------------------------------------- self-tests (서버·한글 불필요)

class _FakeClient:
    def __init__(self, replies):
        self.replies = list(replies)
        self.calls = []

    def chat(self, **kwargs):
        self.calls.append(kwargs)
        reply = self.replies.pop(0)
        if isinstance(reply, Exception):
            raise reply
        return {"message": {"content": reply}}


def _selftest_ask_archive_paths():
    import tempfile, shutil
    d = tempfile.mkdtemp(prefix="_arcqa_")
    db = os.path.join(d, "idx.sqlite")
    try:
        assert ask_archive("?", folders=[], db_path=db)["error"] == "no_folders"
        docs = os.path.join(d, "docs"); os.makedirs(docs)
        open(os.path.join(docs, "2024년 서류심사 결과보고.txt"), "w", encoding="utf-8").write(
            "2024년 서류심사 건수는 총 120건이며 전년 대비 증가하였다.")
        r = ask_archive("완전히 무관한 질문 zzz", folders=[docs], db_path=db, client=_FakeClient(["x"]))
        assert r["error"] == "no_results" and "색인된 문서 1개" in r["reason"], r
        fake = _FakeClient(["- 2024년 서류심사 건수: 120건 [2024년 서류심사 결과보고.txt · 본문]"])
        r = ask_archive("재작년 서류심사 건수 몇 건?", folders=[docs], db_path=db, client=fake)
        assert r["ok"] and "120건" in r["answer"] and r["results"][0]["year"] == 2024, r
        sent = fake.calls[0]["messages"][1]["content"]
        assert sent.startswith("[과거 문서에서 찾은 부분]") and "[2024년 서류심사 결과보고.txt · 본문]" in sent, sent
        r = ask_archive("서류심사", folders=[docs], db_path=db, client=_FakeClient([ConnectionError()]))
        assert r["error"] == "server_unreachable", r
        print("ask_archive(폴더없음/결과없음/정상+출처/미접속) 통과")
    finally:
        shutil.rmtree(d, ignore_errors=True)


if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8")
    _selftest_ask_archive_paths()
