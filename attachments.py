"""attachments.py — 범용 부품 4: 첨부 파일(회의결과, 엑셀, PDF, 텍스트)을
"위치 표시가 붙은 텍스트" 한 가지 규격으로 읽는다(PRD 16-2 부품 4).

source_reader.py는 "숫자검증용 값 추출"(4종 값 목록)이 목적이고,
attachment_preview.py는 "앞부분 맛보기"가 목적이라, 둘 다 모델이 문서 전체
내용을 읽고 답하거나 초안에 반영하는 용도(부품 4)에는 맞지 않았다. 이 모듈은
그 둘의 읽기 방식을 재사용하되 결과를 하나의 규격으로 통일한다:

    {"path": str, "kind": "hwp"|"xlsx"|"pdf"|"text"|None,
     "parts": [{"label": "1페이지"|"시트 '실적'"|"본문", "text": str}, ...],
     "text": 전체를 "[label]" 머리말로 이어 붙인 문자열,
     "truncated": bool, "error": str|None}

- HWP/HWPX: 한글 COM으로 읽되 **반드시 별도 프로세스**에서(source_reader.
  read_hwp_source_isolated과 같은 이유 — 같은 프로세스의 다른 Hwp 인스턴스
  (채팅이 열어둔 보고서)의 COM 연결을 pyhwpx의 __del__이 끊어버림). 무인
  열기 옵션은 hwp_session.open_document(unattended=True).
- 엑셀: 모든 시트, 행을 " | "로 이어 붙임. read_only 스트리밍으로 큰 파일도
  앞쪽 일정 행까지만 읽는다(_MAX_ROWS_PER_SHEET).
- PDF: 쪽 단위(pdfplumber). 텍스트 없는 쪽(스캔)은 건너뜀(OCR 비목표).
- 텍스트/마크다운/CSV: 그대로.
읽기 실패는 예외 대신 error 필드로 알린다 — 다만 "파일이 진짜 비어있는 것"과
"못 연 것"을 구분해 알리는 것이 목적이므로(hwp-report-tool 메모의 OneDrive
온디맨드 사례), 실패 이유를 문자열로 남긴다."""
import json
import os
import subprocess
import sys

_MAX_ROWS_PER_SHEET = 500
_MAX_CHARS_DEFAULT = 60_000
_HWP_TIMEOUT_SEC = 90


def _kind_of(path: str):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".hwp", ".hwpx"):
        return "hwp"
    if ext in (".xlsx", ".xlsm", ".xls"):
        return "xlsx"
    if ext == ".pdf":
        return "pdf"
    if ext in (".txt", ".md", ".csv"):
        return "text"
    return None


def _read_hwp_text_in_this_process(path: str) -> str:
    """별도 프로세스 안에서만 호출된다(아래 _read_hwp_isolated의 스크립트).
    새 숨김 한글 프로세스를 띄워 본문 텍스트를 뽑는다."""
    from hwp_session import new_hwp, open_document
    hwp = new_hwp(visible=False)
    try:
        if not open_document(hwp, path, unattended=True):
            return ""
        return hwp.GetTextFile("TEXT", "") or ""
    finally:
        hwp.quit()


def _read_hwp_isolated(path: str) -> tuple:
    module_dir = os.path.dirname(os.path.abspath(__file__))
    script = (
        "import sys, json; sys.path.insert(0, sys.argv[2]); "
        "from attachments import _read_hwp_text_in_this_process as f; "
        "sys.stdout.reconfigure(encoding='utf-8'); "
        "print(json.dumps(f(sys.argv[1]), ensure_ascii=False))"
    )
    try:
        result = subprocess.run(
            [sys.executable, "-c", script, path, module_dir],
            capture_output=True, text=True, encoding="utf-8", timeout=_HWP_TIMEOUT_SEC,
        )
    except subprocess.TimeoutExpired:
        return "", "한글 읽기가 시간 안에 끝나지 않았습니다(대화상자가 떠 있을 수 있음)"
    except Exception as e:
        return "", f"한글 읽기 프로세스를 시작하지 못했습니다: {e}"
    if result.returncode != 0:
        tail = (result.stderr or "").strip().splitlines()[-1:] or ["알 수 없는 오류"]
        return "", f"한글 읽기 실패: {tail[0]}"
    try:
        return json.loads(result.stdout.strip().splitlines()[-1]), None
    except Exception:
        return "", "한글 읽기 결과를 해석하지 못했습니다"


def _read_xlsx(path: str) -> list:
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    parts = []
    try:
        for ws in wb.worksheets:
            lines = []
            for r, row in enumerate(ws.iter_rows(values_only=True), start=1):
                if r > _MAX_ROWS_PER_SHEET:
                    lines.append(f"… (이하 {ws.max_row - _MAX_ROWS_PER_SHEET}행 생략)")
                    break
                cells = [str(v) for v in row if v is not None and str(v).strip() != ""]
                if cells:
                    lines.append(" | ".join(cells))
            if lines:
                parts.append({"label": f"시트 '{ws.title}'", "text": "\n".join(lines)})
    finally:
        wb.close()
    return parts


def _read_pdf(path: str) -> list:
    import pdfplumber
    parts = []
    with pdfplumber.open(path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            text = (page.extract_text() or "").strip()
            if text:
                parts.append({"label": f"{page_num}페이지", "text": text})
    return parts


_KORDOC_CLI = os.path.join(os.path.dirname(os.path.abspath(__file__)), "node_tools", "node_modules", "kordoc", "dist", "cli.js")
_KORDOC_TIMEOUT_SEC = 120


def kordoc_available() -> bool:
    """kordoc(Node, MIT)이 node_tools에 설치돼 있고 node를 찾을 수 있으면 True."""
    import shutil
    return os.path.exists(_KORDOC_CLI) and shutil.which("node") is not None


def _read_hwp_with_kordoc(path: str) -> tuple:
    """(2026-10-01 도입) kordoc으로 .hwp/.hwpx를 한글 없이 읽는다 — RAG용 청크
    (`--format chunks`: id/type/breadcrumb/text/blockRange/page)를 쪽 단위로 묶어
    parts로 만든다. 표는 HTML/파이프 표 그대로(병합 셀 보존). 목업 실측: 3쪽
    문서 5초(한글 COM 격리 읽기 9.5초), 표·❍ 글머리·위첨자 보존.
    반환: (parts, error)"""
    import shutil, tempfile
    out_dir = tempfile.mkdtemp(prefix="_kordoc_")
    out_path = os.path.join(out_dir, "out.json")
    try:
        try:
            proc = subprocess.run(
                [shutil.which("node"), _KORDOC_CLI, "--silent", "--no-images", "--format", "chunks",
                 "-o", out_path, path],
                capture_output=True, text=True, encoding="utf-8", errors="replace",
                timeout=_KORDOC_TIMEOUT_SEC, cwd=os.path.dirname(_KORDOC_CLI),
            )
        except subprocess.TimeoutExpired:
            return [], "kordoc 읽기가 시간 안에 끝나지 않았습니다"
        except Exception as e:
            return [], f"kordoc 실행 실패: {e}"
        if proc.returncode != 0 or not os.path.exists(out_path):
            tail = (proc.stderr or proc.stdout or "").strip().splitlines()[-1:] or ["알 수 없는 오류"]
            return [], f"kordoc 읽기 실패: {tail[0][:200]}"
        with open(out_path, encoding="utf-8") as f:
            data = json.load(f)
        chunks = data if isinstance(data, list) else data.get("chunks", [])
        by_page = {}
        for c in chunks:
            text = (c.get("text") or "").strip()
            if not text:
                continue
            by_page.setdefault(c.get("page") or 0, []).append(text)
        parts = [{"label": f"{page}페이지" if page else "본문", "text": "\n".join(texts)}
                 for page, texts in sorted(by_page.items())]
        return parts, None
    finally:
        shutil.rmtree(out_dir, ignore_errors=True)


def read_attachment(path: str, max_chars: int = _MAX_CHARS_DEFAULT, reader: str = "auto") -> dict:
    """부품 4. 첨부 파일 하나를 위치 표시가 붙은 텍스트로 읽는다.
    reader: "auto"(kordoc이 있으면 kordoc, 없으면 한글 COM) | "kordoc" | "hwp"(한글 COM 강제)."""
    result = {"path": path, "kind": _kind_of(path), "parts": [], "text": "", "truncated": False, "error": None,
              "reader": None}
    if not os.path.exists(path):
        result["error"] = "파일이 없습니다"
        return result
    if result["kind"] is None:
        result["error"] = "지원하지 않는 형식입니다(hwp/hwpx/xlsx/pdf/txt/md/csv만)"
        return result
    try:
        if result["kind"] == "hwp":
            use_kordoc = reader == "kordoc" or (reader == "auto" and kordoc_available())
            if use_kordoc:
                parts, err = _read_hwp_with_kordoc(path)
                result["reader"] = "kordoc"
                if err and reader == "auto":
                    # kordoc이 못 읽는 파일(암호·DRM 등)은 한글 COM으로 한 번 더
                    text, err = _read_hwp_isolated(path)
                    result["reader"] = "hwp"
                    parts = [{"label": "본문", "text": text.replace("\r\n", "\n").strip()}] if text else []
            else:
                text, err = _read_hwp_isolated(path)
                result["reader"] = "hwp"
                text = text.replace("\r\n", "\n").strip()
                parts = [{"label": "본문", "text": text}] if text else []
            if err:
                result["error"] = err
                return result
            result["parts"] = [p for p in parts if p["text"]]
        elif result["kind"] == "xlsx":
            result["parts"] = _read_xlsx(path)
        elif result["kind"] == "pdf":
            result["parts"] = _read_pdf(path)
        else:
            with open(path, encoding="utf-8", errors="replace") as f:
                text = f.read().strip()
            result["parts"] = [{"label": "본문", "text": text}] if text else []
    except PermissionError:
        result["error"] = "파일을 열 권한이 없습니다(OneDrive 온라인 전용 파일이면 '항상 이 디바이스에 유지'를 켜세요)"
        return result
    except Exception as e:
        result["error"] = f"읽기 실패: {e}"
        return result

    name = os.path.basename(path)
    chunks, used = [], 0
    for part in result["parts"]:
        block = f"[{name} · {part['label']}]\n{part['text']}"
        if used + len(block) > max_chars:
            remain = max_chars - used
            if remain > 200:
                chunks.append(block[:remain] + "\n…(잘림)")
            result["truncated"] = True
            break
        chunks.append(block)
        used += len(block) + 2
    result["text"] = "\n\n".join(chunks)
    return result


def read_attachments(paths: list, max_chars_each: int = _MAX_CHARS_DEFAULT) -> list:
    """폴더가 섞여 있으면 안의 지원 파일을 펼쳐서 전부 읽는다."""
    expanded = []
    for p in paths:
        if os.path.isdir(p):
            for root, _dirs, files in os.walk(p):
                for f in sorted(files):
                    fp = os.path.join(root, f)
                    if _kind_of(fp) is not None:
                        expanded.append(fp)
        else:
            expanded.append(p)
    seen, ordered = set(), []
    for p in expanded:
        key = os.path.normcase(os.path.abspath(p))
        if key not in seen:
            seen.add(key)
            ordered.append(p)
    return [read_attachment(p, max_chars_each) for p in ordered]


# ---------------------------------------------------------------- self-tests

def _selftest_text_xlsx_pdf_and_errors():
    import tempfile
    d = tempfile.mkdtemp(prefix="_att_")
    try:
        txt = os.path.join(d, "메모.txt")
        with open(txt, "w", encoding="utf-8") as f:
            f.write("회의 결과: 현장점검 3월 착수\n예산 1,850,000원")
        r = read_attachment(txt)
        assert r["kind"] == "text" and r["error"] is None and "[메모.txt · 본문]" in r["text"] and "1,850,000" in r["text"], r

        import openpyxl
        xlsx = os.path.join(d, "실적.xlsx")
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = "서류심사"
        ws.append(["연도", "건수"]); ws.append([2024, 120]); ws.append([2025, 135])
        ws2 = wb.create_sheet("현장점검"); ws2.append(["구분", "회수"]); ws2.append(["상반기", 7])
        wb.save(xlsx)
        r = read_attachment(xlsx)
        assert r["kind"] == "xlsx" and len(r["parts"]) == 2, r
        assert r["parts"][0]["label"] == "시트 '서류심사'" and "2024 | 120" in r["parts"][0]["text"], r["parts"]
        assert "[실적.xlsx · 시트 '현장점검']" in r["text"]

        from fpdf import FPDF
        pdf_path = os.path.join(d, "공문.pdf")
        pdf = FPDF(); pdf.add_font("Malgun", fname="C:/Windows/Fonts/malgun.ttf")
        pdf.add_page(); pdf.set_font("Malgun", size=12); pdf.cell(0, 10, "1쪽 내용 예산 1,850,000원")
        pdf.add_page(); pdf.set_font("Malgun", size=12); pdf.cell(0, 10, "2쪽 내용 건수 120건")
        pdf.output(pdf_path)
        r = read_attachment(pdf_path)
        assert r["kind"] == "pdf" and [p["label"] for p in r["parts"]] == ["1페이지", "2페이지"], r["parts"]
        assert "120" in r["parts"][1]["text"]

        r = read_attachment(os.path.join(d, "없음.xlsx"))
        assert r["error"] == "파일이 없습니다", r
        docx = os.path.join(d, "x.docx")
        open(docx, "wb").close()
        r = read_attachment(docx)
        assert r["kind"] is None and "지원하지" in r["error"], r

        # 잘림
        big = os.path.join(d, "big.txt")
        with open(big, "w", encoding="utf-8") as f:
            f.write("가" * 5000)
        r = read_attachment(big, max_chars=1000)
        assert r["truncated"] and r["text"].endswith("…(잘림)") and len(r["text"]) <= 1100, (r["truncated"], len(r["text"]))

        # 폴더 펼치기 + 중복 제거
        rs = read_attachments([d, txt])
        assert sorted(os.path.basename(x["path"]) for x in rs) == ["big.txt", "공문.pdf", "메모.txt", "실적.xlsx"], [x["path"] for x in rs]  # x.docx는 지원 형식이 아니라 폴더 펼치기에서 제외
        print("read_attachment(text/xlsx/pdf/오류/잘림/폴더) 통과")
    finally:
        import shutil
        shutil.rmtree(d, ignore_errors=True)


def _selftest_hwp_isolated_reads_body_without_touching_live_instance():
    """가짜 .hwp를 별도 프로세스로 읽고, 이 프로세스에 살아있는 Hwp 인스턴스가
    그 뒤에도 멀쩡한지(격리의 존재 이유) 확인한다."""
    import tempfile
    from hwp_session import new_hwp
    path = os.path.join(tempfile.gettempdir(), "_test_attachments_회의결과.hwp")
    setup = new_hwp(visible=False)
    setup.insert_text("회의 결과"); setup.BreakPara(); setup.insert_text("현장점검은 3월에 착수한다.")
    setup.save_as(path); setup.quit()
    live = new_hwp(visible=False)
    try:
        live.insert_text("살아있는 문서")
        r = read_attachment(path)  # auto: kordoc이 있으면 kordoc(쪽 단위 라벨), 없으면 한글 COM
        assert r["kind"] == "hwp" and r["error"] is None, r
        assert r["reader"] in ("kordoc", "hwp") and r["parts"][0]["label"] in ("본문", "1페이지"), r
        assert "현장점검은 3월에 착수한다." in r["text"] and "[_test_attachments_회의결과.hwp · " in r["text"], r["text"][:200]
        r = read_attachment(path, reader="hwp")  # 한글 COM 경로는 항상 검증
        assert r["reader"] == "hwp" and r["parts"][0]["label"] == "본문", r
        assert "현장점검은 3월에 착수한다." in r["parts"][0]["text"], r["parts"]
        assert live.get_selected_text() is not None  # COM 연결 생존 확인
        live.SelectAll()
        assert "살아있는 문서" in (live.get_selected_text() or "")
        r = read_attachment(os.path.join(tempfile.gettempdir(), "_없는파일.hwp"))
        assert r["error"] == "파일이 없습니다"
        print("read_attachment(hwp 격리 읽기 + 살아있는 인스턴스 보존) 통과")
    finally:
        live.quit()
        setup_keep = setup  # noqa: F841 — __del__ 타이밍 문제 회피(hwp-report-tool 메모)
        if os.path.exists(path):
            os.remove(path)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    _selftest_text_xlsx_pdf_and_errors()
    _selftest_hwp_isolated_reads_body_without_touching_live_instance()
