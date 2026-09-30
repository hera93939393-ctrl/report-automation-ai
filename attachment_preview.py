"""attachment_preview.py — "+"로 첨부한 문서를 첨부 즉시 작은 미리보기로
보여주기 위한 도구. HWP/HWPX는 이미지 미리보기(1페이지만), 엑셀/PDF는
텍스트 미리보기(상위 몇 줄)를 만든다. source_reader.py의
read_hwp_source_isolated()와 같은 이유로, HWP 미리보기는 반드시 별도
프로세스에서 만든다(같은 프로세스 안의 다른 Hwp 인스턴스가 채팅 세션이
이미 열어둔 HwpReport의 COM 연결을 깨뜨리는 pyhwpx의 한계, source_reader.py
주석 참고).

실측 확인(2026-09-07): create_page_image(path, pgno=1, resolution=60,
format="gif")로 만든 1페이지 미리보기는 문서 한 장 기준 수 KB 수준이다
(직접 측정: 짧은 문서 1장 gif@60dpi 약 6.8KB) - bmp(기본 포맷)는 무압축이라
96dpi 한 장에도 3.5MB나 나가 첨부할 때마다 용량이 급격히 쌓이므로, 이
도구는 반드시 gif+저해상도 조합을 쓴다(사용자가 명시적으로 우려한 용량
관리 요구사항).

(2026-09-17 재조정) 채팅창의 작은 썸네일은 원래 60dpi로 충분했지만,
채팅창에서 클릭해서 확대해 보는 기능이 추가되면서 60dpi 그대로 늘리면
글씨가 깨지는 문제가 실사용으로 재현됐다. 처음엔 "확대할 때만 더 높은
해상도로 다시 렌더링"하는 방식을 시도했는데, 직접 시간을 재보니 이
느림의 원인이 해상도가 아니라 매번 한글 COM 인스턴스를 새로 여는
오버헤드(약 8~9초, 60dpi든 150dpi든 큰 차이 없음)였다 - 그래서 클릭할
때마다 또 기다리게 됐다(실사용 재현 — "한글은 너무 느리다"). 대신
**처음 첨부할 때 딱 한 번만 150dpi로 만들고, 썸네일은 그걸 줄여서
보여주고 확대창은 그 파일을 그대로 재사용**하는 쪽으로 바꿨다 - 첨부할
때 기다리는 시간은 똑같지만(원래도 COM을 열어야 해서 몇 초 걸렸음),
그 뒤로는 몇 번을 클릭해서 확대해도 다시 기다릴 필요가 없다. 150dpi
파일 용량도 실측 15KB 수준이라(짧은 문서 기준) 용량 관리 요구사항과
크게 충돌하지 않는다."""
import glob
import os
import subprocess
import sys
import tempfile

_PREVIEW_DIR = os.path.join(tempfile.gettempdir(), "_hwp_attachment_previews")
_MAX_PREVIEWS = 5
_PREVIEW_RESOLUTION = 150
from hwp_session import new_hwp, open_document


def generate_text_preview(path: str, max_lines: int = 5) -> str:
    """엑셀/PDF의 상위 max_lines줄만 텍스트로 미리보기를 만든다. 이미지이
    아니라 문자열을 그대로 반환한다(채팅 말풍선에 바로 표시할 수 있게).
    지원하지 않는 형식이거나 읽기 실패하면 빈 문자열(예외 없음, 이 모듈의
    다른 함수들과 같은 관례)."""
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext in (".xlsx", ".xls"):
            import openpyxl
            wb = openpyxl.load_workbook(path, data_only=True)
            ws = wb.active
            lines = []
            for row in ws.iter_rows(max_row=max_lines, values_only=True):
                lines.append(", ".join(str(v) for v in row if v is not None))
            return "\n".join(lines)
        if ext == ".pdf":
            import pdfplumber
            with pdfplumber.open(path) as pdf:
                text = pdf.pages[0].extract_text() or ""
            return "\n".join(text.split("\n")[:max_lines])
    except Exception:
        return ""
    return ""


_EXCEL_PREVIEW_MAX_ROWS = 15
_EXCEL_PREVIEW_MAX_COLS = 8
_EXCEL_PREVIEW_CELL_MAX_CHARS = 18  # 이보다 길면 "…"로 잘라서 칸이 무한정 안 넓어지게 함
_EXCEL_PREVIEW_FONT_PATH = "C:/Windows/Fonts/malgun.ttf"  # 이 프로젝트의 다른 곳(PDF self-test)과 같은 폰트
_EXCEL_PREVIEW_FONT_SIZE = 14
_EXCEL_PREVIEW_PADDING = 6
_EXCEL_PREVIEW_HEADER_FILL = (230, 241, 251)  # 이 앱 채팅창의 연파랑 계열과 통일(_BUBBLE_STYLE 참고)


def generate_excel_preview_image(path: str) -> str | None:
    """엑셀 앞쪽 몇 행 x 몇 열만(전체가 아니라 "맛보기") 표 모양으로 그려
    작은 이미지 파일로 저장한다. 파일이 아무리 커도(수만 행이어도) 실제로
    읽고 그리는 범위는 항상 앞쪽 _EXCEL_PREVIEW_MAX_ROWS행뿐이라, 속도가
    원본 파일 크기와 무관하게 항상 일정하다(2026-09-17, 사용자 요청 —
    "엑셀 켜서 보는 번거로움을 줄이자"는 목적에 실제 파일 전체를 그릴
    필요는 없다는 점 확인). 잘려서 안 보이는 행/열이 있으면 마지막 줄에
    "...더 있음" 표시를 남긴다. 지원하지 않는 형식이거나 읽기 실패하면
    None(예외 없음, 이 모듈의 다른 함수들과 같은 관례)."""
    try:
        import openpyxl
        from PIL import Image as PILImage, ImageDraw, ImageFont

        # read_only=True는 openpyxl이 XML을 스트리밍으로 훑게 해서, 뒤에서
        # max_row로 앞쪽 몇 줄만 자르더라도 파일 전체를 먼저 메모리에 다
        # 올리지 않는다 - 파일이 몇만 행이어도 미리보기 속도가 행 개수와
        # 무관하게 일정하다는 요구사항(사용자가 직접 우려한 부분)의 핵심.
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        try:
            ws = wb.active
            total_rows = ws.max_row or 0
            total_cols = ws.max_column or 0
            rows = []
            for row in ws.iter_rows(max_row=_EXCEL_PREVIEW_MAX_ROWS,
                                     max_col=_EXCEL_PREVIEW_MAX_COLS, values_only=True):
                cells = []
                for v in row:
                    text = "" if v is None else str(v)
                    if len(text) > _EXCEL_PREVIEW_CELL_MAX_CHARS:
                        text = text[:_EXCEL_PREVIEW_CELL_MAX_CHARS - 1] + "…"
                    cells.append(text)
                rows.append(cells)
        finally:
            # read_only 워크북은 명시적으로 닫아야 파일 핸들이 풀린다 - 안
            # 닫으면 이 함수가 끝난 뒤에도 원본 엑셀 파일이 "다른 프로세스가
            # 사용 중"이라며 잠겨있는 상태가 되는 게 실측 확인됐다(self-test의
            # 정리 단계에서 PermissionError로 재현됨).
            wb.close()
        if not rows:
            return None

        truncated = total_rows > _EXCEL_PREVIEW_MAX_ROWS or total_cols > _EXCEL_PREVIEW_MAX_COLS
        if truncated:
            rows.append([f"... 더 있음 (전체 {total_rows}행 x {total_cols}열)"] +
                        [""] * (len(rows[0]) - 1))

        font = ImageFont.truetype(_EXCEL_PREVIEW_FONT_PATH, _EXCEL_PREVIEW_FONT_SIZE)
        n_cols = max(len(r) for r in rows)
        col_widths = []
        for c in range(n_cols):
            widest = max((font.getlength(r[c]) if c < len(r) else 0) for r in rows)
            col_widths.append(int(widest) + _EXCEL_PREVIEW_PADDING * 2)
        row_height = _EXCEL_PREVIEW_FONT_SIZE + _EXCEL_PREVIEW_PADDING * 2

        img_width = sum(col_widths)
        img_height = row_height * len(rows)
        image = PILImage.new("RGB", (img_width, img_height), "white")
        draw = ImageDraw.Draw(image)

        for r_idx, row in enumerate(rows):
            y = r_idx * row_height
            if r_idx == 0 and not (truncated and r_idx == len(rows) - 1):
                draw.rectangle([0, y, img_width, y + row_height], fill=_EXCEL_PREVIEW_HEADER_FILL)
            x = 0
            for c_idx, width in enumerate(col_widths):
                text = row[c_idx] if c_idx < len(row) else ""
                draw.text((x + _EXCEL_PREVIEW_PADDING, y + _EXCEL_PREVIEW_PADDING),
                          text, fill="black", font=font)
                x += width
            draw.line([0, y, img_width, y], fill=(221, 221, 221))
        draw.rectangle([0, 0, img_width - 1, img_height - 1], outline=(200, 200, 200))

        os.makedirs(_PREVIEW_DIR, exist_ok=True)
        out_path = os.path.join(_PREVIEW_DIR, f"{os.path.basename(path)}_{os.getpid()}.gif")
        image.save(out_path)
        _prune_old_previews()
        return out_path
    except Exception:
        return None


def _generate_hwp_preview(path: str, out_path: str, resolution: int = _PREVIEW_RESOLUTION) -> bool:
    """path(.hwp/.hwpx)의 1페이지를 out_path(gif)로 저장한다. 이 함수 자체는
    현재 프로세스에서 새 Hwp 인스턴스를 만들므로, 채팅 세션이 이미 다른
    Hwp를 열어둔 상태에서 절대 직접 호출하면 안 된다 - 반드시
    generate_hwp_preview_isolated()를 통해 별도 프로세스에서만 실행한다."""
    from pyhwpx import Hwp
    hwp = None
    try:
        # (2026-10-01) 새 프로세스 강제 + 보안모듈 등록 확인 + 무인 열기 옵션
        # (forceopen, suspendpassword)을 hwp_session이 한 곳에서 맡는다.
        hwp = new_hwp(visible=False)
        if not open_document(hwp, path, unattended=True):
            return False
        return bool(hwp.create_page_image(out_path, pgno=1, resolution=resolution, format="gif"))
    except Exception:
        return False
    finally:
        if hwp is not None:
            hwp.quit()


def generate_hwp_preview_isolated(path: str, resolution: int = _PREVIEW_RESOLUTION) -> str | None:
    """_generate_hwp_preview()를 별도 프로세스에서 실행해, 성공하면 저장된
    미리보기 이미지 경로를, 실패하면 None을 반환한다. source_reader.py의
    read_hwp_source_isolated()와 동일한 subprocess 격리 패턴. 호출 성공
    시마다 _prune_old_previews()로 오래된 미리보기를 정리한다.

    (2026-09-17, 실사용 재현) resolution을 기본값(썸네일용, 저용량)보다
    높여서 부를 수 있게 열어둔다 - 클릭해서 확대해 보는 화면은 저해상도
    원본을 그냥 늘리면 글씨가 깨지는 게 실제로 재현됐다(용량을 아끼려고
    60dpi로 만든 썸네일을 780px까지 늘려서 생긴 문제) - 확대해서 볼
    때만 더 높은 해상도로 새로 만든다."""
    os.makedirs(_PREVIEW_DIR, exist_ok=True)
    out_path = os.path.join(_PREVIEW_DIR, f"{os.path.basename(path)}_{os.getpid()}_{resolution}.gif")
    module_dir = os.path.dirname(os.path.abspath(__file__))
    script = (
        "import sys; sys.path.insert(0, sys.argv[3]); "
        "from attachment_preview import _generate_hwp_preview; "
        "ok = _generate_hwp_preview(sys.argv[1], sys.argv[2], int(sys.argv[4])); "
        "sys.exit(0 if ok else 1)"
    )
    try:
        result = subprocess.run(
            [sys.executable, "-c", script, path, out_path, module_dir, str(resolution)],
            capture_output=True, text=True, timeout=60,
        )
        if result.returncode != 0 or not os.path.exists(out_path):
            return None
        _prune_old_previews()
        return out_path
    except Exception:
        return None


def _prune_old_previews(keep: int = _MAX_PREVIEWS) -> None:
    """_PREVIEW_DIR에 쌓인 미리보기 이미지 중 오래된 것부터 지워 최근
    keep개만 남긴다(용량 관리 요구사항). 파일 개수가 keep 이하면 아무 것도
    안 한다."""
    if not os.path.isdir(_PREVIEW_DIR):
        return
    files = sorted(glob.glob(os.path.join(_PREVIEW_DIR, "*.gif")), key=os.path.getmtime)
    for old_file in files[:-keep] if len(files) > keep else []:
        try:
            os.remove(old_file)
        except OSError:
            pass


def _remove_test_previews(source_path: str) -> None:
    """self-test가 만든 미리보기만 지운다. (2026-10-01) 예전엔 미리보기 폴더의
    *.gif를 전부 지웠는데, 이 폴더는 실행 중인 채팅창(다른 PID)이 사용자
    첨부의 미리보기를 두는 곳이기도 해서, 채팅창이 열어둔 파일을 지우려다
    PermissionError로 테스트 전체가 멈추고(실제 발생) 남의 세션 파일까지
    지워버릴 수 있었다. 파일명이 '<원본 basename>_<pid>...gif' 규칙이므로
    이 테스트의 원본 basename으로 시작하는 것만 지우고, 잠긴 파일은 건너뛴다."""
    pattern = os.path.join(_PREVIEW_DIR, glob.escape(os.path.basename(source_path)) + "_*.gif")
    for f in glob.glob(pattern):
        try:
            os.remove(f)
        except PermissionError:
            pass


def _selftest_generate_hwp_preview_isolated_creates_small_file():
    """미리보기 이미지가 실제로 생성되고, gif+저해상도 조합 덕분에 충분히
    작은지(용량 관리 요구사항) 확인한다."""
    import time
    from pyhwpx import Hwp

    path = os.path.join(tempfile.gettempdir(), "_test_미리보기_원본.hwp")
    setup = Hwp(visible=False, new=True)
    setup.insert_text("미리보기 테스트 문서입니다.")
    setup.save_as(path)
    setup.quit()
    time.sleep(2)

    try:
        preview_path = generate_hwp_preview_isolated(path)
        assert preview_path is not None
        assert os.path.exists(preview_path)
        size = os.path.getsize(preview_path)
        assert size < 200_000, f"미리보기 용량이 예상보다 큼: {size} bytes"
        print("generate_hwp_preview_isolated 통과: 용량", size, "bytes")
    finally:
        os.remove(path)
        _remove_test_previews(path)


def _selftest_generate_hwp_preview_isolated_missing_file_returns_none():
    result = generate_hwp_preview_isolated("존재하지_않는_미리보기.hwp")
    assert result is None, result
    print("generate_hwp_preview_isolated(파일없음) 통과")


def _selftest_generate_excel_preview_image_creates_small_file():
    """작은 엑셀은 잘림 없이, 이미지가 실제로 생성되고 gif라 용량이 작은지
    확인한다."""
    import openpyxl
    path = os.path.join(tempfile.gettempdir(), "_test_엑셀미리보기.xlsx")
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append(["항목", "금액"])
    ws.append(["인건비", 1000000])
    wb.save(path)
    try:
        preview_path = generate_excel_preview_image(path)
        assert preview_path is not None
        assert os.path.exists(preview_path)
        size = os.path.getsize(preview_path)
        assert size < 200_000, f"미리보기 용량이 예상보다 큼: {size} bytes"
        print("generate_excel_preview_image(작은 파일) 통과: 용량", size, "bytes")
    finally:
        os.remove(path)
        _remove_test_previews(path)


def _selftest_generate_excel_preview_image_large_file_stays_fast_and_capped():
    """(2026-09-17, 사용자가 직접 우려한 부분) 행이 아주 많은 엑셀이라도
    앞쪽 _EXCEL_PREVIEW_MAX_ROWS행만 읽고 그려서, 걸리는 시간이 행 개수와
    무관하게 일정해야 한다 — 1만 행짜리 파일로 직접 재현해서 확인한다."""
    import time
    import openpyxl
    path = os.path.join(tempfile.gettempdir(), "_test_엑셀미리보기_대용량.xlsx")
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append(["번호", "항목"])
    for i in range(10000):
        ws.append([i, f"항목{i}"])
    wb.save(path)
    try:
        t0 = time.time()
        preview_path = generate_excel_preview_image(path)
        elapsed = time.time() - t0
        assert preview_path is not None
        assert elapsed < 5, f"1만 행 파일인데 {elapsed:.1f}초나 걸림 - 앞쪽만 읽는지 확인 필요"
        print(f"generate_excel_preview_image(1만 행) 통과: {elapsed:.2f}초")
    finally:
        os.remove(path)
        _remove_test_previews(path)


def _selftest_generate_text_preview_excel():
    import openpyxl
    path = os.path.join(tempfile.gettempdir(), "_test_미리보기.xlsx")
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append(["항목", "금액"])
    ws.append(["인건비", 1000000])
    wb.save(path)
    try:
        preview = generate_text_preview(path, max_lines=5)
        assert "항목" in preview and "인건비" in preview, preview
        print("generate_text_preview(엑셀) 통과:", repr(preview))
    finally:
        os.remove(path)


def _selftest_generate_text_preview_pdf():
    from fpdf import FPDF
    path = os.path.join(tempfile.gettempdir(), "_test_미리보기.pdf")
    pdf = FPDF()
    pdf.add_font("Malgun", fname="C:/Windows/Fonts/malgun.ttf")
    pdf.add_page(); pdf.set_font("Malgun", size=12)
    pdf.cell(0, 10, "예산은 1,850,000원입니다")
    pdf.output(path)
    try:
        preview = generate_text_preview(path, max_lines=5)
        assert "1,850,000" in preview or "1850000" in preview, preview
        print("generate_text_preview(PDF) 통과:", repr(preview))
    finally:
        os.remove(path)


def _selftest_prune_old_previews_keeps_recent_n():
    """미리보기가 keep개보다 많이 쌓이면 오래된 것부터 지워 최근 keep개만
    남아야 한다."""
    os.makedirs(_PREVIEW_DIR, exist_ok=True)
    # (2026-10-01) 이 폴더는 실행 중인 채팅창(다른 PID)의 미리보기도 두는 곳이라
    # 비어 있다고 가정할 수 없다 — 그 파일들은 아래 테스트 파일(mtime을 1970년
    # 근처로 강제)보다 항상 최신이라 keep 자리를 먼저 차지한다. 그 수만큼
    # keep을 늘려 "테스트 파일 중 최근 5개만 남는다"는 검증을 그대로 유지한다.
    foreign_count = len(glob.glob(os.path.join(_PREVIEW_DIR, "*.gif")))
    test_files = []
    for i in range(8):
        p = os.path.join(_PREVIEW_DIR, f"_test_prune_{i}.gif")
        with open(p, "w") as f:
            f.write("dummy")
        os.utime(p, (i, i))  # 오래된 순서대로 mtime을 강제로 다르게 설정
        test_files.append(p)
    try:
        _prune_old_previews(keep=5 + foreign_count)
        remaining = sorted(glob.glob(os.path.join(_PREVIEW_DIR, "_test_prune_*.gif")))
        assert len(remaining) == 5, remaining
        # 가장 오래된 3개(_0,_1,_2)는 지워지고 최근 5개(_3~_7)만 남아야 함
        for i in range(3):
            assert os.path.join(_PREVIEW_DIR, f"_test_prune_{i}.gif") not in remaining
        print("_prune_old_previews(최근 5개만 유지) 통과")
    finally:
        for p in test_files:
            if os.path.exists(p):
                os.remove(p)


if __name__ == "__main__":
    _selftest_generate_hwp_preview_isolated_creates_small_file()
    _selftest_generate_hwp_preview_isolated_missing_file_returns_none()
    _selftest_generate_excel_preview_image_creates_small_file()
    _selftest_generate_excel_preview_image_large_file_stays_fast_and_capped()
    _selftest_generate_text_preview_excel()
    _selftest_generate_text_preview_pdf()
    _selftest_prune_old_previews_keeps_recent_n()
