"""weekly_report_tool.py — 여러 사람이 보낸 주간업무보고(고정 서식: 표 하나,
"이번주(...)"/"다음주(...)" 2열 헤더) 문서를 읽어, 그 표 안에 파란색으로
써넣어진 내용만 뽑아 이미 열려있는 대상 보고서의 같은 칸으로 옮겨 붙이는 도구.
verify_tool.py/table_tool.py/polish_tool.py와 나란한, 도구 하나당 파일 하나
관례를 따른다.

실측 확인(2026-09-07, Hwp(visible=False, new=True)로 직접 테스트 — 계획서
"사전조사 실측 확인 요약" 참고):
- 표 안 특정 칸의 텍스트는 커서를 그 칸으로 옮긴 뒤 hwp.get_selected_text()
  하나로 읽힌다.
- 칸 안 특정 위치의 글자색은 hwp.find(문단텍스트, direction="Forward") 후
  hwp.CharShape.Item("TextColor")로 읽으면 된다 — hwp_report.py의
  get_char_color_at()와 완전히 같은 방식이 표 안에서도 그대로 통한다.
- "이번주"/"다음주" 헤더로 표를 찾는 전용 API는 pyhwpx에 없다 — 문서를
  hwp.find(키워드, direction="Forward")로 훑으면서, 찾은 위치가 is_cell()이고
  get_cell_addr()의 행 번호가 1(헤더 행)인 첫 occurrence를 헤더로 채택한다.
- 표 하나에 헤더 1행 + 데이터 1행(2열)만 있는 고정 구조라는 전제(사용자가
  제시한 실제 양식 스크린샷 기준)를 그대로 따른다.
- 자리가 겹칠 때(이미 다른 사람 내용이 있을 때) 덮어쓰지 않고 이어붙이는
  안전한 방법: TableCellBlock()(칸 전체 선택) + Cancel()은 커서를 선택
  영역의 "시작점"에 남기므로(기존 코드가 문서화한 "find()+Cancel()은
  끝점에 남는다"는 규칙과 정반대 — 표 안 칸 선택에는 그 규칙이 적용되지
  않음을 이번에 실측으로 새로 확인함), 그 상태로 삽입하면 기존 내용 앞에
  끼어드는 버그가 재현된다. 대신 기존 내용의 마지막 줄을 find(Forward)로
  다시 찾아 Cancel()하면 커서가 그 줄의 끝에 남으므로, 그 뒤에
  BreakPara+insert_text로 이어붙여야 안전하다.
"""
import json
import os
import re
import subprocess
import sys


def _is_bluish(rgb: tuple[int, int, int]) -> bool:
    """RGB가 "파란색으로 써넣은 내용"으로 볼 만큼 충분히 파란지 판별한다.
    (0,0,255) 정확히 일치가 아니라 여유를 두는 이유: 사람마다 한/글에서
    고르는 "파랑"이 팔레트상 조금씩 다를 수 있어서다(순정 파랑이 아닌
    남색 계열도 있을 수 있음). B값이 충분히 밝고 R/G보다 뚜렷하게 클 때만
    파란색으로 인정한다."""
    r, g, b = rgb
    return b >= 120 and (b - max(r, g)) >= 60


def _goto_header_cell(hwp, keyword: str) -> bool:
    """문서 안에서 keyword("이번주" 또는 "다음주")가 들어간 표의 헤더 셀
    (1행)을 찾아 커서를 그 칸에 둔다. keyword가 표 밖(본문)에도 등장할 수
    있으므로, 찾은 위치가 반드시 칸 안이고 1행이어야 진짜 헤더로 인정한다 -
    아니면 다음 occurrence로 계속 넘어간다. 못 찾으면 False.

    (2026-09-07, 실측으로 새로 발견) 문서 전체가 표 하나뿐이면(본문 문단이
    실질적으로 표 앞 빈 문단 하나뿐인 이 기능의 고정 서식 특성상 흔함),
    MoveDocBegin() 직후 첫 검색을 direction="Forward"로 하면 실제로 존재하는
    텍스트도 못 찾고 False가 나는 현상이 실측으로 재현됨(같은 문서에서
    direction="AllDoc"으로는 정상적으로 찾아짐 — pyhwpx/한글의 "Forward"
    탐색이 문서 맨 처음 위치에서 표로 곧장 이어지는 경우를 못 잡는 것으로
    보임). 반면 그 다음 occurrence부터는(커서가 이미 진짜 위치로 옮겨진
    뒤라서) direction="Forward"가 정상적으로 동작하고 문서 끝에서 올바르게
    False로 멈춘다(직접 재현 확인 — AllDoc은 문서 끝에서 처음으로 되감기
    때문에 반복문에 계속 쓰면 무한루프 위험이 있어 최초 1회만 쓴다). 그래서
    첫 번째 검색만 "AllDoc", 이후 반복은 "Forward"로 나눈다."""
    hwp.MoveDocBegin()
    found = hwp.find(keyword, direction="AllDoc")
    while found:
        if hwp.is_cell():
            addr = hwp.get_cell_addr()
            row = int(re.search(r'\d+', addr).group())
            if row == 1:
                return True
        found = hwp.find(keyword, direction="Forward")
    return False


def _read_blue_lines_at_cursor(hwp) -> list[str]:
    """현재 커서가 있는 칸(표 데이터 칸)의 내용을 문단(줄) 단위로 읽어, 그
    중 파란색으로 쓰인 문단만 순서대로 리스트로 반환한다. 칸이 비어있으면
    빈 리스트.

    (2026-09-07, 실측으로 새로 발견) get_selected_text() 호출 직후의 커서는
    "막 재설정된 경계 위치"라서, _goto_header_cell()에서 이미 확인된 것과
    같은 pyhwpx/한글 특성(그런 위치에서의 첫 direction="Forward" 검색은
    실제로 존재하는 텍스트도 못 찾을 수 있음)이 여기서도 재현됨 — 첫 줄
    검색만 "AllDoc", 이후 줄은 "Forward"로 나눈다. 그리고 find()가 True를
    반환해도 CharShape.Item("TextColor")가 None을 반환하는 경우가 실측으로
    확인됐다(정확한 재현 조건은 못 밝혔으나 드물지 않게 발생) — 이 경우
    "파란색인지 확정할 수 없음"으로 보고 크래시 대신 건너뛴다(대조불가
    쪽으로 안전하게 치우침, 이 프로젝트의 검증 도구들이 일관되게 따르는
    "애매하면 사람 판단으로 넘긴다" 원칙과 같은 방향)."""
    cell_text = hwp.get_selected_text()
    if not cell_text:
        return []
    lines = [ln for ln in cell_text.replace("\r\n", "\n").split("\n") if ln]
    blue_lines = []
    for i, line in enumerate(lines):
        direction = "AllDoc" if i == 0 else "Forward"
        if not hwp.find(line, direction=direction):
            continue
        color_value = hwp.CharShape.Item("TextColor")
        if color_value is None:
            continue
        rgb = (color_value & 0xFF, (color_value >> 8) & 0xFF, (color_value >> 16) & 0xFF)
        if _is_bluish(rgb):
            blue_lines.append(line)
    return blue_lines


def read_weekly_content(path: str) -> dict:
    """path(.hwp/.hwpx)를 안 보이게 열어 "이번주"/"다음주" 표에서 파란색
    내용만 뽑아 {"this_week": [...], "next_week": [...]} 형태로 반환한다.
    각 리스트는 그 칸에서 파란색으로 확인된 문단(줄) 순서 그대로다 - 파란색
    문단이 하나도 없으면 빈 리스트(병합 대상 없음이라는 뜻, 오류 아님).
    read_hwp_source()와 동일한 관례로, 열 수 없는 파일은 예외 없이 빈
    결과를 반환한다."""
    from pyhwpx import Hwp
    hwp = None
    try:
        hwp = Hwp(visible=False, new=True)
        if not hwp.open(path):
            return {"this_week": [], "next_week": []}
        result = {}
        for key, keyword in (("this_week", "이번주"), ("next_week", "다음주")):
            if _goto_header_cell(hwp, keyword):
                # (2026-09-07, 실측으로 새로 발견) TableColBegin은 "지금 있는
                # 열의 맨 위"가 아니라 "표의 맨 왼쪽 열(A열)"로 이동하는
                # 액션이다 — table_tool.py의 build_table()이 TableColBegin
                # 다음에 TableColPageUp을 같이 써서 "표의 A1(맨 위 첫 칸)"로
                # 되돌아가는 데 쓰는 것과 같은 의미. _goto_header_cell()이
                # 이미 정확한 열의 헤더 행(1행)에 커서를 둔 상태이므로,
                # 여기서 TableColBegin을 부르면 다음주(B열) 헤더를 찾은
                # 뒤에도 A열로 튕겨나가버린다(직접 재현 확인 — next_week 결과가
                # this_week와 똑같이 나오는 버그였음). MoveDown만으로 충분하다.
                hwp.Run("MoveDown")
                result[key] = _read_blue_lines_at_cursor(hwp)
            else:
                result[key] = []
        return result
    except Exception:
        return {"this_week": [], "next_week": []}
    finally:
        if hwp is not None:
            hwp.quit()


def read_weekly_content_isolated(path: str) -> dict:
    """read_weekly_content()를 별도 프로세스에서 실행해 결과를 받아온다 -
    source_reader.py의 read_hwp_source_isolated()와 완전히 같은 이유(같은
    프로세스 안에 이미 열려있는 대상 문서의 HwpReport COM 연결이 pyhwpx
    Hwp.__del__의 CoUninitialize()로 함께 끊기는 것을 방지)로 반드시 격리
    실행해야 한다."""
    module_dir = os.path.dirname(os.path.abspath(__file__))
    script = (
        "import sys; sys.path.insert(0, sys.argv[2]); "
        "import json; from weekly_report_tool import read_weekly_content; "
        "print(json.dumps(read_weekly_content(sys.argv[1])))"
    )
    try:
        result = subprocess.run(
            [sys.executable, "-c", script, path, module_dir],
            capture_output=True, text=True, timeout=60,
        )
        if result.returncode != 0:
            return {"this_week": [], "next_week": []}
        return json.loads(result.stdout)
    except Exception:
        return {"this_week": [], "next_week": []}


def _append_lines_to_data_cell(hwp, keyword: str, lines: list[str]) -> bool:
    """대상 문서에서 keyword 헤더 밑 데이터 칸에 lines를 이어붙인다. 칸에
    이미 내용이 있으면 그 뒤에(자리가 겹쳐도 덮어쓰지 않고) 새 문단으로
    추가한다 - 실측 확인된 안전한 방법(모듈 docstring 참고): 기존 내용의
    마지막 줄을 find()로 다시 찾아 Cancel()로 그 줄 끝에 커서를 남긴 뒤에만
    BreakPara+insert_text로 이어붙인다. lines가 비어있으면 아무 것도 하지
    않고 False."""
    if not lines:
        return False
    if not _goto_header_cell(hwp, keyword):
        return False
    # (2026-09-07, 실측으로 새로 발견) TableColBegin은 "표의 맨 왼쪽 열(A열)"로
    # 이동하는 액션이라 여기서 부르면 안 됨(read_weekly_content() 주석 참고,
    # "다음주" 칸에 이어붙이려다 "이번주" 칸이 건드려지는 버그로 재현됨).
    hwp.Run("MoveDown")

    existing = hwp.get_selected_text()
    if existing:
        existing_lines = [ln for ln in existing.replace("\r\n", "\n").split("\n") if ln]
        last_line = existing_lines[-1] if existing_lines else existing
        # get_selected_text() 호출 후 커서는 칸 시작 위치로 돌아와 있다 -
        # 다시 칸에 들어가 마지막 줄을 찾는다.
        if not _goto_header_cell(hwp, keyword):
            return False
        hwp.Run("MoveDown")
        # AllDoc을 쓰는 이유: 막 재설정된 커서 위치(MoveDown 직후)에서의 첫
        # direction="Forward" 검색이 실제로 존재하는 텍스트도 못 찾는 현상이
        # 실측으로 재현됨(_goto_header_cell 주석 참고, 같은 pyhwpx/한글 특성).
        hwp.find(last_line, direction="AllDoc")
        hwp.Cancel()  # find()로 만든 선택 - Cancel() 후 커서는 그 줄의 끝에 남음
        hwp.Run("BreakPara")

    for i, line in enumerate(lines):
        if i > 0:
            hwp.Run("BreakPara")
        hwp.insert_text(line)
    return True


def merge_weekly_reports(report, source_paths: list[str]) -> dict:
    """source_paths 중 .hwp/.hwpx 파일들을 각각 읽어, 파란색 내용을 대상
    문서(이미 열려있는 report: HwpReport)의 "이번주"/"다음주" 칸으로 옮겨
    붙인다. 자리가 겹치면 덮어쓰지 않고 그 아래에 이어붙인다
    (_append_lines_to_data_cell 참고). 엑셀/PDF/폴더 등 다른 형식은 이
    도구의 대상이 아니므로 조용히 건너뛴다.

    반환: {"merged_files": [...], "no_content_files": [...],
           "skipped_files": [...]}
    - merged_files: 파란색 내용이 하나라도 있어 실제로 옮겨진 소스 파일
    - no_content_files: .hwp/.hwpx이지만 파란색 내용이 하나도 없던 파일
    - skipped_files: 애초에 .hwp/.hwpx가 아니어서(폴더 포함) 건드리지 않은 경로
    """
    hwp_paths = [p for p in source_paths if os.path.isfile(p) and p.lower().endswith((".hwp", ".hwpx"))]
    merged_files, no_content_files = [], []

    for path in hwp_paths:
        content = read_weekly_content_isolated(path)
        this_week = content.get("this_week", [])
        next_week = content.get("next_week", [])
        if not this_week and not next_week:
            no_content_files.append(path)
            continue
        _append_lines_to_data_cell(report.hwp, "이번주", this_week)
        _append_lines_to_data_cell(report.hwp, "다음주", next_week)
        merged_files.append(path)

    return {
        "merged_files": merged_files,
        "no_content_files": no_content_files,
        "skipped_files": [p for p in source_paths if p not in hwp_paths],
    }


def _selftest_read_weekly_content_extracts_only_blue_lines():
    """소스 문서의 "이번주"/"다음주" 데이터 칸에 파란색 실제 내용과 검정색
    안내문구가 섞여 있어도, 파란색 문단만 뽑아야 한다(실측 확인: 칸 안에서도
    find()+CharShape.Item("TextColor")로 문단별 색을 정확히 구분할 수 있음,
    모듈 docstring 참고)."""
    import os, tempfile, time
    from pyhwpx import Hwp
    import pandas as pd

    path = os.path.join(tempfile.gettempdir(), "_test_주간보고_원본.hwp")
    setup = Hwp(visible=False, new=True)
    df = pd.DataFrame({"이번주(9/1~9/5)": [""], "다음주(9/8~9/12)": [""]})
    setup.table_from_data(df, header=True, index=False, header_bold=True)

    setup.Run("TableColBegin")
    setup.Run("MoveDown")  # A2 (이번주 데이터 칸)
    setup.insert_text("- 이번주 실제 항목")
    setup.TableCellBlock()
    setup.set_font(TextColor=setup.RGBColor(0, 0, 255))
    setup.Cancel()
    setup.find("이번주 실제 항목", direction="Forward")
    setup.Cancel()  # find()로 만든 선택 - Cancel() 후 커서는 그 줄의 끝에 남음
    setup.HAction.Run("BreakPara")
    setup.insert_text("(안내문구)")  # 삽입 직후엔 직전 글자색(파랑)을 그대로 물려받음
    setup.find("(안내문구)", direction="Forward")
    setup.set_font(TextColor=setup.RGBColor(0, 0, 0))  # 검정으로 되돌려 안내문구처럼 만듦
    setup.Cancel()

    setup.Run("TableRightCell")  # B2 (다음주 데이터 칸)
    setup.insert_text("- 다음주 실제 항목")
    setup.TableCellBlock()
    setup.set_font(TextColor=setup.RGBColor(0, 0, 255))
    setup.Cancel()

    setup.save_as(path)
    setup.quit()
    time.sleep(2)

    try:
        result = read_weekly_content(path)
        assert result["this_week"] == ["- 이번주 실제 항목"], result
        assert result["next_week"] == ["- 다음주 실제 항목"], result
        print("read_weekly_content(파란색만 추출) 통과:", result)
    finally:
        os.remove(path)


def _selftest_read_weekly_content_isolated_basic():
    """read_weekly_content_isolated()가 별도 프로세스를 통해서도
    read_weekly_content()와 동일한 결과를 돌려주는지 확인한다.
    source_reader.py의 read_hwp_source_isolated()와 같은 이유(같은 프로세스
    안에 이미 열려있는 대상 문서의 HwpReport COM 연결이 pyhwpx Hwp.__del__의
    CoUninitialize()로 함께 끊기는 것을 방지)로 반드시 격리 실행해야 한다."""
    import os, tempfile, time
    from pyhwpx import Hwp
    import pandas as pd

    path = os.path.join(tempfile.gettempdir(), "_test_주간보고_격리.hwp")
    setup = Hwp(visible=False, new=True)
    df = pd.DataFrame({"이번주(9/1~9/5)": [""], "다음주(9/8~9/12)": [""]})
    setup.table_from_data(df, header=True, index=False, header_bold=True)
    setup.Run("TableColBegin")
    setup.Run("MoveDown")
    setup.insert_text("이번주 항목")
    setup.TableCellBlock()
    setup.set_font(TextColor=setup.RGBColor(0, 0, 255))
    setup.Cancel()
    setup.Run("TableRightCell")
    setup.insert_text("다음주 항목")
    setup.TableCellBlock()
    setup.set_font(TextColor=setup.RGBColor(0, 0, 255))
    setup.Cancel()
    setup.save_as(path)
    setup.quit()
    time.sleep(2)

    try:
        result = read_weekly_content_isolated(path)
        assert result == {"this_week": ["이번주 항목"], "next_week": ["다음주 항목"]}, result
        print("read_weekly_content_isolated 통과:", result)
    finally:
        os.remove(path)


def _selftest_read_weekly_content_isolated_missing_file_no_crash():
    """존재하지 않는 파일도 예외 없이 빈 결과를 반환해야 한다."""
    result = read_weekly_content_isolated("존재하지_않는_주간보고.hwp")
    assert result == {"this_week": [], "next_week": []}, result
    print("read_weekly_content_isolated(파일없음) 통과")


def _selftest_append_lines_to_data_cell_into_empty_cell():
    import os, tempfile, time
    from pyhwpx import Hwp
    import pandas as pd

    path = os.path.join(tempfile.gettempdir(), "_test_대상_빈칸.hwp")
    setup = Hwp(visible=False, new=True)
    df = pd.DataFrame({"이번주(9/1~9/5)": [""], "다음주(9/8~9/12)": [""]})
    setup.table_from_data(df, header=True, index=False, header_bold=True)
    setup.save_as(path)
    setup.quit()
    time.sleep(2)

    target = Hwp(visible=False, new=True)
    target.open(path)
    try:
        ok = _append_lines_to_data_cell(target, "이번주", ["새 항목1", "새 항목2"])
        assert ok is True

        assert _goto_header_cell(target, "이번주")
        target.Run("TableColBegin")
        target.Run("MoveDown")
        cell_text = target.get_selected_text()
        assert cell_text == "새 항목1\r\n새 항목2", repr(cell_text)
        print("_append_lines_to_data_cell(빈 칸) 통과:", repr(cell_text))
    finally:
        target.quit()
        os.remove(path)


def _selftest_append_lines_to_data_cell_preserves_existing_content():
    """핵심 안전규칙 회귀 테스트: 자리가 겹치면(이미 다른 사람 내용이 있으면)
    덮어쓰지 않고 그 아래에 이어붙여야 한다. TableCellBlock()+Cancel()로
    커서를 옮기면 선택 영역의 "시작점"에 남아 기존 내용 앞에 끼어드는 회귀가
    실제로 재현됐었다(모듈 docstring "사전조사 실측 확인" 참고) -
    find(마지막 줄)+Cancel() 방식만 옳게 동작한다."""
    import os, tempfile, time
    from pyhwpx import Hwp
    import pandas as pd

    path = os.path.join(tempfile.gettempdir(), "_test_대상_겹침.hwp")
    setup = Hwp(visible=False, new=True)
    df = pd.DataFrame({"이번주(9/1~9/5)": [""], "다음주(9/8~9/12)": [""]})
    setup.table_from_data(df, header=True, index=False, header_bold=True)
    setup.Run("TableColBegin")
    setup.Run("MoveDown")
    setup.insert_text("기존 내용1")
    setup.HAction.Run("BreakPara")
    setup.insert_text("기존 내용2")
    setup.save_as(path)
    setup.quit()
    time.sleep(2)

    target = Hwp(visible=False, new=True)
    target.open(path)
    try:
        ok = _append_lines_to_data_cell(target, "이번주", ["새 항목"])
        assert ok is True

        assert _goto_header_cell(target, "이번주")
        target.Run("TableColBegin")
        target.Run("MoveDown")
        cell_text = target.get_selected_text()
        assert cell_text == "기존 내용1\r\n기존 내용2\r\n새 항목", repr(cell_text)
        print("_append_lines_to_data_cell(자리 겹침, 이어붙이기) 통과:", repr(cell_text))
    finally:
        target.quit()
        os.remove(path)


def _selftest_append_lines_to_data_cell_empty_lines_is_noop():
    """lines가 비어있으면 아무 것도 안 하고 False를 반환해야 한다(병합할
    내용이 없는 소스 문서를 건너뛸 때 이 계약에 의존함, merge_weekly_reports
    참고)."""
    import os, tempfile, time
    from pyhwpx import Hwp
    import pandas as pd

    path = os.path.join(tempfile.gettempdir(), "_test_대상_빈리스트.hwp")
    setup = Hwp(visible=False, new=True)
    df = pd.DataFrame({"이번주(9/1~9/5)": [""], "다음주(9/8~9/12)": [""]})
    setup.table_from_data(df, header=True, index=False, header_bold=True)
    setup.save_as(path)
    setup.quit()
    time.sleep(2)

    target = Hwp(visible=False, new=True)
    target.open(path)
    try:
        ok = _append_lines_to_data_cell(target, "이번주", [])
        assert ok is False
        print("_append_lines_to_data_cell(빈 리스트, no-op) 통과")
    finally:
        target.quit()
        os.remove(path)


def _selftest_merge_weekly_reports_end_to_end():
    """전체 흐름: 소스 2개(각각 이번주/다음주에 서로 다른 파란색 항목),
    파란색 내용이 없는 소스 1개, 비-hwp 파일 1개를 섞어 첨부했을 때 -
    merged_files/no_content_files/skipped_files가 올바르게 분류되고, 대상
    문서의 "이번주" 칸(이미 다른 담당자 내용이 있는 자리 겹침 시나리오)에
    두 소스 내용이 순서대로 이어붙는지 확인한다."""
    import os, shutil, tempfile, time
    from pyhwpx import Hwp
    from hwp_report import HwpReport
    import pandas as pd

    test_dir = os.path.join(tempfile.gettempdir(), "_test_주간보고_취합")
    os.makedirs(test_dir, exist_ok=True)

    def make_source(filename, this_week_line, next_week_line):
        p = os.path.join(test_dir, filename)
        h = Hwp(visible=False, new=True)
        df = pd.DataFrame({"이번주(9/1~9/5)": [""], "다음주(9/8~9/12)": [""]})
        h.table_from_data(df, header=True, index=False, header_bold=True)
        h.Run("TableColBegin")
        h.Run("MoveDown")
        h.insert_text(this_week_line)
        h.TableCellBlock()
        h.set_font(TextColor=h.RGBColor(0, 0, 255))
        h.Cancel()
        h.Run("TableRightCell")
        h.insert_text(next_week_line)
        h.TableCellBlock()
        h.set_font(TextColor=h.RGBColor(0, 0, 255))
        h.Cancel()
        h.save_as(p)
        h.quit()
        return p

    source1 = make_source("김대리.hwp", "- 예산안 검토", "- 결과보고 작성")
    time.sleep(2)
    source2 = make_source("박주임.hwp", "- 행사 준비", "- 정산 마무리")
    time.sleep(2)

    no_content_path = os.path.join(test_dir, "빈내용.hwp")
    h_empty = Hwp(visible=False, new=True)
    df_empty = pd.DataFrame({"이번주(9/1~9/5)": [""], "다음주(9/8~9/12)": [""]})
    h_empty.table_from_data(df_empty, header=True, index=False, header_bold=True)
    h_empty.save_as(no_content_path)
    h_empty.quit()
    time.sleep(2)

    non_hwp_path = os.path.join(test_dir, "무관.txt")
    with open(non_hwp_path, "w", encoding="utf-8") as f:
        f.write("이 파일은 취합 대상이 아님")

    target_path = os.path.join(test_dir, "대상보고서.hwp")
    setup = Hwp(visible=False, new=True)
    df_target = pd.DataFrame({"이번주(9/1~9/5)": [""], "다음주(9/8~9/12)": [""]})
    setup.table_from_data(df_target, header=True, index=False, header_bold=True)
    setup.Run("TableColBegin")
    setup.Run("MoveDown")
    setup.insert_text("- 기존 담당자 항목")  # 자리 겹침 시나리오
    setup.save_as(target_path)
    setup.quit()
    time.sleep(2)

    report = None
    try:
        report = HwpReport(target_path)
        result = merge_weekly_reports(
            report, source_paths=[source1, source2, no_content_path, non_hwp_path]
        )
        assert sorted(result["merged_files"]) == sorted([source1, source2]), result
        assert result["no_content_files"] == [no_content_path], result
        assert result["skipped_files"] == [non_hwp_path], result

        assert _goto_header_cell(report.hwp, "이번주")
        report.hwp.Run("MoveDown")
        this_week_text = report.hwp.get_selected_text()
        assert this_week_text == "- 기존 담당자 항목\r\n- 예산안 검토\r\n- 행사 준비", repr(this_week_text)

        assert _goto_header_cell(report.hwp, "다음주")
        report.hwp.Run("MoveDown")
        next_week_text = report.hwp.get_selected_text()
        assert next_week_text == "- 결과보고 작성\r\n- 정산 마무리", repr(next_week_text)
        print("merge_weekly_reports(전체 흐름) 통과")
    finally:
        if report is not None:
            report.close(save=False)
        shutil.rmtree(test_dir)


if __name__ == "__main__":
    _selftest_read_weekly_content_extracts_only_blue_lines()
    _selftest_read_weekly_content_isolated_basic()
    _selftest_read_weekly_content_isolated_missing_file_no_crash()
    _selftest_append_lines_to_data_cell_into_empty_cell()
    _selftest_append_lines_to_data_cell_preserves_existing_content()
    _selftest_append_lines_to_data_cell_empty_lines_is_noop()
    _selftest_merge_weekly_reports_end_to_end()
