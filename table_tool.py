"""table_tool.py — 원본자료(엑셀) 또는 문서에서 선택한 텍스트를 읽어 한글 표로
변환해 커서 위치에 삽입하는 도구. verify_tool.py/polish_tool.py와 나란한 세 번째
도구. PRD 14-2에 따라 이번 라운드는 표 스타일을 배경색(cell_fill)+헤더 굵게
조합으로만 한정한다(테두리 스타일 등은 다루지 않음).

(2026-09-03, 실사용 피드백 반영) 처음 버전은 문서 선택 여부와 무관하게 항상
첨부된 엑셀 데이터로만 표를 만들었다 — 사용자가 실제로 문서에서 텍스트를
블록선택한 뒤 "표 만들어줘"를 실행했을 때, 선택 내용이 무시되고 엉뚱한(첨부된
엑셀의) 내용으로 표가 채워지는 문제가 실사용에서 확인됐다. 지금은 선택
여부로 분기한다: 선택돼 있으면 그 선택 텍스트를 표로, 선택이 없으면 기존대로
첨부 엑셀 데이터를 표로 만든다."""
import os
import re
import pandas as pd

from hwp_report import HwpReport

# (2026-09-04, 실사용 피드백) "표 만들어줘"로 선택한 텍스트를 표로 바꿀 때,
# 원래는 한 줄을 통째로 "내용" 한 칸에 넣었다 — 사용자가 "(1) 사전조사 :
# 하루에 한번씩" 같은 줄을 번호/항목/설명 3칸으로 나눠 넣어달라고 요청함.
# 줄 맨 앞의 흔한 번호/기호 패턴을 인식한다: 원문자(①~⑳), "(1)", "1)", "1.",
# 그리고 "-"/"•" 같은 일반 불릿. 문맥을 읽고 판단하는 방식(로컬 AI)도
# 고려했으나, 이 프로젝트의 로컬 모델(qwen3.5:2b)이 이런 구조화 응답에서
# 이미 여러 번 멈추거나 빈 응답을 낸 전례가 있어(_generate_formal_style
# 참고) 규칙 기반으로 먼저 만들고, 모델이 더 좋아지면(하드웨어 업그레이드
# 논의 참고) AI 기반 분리를 옵션으로 추가하기로 사용자와 합의함.
_BULLET_PATTERN = re.compile(
    r'^\s*((?:[①-⑳])|(?:\([0-9]+\))|(?:[0-9]+[.)])|(?:[-•]))\s*'
)


def _split_line_into_columns(line: str) -> tuple[str, str, str]:
    """한 줄을 (번호, 항목, 설명) 3칸으로 나눈다.

    번호: 줄 맨 앞의 흔한 번호/기호(_BULLET_PATTERN) — 없으면 빈 칸(사용자
    요청: "칸1에는 숫자가 들어가, 없는 경우엔 칸2에 들어가는 거지" — 번호가
    없다고 번호 칸에 본문을 채우지 않고, 항목 칸부터 채운다).
    항목/설명: 번호를 뗀 나머지를 콜론(반각 : 또는 전각 ：) 기준으로 앞/뒤
    분리한다. 둘 다 있으면 더 먼저(왼쪽에) 나오는 콜론을 기준으로 삼는다.
    콜론이 아예 없으면 나머지 전체를 항목 칸에 넣고 설명 칸은 비운다."""
    bullet_match = _BULLET_PATTERN.match(line)
    if bullet_match:
        bullet = bullet_match.group(1)
        rest = line[bullet_match.end():]
    else:
        bullet = ""
        rest = line

    colon_positions = [p for p in (rest.find(":"), rest.find("：")) if p != -1]
    if colon_positions:
        sep_pos = min(colon_positions)
        title = rest[:sep_pos].strip()
        description = rest[sep_pos + 1:].strip()
    else:
        title = rest.strip()
        description = ""
    return bullet, title, description

# PRD 14-2가 "표 스타일은 배경색+헤더 굵게 2~3종으로 한정"이라고 명시했던
# 범위 제한이었고, 사용자가 이후 "표 모양을 더 다양하게" 해달라고 요청해
# 2026-09-03 이 라운드에서 확장함 — 미리보기 팝업(chat_assistant.py)에 쓰이는
# 값과 반드시 일치해야 한다. "style" 파라미터는 이 dict의 key만 허용한다.
#
# 각 스타일은 header_fill(헤더 행 배경색, None이면 채우지 않음)과
# stripe_fill(데이터 행 홀수번째에 입힐 옅은 배경색, None이면 줄무늬 없음)
# 두 축을 조합한다 — 이번 확장에서 "헤더 강조"와 "행 구분(줄무늬)"이 서로
# 다른 축임을 확인했고(build_table 참고), 굳이 서로 배타적일 필요가 없어
# 나중에 "파란 헤더 + 줄무늬" 같은 조합도 쉽게 추가할 수 있게 열어뒀다
# (지금은 조합 스타일까지는 만들지 않고 6종만 우선 제공).
TABLE_STYLES = {
    "default": {"label": "기본형 (테두리만)", "header_fill": None, "stripe_fill": None},
    "gray_header": {"label": "회색 헤더", "header_fill": (217, 217, 217), "stripe_fill": None},
    "blue_header": {"label": "파란 헤더", "header_fill": (198, 224, 241), "stripe_fill": None},
    "green_header": {"label": "초록 헤더", "header_fill": (200, 230, 201), "stripe_fill": None},
    "amber_header": {"label": "주황 헤더", "header_fill": (255, 224, 178), "stripe_fill": None},
    "striped": {"label": "줄무늬(홀짝 구분)", "header_fill": None, "stripe_fill": (240, 240, 240)},
}


def _first_excel_source(source_paths: list[str]) -> str | None:
    """PRD 14-2: 원본자료가 여러 개면 첫 번째 엑셀 파일을 기본으로 선택한다
    (여러 개 중 고르는 UI는 이번 라운드에 만들지 않음). source_paths에는 파일
    경로와 폴더 경로가 섞여 있을 수 있다(F12 1단계의 "+" 첨부 방식) — 폴더는
    이 함수에서는 건드리지 않고 개별 파일 경로만 검사한다(폴더까지 훑는 건
    source_reader.read_source_files의 책임이라 여기서 중복하지 않음).
    """
    for path in source_paths:
        if os.path.isfile(path) and path.lower().endswith((".xlsx", ".xls")):
            return path
    return None


def insert_table_from_source(report: HwpReport, source_paths: list[str], style: str) -> dict:
    """원본자료 중 첫 엑셀 파일을 읽어 커서 위치에 표로 삽입한다.

    style은 TABLE_STYLES의 key 중 하나여야 한다. header_fill이 있으면
    pyhwpx의 table_from_data() 자체에 내장된 cell_fill 파라미터로 헤더 행
    배경색을 적용한다 — table_from_data 소스(site-packages/pyhwpx/core.py,
    5087번 줄 근처)를 직접 확인한 결과, header_bold/cell_fill 처리가 이미
    함수 내부에 구현되어 있다:

        self.TableColBegin(); self.TableColPageUp()
        self.TableCellBlockExtendAbs(); self.TableColEnd()
        if header_bold: self.CharShapeBold()
        if cell_fill: self.cell_fill(cell_fill)  # cell_fill이 튜플이면 그 색 사용
        self.TableColBegin(); self.Cancel()

    즉 표 삽입 직후 커서가 첫 열 맨 위(제목행 포함 첫 칸)로 돌아가 있고,
    선택은 Cancel()로 이미 해제된 상태 — 헤더 행을 별도로 다시 블록
    선택할 필요가 없다(계획서가 우려했던 TableCellBlockRow 같은 이름의
    "행 전체 선택" 메서드는 pyhwpx에 존재하지 않고, table_from_data가
    TableColPageUp+TableCellBlockExtendAbs+TableColEnd 조합으로 이미
    같은 일을 하고 있었다 — grep으로 직접 확인함).

    원본자료에 엑셀 파일이 하나도 없으면 문서를 건드리지 않고
    inserted=False를 반환한다.

    회귀 테스트로 확인된 크래시 방지: 문서에 텍스트가 선택된 상태로 이 함수를
    호출해도 크래시하지 않는다. table_from_data() 호출 직전에 항상
    report.hwp.Cancel()로 선택을 해제하기 때문이다(선택이 없었으면 no-op).
    이 가드가 없으면, 선택된 상태에서 table_from_data()가 내부적으로 실행하는
    "TableCreate" HAction이 COM 에러를 던지고 pyhwpx의 create_table()이
    ctrl=None에 .Properties를 대입하려다 AttributeError로 이어지는 크래시가
    결정적으로(재현성 100%) 발생한다 — numbering_tool.py의
    insert_numbering_prefix()가 이미 같은 철학(삽입 직전 Cancel())으로 방어한
    것과 동일한 안전장치를 여기에도 맞춘 것이다. numbering_tool.py의 텍스트
    프리픽스 케이스와 달리 표는 새로 삽입되는 콘텐츠라서(기존 텍스트를
    대체하는 게 아니라 커서 위치에 새 표를 끼워넣음), 선택했던 텍스트가
    사라지거나 대체될 위험 자체가 원래 없었다 — Cancel()은 오직
    table_from_data/create_table이 COM 에러 없이 안정적으로 동작하게 만들기
    위한 방어일 뿐이다. Cancel() 이후 커서는 선택 영역의 끝점에 남고, 표는
    바로 그 커서 위치(선택했던 콘텐츠 바로 뒤)에 삽입된다 — 선택했던 텍스트
    자체는 문서에 그대로 남는다(직접 재현 테스트로 확인,
    _selftest_insert_table_from_source_with_selection_does_not_crash 참고).
    """
    if style not in TABLE_STYLES:
        raise ValueError(f"알 수 없는 표 스타일: {style}")

    # (2026-09-03, 실사용 피드백) 문서에 선택된 텍스트가 있으면, 그 선택
    # 내용을 표로 만든다 — 첨부된 엑셀 데이터는 이 경우 쓰지 않는다.
    # "표 만들어줘"를 실행하기 전에 사용자가 문서에서 직접 블록선택해둔
    # 것이 있다면, 그게 첨부 엑셀보다 우선한다는 뜻으로 해석했다(사용자가
    # 명시적으로 "둘 다 지원"을 요청함).
    if report.hwp.SelectionMode != 0:
        return _insert_table_from_selected_text(report, style)

    excel_path = _first_excel_source(source_paths)
    if excel_path is None:
        return {"inserted": False, "reason": "원본자료 중 엑셀 파일이 없습니다"}

    build_table(report, excel_path, style)
    return {"inserted": True, "style": style, "source_file": excel_path}


def _insert_table_from_selected_text(report: HwpReport, style: str) -> dict:
    """문서에서 현재 선택된 텍스트(여러 문단 가능)를 한 줄당 한 행씩 표로
    변환해 선택 영역 바로 뒤에 삽입한다.

    선택했던 원본 텍스트는 지우지 않고 그대로 둔다 — insert_text("")로
    선택을 비워보려는 시도를 직접 테스트해봤으나 신뢰할 수 없게 동작했다
    (선택이 실제로 지워지지 않고 SelectionMode도 0으로 안 돌아오는 경우가
    실측됨). 반면 Cancel()은 이미 이 프로젝트에서 여러 번 검증된 안전한
    선택 해제 방법이라(table_tool.py/numbering_tool.py의 다른 크래시 수정
    참고), 원본을 지우려 하지 않고 "선택 내용 바로 뒤에 표를 추가로
    삽입"하는 더 안전한 방식을 택했다. 사용자가 원본 줄을 지우고 싶으면
    표 삽입 후 직접 지우면 된다(자동저장 안 하므로 Ctrl+Z로도 되돌릴 수 있음).
    """
    selected = report.hwp.get_selected_text(keep_select=True)
    lines = [line.strip() for line in selected.replace("\r\n", "\n").split("\n") if line.strip()]
    if not lines:
        report.hwp.Cancel()
        return {"inserted": False, "reason": "선택한 텍스트가 비어있습니다"}

    # (2026-09-04, 실사용 피드백) "(1) 사전조사 : 하루에 한번씩" 같은 줄은
    # 번호/항목/설명 3칸으로 나눈다(_split_line_into_columns). 다만 선택한
    # 줄 전부가 번호도 콜론도 없는 평범한 목록(예: 그냥 항목 나열)이면,
    # 번호/설명 칸이 모든 행에서 항상 비어있는 표가 되어 예전(1칸)보다
    # 오히려 정보가 없는 빈 칸만 늘어난다 — 그래서 하나라도 번호나 콜론이
    # 실제로 인식된 줄이 있을 때만 3칸으로 만들고, 전혀 없으면 기존처럼
    # "내용" 1칸짜리 표로 만든다(하위호환).
    parsed = [_split_line_into_columns(line) for line in lines]
    has_structure = any(bullet or description for bullet, _title, description in parsed)
    if has_structure:
        df = pd.DataFrame(parsed, columns=["번호", "항목", "설명"])
    else:
        df = pd.DataFrame({"내용": lines})
    build_table(report, df, style)
    return {"inserted": True, "style": style, "source": "selection", "row_count": len(lines)}


def build_table(report: HwpReport, data, style: str) -> None:
    """table_from_data 호출을 감싸는 공통 헬퍼. data는 엑셀 파일 경로(str) 또는
    pandas DataFrame — table_from_data가 둘 다 받는다(pyhwpx 소스 확인됨).

    코드품질 검토에서 발견된 크래시 버그 수정: 문서에 텍스트가 선택된 채로
    table_from_data()를 호출하면, 그 내부의 create_table()이 실행하는
    "TableCreate" HAction이 COM 에러(pywintypes.com_error, -2147417851)를
    던지고, 이어서 pyhwpx core.py의 create_table() finally 블록에서
    `ctrl = self.hwp.CurSelectedCtrl or self.hwp.ParentCtrl` 결과가 None인
    채로 `ctrl.Properties = pset`를 실행하다 AttributeError로 이어져
    프로그램이 크래시한다(직접 재현 확인, 두 번 연속 동일 스택트레이스 —
    간헐적 COM 불안정성이 아니라 선택 상태에 의해 결정적으로 재현됨).
    numbering_tool.py의 insert_numbering_prefix()가 이미 같은 부류의
    "선택된 콘텐츠" 위험에 대해 삽입 직전 Cancel()로 방어한 것과 동일한
    철학으로, 여기서도 표 생성 직전에 선택을 항상 해제한다 — 선택이
    없었으면 no-op이고, 선택이 있었으면(_insert_table_from_selected_text가
    이미 텍스트를 읽어둔 뒤이므로) 그 콘텐츠는 그대로 보존된 채 선택만
    풀린다. 표는 커서 위치(Cancel() 이후 선택 영역의 끝점)에 새로
    삽입되므로, 선택했던 텍스트가 지워지거나 대체될 위험 자체가 없다.

    (2026-09-03, 실사용 피드백) 표 삽입 후 모든 셀을 가운데 정렬한다 —
    사용자가 "표로 만들었으면 가운데 줄 정렬(디폴트)"을 요청함. 표 전체
    셀을 블록선택하는 전용 API가 pyhwpx에 없어서, table_from_data 내부가
    헤더 행을 선택할 때 쓰는 것과 같은 조합(TableColBegin/TableCellBlockExtendAbs/
    TableColEnd)에 TableColPageDown()을 추가해 마지막 행까지 확장한다 —
    직접 테스트로 SelectionMode가 셀 블록 선택 상태(19)로 바뀌고 스크린샷
    상 모든 셀(헤더/데이터 전부)이 가운데 정렬되는 것을 확인했다.
    ParagraphShapeAlignCenter()는 별도 래퍼가 pyhwpx에 없지만, pyhwpx가
    HAction 이름과 일치하는 메서드를 동적으로 만들어주므로 그대로 호출 가능
    하다(ParagraphShapeAlignLeft/Right/Justify 등과 동일한 패턴).

    (2026-09-03, "표 모양을 더 다양하게" 요청 반영) stripe_fill이 있는
    스타일("striped")은 데이터 행 중 홀수번째(0, 2, 4...)에 옅은 배경색을
    입혀 줄무늬를 만든다. 표 전체를 한 번에 블록선택하는 가운데 정렬과
    달리, 줄무늬는 "한 행씩" 골라 칠해야 해서 다른 방식이 필요했다 — 직접
    테스트로 확인한 절차: 표 맨 위(헤더) 첫 칸으로 이동 → 데이터 행 수만큼
    MoveDown(한 칸 아래로, 표 안에서는 같은 열의 다음 행으로 이동하는
    표준 HAction)을 반복하면서, 그 행 전체를 TableColBegin→
    TableCellBlockExtendAbs→TableColEnd로 선택해 cell_fill() 적용 후
    Cancel()로 선택 해제 — 이 조합이 실제로 지정한 행에만(전체가 아니라)
    정확히 칠해지는 걸 스크린샷으로 확인했다(`.tmp/screenshots/striped_table_test.png`
    참고, 커밋 시 정식 self-test로도 재확인). 가운데 정렬 블록 이후에는
    커서/선택이 표의 마지막 셀에 남아있으므로, 줄무늬를 칠하기 전에
    TableColBegin+TableColPageUp으로 맨 위 첫 칸으로 다시 이동해야 한다 —
    이 재이동을 빠뜨리면 MoveDown이 표 밖으로 나가버려 줄무늬가 하나도
    안 칠해지거나 엉뚱한 셀에 칠해진다."""
    style_info = TABLE_STYLES[style]
    header_fill = style_info["header_fill"]
    stripe_fill = style_info.get("stripe_fill")
    row_count = _data_row_count(data)

    report.hwp.Cancel()
    report.hwp.table_from_data(
        data,
        header=True,
        index=False,
        header_bold=True,
        cell_fill=header_fill if header_fill is not None else False,
    )
    report.hwp.TableColBegin()
    report.hwp.TableColPageUp()
    report.hwp.TableCellBlockExtendAbs()
    report.hwp.TableColPageDown()
    report.hwp.TableColEnd()
    report.hwp.ParagraphShapeAlignCenter()
    report.hwp.Cancel()

    if stripe_fill is not None:
        report.hwp.TableColBegin()
        report.hwp.TableColPageUp()  # 맨 위(헤더) 첫 칸으로 확실히 재이동
        for i in range(row_count):
            report.hwp.Run("MoveDown")
            if i % 2 == 0:
                report.hwp.TableColBegin()
                report.hwp.TableCellBlockExtendAbs()
                report.hwp.TableColEnd()
                report.hwp.cell_fill(stripe_fill)
                report.hwp.Cancel()
                report.hwp.TableColBegin()
        report.hwp.Cancel()


def _data_row_count(data) -> int:
    """table_from_data에 넘길 data(엑셀 경로 또는 DataFrame)의 실제 데이터
    행 수(헤더 제외)를 미리 계산한다. table_from_data 내부에도 똑같은
    변환 로직이 있지만(엑셀 경로면 pd.read_excel/read_csv), pyhwpx가 표
    생성 후 행 수를 다시 조회하는 공개 API를 제공하지 않아 줄무늬를 몇
    번 반복할지 미리 알 방법이 없다 — 그래서 표를 만들기 전에 같은 방식으로
    한 번 더 읽는다(작은 표 하나 분량이라 중복 비용은 무시할 만함)."""
    if isinstance(data, str):
        return len(pd.read_excel(data) if ".xls" in data.lower() else pd.read_csv(data))
    return len(data)


def _selftest_insert_table_from_source_default_style():
    """엑셀 원본자료 하나를 기본 스타일(배경색 없음, 헤더만 굵게)로 표 삽입하면,
    문서에 표가 실제로 생기고 헤더/데이터 값이 포함되는지 확인한다."""
    import os, tempfile, openpyxl
    from pyhwpx import Hwp
    from hwp_report import HwpReport

    test_dir = os.path.join(tempfile.gettempdir(), "_test_표원본")
    os.makedirs(test_dir, exist_ok=True)
    source_path = os.path.join(test_dir, "원본.xlsx")
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Sheet1"
    ws.append(["항목", "금액"])
    ws.append(["인건비", 1000000])
    ws.append(["운영비", 500000])
    wb.save(source_path)

    report_path = os.path.join(tempfile.gettempdir(), "_test_보고서_표.hwp")
    setup = Hwp(visible=False, new=True)
    setup.save_as(report_path)  # 빈 문서
    setup.quit()

    report = None
    try:
        report = HwpReport(report_path)
        result = insert_table_from_source(report, source_paths=[source_path], style="default")
        assert result["inserted"] is True, result

        final_text = report.get_text()
        assert "항목" in final_text, final_text
        assert "인건비" in final_text, final_text
        assert "1000000" in final_text, final_text
        print("insert_table_from_source(기본 스타일) 통과")
    finally:
        if report is not None:
            report.close(save=False)
        import shutil
        shutil.rmtree(test_dir)
        os.remove(report_path)


def _selftest_insert_table_from_source_colored_header():
    """gray_header 스타일을 쓰면 헤더 행 셀의 배경색이 실제로 바뀌는지 확인한다."""
    import os, tempfile, openpyxl
    from pyhwpx import Hwp
    from hwp_report import HwpReport

    test_dir = os.path.join(tempfile.gettempdir(), "_test_표스타일")
    os.makedirs(test_dir, exist_ok=True)
    source_path = os.path.join(test_dir, "원본.xlsx")
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Sheet1"
    ws.append(["항목", "금액"])
    ws.append(["인건비", 1000000])
    wb.save(source_path)

    report_path = os.path.join(tempfile.gettempdir(), "_test_보고서_표스타일.hwp")
    setup = Hwp(visible=False, new=True)
    setup.save_as(report_path)
    setup.quit()

    report = None
    try:
        report = HwpReport(report_path)
        result = insert_table_from_source(report, source_paths=[source_path], style="gray_header")
        assert result["inserted"] is True, result
        # 셀 배경색 자체를 읽는 API가 이 프로젝트엔 아직 없으므로, 여기서는
        # "예외 없이 완료됨 + 표 내용은 정상"만 확인한다 — 실제 색상이 눈으로
        # 봤을 때 맞는지는 아침에 스크린샷으로 사람이 확인해야 한다(PRD 14-6).
        final_text = report.get_text()
        assert "인건비" in final_text, final_text
        print("insert_table_from_source(gray_header) 통과 (배경색 육안 확인은 아침에 필요)")
    finally:
        if report is not None:
            report.close(save=False)
        import shutil
        shutil.rmtree(test_dir)
        os.remove(report_path)


def _selftest_insert_table_from_source_no_excel_source():
    """원본자료에 엑셀 파일이 하나도 없으면(예: 이미지/한글 파일만 첨부됐거나
    원본자료 자체를 안 붙인 경우) 예외 없이 inserted=False를 반환하고 문서를
    전혀 건드리지 않는지 확인한다 (코드품질 검토에서 지적된 회귀테스트 공백을
    메움 — _first_excel_source 자체의 None 반환은 이미 확인됐지만, 그 결과를
    받은 insert_table_from_source가 실제로 문서를 안전하게 스킵하는지는
    아직 테스트가 없었음)."""
    import os, tempfile
    from pyhwpx import Hwp
    from hwp_report import HwpReport

    report_path = os.path.join(tempfile.gettempdir(), "_test_보고서_표없음.hwp")
    setup = Hwp(visible=False, new=True)
    setup.save_as(report_path)  # 빈 문서
    setup.quit()

    report = None
    try:
        report = HwpReport(report_path)
        # 엑셀이 하나도 없는 상황(빈 리스트, 그리고 비엑셀 파일만 있는 경우 둘 다)
        result_empty = insert_table_from_source(report, source_paths=[], style="default")
        assert result_empty == {"inserted": False, "reason": "원본자료 중 엑셀 파일이 없습니다"}, result_empty

        non_excel_path = os.path.join(tempfile.gettempdir(), "_test_표없음_원본.txt")
        with open(non_excel_path, "w", encoding="utf-8") as f:
            f.write("엑셀 아님")
        try:
            result_non_excel = insert_table_from_source(
                report, source_paths=[non_excel_path], style="default"
            )
            assert result_non_excel["inserted"] is False, result_non_excel
        finally:
            os.remove(non_excel_path)

        # 문서가 실제로 전혀 건드려지지 않았는지 확인 — 표가 생겼다면 get_text()에
        # 뭔가 내용이 생겼을 것이다.
        assert report.get_text() == "", repr(report.get_text())
        print("insert_table_from_source(엑셀 없음) 통과: 문서 안 건드리고 inserted=False")
    finally:
        if report is not None:
            report.close(save=False)
        os.remove(report_path)


def _selftest_insert_table_from_source_with_single_line_selection_does_not_crash():
    """회귀 테스트(코드품질 검토에서 발견된 크래시 버그): 문서에 텍스트가 선택된
    채로 insert_table_from_source()를 호출해도 크래시하면 안 된다.

    선택된 상태로 report.hwp.table_from_data(...)를 호출하면 내부적으로 실행하는
    "TableCreate" HAction이 COM 에러(-2147417851)를 던지고, 이어서 pyhwpx의
    create_table() 내부 finally 블록(site-packages/pyhwpx/core.py, 5916번 줄 근처)의
    `ctrl = self.hwp.CurSelectedCtrl or self.hwp.ParentCtrl; ... ctrl.Properties = pset`
    에서 표 생성 자체가 실패해 ctrl이 None인 채로 남아 AttributeError로 이어지는
    크래시가 재현된다(직접 재현 확인, 두 번 연속 동일 스택트레이스 — 결정적 버그).

    (2026-09-03 갱신) 이 테스트는 원래 "선택이 있어도 첨부 엑셀로 표를 만든다"는
    옛 동작을 검증했다. 지금은 선택이 있으면 그 선택 내용으로 표를 만드는 게
    새 기본 동작이라(실사용 피드백 반영), 이 테스트의 목적(크래시 안 남)은
    그대로 두고 기대값만 새 동작(선택 텍스트가 표에 들어감, 엑셀 내용은 안 들어감)에
    맞게 갱신한다. 여러 문단을 선택했을 때 각 줄이 표의 행이 되는지는
    `_selftest_insert_table_from_selected_text_multiline`에서 별도로 검증한다."""
    import os, tempfile, openpyxl
    from pyhwpx import Hwp
    from hwp_report import HwpReport

    test_dir = os.path.join(tempfile.gettempdir(), "_test_표선택유지")
    os.makedirs(test_dir, exist_ok=True)
    source_path = os.path.join(test_dir, "원본.xlsx")
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Sheet1"
    ws.append(["항목", "금액"])
    ws.append(["인건비", 1000000])
    wb.save(source_path)

    report_path = os.path.join(tempfile.gettempdir(), "_test_보고서_표선택.hwp")
    setup = Hwp(visible=False, new=True)
    setup.insert_text("가나다라마바사")
    setup.save_as(report_path)
    setup.quit()

    report = None
    try:
        report = HwpReport(report_path)
        found = report.hwp.find("다라마", direction="AllDoc")
        assert found, "테스트 문장에서 '다라마'를 못 찾음"
        assert report.hwp.SelectionMode != 0, "선택이 안 된 상태로 테스트가 시작됨"

        result = insert_table_from_source(report, source_paths=[source_path], style="default")
        assert result["inserted"] is True, result
        assert result.get("source") == "selection", result

        final_text = report.get_text()
        assert "다라마" in final_text, (
            "선택된 콘텐츠('다라마')가 사라짐: " + repr(final_text)
        )
        assert "인건비" not in final_text, (
            "선택이 있었는데도 엑셀 데이터가 표에 들어감(옛 동작으로 회귀됨): " + repr(final_text)
        )
        print("insert_table_from_source(선택 있음, 크래시 없음, 선택내용 우선) 통과:", repr(final_text))
    finally:
        if report is not None:
            report.close(save=False)
        import shutil
        shutil.rmtree(test_dir)
        os.remove(report_path)


def _selftest_insert_table_from_selected_text_multiline():
    """(2026-09-03, 실사용 피드백 반영 신규 테스트) 문서에서 여러 문단(줄)을
    블록선택한 채로 "표 만들어줘"에 해당하는 insert_table_from_source()를
    호출하면, 선택된 각 줄이 표의 한 행씩으로 들어가는지 확인한다.

    find()로는 이 시나리오를 재현할 수 없었다 — 실제로 시도해본 결과 pyhwpx의
    find()는 "\\r\\n"(문단 구분자)가 포함된 여러 문단짜리 검색어를 문단 경계
    너머로는 찾지 못한다(직접 재현: find() 실패, SelectionMode 0으로 남음).
    반면 사용자가 실제로 문서에서 마우스로 여러 줄을 드래그해 선택하는 것은
    문단 경계와 무관하게 항상 되므로, 이 테스트에서는 find() 대신
    hwp.select_text(spara, spos, epara, epos)로 문단 번호 기준 범위 선택을
    재현한다(직접 프로브로 정상 동작 확인됨) — 실사용 시나리오(마우스 드래그
    선택)를 더 정확히 흉내낸다."""
    import os, tempfile
    from pyhwpx import Hwp
    from hwp_report import HwpReport

    report_path = os.path.join(tempfile.gettempdir(), "_test_표_여러줄선택.hwp")
    setup = Hwp(visible=False, new=True)
    setup.insert_text("앞 문단입니다"); setup.BreakPara()
    setup.insert_text("이슈사항 논의"); setup.BreakPara()
    setup.insert_text("이슈에 따른 결론"); setup.BreakPara()
    setup.insert_text("8월달 논의사항 정리"); setup.BreakPara()
    setup.insert_text("처장님 이사말씀"); setup.BreakPara()
    setup.insert_text("뒤 문단입니다")
    setup.save_as(report_path)
    setup.quit()

    report = None
    try:
        report = HwpReport(report_path)
        # 문단 번호(0-based): 0=앞 문단, 1~4=목록 4줄, 5=뒤 문단
        ok = report.hwp.select_text(1, 0, 4, -1)
        assert ok, "문단 1~4 선택 실패"
        assert report.hwp.SelectionMode != 0, "선택이 안 된 상태로 테스트가 시작됨"

        result = insert_table_from_source(report, source_paths=[], style="blue_header")
        assert result["inserted"] is True, result
        assert result.get("row_count") == 4, result

        final_text = report.get_text()
        for line in ["이슈사항 논의", "이슈에 따른 결론", "8월달 논의사항 정리", "처장님 이사말씀"]:
            assert line in final_text, (line, final_text)
        assert "앞 문단입니다" in final_text, "앞 문단이 사라짐: " + repr(final_text)
        assert "뒤 문단입니다" in final_text, "뒤 문단이 사라짐: " + repr(final_text)
        print("insert_table_from_source(여러 줄 선택 → 표) 통과:", repr(final_text))
    finally:
        if report is not None:
            report.close(save=False)
        os.remove(report_path)


def _selftest_split_line_into_columns():
    """_split_line_into_columns()가 번호/콜론 유무에 따라 올바르게 3칸으로
    나누는지 확인한다(2026-09-04, 실사용 피드백 — 사용자가 준 예시 포함)."""
    assert _split_line_into_columns("(1) 사전조사 : 하루에 한번씩") == ("(1)", "사전조사", "하루에 한번씩"), \
        _split_line_into_columns("(1) 사전조사 : 하루에 한번씩")
    assert _split_line_into_columns("① 결과보고") == ("①", "결과보고", ""), \
        _split_line_into_columns("① 결과보고")
    assert _split_line_into_columns("일반사항 검토 필요") == ("", "일반사항 검토 필요", ""), \
        _split_line_into_columns("일반사항 검토 필요")
    assert _split_line_into_columns("2) 예산 편성 : 부서별로 배정") == ("2)", "예산 편성", "부서별로 배정"), \
        _split_line_into_columns("2) 예산 편성 : 부서별로 배정")
    print("_split_line_into_columns 통과")


def _selftest_insert_table_from_selected_text_splits_structured_lines():
    """(2026-09-04, 실사용 피드백) 선택한 줄에 번호/콜론 구조가 있으면
    "내용" 1칸이 아니라 번호/항목/설명 3칸 표로 만든다. 사용자가 준 예시
    ("(1) 사전조사 : 하루에 한번씩")를 포함한 두 줄로 확인한다 — 하나는
    번호+콜론이 다 있고, 하나는 둘 다 없는 경우를 섞어 has_structure
    판단(하나라도 구조가 있으면 3칸)과 빈 칸 처리(번호 없는 줄은 번호 칸이
    빈 채로 항목 칸부터 채워짐)를 함께 검증한다."""
    import os, tempfile
    from pyhwpx import Hwp
    from hwp_report import HwpReport

    report_path = os.path.join(tempfile.gettempdir(), "_test_표_구조분리.hwp")
    setup = Hwp(visible=False, new=True)
    setup.insert_text("앞 문단"); setup.BreakPara()
    setup.insert_text("(1) 사전조사 : 하루에 한번씩"); setup.BreakPara()
    setup.insert_text("일반사항 검토 필요"); setup.BreakPara()
    setup.insert_text("뒤 문단")
    setup.save_as(report_path)
    setup.quit()

    report = None
    try:
        report = HwpReport(report_path)
        ok = report.hwp.select_text(1, 0, 2, -1)
        assert ok, "문단 1~2 선택 실패"

        result = insert_table_from_source(report, source_paths=[], style="default")
        assert result["inserted"] is True, result
        assert result.get("row_count") == 2, result

        final_text = report.get_text()
        for expected in ["번호", "항목", "설명", "(1)", "사전조사", "하루에 한번씩", "일반사항 검토 필요"]:
            assert expected in final_text, (expected, final_text)
        print("insert_table_from_source(구조화된 줄 → 번호/항목/설명 3칸) 통과:", repr(final_text))
    finally:
        if report is not None:
            report.close(save=False)
        os.remove(report_path)


def _selftest_build_table_striped_style_colors_odd_rows_only():
    """(2026-09-03, "표 모양을 더 다양하게" 요청 반영) "striped" 스타일은
    데이터 행 중 0,2번째(1행,3행)에만 옅은 회색(240,240,240)이 칠해지고
    1번째(2행)는 안 칠해져야 한다. HParameterSet.HCellBorderFill의
    FillAttr.WinBrushFaceColor를 직접 읽어 확인한다 — 0xBBGGRR 순서 정수로
    인코딩되며(직접 프로브로 확인, 240,240,240 → 0xF0F0F0), 안 칠해진
    셀은 0으로 나온다."""
    import os, tempfile
    from pyhwpx import Hwp
    from hwp_report import HwpReport

    report_path = os.path.join(tempfile.gettempdir(), "_test_표_줄무늬.hwp")
    setup = Hwp(visible=False, new=True)
    setup.save_as(report_path)
    setup.quit()

    report = None
    try:
        report = HwpReport(report_path)
        df = pd.DataFrame({"항목": ["A", "B", "C"], "금액": [1, 2, 3]})
        build_table(report, df, "striped")

        def fill_color_at(text: str) -> int:
            report.hwp.find(text, direction="AllDoc")
            pset = report.hwp.HParameterSet.HCellBorderFill
            report.hwp.HAction.GetDefault("CellFill", pset.HSet)
            return pset.FillAttr.WinBrushFaceColor

        assert fill_color_at("A") == 0xF0F0F0, fill_color_at("A")
        assert fill_color_at("B") == 0, fill_color_at("B")
        assert fill_color_at("C") == 0xF0F0F0, fill_color_at("C")
        print("build_table(striped, 홀수 데이터행만 줄무늬) 통과")
    finally:
        if report is not None:
            report.close(save=False)
        os.remove(report_path)


def _selftest_build_table_centers_all_cells():
    """(2026-09-03, 실사용 피드백) build_table()로 만든 표는 헤더/데이터 셀
    전부 가운데 정렬이어야 한다. HWP의 문단모양 Alignment 값은
    0=양쪽/1=왼쪽/2=오른쪽/3=가운데/4=배분/5=나눔이다(HParameterSet.HParaShape
    직접 읽어 확인) — find()로 커서를 각 셀로 옮긴 뒤 AlignType을 읽어
    검증한다."""
    import os, tempfile
    from pyhwpx import Hwp
    from hwp_report import HwpReport

    report_path = os.path.join(tempfile.gettempdir(), "_test_표_가운데정렬.hwp")
    setup = Hwp(visible=False, new=True)
    setup.save_as(report_path)
    setup.quit()

    report = None
    try:
        report = HwpReport(report_path)
        df = pd.DataFrame({"항목": ["인건비", "운영비"], "금액": [1000000, 500000]})
        build_table(report, df, "default")

        def align_type_at(text: str) -> int:
            report.hwp.find(text, direction="AllDoc")
            pset = report.hwp.HParameterSet.HParaShape
            report.hwp.HAction.GetDefault("ParagraphShape", pset.HSet)
            return pset.AlignType

        for text in ["항목", "금액", "인건비", "1000000"]:
            assert align_type_at(text) == 3, f"{text} 셀이 가운데 정렬이 아님"
        print("build_table(모든 셀 가운데 정렬) 통과")
    finally:
        if report is not None:
            report.close(save=False)
        os.remove(report_path)


if __name__ == "__main__":
    _selftest_insert_table_from_source_default_style()
    _selftest_insert_table_from_source_colored_header()
    _selftest_insert_table_from_source_no_excel_source()
    _selftest_insert_table_from_source_with_single_line_selection_does_not_crash()
    _selftest_insert_table_from_selected_text_multiline()
    _selftest_build_table_centers_all_cells()
    _selftest_build_table_striped_style_colors_odd_rows_only()
    _selftest_split_line_into_columns()
    _selftest_insert_table_from_selected_text_splits_structured_lines()
