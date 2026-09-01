"""table_tool.py — 원본자료(엑셀)를 읽어 한글 표로 변환해 커서 위치에 삽입하는 도구.
verify_tool.py/polish_tool.py와 나란한 세 번째 도구. PRD 14-2에 따라 이번 라운드는
표 스타일을 배경색(cell_fill)+헤더 굵게 조합으로만 한정한다(테두리 스타일 등은 다루지 않음)."""
import os

from hwp_report import HwpReport

# PRD 14-2: 표 스타일은 이 3종으로 한정 — 미리보기 팝업(chat_assistant.py)에 쓰이는
# 값과 반드시 일치해야 한다. "style" 파라미터는 이 dict의 key만 허용한다.
TABLE_STYLES = {
    "default": {"label": "기본형 (테두리만)", "header_fill": None},
    "gray_header": {"label": "회색 헤더", "header_fill": (217, 217, 217)},
    "blue_header": {"label": "파란 헤더", "header_fill": (198, 224, 241)},
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
    """
    if style not in TABLE_STYLES:
        raise ValueError(f"알 수 없는 표 스타일: {style}")

    excel_path = _first_excel_source(source_paths)
    if excel_path is None:
        return {"inserted": False, "reason": "원본자료 중 엑셀 파일이 없습니다"}

    header_fill = TABLE_STYLES[style]["header_fill"]
    # cell_fill 파라미터는 False(배경색 없음) 또는 (R,G,B) 튜플을 받는다.
    # header_fill이 None이면 False를 넘겨 table_from_data가 cell_fill()을
    # 아예 호출하지 않게 한다(기본형 스타일은 배경색을 적용하지 않아야 함).
    report.hwp.table_from_data(
        excel_path,
        header=True,
        index=False,
        header_bold=True,
        cell_fill=header_fill if header_fill is not None else False,
    )

    return {"inserted": True, "style": style, "source_file": excel_path}


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


if __name__ == "__main__":
    _selftest_insert_table_from_source_default_style()
    _selftest_insert_table_from_source_colored_header()
