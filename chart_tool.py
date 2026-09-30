"""chart_tool.py — 원본자료(엑셀)를 읽어 막대/꺾은선/원형 그래프 이미지를
만들어 커서 위치에 삽입하는 도구. table_tool.py/numbering_tool.py와 나란한
네 번째 도구. 표(table_tool)와 달리 그래프는 "선택한 텍스트로 만들기"가
성립하지 않는다(자유 텍스트에는 그래프로 그릴 수치 구조가 없음) — 그래서
이 도구는 항상 첨부된 엑셀 원본자료만 대상으로 한다.

(2026-09-30) matplotlib 기본 폰트(DejaVu Sans)는 한글 글리프가 없어
"항목"/"금액" 같은 라벨이 네모 박스(tofu)로 깨진다 - 반드시 이 모듈을
불러올 때 맑은 고딕으로 폰트를 맞춰야 한다(attachment_preview.py가 PIL
이미지에 쓰는 폰트와 같은 이유, 같은 폰트 파일). 이 앱은 tkinter GUI
프로세스 안에서 실행되므로, matplotlib이 기본으로 고르는 GUI 백엔드
(TkAgg)가 메인 이벤트루프와 충돌할 위험이 있어 파일 저장 전용 Agg
백엔드를 명시적으로 고정한다 - 반드시 pyplot을 처음 import하기 전에
지정해야 한다."""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

import os
import tempfile
import pandas as pd

from hwp_report import HwpReport
from table_tool import _first_excel_source

plt.rcParams["font.family"] = "Malgun Gothic"
plt.rcParams["axes.unicode_minus"] = False  # 한글 폰트로 바꾸면 마이너스 기호가 네모로 깨지는 문제 방지

CHART_TYPES = {
    "bar": {"label": "막대그래프"},
    "line": {"label": "꺾은선그래프"},
    "pie": {"label": "원형그래프"},
}


def insert_chart_from_source(report: HwpReport, source_paths: list[str], chart_type: str) -> dict:
    """원본자료 중 첫 엑셀 파일을 읽어 커서 위치에 그래프로 삽입한다.

    chart_type은 CHART_TYPES의 key 중 하나여야 한다. 표(table_tool)와 달리
    "문서에서 선택한 텍스트로 만들기"는 지원하지 않는다 - 자유 텍스트에는
    그래프로 그릴 수치 구조(라벨+값)가 없어 표처럼 한 줄=한 행으로 기계적
    변환할 수 없기 때문이다. 원본자료에 엑셀이 하나도 없으면 문서를
    건드리지 않고 inserted=False를 반환한다(table_tool의 같은 관례)."""
    if chart_type not in CHART_TYPES:
        raise ValueError(f"알 수 없는 그래프 종류: {chart_type}")

    excel_path = _first_excel_source(source_paths)
    if excel_path is None:
        return {"inserted": False, "reason": "원본자료 중 엑셀 파일이 없습니다"}

    build_chart(report, excel_path, chart_type)
    return {"inserted": True, "chart_type": chart_type, "source_file": excel_path}


def build_chart(report: HwpReport, data, chart_type: str) -> None:
    """data(엑셀 경로 또는 DataFrame)의 첫 열을 라벨(x축/파이 조각 이름)로,
    나머지 숫자 열을 값으로 삼아 그래프 이미지를 그려 커서 위치에 삽입한다.

    원형그래프는 조각이 하나의 값 계열이어야 뜻이 통하므로(여러 계열을
    파이 하나에 같이 그리면 비율이 뒤섞여 의미가 없어짐), 숫자 열이
    여러 개여도 첫 번째 숫자 열만 쓴다 - 막대/꺾은선은 여러 계열을 그대로
    다 그린다(df.plot이 알아서 계열별 막대/선을 나눠 그림).

    표(table_tool.build_table)와 달리 선택 텍스트 방지용 Cancel() 가드가
    필요 없다 - insert_picture()는 표 생성(create_table)과 달리 선택된
    콘텐츠가 있어도 크래시하는 사례가 report.hwp.insert_picture() 자체
    문서(및 pyhwpx 소스)에 보고돼있지 않고, 그림은 표처럼 선택 영역을
    대체하는 게 아니라 캐럿 위치에 글자처럼(treat_as_char=True) 끼워
    넣는 것이라 표의 위험(TableCreate HAction 실패)과 애초에 다른
    코드경로다."""
    if isinstance(data, str):
        df = pd.read_excel(data) if data.lower().endswith((".xls", ".xlsx")) else pd.read_csv(data)
    else:
        df = data

    label_col = df.columns[0]
    value_cols = [c for c in df.columns[1:] if pd.api.types.is_numeric_dtype(df[c])]
    if not value_cols:
        raise ValueError("그래프로 그릴 숫자 열이 없습니다 (첫 열을 뺀 나머지가 전부 숫자가 아님)")

    fig, ax = plt.subplots(figsize=(6, 4))
    if chart_type == "bar":
        df.plot(kind="bar", x=label_col, y=value_cols, ax=ax)
    elif chart_type == "line":
        df.plot(kind="line", x=label_col, y=value_cols, ax=ax, marker="o")
    elif chart_type == "pie":
        ax.pie(df[value_cols[0]], labels=df[label_col], autopct="%1.1f%%")
        ax.set_ylabel("")
    fig.tight_layout()

    fd, png_path = tempfile.mkstemp(suffix=".png", prefix="chart_")
    os.close(fd)
    try:
        fig.savefig(png_path)
        plt.close(fig)
        report.hwp.insert_picture(png_path, treat_as_char=True, sizeoption=0)
    finally:
        if os.path.exists(png_path):
            os.remove(png_path)


def _count_picture_ctrls(report: HwpReport) -> int:
    """문서 안의 그림(gso) 컨트롤 개수를 센다. pyhwpx의
    get_ctrl_by_ctrl_id()는 ctrl.UserDesc(사람이 읽는 설명, 그림은 '그림')를
    ctrl_id 코드('gso')와 직접 비교하는 버그가 있어(직접 프로브로 확인:
    ctrl.CtrlID='gso'인데 ctrl.UserDesc='그림'이라 서로 절대 안 맞음) 항상
    빈 리스트만 반환한다 - 그래서 여기서는 CtrlID를 직접 비교해 순회한다."""
    count = 0
    ctrl = report.hwp.HeadCtrl.Next.Next
    while ctrl:
        if ctrl.CtrlID == "gso":
            count += 1
        ctrl = ctrl.Next
    return count


def _selftest_insert_chart_from_source_bar():
    """엑셀 원본자료 하나를 막대그래프로 삽입하면, 문서에 그림 컨트롤이
    실제로 하나 생기는지 확인한다."""
    import openpyxl
    from pyhwpx import Hwp

    test_dir = os.path.join(tempfile.gettempdir(), "_test_그래프원본")
    os.makedirs(test_dir, exist_ok=True)
    source_path = os.path.join(test_dir, "원본.xlsx")
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Sheet1"
    ws.append(["항목", "금액"])
    ws.append(["인건비", 1000000])
    ws.append(["운영비", 500000])
    wb.save(source_path)

    report_path = os.path.join(tempfile.gettempdir(), "_test_보고서_그래프.hwp")
    setup = Hwp(visible=False, new=True)
    setup.save_as(report_path)  # 빈 문서
    setup.quit()

    report = None
    try:
        report = HwpReport(report_path)
        before = _count_picture_ctrls(report)
        result = insert_chart_from_source(report, source_paths=[source_path], chart_type="bar")
        assert result["inserted"] is True, result
        after = _count_picture_ctrls(report)
        assert after == before + 1, f"그림 컨트롤이 안 늘어남 (before={before}, after={after})"
        print("insert_chart_from_source(막대그래프) 통과")
    finally:
        if report is not None:
            report.close(save=False)
        import shutil
        shutil.rmtree(test_dir)
        os.remove(report_path)


def _selftest_insert_chart_from_source_line_and_pie():
    """꺾은선/원형 그래프도 같은 방식으로 삽입되는지 확인한다(종류별로
    다른 matplotlib 호출 경로를 타므로 각각 실제로 예외 없이 도는지가
    핵심 - 렌더링 결과 자체의 시각적 정확성은 사람이 눈으로 봐야 함)."""
    import openpyxl
    from pyhwpx import Hwp

    test_dir = os.path.join(tempfile.gettempdir(), "_test_그래프원본2")
    os.makedirs(test_dir, exist_ok=True)
    source_path = os.path.join(test_dir, "원본.xlsx")
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Sheet1"
    ws.append(["분기", "실적", "목표"])
    ws.append(["1분기", 100, 120])
    ws.append(["2분기", 150, 140])
    wb.save(source_path)

    report_path = os.path.join(tempfile.gettempdir(), "_test_보고서_그래프2.hwp")
    setup = Hwp(visible=False, new=True)
    setup.save_as(report_path)
    setup.quit()

    report = None
    try:
        report = HwpReport(report_path)
        for chart_type in ("line", "pie"):
            before = _count_picture_ctrls(report)
            result = insert_chart_from_source(report, source_paths=[source_path], chart_type=chart_type)
            assert result["inserted"] is True, result
            after = _count_picture_ctrls(report)
            assert after == before + 1, f"{chart_type}: 그림 컨트롤이 안 늘어남"
        print("insert_chart_from_source(꺾은선/원형) 통과")
    finally:
        if report is not None:
            report.close(save=False)
        import shutil
        shutil.rmtree(test_dir)
        os.remove(report_path)


def _selftest_insert_chart_from_source_no_excel_source():
    """원본자료에 엑셀 파일이 없으면 예외 없이 inserted=False를 반환하고
    문서를 전혀 건드리지 않는지 확인한다(table_tool의 같은 회귀테스트와
    동일한 목적)."""
    from pyhwpx import Hwp

    report_path = os.path.join(tempfile.gettempdir(), "_test_보고서_그래프없음.hwp")
    setup = Hwp(visible=False, new=True)
    setup.save_as(report_path)
    setup.quit()

    report = None
    try:
        report = HwpReport(report_path)
        result = insert_chart_from_source(report, source_paths=[], chart_type="bar")
        assert result == {"inserted": False, "reason": "원본자료 중 엑셀 파일이 없습니다"}, result
        assert report.get_text() == "", repr(report.get_text())
        print("insert_chart_from_source(엑셀 없음) 통과: 문서 안 건드리고 inserted=False")
    finally:
        if report is not None:
            report.close(save=False)
        os.remove(report_path)


def _selftest_build_chart_pie_uses_only_first_value_column():
    """숫자 열이 여러 개인 데이터로 원형그래프를 그려도 예외 없이 동작하는지
    확인한다(첫 번째 숫자 열만 쓰고 나머지는 무시하는 설계 - 여러 계열을
    파이 하나에 같이 그리면 의미가 없어 의도적으로 하나만 씀)."""
    from pyhwpx import Hwp

    report_path = os.path.join(tempfile.gettempdir(), "_test_보고서_파이여러열.hwp")
    setup = Hwp(visible=False, new=True)
    setup.save_as(report_path)
    setup.quit()

    report = None
    try:
        report = HwpReport(report_path)
        df = pd.DataFrame({"항목": ["A", "B", "C"], "실적": [10, 20, 30], "목표": [15, 15, 15]})
        before = _count_picture_ctrls(report)
        build_chart(report, df, "pie")
        after = _count_picture_ctrls(report)
        assert after == before + 1, "원형그래프 삽입 실패"
        print("build_chart(원형그래프, 숫자열 여러개) 통과")
    finally:
        if report is not None:
            report.close(save=False)
        os.remove(report_path)


if __name__ == "__main__":
    _selftest_insert_chart_from_source_bar()
    _selftest_insert_chart_from_source_line_and_pie()
    _selftest_insert_chart_from_source_no_excel_source()
    _selftest_build_chart_pie_uses_only_first_value_column()
