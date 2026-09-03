"""numbering_tool.py — 번호서식(문단번호 스타일) 프리픽스를 커서 위치에 삽입하는 도구.

PRD 14-2: 한/글의 "진짜" 개요번호(자동 채번) 객체가 아니라, 텍스트 프리픽스를
직접 삽입하는 방식이다 — 순서가 바뀌어도 자동으로 재채번되지 않는다는 한계가
있음을 명시적으로 문서화해둔다(다음 라운드 재검토 대상).

(2026-09-03, 실사용 피드백 반영) 사용자가 실제로 "번호 매겨줘"를 써보고 준
피드백: 원한 건 커서에 "1." 하나만 넣는 게 아니라, 선택한 목록 전체를
번호(1·2·3·4) 붙은 표로 만드는 것이었다. 그래서 이제 이 모듈은 두 가지를
다 제공한다 — insert_numbering_prefix()(선택 없을 때, 커서에 프리픽스 삽입,
기존 그대로)와 insert_numbered_table_from_selection()(선택 있을 때, 번호/내용
2열 표로 변환, 신규). 어느 걸 부를지는 chat_assistant.py가 선택 여부를 보고
결정한다."""
import pandas as pd

from hwp_report import HwpReport
from table_tool import TABLE_STYLES, build_table

# PRD 14-2: 미리보기 팝업(chat_assistant.py)에 쓰이는 값과 반드시 일치해야 한다.
NUMBERING_STYLES = {
    "arabic_dot": {"label": "1. 아라비아숫자", "prefix": "1. "},
    "arabic_paren": {"label": "1) 괄호숫자", "prefix": "1) "},
    "double_paren": {"label": "(1) 이중괄호", "prefix": "(1) "},
    "circled": {"label": "① 원문자", "prefix": "① "},
}


def insert_numbering_prefix(report: HwpReport, style: str) -> dict:
    """style에 해당하는 프리픽스 문자열을 커서 위치에 삽입한다.

    insert_text()는 문서에 선택된 텍스트가 있으면 그 선택 영역을 통째로 지우고
    대체해버리는 부작용이 있다(F12 1단계 polish_tool.py에서 이미 검증된 pyhwpx
    동작 — 직접 재현: "가나다라마바사"에서 "다라마"를 선택한 채 이 함수를
    호출하면 "다라마"가 사라지고 "가나1. 바사"가 됨). 미리보기 팝업 플로우
    (PRD 14-2, chat_assistant.py)는 사용자가 문서에서 뭔가 드래그해 선택해둔
    채로 이 도구를 실행하는 상황을 막지 않으므로, 이 함수가 스스로 방어해야
    한다.

    그래서 삽입 직전에 항상 report.hwp.Cancel()로 선택을 해제한다 — 선택이
    없었으면 no-op이고, 선택이 있었으면 그 콘텐츠는 그대로 보존된 채 선택만
    풀린다. 직접 테스트로 확인한 바, Cancel() 이후 커서는 선택 영역의
    "끝점"에 남는다(시작점이 아님) — 예: "다라마"를 선택하고 Cancel() 하면
    커서는 "마"와 "바" 사이에 위치한다. 따라서 프리픽스는 선택했던 콘텐츠
    바로 뒤에 삽입된다.
    """
    if style not in NUMBERING_STYLES:
        raise ValueError(f"알 수 없는 번호서식: {style}")

    prefix = NUMBERING_STYLES[style]["prefix"]
    report.hwp.Cancel()
    applied = report.hwp.insert_text(prefix)
    return {"inserted": bool(applied), "style": style}


def insert_numbered_table_from_selection(report: HwpReport, style: str) -> dict:
    """문서에서 현재 선택된 텍스트(여러 문단 가능)를 "번호"(1,2,3,4...)와
    "내용" 2열 표로 변환해 선택 영역 바로 뒤에 삽입한다.

    style은 이 모듈의 NUMBERING_STYLES가 아니라 table_tool.TABLE_STYLES의
    key(표 배경색 스타일: default/gray_header/blue_header)여야 한다 — 선택된
    텍스트가 있을 때는 결과물이 "텍스트 프리픽스"가 아니라 "표"이므로,
    미리보기 팝업도 표 스타일 3종을 보여주는 게 자연스럽다(chat_assistant.py의
    _show_numbering_style_picker가 선택 여부를 보고 어느 스타일 집합/픽커를
    보여줄지 결정한다).

    table_tool.py의 insert_table_from_source()가 이미 검증한 것과 완전히
    같은 안전장치(build_table 내부의 Cancel())를 재사용한다 — 별도로 다시
    구현하지 않고 그대로 import해서 쓴다(같은 라운드에 두 도구가 사실상
    같은 "선택→표" 동작을 하게 됐으므로, 중복 구현 대신 재사용이 맞다고
    판단함).
    """
    if style not in TABLE_STYLES:
        raise ValueError(f"알 수 없는 표 스타일: {style}")

    selected = report.hwp.get_selected_text(keep_select=True)
    lines = [line.strip() for line in selected.replace("\r\n", "\n").split("\n") if line.strip()]
    if not lines:
        report.hwp.Cancel()
        return {"inserted": False, "reason": "선택한 텍스트가 비어있습니다"}

    df = pd.DataFrame({"번호": list(range(1, len(lines) + 1)), "내용": lines})
    build_table(report, df, style)
    return {"inserted": True, "style": style, "source": "selection", "row_count": len(lines)}


def _selftest_insert_numbering_prefix():
    """선택된 번호서식 프리픽스가 커서 위치에 삽입되는지 확인한다."""
    import os, tempfile
    from pyhwpx import Hwp
    from hwp_report import HwpReport

    report_path = os.path.join(tempfile.gettempdir(), "_test_보고서_번호.hwp")
    setup = Hwp(visible=False, new=True)
    setup.save_as(report_path)
    setup.quit()

    report = None
    try:
        report = HwpReport(report_path)
        result = insert_numbering_prefix(report, style="arabic_dot")
        assert result["inserted"] is True, result

        final_text = report.get_text()
        assert final_text.startswith("1. "), repr(final_text)
        print("insert_numbering_prefix(arabic_dot) 통과:", repr(final_text))
    finally:
        if report is not None:
            report.close(save=False)
        os.remove(report_path)


def _selftest_all_styles_have_correct_prefix():
    """NUMBERING_STYLES의 4종 모두 실제로 올바른 프리픽스로 삽입되는지 확인."""
    import os, tempfile
    from pyhwpx import Hwp
    from hwp_report import HwpReport

    expected = {
        "arabic_dot": "1. ",
        "arabic_paren": "1) ",
        "double_paren": "(1) ",
        "circled": "① ",
    }
    for style, prefix in expected.items():
        report_path = os.path.join(tempfile.gettempdir(), f"_test_번호_{style}.hwp")
        setup = Hwp(visible=False, new=True)
        setup.save_as(report_path)
        setup.quit()

        report = None
        try:
            report = HwpReport(report_path)
            result = insert_numbering_prefix(report, style=style)
            assert result["inserted"] is True, result
            final_text = report.get_text()
            assert final_text.startswith(prefix), (style, repr(final_text))
        finally:
            if report is not None:
                report.close(save=False)
            os.remove(report_path)
    print("insert_numbering_prefix 4종 스타일 전부 통과")


def _selftest_insert_numbering_prefix_preserves_selection():
    """회귀 테스트(스펙 검토에서 발견된 버그): 문서에 선택된 텍스트가 있는 채로
    insert_numbering_prefix()를 호출해도 그 선택된 콘텐츠가 사라지면 안 된다.

    insert_text()는 선택 영역이 있으면 그 선택 영역을 통째로 지우고 대체해버리는
    부작용이 있다(F12 1단계 polish_tool.py에서 이미 검증된 pyhwpx 동작). 이 함수의
    이전 구현은 이 부작용을 막는 가드가 전혀 없어서, 사용자가 문서에서 뭔가
    드래그해 선택해둔 채로 "번호 매겨줘"를 실행하면 선택된 콘텐츠가 조용히
    사라지는 실사용 위험이 있었다(직접 재현: "가나다라마바사"에서 "다라마"를
    선택한 채 호출하면 "가나1) 바사"가 되어 "다라마"가 사라짐).
    """
    import os, tempfile
    from pyhwpx import Hwp
    from hwp_report import HwpReport

    report_path = os.path.join(tempfile.gettempdir(), "_test_보고서_번호_선택유지.hwp")
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

        result = insert_numbering_prefix(report, style="arabic_dot")
        assert result["inserted"] is True, result

        final_text = report.get_text()
        assert "다라마" in final_text, (
            "선택된 콘텐츠('다라마')가 사라짐 — insert_text()가 선택 영역을 "
            f"대체해버리는 버그가 재현됨: {final_text!r}"
        )
        assert "1. " in final_text, final_text
        print("insert_numbering_prefix(선택 있음, 콘텐츠 보존) 통과:", repr(final_text))
    finally:
        if report is not None:
            report.close(save=False)
        os.remove(report_path)


def _selftest_insert_numbered_table_from_selection():
    """(2026-09-03, 실사용 피드백 반영 신규 테스트) 여러 문단을 선택한 채로
    insert_numbered_table_from_selection()을 호출하면, 각 줄이 "번호"(1,2,3,4)와
    "내용" 2열 표로 들어가는지 확인한다. table_tool.py의 동명 테스트와 같은
    이유로 find() 대신 select_text()로 선택을 재현한다(find()는 문단 경계를
    넘는 여러 줄짜리 검색어를 못 찾음 — 직접 재현 확인됨)."""
    import os, tempfile
    from pyhwpx import Hwp
    from hwp_report import HwpReport

    report_path = os.path.join(tempfile.gettempdir(), "_test_번호표_여러줄선택.hwp")
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
        ok = report.hwp.select_text(1, 0, 4, -1)
        assert ok, "문단 1~4 선택 실패"
        assert report.hwp.SelectionMode != 0, "선택이 안 된 상태로 테스트가 시작됨"

        result = insert_numbered_table_from_selection(report, style="gray_header")
        assert result["inserted"] is True, result
        assert result.get("row_count") == 4, result

        final_text = report.get_text()
        assert "번호" in final_text, final_text
        for n in ["1", "2", "3", "4"]:
            assert n in final_text, (n, final_text)
        for line in ["이슈사항 논의", "이슈에 따른 결론", "8월달 논의사항 정리", "처장님 이사말씀"]:
            assert line in final_text, (line, final_text)
        assert "앞 문단입니다" in final_text, "앞 문단이 사라짐: " + repr(final_text)
        assert "뒤 문단입니다" in final_text, "뒤 문단이 사라짐: " + repr(final_text)
        print("insert_numbered_table_from_selection(여러 줄 선택 → 번호붙은 표) 통과:", repr(final_text))
    finally:
        if report is not None:
            report.close(save=False)
        os.remove(report_path)


def _selftest_insert_numbered_table_wrong_style_raises():
    """insert_numbered_table_from_selection()에 NUMBERING_STYLES의 key(예:
    "arabic_dot")를 잘못 넘기면 ValueError가 나야 한다 — 이 함수는
    table_tool.TABLE_STYLES의 key를 받는다는 걸 실수로 헷갈리기 쉬운 지점이라
    (두 스타일 딕셔너리 모두 이 모듈/table_tool.py에서 "style"이라는 같은
    파라미터명을 쓰지만 값 집합이 다름), 오용 시 조용히 이상하게 동작하는 대신
    바로 드러나는지 확인한다."""
    import os, tempfile
    from pyhwpx import Hwp
    from hwp_report import HwpReport

    report_path = os.path.join(tempfile.gettempdir(), "_test_번호표_잘못된스타일.hwp")
    setup = Hwp(visible=False, new=True)
    setup.insert_text("가나다"); setup.BreakPara()
    setup.insert_text("라마바")
    setup.save_as(report_path)
    setup.quit()

    report = None
    try:
        report = HwpReport(report_path)
        report.hwp.select_text(0, 0, 1, -1)
        try:
            insert_numbered_table_from_selection(report, style="arabic_dot")
            raise AssertionError("ValueError가 나야 하는데 나지 않음")
        except ValueError:
            pass
        print("insert_numbered_table_from_selection(잘못된 스타일 거부) 통과")
    finally:
        if report is not None:
            report.close(save=False)
        os.remove(report_path)


if __name__ == "__main__":
    _selftest_insert_numbering_prefix()
    _selftest_all_styles_have_correct_prefix()
    _selftest_insert_numbering_prefix_preserves_selection()
    _selftest_insert_numbered_table_from_selection()
    _selftest_insert_numbered_table_wrong_style_raises()
