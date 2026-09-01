"""numbering_tool.py — 번호서식(문단번호 스타일) 프리픽스를 커서 위치에 삽입하는 도구.

PRD 14-2: 한/글의 "진짜" 개요번호(자동 채번) 객체가 아니라, 텍스트 프리픽스를
직접 삽입하는 방식이다 — 순서가 바뀌어도 자동으로 재채번되지 않는다는 한계가
있음을 명시적으로 문서화해둔다(다음 라운드 재검토 대상)."""

from hwp_report import HwpReport

# PRD 14-2: 미리보기 팝업(chat_assistant.py)에 쓰이는 값과 반드시 일치해야 한다.
NUMBERING_STYLES = {
    "arabic_dot": {"label": "1. 아라비아숫자", "prefix": "1. "},
    "arabic_paren": {"label": "1) 괄호숫자", "prefix": "1) "},
    "double_paren": {"label": "(1) 이중괄호", "prefix": "(1) "},
    "circled": {"label": "① 원문자", "prefix": "① "},
}


def insert_numbering_prefix(report: HwpReport, style: str) -> dict:
    """style에 해당하는 프리픽스 문자열을 커서 위치에 삽입한다.
    선택된 텍스트가 있으면(F12 1단계 polish_to_formal_style과 같은 관례로)
    그 앞에 프리픽스만 덧붙이는 것이 아니라, 프리픽스를 그대로 삽입 지점에
    끼워넣는다 — insert_text()의 기본 동작(선택 영역 앞에 삽입되지 않고
    선택 영역을 대체함)과 혼동하지 않도록, 이 함수는 항상 "선택 없음" 상태를
    전제로 한다(미리보기 팝업 플로우에서 선택 여부를 별도로 다루지 않음,
    PRD 14-2 범위 밖)."""
    if style not in NUMBERING_STYLES:
        raise ValueError(f"알 수 없는 번호서식: {style}")

    prefix = NUMBERING_STYLES[style]["prefix"]
    applied = report.hwp.insert_text(prefix)
    return {"inserted": bool(applied), "style": style}


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


if __name__ == "__main__":
    _selftest_insert_numbering_prefix()
    _selftest_all_styles_have_correct_prefix()
