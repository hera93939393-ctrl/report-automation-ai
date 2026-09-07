"""fit_to_page_tool.py — 문서 전체를 1페이지에 맞추기 위해 행간→자간→
글자크기 순서로 아주 조금씩 줄여가며 페이지 수를 확인하는 도구.
verify_tool.py/table_tool.py처럼 도구 하나당 파일 하나 관례를 따른다.

실측 확인(2026-09-07): hwp.PageCount는 SelectAll() 후
hwp.set_linespacing()/hwp.set_font(Spacing=...)/hwp.set_font(Height=...)를
적용하면 바로 갱신되어 읽힌다. **자간 API는 사전 확인 내용과 실제로
달랐다** - hwp.set_para(Spacing=...)라는 API 자체가 pyhwpx에 없다
(set_para()의 실제 파라미터는 Condense 등이고 Spacing이 없음, pyhwpx
소스(core.py set_para 시그니처)로 직접 확인함). 자간은
hwp.set_font(Spacing=값)(-50~50, set_font()의 파라미터)가 맞다. 글자크기는
"축소확대%"인 Size(10~250)가 아니라, 포인트 단위로 직접 지정하는 Height
파라미터를 쓴다 - 현재 값은 hwp.HwpUnitToPoint(hwp.CharShape.Item("Height"))
로 읽는다.

칼리브레이션(42개 짧은 문단으로 만든, 1페이지를 살짝 넘는 문서 기준 직접
측정): 행간 160%→150%(2단계)만으로 1페이지로 돌아옴 - "아주 조금 넘친"
현실적인 경우엔 첫 단계(행간)만으로 대부분 해결된다는 뜻. 반면 훨씬 많이
넘치는 문서(문장을 3배로 늘린 문단 40개, 3페이지)에서는 행간을 하한(130%)
까지 줄여도, 자간을 하한(-20)까지 줄여도 1페이지가 안 되고 글자크기까지
줄여야 하는 경우도 실측으로 확인했다 - 그래서 세 단계 전부가 실제로
필요할 수 있다."""

_LINESPACING_START = 160
_LINESPACING_FLOOR = 130
_LINESPACING_STEP = 5

_SPACING_START = 0
_SPACING_FLOOR = -20
_SPACING_STEP = 3

_FONT_SIZE_FLOOR = 8.0
_FONT_SIZE_STEP = 0.5


def fit_to_one_page(
    report,
    linespacing_floor: int = _LINESPACING_FLOOR,
    spacing_floor: int = _SPACING_FLOOR,
    font_size_floor: float = _FONT_SIZE_FLOOR,
) -> dict:
    """report(HwpReport)의 문서 전체를 1페이지에 맞춘다. 이미 1페이지면
    아무 것도 안 하고 바로 성공을 반환한다. 행간을 linespacing_floor까지,
    그래도 안 되면 자간을 spacing_floor까지, 그래도 안 되면 글자크기를
    font_size_floor까지 조금씩 줄여가며 매 단계마다 PageCount를 확인한다.
    세 단계를 전부 시도해도 1페이지가 안 되면 실패로 알린다(문서를 건드린
    상태 그대로 둔다 - 자동저장 안 하므로 Ctrl+Z로 되돌릴 수 있음).
    linespacing_floor/spacing_floor/font_size_floor 파라미터는 기본값을
    쓰는 일반 호출에서는 그대로 두면 되고, self-test에서 특정 단계만
    강제로 시험하기 위해 좁혀서 넘길 수 있게 열어뒀다."""
    hwp = report.hwp
    if hwp.PageCount <= 1:
        return {"fitted": True, "method": "already_one_page", "page_count": hwp.PageCount}

    linespacing = _LINESPACING_START
    while linespacing > linespacing_floor and hwp.PageCount > 1:
        linespacing -= _LINESPACING_STEP
        hwp.SelectAll()
        hwp.set_linespacing(linespacing, method="Percent")
        if hwp.PageCount <= 1:
            hwp.Cancel()
            return {"fitted": True, "method": "linespacing", "page_count": hwp.PageCount}

    spacing = _SPACING_START
    while spacing > spacing_floor and hwp.PageCount > 1:
        spacing -= _SPACING_STEP
        hwp.SelectAll()
        hwp.set_font(Spacing=spacing)
        if hwp.PageCount <= 1:
            hwp.Cancel()
            return {"fitted": True, "method": "spacing", "page_count": hwp.PageCount}

    height = hwp.HwpUnitToPoint(hwp.CharShape.Item("Height"))
    while height > font_size_floor and hwp.PageCount > 1:
        height -= _FONT_SIZE_STEP
        hwp.SelectAll()
        hwp.set_font(Height=height)
        if hwp.PageCount <= 1:
            hwp.Cancel()
            return {"fitted": True, "method": "font_size", "page_count": hwp.PageCount}

    hwp.Cancel()
    return {"fitted": False, "method": None, "page_count": hwp.PageCount}


def _selftest_fit_to_one_page_already_one_page_is_noop():
    """이미 1페이지인 문서는 아무 것도 바꾸지 않고 바로 성공해야 한다."""
    from pyhwpx import Hwp
    from hwp_report import HwpReport
    import os, tempfile, time

    path = os.path.join(tempfile.gettempdir(), "_test_맞춤_이미1페이지.hwp")
    setup = Hwp(visible=False, new=True)
    setup.insert_text("한 페이지 문서입니다.")
    setup.save_as(path)
    setup.quit()
    time.sleep(2)

    report = None
    try:
        report = HwpReport(path)
        result = fit_to_one_page(report)
        assert result == {"fitted": True, "method": "already_one_page", "page_count": 1}, result
        print("fit_to_one_page(이미 1페이지) 통과:", result)
    finally:
        if report is not None:
            report.close(save=False)
        os.remove(path)


def _selftest_fit_to_one_page_default_floors_uses_linespacing():
    """1페이지를 살짝 넘는 문서(42개 짧은 문단, 직접 측정으로 확인된
    경계값)는 기본 floor 설정에서 행간 조정만으로 1페이지가 돼야 한다."""
    from pyhwpx import Hwp
    from hwp_report import HwpReport
    import os, tempfile, time

    path = os.path.join(tempfile.gettempdir(), "_test_맞춤_행간.hwp")
    setup = Hwp(visible=False, new=True)
    for i in range(42):
        setup.insert_text(f"{i}번째 문단입니다. 이것은 페이지 채우기용 테스트 문장입니다.")
        setup.HAction.Run("BreakPara")
    setup.save_as(path)
    setup.quit()
    time.sleep(2)

    report = None
    try:
        report = HwpReport(path)
        assert report.hwp.PageCount == 2, report.hwp.PageCount
        result = fit_to_one_page(report)
        assert result["fitted"] is True, result
        assert result["method"] == "linespacing", result
        assert result["page_count"] == 1, result
        print("fit_to_one_page(행간 조정만으로 해결) 통과:", result)
    finally:
        if report is not None:
            report.close(save=False)
        os.remove(path)


def _selftest_fit_to_one_page_falls_through_to_font_size():
    """행간과 자간 조정 여지를 강제로 좁혀두면(linespacing_floor=158로 한
    단계만, spacing_floor=3으로 0단계) 글자크기 조정 단계까지 내려가
    성공해야 한다 - 세 단계 사이 폴백(fallthrough) 로직 자체를 검증한다."""
    from pyhwpx import Hwp
    from hwp_report import HwpReport
    import os, tempfile, time

    path = os.path.join(tempfile.gettempdir(), "_test_맞춤_글자크기.hwp")
    setup = Hwp(visible=False, new=True)
    for i in range(42):
        setup.insert_text(f"{i}번째 문단입니다. 이것은 페이지 채우기용 테스트 문장입니다.")
        setup.HAction.Run("BreakPara")
    setup.save_as(path)
    setup.quit()
    time.sleep(2)

    report = None
    try:
        report = HwpReport(path)
        result = fit_to_one_page(report, linespacing_floor=158, spacing_floor=3)
        assert result["fitted"] is True, result
        assert result["method"] == "font_size", result
        assert result["page_count"] == 1, result
        print("fit_to_one_page(글자크기까지 폴백) 통과:", result)
    finally:
        if report is not None:
            report.close(save=False)
        os.remove(path)


def _selftest_fit_to_one_page_all_floors_too_tight_fails_honestly():
    """세 단계 모두 여지를 0으로 막아두면(각 floor를 시작값과 같게) 문서를
    건드리지 못하고 fitted=False를 정직하게 반환해야 한다(무한루프 없이
    종료되는지도 함께 확인)."""
    from pyhwpx import Hwp
    from hwp_report import HwpReport
    import os, tempfile, time

    path = os.path.join(tempfile.gettempdir(), "_test_맞춤_실패.hwp")
    setup = Hwp(visible=False, new=True)
    for i in range(42):
        setup.insert_text(f"{i}번째 문단입니다. 이것은 페이지 채우기용 테스트 문장입니다.")
        setup.HAction.Run("BreakPara")
    setup.save_as(path)
    setup.quit()
    time.sleep(2)

    report = None
    try:
        report = HwpReport(path)
        result = fit_to_one_page(report, linespacing_floor=160, spacing_floor=0, font_size_floor=10.0)
        assert result == {"fitted": False, "method": None, "page_count": 2}, result
        print("fit_to_one_page(여지 없음, 정직한 실패) 통과:", result)
    finally:
        if report is not None:
            report.close(save=False)
        os.remove(path)


if __name__ == "__main__":
    _selftest_fit_to_one_page_already_one_page_is_noop()
    _selftest_fit_to_one_page_default_floors_uses_linespacing()
    _selftest_fit_to_one_page_falls_through_to_font_size()
    _selftest_fit_to_one_page_all_floors_too_tight_fails_honestly()
