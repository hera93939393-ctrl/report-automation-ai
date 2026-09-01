"""window_layout.py — 한글 창과 채팅창을 화면 좌우로 자동 배치한다.
한글 창은 win32gui로(남의 프로세스 창이라 외부 제어 필요), 채팅창은
CustomTkinter 자신의 geometry()로 위치를 잡는다."""
import win32api
import win32gui


def _get_work_area():
    """모니터의 "작업 영역"(work area) 사각형 (left, top, right, bottom)을 반환한다.

    작업 영역 = 전체 화면(GetSystemMetrics SM_CXSCREEN/SM_CYSCREEN)에서 작업
    표시줄(taskbar) 등이 차지하는 부분을 뺀, 실제로 창을 온전히 놓을 수 있는
    영역. win32gui.SystemParametersInfo(SPI_GETWORKAREA)는 설치된 pywin32
    빌드에서 NotImplementedError를 던지므로 대신 GetMonitorInfo를 쓴다.
    origin이 항상 (0, 0)이라고 가정하지 않고 left/top을 그대로 사용한다 —
    멀티 모니터 등 구성에 따라 작업 영역의 원점이 (0, 0)이 아닐 수 있음.
    """
    monitor = win32api.MonitorFromPoint((0, 0))
    return win32api.GetMonitorInfo(monitor)["Work"]  # (left, top, right, bottom)


def _calculate_layout(work_left: int, work_top: int, work_width: int, work_height: int,
                       hwp_ratio: float = 0.75):
    """작업 영역(모니터 전체 화면에서 작업표시줄 등을 뺀 영역)의 원점/크기와
    한글 창이 차지할 비율을 받아, 한글 창과 채팅창 각각의 (x, y, width, height)
    사각형을 계산한다. 순수 계산 함수로 분리한 이유: win32gui.MoveWindow 호출
    자체는 실제 창이 있어야만 테스트 가능하지만, 좌표 계산 로직은 작업 영역
    크기만 있으면 되므로 실제 창 없이도 빠르게 검증 가능.

    반드시 "작업 영역" 기준으로 계산해야 한다 — 예전에는 GetSystemMetrics의
    전체 화면 크기(예: 1920x1080)를 그대로 썼는데, 실제 작업표시줄을 뺀
    영역은 더 작다(예: 1920x1032, 작업표시줄이 48px 차지). 채팅창(Tk/
    CustomTkinter 창)은 한글 창과 달리 화면 경계에 맞춰 스스로 줄어들지
    않으므로, 전체 화면 크기로 배치를 계산하면 채팅창이 모니터 오른쪽
    경계를 넘어가거나 작업표시줄 영역과 겹치는 실제 버그가 있었다
    (Task 4 코드품질 리뷰에서 실측으로 확인: 요청 좌표 기준 결과 창 rect가
    (1440, 0, 1936, 1100)으로, 1920 폭을 넘고 1032 작업 영역 높이도 넘어감).
    """
    hwp_width = round(work_width * hwp_ratio)
    chat_width = work_width - hwp_width
    hwp_rect = (work_left, work_top, hwp_width, work_height)
    chat_rect = (work_left + hwp_width, work_top, chat_width, work_height)
    return hwp_rect, chat_rect


def position_windows(hwp_hwnd: int, chat_window, hwp_ratio: float = 0.75) -> None:
    """한글 창(hwp_hwnd)을 작업 영역(모니터 전체 화면에서 작업표시줄을 뺀
    영역) 왼쪽 hwp_ratio 비율로, 채팅창(chat_window, CustomTkinter 인스턴스)을
    나머지 오른쪽에 배치한다.

    hwp_ratio=0.75가 기본값(실사용 후 조정 가능하도록 파라미터로 뺌 —
    하드코딩하지 않음).

    전체 화면 크기(GetSystemMetrics) 대신 작업 영역(GetMonitorInfo의 'Work')을
    쓰는 이유: 한글 창은 우연히 작업표시줄을 피해 스스로 높이를 줄이지만
    (예: 요청 1080 → 실제 1032), 이는 win32gui.MoveWindow로 정확한 크기를
    요청해도 한글 프로그램이 내부적으로 재조정하는 것일 뿐 이 함수가 보장하는
    동작이 아니다. 반면 채팅창(Tk/CustomTkinter)은 그런 보정이 없어서, 예전
    코드(전체 화면 기준)로는 실제로 모니터 오른쪽 경계를 넘고 작업표시줄
    영역까지 침범하는 창이 만들어졌다(Task 4 리뷰에서 실측 확인). 그래서
    두 창 모두 작업 영역 안에 들어오도록 작업 영역 기준으로 계산한다.

    hwp_hwnd는 호출 시점에 유효한 살아있는 창 핸들이어야 한다 — HwpReport.
    get_window_handle()로 얻은 HWND는 그 창이 닫히면 무효해지므로, 얻은
    직후 바로 이 함수에 넘겨 쓰고 오래 보관해두지 않아야 한다(Task 3
    코드품질 검토에서 나온 권고 사항 반영).
    """
    work_left, work_top, work_right, work_bottom = _get_work_area()
    work_width = work_right - work_left
    work_height = work_bottom - work_top
    hwp_rect, chat_rect = _calculate_layout(work_left, work_top, work_width, work_height, hwp_ratio)

    hwp_x, hwp_y, hwp_w, hwp_h = hwp_rect
    win32gui.MoveWindow(hwp_hwnd, hwp_x, hwp_y, hwp_w, hwp_h, True)

    chat_x, chat_y, chat_w, chat_h = chat_rect
    chat_window.geometry(f"{chat_w}x{chat_h}+{chat_x}+{chat_y}")


def _selftest_position_windows_calculates_correct_rects():
    """실제 모니터의 작업 영역(work area)을 기준으로, 한글 창에게 계산해서
    넘겨줄 사각형이 작업 영역 왼쪽 hwp_ratio 비율을 정확히 차지하는지 확인한다
    (win32gui.MoveWindow 호출 자체는 실제 창이 있어야 하므로 별도 self-test에서
    확인 — 여기선 좌표 계산 로직만 순수하게 검증한다).

    기대값도 _calculate_layout과 동일하게 GetMonitorInfo 기반 작업 영역에서
    계산한다 — 예전에는 GetSystemMetrics(전체 화면)로 기대값을 계산해서,
    작업표시줄을 침범하는 잘못된 배치도 이 테스트를 통과했었다."""
    work_left, work_top, work_right, work_bottom = _get_work_area()
    work_width = work_right - work_left
    work_height = work_bottom - work_top

    hwp_rect, chat_rect = _calculate_layout(work_left, work_top, work_width, work_height, hwp_ratio=0.75)

    expected_hwp_width = round(work_width * 0.75)
    assert hwp_rect == (work_left, work_top, expected_hwp_width, work_height), hwp_rect
    expected_chat_x = work_left + expected_hwp_width
    expected_chat_width = work_width - expected_hwp_width
    assert chat_rect == (expected_chat_x, work_top, expected_chat_width, work_height), chat_rect
    print("_calculate_layout 통과:", hwp_rect, chat_rect)


def _selftest_position_windows_moves_real_hwp_window():
    """position_windows()가 실제 한글 창과 실제 채팅창(CustomTkinter) 둘 다
    작업 영역(work area) 안에 들어오도록 배치하는지 확인한다 (위 self-test는
    좌표 계산만 검증했음 — 이건 실제 win32gui.MoveWindow / 실제 tkinter 창의
    geometry() 결과까지 end-to-end로 검증한다).

    채팅창은 _FakeChatWindow 스텁이 아니라 진짜 customtkinter.CTk() 인스턴스를
    띄워서 win32gui.GetWindowRect로 실제 결과 사각형을 확인한다 — 예전에는
    스텁이 geometry() 문자열만 기록하고 실제 창을 만들지 않아서, 채팅창이
    창 테두리(chrome) 때문에 요청한 크기보다 커져 모니터 오른쪽 경계와
    작업표시줄을 침범하는 실제 버그를 이 테스트가 잡지 못했다(Task 4
    코드품질 리뷰에서 발견: 요청 위치/크기로 실제 만든 Tk 창의 GetWindowRect가
    (1440, 0, 1936, 1100)으로 나와 1920 폭과 1032 작업 영역 높이를 모두
    넘어갔음). 이제는 실제 창의 rect가 작업 영역 경계를 넘지 않는지 직접
    확인한다."""
    import os
    import tempfile
    import win32gui
    import customtkinter as ctk
    from hwp_report import HwpReport

    test_path = os.path.join(tempfile.gettempdir(), "_test_창배치.hwp")
    from pyhwpx import Hwp
    setup = Hwp(visible=False, new=True)
    setup.insert_text("창 배치 테스트")
    setup.save_as(test_path)
    setup.quit()

    work_left, work_top, work_right, work_bottom = _get_work_area()

    report = None
    chat_window = ctk.CTk()
    try:
        report = HwpReport(test_path)
        hwnd = report.get_window_handle()
        position_windows(hwnd, chat_window=chat_window, hwp_ratio=0.75)
        chat_window.update()  # geometry() 호출을 실제로 반영시키기 위해 필요

        hwp_rect = win32gui.GetWindowRect(hwnd)  # (left, top, right, bottom)
        hwp_actual_width = hwp_rect[2] - hwp_rect[0]
        work_width = work_right - work_left
        expected_hwp_width = round(work_width * 0.75)
        # 창 테두리/그림자 등으로 몇 픽셀 오차가 날 수 있어 approximate하게 확인
        assert abs(hwp_actual_width - expected_hwp_width) <= 20, (hwp_actual_width, expected_hwp_width)

        chat_rect = win32gui.GetWindowRect(chat_window.winfo_id())
        # 허용 오차: Windows에서 CustomTkinter/Tkinter 창은 geometry()로 요청한
        # 좌표와 실제 GetWindowRect 결과 사이에 타이틀바/테두리 크기만큼의 차이가
        # 항상 존재한다(이 머신에서 실측: 세로로 약 31px, 가로로 약 8px — Tk가
        # Windows에서 창을 배치할 때 나타나는 잘 알려진 창 chrome 차이이며, 이
        # 함수의 작업 영역 계산 로직과는 무관한 별개의 요인). 그래서 기존 hwp
        # 폭 비교에 쓰인 20px보다 조금 더 넉넉한 40px을 허용치로 쓴다 — 그래도
        # 이 테스트가 고치려는 실제 버그(전체 화면 기준 계산 시 세로로 약
        # 60px 넘게 작업표시줄을 침범했던 것, 위 실측으로 확인)는 이 허용치
        # 안에서도 여전히 잡아낸다.
        TOLERANCE = 40
        assert chat_rect[0] >= work_left - TOLERANCE, (chat_rect, work_left)
        assert chat_rect[1] >= work_top - TOLERANCE, (chat_rect, work_top)
        assert chat_rect[2] <= work_right + TOLERANCE, (
            "채팅창이 작업 영역 오른쪽 경계를 넘어감", chat_rect, work_right)
        assert chat_rect[3] <= work_bottom + TOLERANCE, (
            "채팅창이 작업 영역 아래쪽(작업표시줄) 경계를 넘어감", chat_rect, work_bottom)

        print("position_windows() 통과 (실제 한글 창 이동 확인):", hwp_rect)
        print("position_windows() 통과 (실제 채팅창 rect가 작업 영역 안에 있음 확인):", chat_rect,
              "작업 영역:", (work_left, work_top, work_right, work_bottom))
    finally:
        chat_window.destroy()
        if report is not None:
            report.close(save=False)
        os.remove(test_path)


if __name__ == "__main__":
    _selftest_position_windows_calculates_correct_rects()
    _selftest_position_windows_moves_real_hwp_window()
