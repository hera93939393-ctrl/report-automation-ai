"""window_layout.py — 한글 창과 채팅창을 화면 좌우로 자동 배치한다.
한글 창은 win32gui로(남의 프로세스 창이라 외부 제어 필요), 채팅창은
CustomTkinter 자신의 geometry()로 위치를 잡는다."""
import win32api
import win32con
import win32gui


def _calculate_layout(screen_width: int, screen_height: int, hwp_ratio: float = 0.75):
    """화면 크기와 한글 창이 차지할 비율을 받아, 한글 창과 채팅창 각각의
    (x, y, width, height) 사각형을 계산한다. 순수 계산 함수로 분리한 이유:
    win32gui.MoveWindow 호출 자체는 실제 창이 있어야만 테스트 가능하지만,
    좌표 계산 로직은 화면 크기만 있으면 되므로 실제 창 없이도 빠르게 검증 가능.
    """
    hwp_width = round(screen_width * hwp_ratio)
    chat_width = screen_width - hwp_width
    hwp_rect = (0, 0, hwp_width, screen_height)
    chat_rect = (hwp_width, 0, chat_width, screen_height)
    return hwp_rect, chat_rect


def position_windows(hwp_hwnd: int, chat_window, hwp_ratio: float = 0.75) -> None:
    """한글 창(hwp_hwnd)을 화면 왼쪽 hwp_ratio 비율로, 채팅창(chat_window,
    CustomTkinter 인스턴스)을 나머지 오른쪽에 배치한다.

    hwp_ratio=0.75가 기본값(실사용 후 조정 가능하도록 파라미터로 뺌 —
    하드코딩하지 않음).

    hwp_hwnd는 호출 시점에 유효한 살아있는 창 핸들이어야 한다 — HwpReport.
    get_window_handle()로 얻은 HWND는 그 창이 닫히면 무효해지므로, 얻은
    직후 바로 이 함수에 넘겨 쓰고 오래 보관해두지 않아야 한다(Task 3
    코드품질 검토에서 나온 권고 사항 반영).
    """
    screen_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
    screen_height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
    hwp_rect, chat_rect = _calculate_layout(screen_width, screen_height, hwp_ratio)

    hwp_x, hwp_y, hwp_w, hwp_h = hwp_rect
    win32gui.MoveWindow(hwp_hwnd, hwp_x, hwp_y, hwp_w, hwp_h, True)

    chat_x, chat_y, chat_w, chat_h = chat_rect
    chat_window.geometry(f"{chat_w}x{chat_h}+{chat_x}+{chat_y}")


def _selftest_position_windows_calculates_correct_rects():
    """실제 화면 해상도를 기준으로, 한글 창에게 계산해서 넘겨줄 사각형이
    화면 왼쪽 hwp_ratio 비율을 정확히 차지하는지 확인한다 (win32gui.MoveWindow
    호출 자체는 실제 창이 있어야 하므로 별도 self-test에서 확인 — 여기선
    좌표 계산 로직만 순수하게 검증한다)."""
    screen_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
    screen_height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)

    hwp_rect, chat_rect = _calculate_layout(screen_width, screen_height, hwp_ratio=0.75)

    assert hwp_rect == (0, 0, round(screen_width * 0.75), screen_height), hwp_rect
    expected_chat_x = round(screen_width * 0.75)
    assert chat_rect == (expected_chat_x, 0, screen_width - expected_chat_x, screen_height), chat_rect
    print("_calculate_layout 통과:", hwp_rect, chat_rect)


def _selftest_position_windows_moves_real_hwp_window():
    """position_windows()가 실제 한글 창을 진짜로 왼쪽 75%로 이동시키는지
    확인한다 (위 self-test는 좌표 계산만 검증했음 — 이건 실제 win32gui.
    MoveWindow 호출까지 end-to-end로 검증한다)."""
    import os
    import tempfile
    import win32gui
    from hwp_report import HwpReport

    test_path = os.path.join(tempfile.gettempdir(), "_test_창배치.hwp")
    from pyhwpx import Hwp
    setup = Hwp(visible=False, new=True)
    setup.insert_text("창 배치 테스트")
    setup.save_as(test_path)
    setup.quit()

    report = None
    try:
        report = HwpReport(test_path)
        hwnd = report.get_window_handle()
        position_windows(hwnd, chat_window=_FakeChatWindow(), hwp_ratio=0.75)

        rect = win32gui.GetWindowRect(hwnd)  # (left, top, right, bottom)
        actual_width = rect[2] - rect[0]
        screen_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
        expected_width = round(screen_width * 0.75)
        # 창 테두리/그림자 등으로 몇 픽셀 오차가 날 수 있어 approximate하게 확인
        assert abs(actual_width - expected_width) <= 20, (actual_width, expected_width)
        print("position_windows() 통과 (실제 창 이동 확인):", rect)
    finally:
        if report is not None:
            report.close(save=False)
        os.remove(test_path)


class _FakeChatWindow:
    """채팅창 없이 position_windows()를 테스트하기 위한 최소 스텁 —
    geometry() 호출만 기록하고 실제 tkinter 창은 띄우지 않는다."""
    def __init__(self):
        self.last_geometry = None

    def geometry(self, spec: str):
        self.last_geometry = spec


if __name__ == "__main__":
    _selftest_position_windows_calculates_correct_rects()
    _selftest_position_windows_moves_real_hwp_window()
