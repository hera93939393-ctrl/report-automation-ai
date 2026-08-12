# -*- coding: utf-8 -*-
"""자체 테스트용: tkinter 텍스트 창을 만들어 선택->단축키 전송->결과 확인을 자동으로 검증한다."""
import sys
import time
import ctypes
import tkinter as tk
import keyboard

HOTKEY = sys.argv[1] if len(sys.argv) > 1 else "ctrl+shift+g"
SAMPLE = sys.argv[2] if len(sys.argv) > 2 else "이번 회의를 여는 이유는 최근 민원이 늘어서 그거 논의하려고"
WAIT_MS = int(sys.argv[3]) if len(sys.argv) > 3 else 25000


def main():
    root = tk.Tk()
    root.title("SELFTEST")
    root.geometry("500x250+50+50")
    text = tk.Text(root, font=("맑은 고딕", 13))
    text.pack(fill="both", expand=True)
    text.insert("1.0", SAMPLE)
    root.update()
    root.lift()
    root.attributes("-topmost", True)
    root.focus_force()
    text.focus_force()
    text.tag_add("sel", "1.0", "end-1c")
    root.update()

    hwnd = root.winfo_id()
    # winfo_id()는 자식 위젯 핸들일 수 있어 최상위 창 핸들을 구해 강제로 포그라운드로 올림
    top_hwnd = ctypes.windll.user32.GetAncestor(hwnd, 2)  # GA_ROOT
    ctypes.windll.user32.ShowWindow(top_hwnd, 9)
    ctypes.windll.user32.SetForegroundWindow(top_hwnd)

    def trigger():
        # 실제로 포그라운드가 됐는지 재확인 후 한 번 더 강제 시도
        ctypes.windll.user32.SetForegroundWindow(top_hwnd)
        root.focus_force()
        text.focus_force()
        time.sleep(0.2)
        print("BEFORE:", repr(text.get("1.0", "end")), flush=True)
        print("foreground hwnd match:", ctypes.windll.user32.GetForegroundWindow() == top_hwnd, flush=True)
        keyboard.send(HOTKEY)

    root.after(1000, trigger)

    def check_and_close():
        content = text.get("1.0", "end")
        print("AFTER:", repr(content), flush=True)
        root.destroy()

    root.after(WAIT_MS, check_and_close)
    root.mainloop()


if __name__ == "__main__":
    main()
