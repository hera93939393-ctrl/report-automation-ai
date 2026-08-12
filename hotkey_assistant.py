# -*- coding: utf-8 -*-
"""
단축키 도우미 (F1~F8을 한글 등 어떤 프로그램에서도 바로 쓰기)
- 사람이 직접 복사(Ctrl+C)-붙여넣기(Ctrl+V) 하는 동작을 그대로 흉내낸다.
- 한글 COM으로 몰래 접근하는 게 아니라, "지금 선택된 부분"을 복사해서 AI로 처리하고
  원문 아래 줄에 결과를 추가하는 방식. 원문은 그대로 남아있어 비교하며 확인할 수 있다.
- 그래서 한글뿐 아니라 워드, 메모장 등 어디서나 동작한다.
- Alt가 들어간 조합은 창 프로그램의 "메뉴 열기"와 얽힐 수 있어 피한다.

등록된 단축키:
  Ctrl+Shift+G : F1 문장을 공문서 개조식으로 변환
  Ctrl+Shift+M : F2 선택한 녹취/텍스트를 회의록으로 요약
  Ctrl+Shift+L : F5 선택한 키워드로 관련 법령 검색
  Ctrl+Shift+R : F7 선택한(또는 전체) 문서를 검토 (결과는 알림창으로, 문서는 안 건드림)
"""
import time
import winsound
import keyboard
import pyperclip
import tkinter as tk
from tkinter import messagebox

import ai_writer
import law_search
import format_checker


def _get_selected_text() -> str | None:
    """지금 선택된 텍스트를 복사해서 읽어온다. 선택된 게 없으면 None."""
    original_clipboard = pyperclip.paste()

    keyboard.release("shift")
    keyboard.release("ctrl")
    keyboard.release("g")
    keyboard.release("m")
    keyboard.release("l")
    keyboard.release("r")
    time.sleep(0.1)

    keyboard.send("ctrl+c")
    time.sleep(0.3)

    selected = pyperclip.paste()
    if not selected or selected == original_clipboard:
        return None
    return selected


def _insert_below(result_text: str):
    """원문은 그대로 두고, 그 아래 줄에 결과를 붙여넣는다."""
    original_clipboard = pyperclip.paste()
    pyperclip.copy(result_text)
    time.sleep(0.2)

    keyboard.send("end")
    time.sleep(0.1)
    keyboard.send("enter")
    time.sleep(0.1)
    keyboard.send("ctrl+v")

    time.sleep(2.0)
    pyperclip.copy(original_clipboard)


def _show_popup(title: str, message: str):
    """문서를 안 건드리는 결과(검토 등)는 별도 알림창으로 보여준다."""
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    messagebox.showinfo(title, message, parent=root)
    root.destroy()


def _run(label: str, work_fn, insert: bool = True):
    winsound.Beep(1000, 150)
    print(f"[{label}] 선택한 부분을 확인합니다...", flush=True)
    selected = _get_selected_text()
    if selected is None:
        print(f"[{label}] 선택된 텍스트가 없습니다. 먼저 드래그로 선택하세요.", flush=True)
        winsound.Beep(400, 300)
        return

    print(f"[{label}] 처리 중:", selected[:30], "...", flush=True)
    try:
        result = work_fn(selected)
    except Exception as e:
        print(f"[{label} 오류]", e, flush=True)
        winsound.Beep(400, 300)
        return

    print(f"[{label} 결과]", result, flush=True)
    if insert:
        _insert_below(result)
    else:
        _show_popup(label, result)
    winsound.Beep(1500, 150)
    print(f"[{label} 완료]", flush=True)


def handle_f1():
    _run("F1 문장변환", ai_writer.to_gaejoshik, insert=True)


def handle_f2():
    _run("F2 회의록요약", ai_writer.summarize_meeting, insert=True)


def handle_f5():
    def work(keyword):
        results = law_search.search_law(keyword, display=5)
        if not results:
            return "(관련 법령을 찾지 못했습니다)"
        return "\n".join(f"- {r['법령명']} ({r['공포일자']}, {r['소관부처']})" for r in results)
    _run("F5 법령검색", work, insert=True)


def handle_f7():
    def work(text):
        fmt = format_checker.run_all_checks(text)
        spell = ai_writer.check_spelling_llm(text)
        lines = [f"항목번호 오류: {fmt['numbering_issues']}",
                 f"계산 오류: {fmt['arithmetic_issues']}",
                 f"맞춤법 검토(참고용):\n{spell}"]
        return "\n".join(lines)
    _run("F7 검토", work, insert=False)


if __name__ == "__main__":
    print("단축키 도우미 시작.", flush=True)
    print("Ctrl+Shift+G: 문장변환 / Ctrl+Shift+M: 회의록요약 / Ctrl+Shift+L: 법령검색 / Ctrl+Shift+R: 검토", flush=True)
    print("(한글/워드/메모장 등에서 문장을 드래그로 선택한 뒤 눌러보세요)", flush=True)
    print("!! 이 창은 클릭하지 마세요 !! 그냥 켜둔 채로 두고, 작업은 한글 등 다른 창에서 하세요.", flush=True)
    print("종료하려면 이 창의 X 버튼을 누르세요 (Ctrl+C는 누르지 마세요).", flush=True)

    keyboard.add_hotkey("ctrl+shift+g", handle_f1)
    keyboard.add_hotkey("ctrl+shift+m", handle_f2)
    keyboard.add_hotkey("ctrl+shift+l", handle_f5)
    keyboard.add_hotkey("ctrl+shift+r", handle_f7)
    try:
        keyboard.wait()
    except KeyboardInterrupt:
        print("\n종료합니다.", flush=True)
