# -*- coding: utf-8 -*-
"""
계획(안) 자동 작성 도구  (입력 창 버전)
- 빈칸에 내용만 입력하고 [한글파일 만들기] 를 누르면
  서식/도형/표가 그대로 유지된 계획(안).hwp 가 생성됩니다.
"""
import os
import sys
import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from fill_core import fill_report, FIELDS, TEMPLATE

# 여러 줄로 입력받을 항목 (나머지는 한 줄)
MULTILINE = {"목적", "내용1", "내용2", "내용3", "내용4", "일정설명"}


class App:
    def __init__(self, root):
        self.root = root
        root.title("계획(안) 자동 작성 도구")
        root.geometry("760x720")

        head = tk.Label(root, text="빈칸에 내용만 채우고 아래 버튼을 누르세요.\n서식·도형·표 모양은 그대로 유지됩니다.",
                        font=("맑은 고딕", 11), fg="#333", justify="left")
        head.pack(anchor="w", padx=16, pady=(14, 6))

        # 스크롤 가능한 입력 영역
        canvas = tk.Canvas(root, highlightthickness=0)
        scroll = ttk.Scrollbar(root, orient="vertical", command=canvas.yview)
        frame = tk.Frame(canvas)
        frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=frame, anchor="nw")
        canvas.configure(yscrollcommand=scroll.set)
        canvas.pack(side="left", fill="both", expand=True, padx=(16, 0))
        scroll.pack(side="right", fill="y")

        self.widgets = {}
        for key, label in FIELDS:
            row = tk.Frame(frame)
            row.pack(fill="x", pady=6, padx=(0, 16))
            tk.Label(row, text=label, font=("맑은 고딕", 10, "bold"),
                     anchor="w", width=40, justify="left", wraplength=300).pack(anchor="w")
            if key in MULTILINE:
                w = tk.Text(row, height=2, font=("맑은 고딕", 10), wrap="word")
            else:
                w = tk.Entry(row, font=("맑은 고딕", 10))
            w.pack(fill="x")
            self.widgets[key] = w

        btn = tk.Button(root, text="한글파일 만들기", font=("맑은 고딕", 12, "bold"),
                        bg="#2b6cb0", fg="white", height=2, command=self.generate)
        btn.pack(fill="x", padx=16, pady=12)

    def _get(self, key):
        w = self.widgets[key]
        if isinstance(w, tk.Text):
            return w.get("1.0", "end").strip()
        return w.get().strip()

    def generate(self):
        if not os.path.exists(TEMPLATE):
            messagebox.showerror("오류", f"서식 뼈대가 없습니다:\n{TEMPLATE}\n\n먼저 make_template.py 를 실행하세요.")
            return
        values = {key: self._get(key) for key, _ in FIELDS}
        title = values.get("제목") or "계획안"
        default_name = f"{title} 계획(안).hwp"
        out = filedialog.asksaveasfilename(
            title="저장 위치 선택",
            initialfile=default_name,
            defaultextension=".hwp",
            filetypes=[("한글 파일", "*.hwp")],
        )
        if not out:
            return
        try:
            fill_report(values, out)
        except Exception as e:
            messagebox.showerror("생성 실패", str(e))
            return
        if messagebox.askyesno("완료", f"만들었습니다:\n{out}\n\n지금 한글로 열어볼까요?"):
            os.startfile(out)


def main():
    root = tk.Tk()
    app = App(root)
    if "--smoke" in sys.argv:
        root.update_idletasks()
        root.update()
        print("SMOKE OK: window built, fields =", list(app.widgets.keys()))
        root.destroy()
        return
    root.mainloop()


if __name__ == "__main__":
    main()
