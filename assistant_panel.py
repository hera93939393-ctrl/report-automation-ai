# -*- coding: utf-8 -*-
"""
도우미 패널 (F1~F8 통합 UI) — 1단계 버전
- 별도 창에서 F1~F8 기능을 버튼/텍스트로 실행하고, 결과를 화면에 보여준다.
- 사용자가 결과를 확인하고 직접 한글 문서에 복사-붙여넣기 하는 방식.
- (다음 단계: 한글 문서에 직접 삽입하는 방식은 매크로 방식 조사 후 시도)
"""
import os
import tkinter as tk
from tkinter import scrolledtext, filedialog, messagebox
import threading

import ai_writer
import stt
import law_search
import excel_extractor
import format_checker
import image_search

IMAGE_INDEX_PATH = "manual_image_index.npz"


class AssistantPanel:
    def __init__(self, root):
        self.root = root
        root.title("보고서 작성 도우미 (F1~F8)")
        root.geometry("560x620")

        # 버튼 영역
        btn_frame = tk.Frame(root)
        btn_frame.pack(fill="x", padx=8, pady=8)

        buttons = [
            ("F1 문장변환", self.run_f1),
            ("F2 회의록요약", self.run_f2),
            ("F3 음성인식", self.run_f3),
            ("F4 이미지검색", self.run_f4),
            ("F5 법령검색", self.run_f5),
            ("F6 엑셀추출", self.run_f6),
            ("F7 검토", self.run_f7),
            ("F8 포스터생성", self.run_f8),
        ]
        for i, (label, handler) in enumerate(buttons):
            b = tk.Button(btn_frame, text=label, width=12, command=handler)
            b.grid(row=i // 4, column=i % 4, padx=3, pady=3)

        # 입력 영역
        tk.Label(root, text="입력 (텍스트를 쓰거나, 파일이 필요한 버튼은 클릭 시 선택창이 뜸)").pack(anchor="w", padx=8)
        self.input_box = tk.Text(root, height=6)
        self.input_box.pack(fill="x", padx=8, pady=4)

        # 결과 영역
        tk.Label(root, text="결과 (복사해서 한글 문서에 붙여넣으세요)").pack(anchor="w", padx=8)
        self.output_box = scrolledtext.ScrolledText(root, height=20, state="disabled")
        self.output_box.pack(fill="both", expand=True, padx=8, pady=4)

        self.status = tk.Label(root, text="준비됨", anchor="w", fg="gray")
        self.status.pack(fill="x", padx=8, pady=4)

    # ---------- 유틸 ----------
    def _get_input(self) -> str:
        return self.input_box.get("1.0", "end").strip()

    def _append_output(self, text: str):
        self.output_box.config(state="normal")
        self.output_box.insert("end", text + "\n\n" + ("-" * 40) + "\n\n")
        self.output_box.see("end")
        self.output_box.config(state="disabled")

    def _set_status(self, text: str):
        self.status.config(text=text)
        self.root.update_idletasks()

    def _run_async(self, work_fn, on_done):
        """모델 실행은 시간이 걸리므로 별도 스레드에서 돌려 화면이 멈추지 않게 한다."""
        def wrapper():
            try:
                result = work_fn()
            except Exception as e:
                result = f"오류 발생: {e}"
            self.root.after(0, lambda: on_done(result))
        self._set_status("처리 중... (시간이 걸릴 수 있어요)")
        threading.Thread(target=wrapper, daemon=True).start()

    # ---------- F1~F8 핸들러 ----------
    def run_f1(self):
        text = self._get_input()
        if not text:
            messagebox.showinfo("안내", "입력창에 자연어 문장을 먼저 써주세요.")
            return
        self._run_async(lambda: ai_writer.to_gaejoshik(text),
                         lambda r: (self._append_output("[F1 결과]\n" + r), self._set_status("완료")))

    def run_f2(self):
        text = self._get_input()
        if not text:
            messagebox.showinfo("안내", "입력창에 회의 녹취(텍스트)를 먼저 써주세요.")
            return
        self._run_async(lambda: ai_writer.summarize_meeting(text),
                         lambda r: (self._append_output("[F2 결과]\n" + r), self._set_status("완료")))

    def run_f3(self):
        path = filedialog.askopenfilename(title="음성 파일 선택", filetypes=[("음성 파일", "*.wav *.mp3 *.m4a")])
        if not path:
            return
        def work():
            return stt.transcribe(path)
        def done(r):
            self.input_box.delete("1.0", "end")
            self.input_box.insert("1.0", r)
            self._append_output("[F3 결과 - 입력창에 넣었어요]\n" + r)
            self._set_status("완료")
        self._run_async(work, done)

    def run_f4(self):
        query = self._get_input()
        if not query:
            messagebox.showinfo("안내", "입력창에 찾고 싶은 내용을 문장으로 써주세요.")
            return
        if not os.path.exists(IMAGE_INDEX_PATH):
            messagebox.showwarning("안내", f"이미지 색인({IMAGE_INDEX_PATH})이 없어요. 먼저 image_search.build_index()로 색인을 만들어야 해요.")
            return
        def work():
            results = image_search.search(query, IMAGE_INDEX_PATH, top_k=5)
            return "\n".join(f"{r['파일명']} (유사도 {r['유사도']:.3f})" for r in results)
        self._run_async(work, lambda r: (self._append_output("[F4 결과]\n" + r), self._set_status("완료")))

    def run_f5(self):
        keyword = self._get_input()
        if not keyword:
            messagebox.showinfo("안내", "입력창에 검색할 짧은 단어를 써주세요 (예: 청소년).")
            return
        def work():
            results = law_search.search_law(keyword, display=5)
            return "\n".join(f"{r['법령명']} ({r['공포일자']}, {r['소관부처']})" for r in results)
        self._run_async(work, lambda r: (self._append_output("[F5 결과]\n" + r), self._set_status("완료")))

    def run_f6(self):
        path = filedialog.askopenfilename(title="엑셀 파일 선택", filetypes=[("엑셀 파일", "*.xlsx")])
        if not path:
            return
        def work():
            values = excel_extractor.extract_label_value_pairs(path)
            draft = excel_extractor.values_to_draft_text(values)
            return ai_writer.to_gaejoshik(draft)
        self._run_async(work, lambda r: (self._append_output("[F6 결과]\n" + r), self._set_status("완료")))

    def run_f7(self):
        text = self._get_input()
        if not text:
            messagebox.showinfo("안내", "입력창에 검토할 문서 내용을 써주세요.")
            return
        def work():
            fmt = format_checker.run_all_checks(text)
            spell = ai_writer.check_spelling_llm(text)
            lines = [f"항목번호 오류: {fmt['numbering_issues']}",
                     f"계산 오류: {fmt['arithmetic_issues']}",
                     f"맞춤법 검토(참고용):\n{spell}"]
            return "\n\n".join(lines)
        self._run_async(work, lambda r: (self._append_output("[F7 결과]\n" + r), self._set_status("완료")))

    def run_f8(self):
        event_info = self._get_input()
        if not event_info:
            messagebox.showinfo("안내", "입력창에 행사 정보를 써주세요.")
            return
        def work():
            import poster_generator
            prompt = poster_generator.to_image_prompt(event_info)
            path = poster_generator.generate_poster_image(prompt, output_path="_포스터결과.png")
            return f"이미지 프롬프트: {prompt}\n저장 위치: {os.path.abspath(path)}"
        self._run_async(work, lambda r: (self._append_output("[F8 결과]\n" + r), self._set_status("완료")))


if __name__ == "__main__":
    root = tk.Tk()
    app = AssistantPanel(root)
    root.mainloop()
