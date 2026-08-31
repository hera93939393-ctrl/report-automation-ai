"""chat_assistant.py — 작고 예쁜 채팅창 + Ollama Qwen3 도구호출 + verify_tool 실행.
항상 위에 떠 있고 드래그 가능한 CustomTkinter 창."""
import customtkinter as ctk
from tkinter import filedialog

ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")


class ChatAssistant(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("보고서 도우미")
        self.geometry("320x480")
        self.attributes("-topmost", True)

        self.report_path = None
        self.source_path = None

        self.report_button = ctk.CTkButton(self, text="보고서 파일 선택", command=self._choose_report)
        self.report_button.pack(pady=(10, 4), padx=10, fill="x")

        self.source_button = ctk.CTkButton(self, text="원본자료 선택 (파일)", command=self._choose_source_file)
        self.source_button.pack(pady=4, padx=10, fill="x")

        self.source_folder_button = ctk.CTkButton(self, text="원본자료 선택 (폴더)", command=self._choose_source_folder)
        self.source_folder_button.pack(pady=4, padx=10, fill="x")

        self.chat_log = ctk.CTkTextbox(self, height=280)
        self.chat_log.pack(pady=10, padx=10, fill="both", expand=True)
        self.chat_log.configure(state="disabled")

        self.input_box = ctk.CTkEntry(self, placeholder_text="예: 숫자 검증해줘")
        self.input_box.pack(pady=(0, 10), padx=10, fill="x")
        self.input_box.bind("<Return>", self._on_submit)

    def _choose_report(self):
        path = filedialog.askopenfilename(filetypes=[("한글 문서", "*.hwp *.hwpx")])
        if path:
            self.report_path = path
            self._log(f"보고서 선택됨: {path}")

    def _choose_source_file(self):
        path = filedialog.askopenfilename(filetypes=[("원본자료", "*.xlsx *.xls *.hwp *.hwpx *.pdf")])
        if path:
            self.source_path = path
            self._log(f"원본자료(파일) 선택됨: {path}")

    def _choose_source_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.source_path = path
            self._log(f"원본자료(폴더) 선택됨: {path}")

    def _log(self, message: str):
        self.chat_log.configure(state="normal")
        self.chat_log.insert("end", message + "\n")
        self.chat_log.configure(state="disabled")
        self.chat_log.see("end")

    def _on_submit(self, event):
        text = self.input_box.get()
        self.input_box.delete(0, "end")
        self._log(f"나: {text}")
        # Task 15에서 Ollama 도구호출로 교체 예정. 지금은 입력이 화면에 찍히는지만 확인.


if __name__ == "__main__":
    app = ChatAssistant()
    app.mainloop()
