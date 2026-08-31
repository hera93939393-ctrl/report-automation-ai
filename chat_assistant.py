"""chat_assistant.py — 작고 예쁜 채팅창 + Ollama Qwen3 도구호출 + verify_tool 실행.
항상 위에 떠 있고 드래그 가능한 CustomTkinter 창."""
import customtkinter as ctk
from tkinter import filedialog
import ollama

ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")

_TOOLS = [{
    "type": "function",
    "function": {
        "name": "verify_numbers",
        "description": "지금 열려있는 한글 보고서의 금액/날짜/시간/전화번호를 원본데이터와 대조해서 틀린 부분을 빨간색으로 표시한다",
        "parameters": {"type": "object", "properties": {}, "required": []},
    },
}]


_VERIFY_KEYWORDS = ["검증", "확인", "대조", "체크", "맞는지", "틀린"]


def route_intent(user_message: str) -> str | None:
    """사용자의 자연어 입력이 verify_numbers 도구를 원하는지 판단한다.
    Ollama qwen3.5:2b의 도구호출을 우선 시도하지만, 실측 결과 이 모델의
    도구호출 신뢰도가 낮아(3회 중 1회만 성공, 각 호출 ~1분45초 소요 — 재시도로
    신뢰도를 높이면 채팅 UI가 5분 이상 멈춘 것처럼 느껴짐) 도구호출이 비어있어도
    명백한 검증 요청 키워드가 있으면 verify_numbers로 보내는 안전망을 둔다.
    지금은 등록된 도구가 하나뿐이라 이 방식이 안전하다 — 도구가 여러 개로 늘어나면
    이 키워드 안전망은 재설계가 필요하다(TODO 아님, 다음 라운드에서 다룰 설계 결정).
    """
    response = ollama.chat(
        model="qwen3.5:2b",
        messages=[{"role": "user", "content": user_message}],
        tools=_TOOLS,
    )
    tool_calls = response.get("message", {}).get("tool_calls") or []
    if tool_calls:
        return tool_calls[0]["function"]["name"]
    if any(keyword in user_message for keyword in _VERIFY_KEYWORDS):
        return "verify_numbers"
    return None


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


def _selftest_route_intent():
    # Ollama가 실제로 설치되어 있어야 통과한다 (Phase 0 사전준비 완료 전제)
    # 1) 검증 요청 -> verify_numbers (LLM 도구호출이 실패해도 키워드 안전망이 잡아줘야 함)
    tool_called = route_intent("숫자 검증해줘")
    assert tool_called == "verify_numbers", tool_called
    print("route_intent 통과 (검증 요청):", tool_called)

    # 2) 무관한 요청 -> None (안전망이 과도하게 넓지 않은지 확인)
    unrelated = route_intent("오늘 날씨 어때")
    assert unrelated is None, unrelated
    print("route_intent 통과 (무관한 요청):", unrelated)


if __name__ == "__main__":
    import sys
    if "--selftest" in sys.argv:
        _selftest_route_intent()
    else:
        app = ChatAssistant()
        app.mainloop()
