"""chat_assistant.py — 작고 예쁜 채팅창 + Ollama Qwen3 도구호출 + verify_tool 실행.
항상 위에 떠 있고 드래그 가능한 CustomTkinter 창."""
import customtkinter as ctk
from tkinter import filedialog
import ollama

from hwp_report import HwpReport
from window_layout import position_windows

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


_VERIFY_KEYWORDS = [
    "검증", "확인", "대조", "체크", "맞는지", "틀린",
    # 2026-08-31 Task14 리뷰 반영: 이 프로젝트의 PRD.md 자체가 "검토"를 이런 요청의
    # 자연어 표현으로 20회 넘게 쓰고 있는데도 원래 목록에 빠져 있었음(예: "형식검토",
    # "최종 검토"). "점검"/"검사"/"오류"도 비개발자가 같은 요청을 할 법한 표현이라 추가.
    "검토", "점검", "검사", "오류",
]

# 오탐(false positive) 방지용 최소 안전장치. bare substring 매칭이라 검증과 무관한
# 문장에도 우연히 걸릴 수 있음이 리뷰에서 실측 확인됨:
#   - "체크카드로 결제했어요" → "체크"가 "체크카드"의 일부로 걸림
#   - "파일 선택 확인했어" → "확인"이 "그냥 선택을 확인했다"는 무관한 진술에 걸림
# 완전한 자연어 이해 없이 이 두 사례만 막기 위해, 매칭 전에 이런 무관한 복합어/구절을
# 먼저 지워버리는 최소 denylist를 둔다("체크"는 살려둬서 "더블체크해줘" 같은 정상
# 요청은 계속 잡히게 함). 이 정도가 "가벼운 안전망" 단계에 맞는 절충이라고 판단함 —
# 완벽한 정확도가 필요해지면(도구가 여러 개로 늘어나는 다음 라운드) 재설계 대상.
_FALSE_POSITIVE_DENYLIST = ["체크카드", "선택 확인", "확인서"]


def route_intent(user_message: str) -> str | None:
    """사용자의 자연어 입력이 verify_numbers 도구를 원하는지 판단한다.
    Ollama qwen3.5:2b의 도구호출을 우선 시도하지만, 실측 결과 이 모델의
    도구호출 신뢰도가 낮아(3회 중 1회만 성공, 각 호출 ~1분45초 소요 — 재시도로
    신뢰도를 높이면 채팅 UI가 5분 이상 멈춘 것처럼 느껴짐) 도구호출이 비어있어도
    명백한 검증 요청 키워드가 있으면 verify_numbers로 보내는 안전망을 둔다.
    지금은 등록된 도구가 하나뿐이라 이 방식이 안전하다 — 도구가 여러 개로 늘어나면
    이 키워드 안전망은 재설계가 필요하다(TODO 아님, 다음 라운드에서 다룰 설계 결정).

    키워드 매칭은 여전히 bare substring 매칭이라 원리적으로 오탐 가능성이 남아있다
    (예: 새로운 복합어). 실측된 오탐 두 건은 `_FALSE_POSITIVE_DENYLIST`로 막았지만,
    이건 화이트리스트가 아니라 알려진 사례에 대한 대응이므로 이 정도 잔여 리스크는
    감내한다 — 최악의 경우도 "검증이 한 번 더 돌고 화면이 빨갛게 표시되는" 정도이고
    (자동저장이 없어 파괴적이지 않음) 사용자가 바로 알아챌 수 있는 수준이다.
    """
    response = ollama.chat(
        model="qwen3.5:2b",
        messages=[{"role": "user", "content": user_message}],
        tools=_TOOLS,
    )
    tool_calls = response.get("message", {}).get("tool_calls") or []
    if tool_calls:
        return tool_calls[0]["function"]["name"]
    cleaned_message = user_message
    for phrase in _FALSE_POSITIVE_DENYLIST:
        cleaned_message = cleaned_message.replace(phrase, "")
    if any(keyword in cleaned_message for keyword in _VERIFY_KEYWORDS):
        return "verify_numbers"
    return None


class ChatAssistant(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("보고서 도우미")
        self.geometry("320x480")
        self.attributes("-topmost", True)

        self.report = None  # HwpReport | None — F12: 문서를 한 번만 열고 계속 재사용
        self.source_paths = []  # list[str] — "+"로 첨부된 파일/폴더 경로 목록 (Task 7에서 실제 채워짐, 이 태스크에선 아직 빈 리스트로만 둠)
        self._busy = False

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
        if not path:
            return
        if self.report is not None:
            # 이미 다른 문서가 열려있으면 먼저 정리 — F12는 문서 핸들을 하나만
            # 유지하는 구조라(Task 2), 새 보고서를 고르면 이전 것과 헷갈리지
            # 않도록 명시적으로 닫는다.
            self.report.close(save=False)
        self.report = None
        # (코드품질 검토 후 수정) self.report를 먼저 None으로 비워두고, 새 문서를
        # 여는 데 성공했을 때만 채운다 — HwpReport(path)는 파일이 없을 때
        # FileNotFoundError를 던지지만, 그 외에도 간헐적인 실제 COM/RPC 오류
        # (pywintypes.com_error 등, 이 프로젝트에서 이미 여러 번 확인된 HWP COM
        # 자동화 환경 불안정성)로도 실패할 수 있다는 게 리뷰에서 실측 재현됐다.
        # 이전 코드처럼 self.report = HwpReport(path)를 먼저 시도해버리면, 이
        # 대입이 예외로 중간에 끊겼을 때 self.report가 방금 닫아버린 죽은
        # 핸들을 계속 가리키게 되는 문제가 있었다 — 지금 구조는 그 문제가 아예
        # 생길 수 없다(성공 전까지 항상 None).
        try:
            new_report = HwpReport(path)
        except Exception as e:
            self._log(f"도우미: 보고서를 열 수 없습니다 - {e}")
            return
        self.report = new_report
        self._log(f"보고서 선택됨: {path}")
        position_windows(self.report.get_window_handle(), self)

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
        if self._busy:
            self._log("도우미: 아직 이전 요청을 처리 중이에요. 잠시만 기다려주세요.")
            return

        text = self.input_box.get()
        self.input_box.delete(0, "end")
        self._log(f"나: {text}")

        if not self.report_path or not self.source_path:
            self._log("도우미: 먼저 보고서 파일과 원본자료를 선택해주세요.")
            return

        self._busy = True
        self.input_box.configure(state="disabled")
        try:
            self._log("도우미: 확인 중입니다... (시간이 좀 걸릴 수 있어요)")
            # 여러 분이 걸릴 수 있는 블로킹 호출(Ollama/HWP COM) 전에 위 메시지가
            # 실제로 화면에 그려지도록 강제로 갱신한다. update_idletasks()가 아니라
            # update()를 쓰는 이유: update_idletasks()는 대기 중인 draw만 처리하고
            # 이벤트 큐는 비우지 않아 일부 환경에서 텍스트가 그려지지 않을 수 있다.
            self.update()

            tool_name = route_intent(text)
            if tool_name == "verify_numbers":
                from verify_tool import run_verification
                result = run_verification(self.report_path, self.source_path, default_year=2026)
                self._log(f"도우미: {result['summary']}")
            else:
                self._log("도우미: 아직 이 요청은 처리할 수 있는 도구가 없어요. '숫자 검증해줘'라고 말씀해보세요.")
        except Exception as e:
            self._log(f"도우미: 오류가 발생했습니다 - {e}")
        finally:
            self._busy = False
            self.input_box.configure(state="normal")


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

    # 3) 2026-08-31 Task14 리뷰 반영: 새로 추가된 키워드(검토/점검/오류/검사)도
    # 안전망에서 잡히는지 확인 (PRD.md가 "검토"를 이런 요청에 20회 넘게 쓰는데
    # 정작 원래 키워드 목록에는 빠져 있었던 커버리지 공백에 대한 회귀 테스트)
    new_keywords = route_intent("이거 검토 점검하고 오류 있는지 검사해줘")
    assert new_keywords == "verify_numbers", new_keywords
    print("route_intent 통과 (검토/점검/오류/검사):", new_keywords)

    # 4) 2026-08-31 Task14 리뷰 반영: 실측된 오탐 두 건이 denylist로 막히는지 확인
    fp_choice = route_intent("파일 선택 확인했어")
    assert fp_choice is None, fp_choice
    print("route_intent 통과 (오탐 방지: 파일 선택 확인했어):", fp_choice)

    fp_card = route_intent("체크카드로 결제했어요")
    assert fp_card is None, fp_card
    print("route_intent 통과 (오탐 방지: 체크카드로 결제했어요):", fp_card)


if __name__ == "__main__":
    import sys
    if "--selftest" in sys.argv:
        _selftest_route_intent()
    else:
        app = ChatAssistant()
        app.mainloop()
