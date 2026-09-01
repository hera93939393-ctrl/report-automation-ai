"""chat_assistant.py — 작고 예쁜 채팅창 + Ollama Qwen3 도구호출 + verify_tool 실행.
항상 위에 떠 있고 드래그 가능한 CustomTkinter 창."""
import os

import customtkinter as ctk
from tkinter import filedialog
import ollama

from hwp_report import HwpReport
from window_layout import position_windows

ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")

_TOOLS = [
    {
        "type": "function",
        "function": {
            "name": "verify_numbers",
            "description": "지금 열려있는 한글 보고서의 금액/날짜/시간/전화번호를 원본데이터와 대조해서 틀린 부분을 빨간색으로 표시한다",
            "parameters": {"type": "object", "properties": {}, "required": []},
        },
    },
    {
        "type": "function",
        "function": {
            "name": "polish_to_formal_style",
            "description": "선택된 문장을 공문서체로 다듬거나, 새 문장을 공문서체로 만들어 커서 위치에 삽입한다",
            "parameters": {"type": "object", "properties": {}, "required": []},
        },
    },
]


_VERIFY_KEYWORDS = [
    "검증", "확인", "대조", "체크", "맞는지", "틀린",
    # 2026-08-31 Task14 리뷰 반영: 이 프로젝트의 PRD.md 자체가 "검토"를 이런 요청의
    # 자연어 표현으로 20회 넘게 쓰고 있는데도 원래 목록에 빠져 있었음(예: "형식검토",
    # "최종 검토"). "점검"/"검사"/"오류"도 비개발자가 같은 요청을 할 법한 표현이라 추가.
    "검토", "점검", "검사", "오류",
]

_POLISH_KEYWORDS = [
    "공문서", "다듬어", "정리해", "써줘", "작성해", "바꿔줘", "고쳐줘",
    "손봐줘", "매끄럽게", "격식있게",
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
    """사용자의 자연어 입력이 verify_numbers/polish_to_formal_style 중 어느
    도구를 원하는지 판단한다. "뜻을 이해"하는 역할은 로컬 LLM(qwen3.5:2b)의
    도구호출이 담당하고, 키워드 목록은 그 도구호출이 실패했을 때의 안전망이다
    (F11에서 실측된 도구호출 성공률 약 33% — 코드만으로 완전한 자유 이해를
    만들 수는 없고, 이는 결국 로컬 모델 성능/하드웨어에 달린 문제).

    키워드 매칭은 bare substring 매칭이라 오탐 가능성이 남아있다. 실측된
    오탐 두 건은 `_FALSE_POSITIVE_DENYLIST`로 막는다(F11에서 이미 검증됨).
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

    # (2026-09-01 Task8 리뷰에서 발견) 한 문장에 두 키워드 목록이 동시에 걸리는
    # 경우(예: "이거 검토해서 공문서체로 다듬어줘" — "검토"와 "다듬어" 둘 다 걸림)
    # 는 verify_numbers가 항상 이긴다 — 순서상 우연이 아니라 의도적으로 정한
    # 우선순위다. 근거: _VERIFY_KEYWORDS는 검증 도구를 놓치는 게 가장 위험하다는
    # 판단(오탐지보다 미탐지가 더 나쁨, F11 12-6 성공기준 참고)이 깔려있고, 이
    # 우선순위 덕에 최악의 경우도 "공문서체 변환 대신 검증이 한 번 더 도는" 정도로
    # 그친다(자동저장 없어 비파괴적). 다만 이 tie-break는 두 키워드가 실제로
    # 겹치는 드문 입력에서만 작동하고, 문장 끝의 동사(예: "다듬어줘")가 진짜
    # 의도를 더 잘 나타내는 경우도 있어 완벽하진 않음 — 다음 라운드에서 도구가
    # 더 늘어나면 재검토 대상.
    if any(keyword in cleaned_message for keyword in _VERIFY_KEYWORDS):
        return "verify_numbers"
    if any(keyword in cleaned_message for keyword in _POLISH_KEYWORDS):
        return "polish_to_formal_style"
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
        self._pending_clarification = None  # str | None — 되묻기 대상이었던 원문

        self.report_button = ctk.CTkButton(self, text="보고서 파일 선택", command=self._choose_report)
        self.report_button.pack(pady=(10, 4), padx=10, fill="x")

        self.chat_log = ctk.CTkTextbox(self, height=280)
        self.chat_log.pack(pady=10, padx=10, fill="both", expand=True)
        self.chat_log.configure(state="disabled")

        self.input_row = ctk.CTkFrame(self, fg_color="transparent")
        self.input_row.pack(pady=(0, 10), padx=10, fill="x")

        self.attach_button = ctk.CTkButton(self.input_row, text="+", width=32, command=self._attach_source)
        self.attach_button.pack(side="left", padx=(0, 6))

        self.input_box = ctk.CTkEntry(self.input_row, placeholder_text="예: 숫자 검증해줘")
        self.input_box.pack(side="left", fill="x", expand=True)
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

    def _attach_source(self):
        """"+" 버튼 클릭 시 파일 여러 개 또는 폴더 중 고르는 작은 선택창을 띄운다.
        tkinter의 파일 대화상자는 "파일이든 폴더든 한 화면에서 고르기"를 지원하지
        않아, 이 작은 중간 선택창으로 두 경로를 하나의 "+" 진입점으로 통합한다."""
        choice_window = ctk.CTkToplevel(self)
        choice_window.title("원본자료 첨부")
        choice_window.geometry("240x110")
        choice_window.attributes("-topmost", True)

        def pick_files():
            choice_window.destroy()
            self._pick_source_files()

        def pick_folder():
            choice_window.destroy()
            self._pick_source_folder()

        ctk.CTkButton(choice_window, text="파일 선택 (여러 개 가능)", command=pick_files).pack(
            pady=(10, 4), padx=10, fill="x")
        ctk.CTkButton(choice_window, text="폴더 선택", command=pick_folder).pack(
            pady=4, padx=10, fill="x")

    def _pick_source_files(self):
        paths = filedialog.askopenfilenames(
            filetypes=[("원본자료", "*.xlsx *.xls *.hwp *.hwpx *.pdf")]
        )
        if not paths:
            return
        self.source_paths = list(paths)
        names = ", ".join(os.path.basename(p) for p in self.source_paths)
        self._log(f"📎 원본자료 {len(self.source_paths)}개 첨부됨: {names}")

    def _pick_source_folder(self):
        path = filedialog.askdirectory()
        if not path:
            return
        self.source_paths = [path]
        self._log(f"📎 원본자료(폴더) 첨부됨: {path}")

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

        if not self.report or not self.source_paths:
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

            if self._pending_clarification is not None:
                # 되묻기 응답 처리: 이전에 애매했던 원문 + 이번 답변을 합쳐
                # 같은 라우팅 로직에 다시 태운다 — 별도의 대화상태 기계 없이
                # "합쳐서 다시 판단"만으로 1회 재질문까지는 충분히 동작함(실사용
                # 검증에서 2메시지 조합 케이스로 확인됨). 다만 이번에도 또
                # 애매하면 아래 else 분기가 이번 원문만(combined 전체가 아니라)
                # 다시 pending으로 남기므로, 1라운드 전 내용은 기억되지 않고
                # "최근 두 메시지"만 유효한 컨텍스트다 — 3회 이상 연속으로
                # 애매하면 처음 의도는 잊혀지고 사용자가 처음부터 다시 말해야
                # 한다(2026-09-01 Task9 리뷰에서 확인, 도구가 늘어나는 다음
                # 라운드에서 누적형으로 재검토할 만함).
                combined = f"{self._pending_clarification} {text}"
                tool_name = route_intent(combined)
                self._pending_clarification = None
            else:
                tool_name = route_intent(text)

            if tool_name == "verify_numbers":
                from verify_tool import run_verification
                result = run_verification(self.report, self.source_paths, default_year=2026)
                self._log(f"도우미: {result['summary']}")
            elif tool_name == "polish_to_formal_style":
                from polish_tool import polish_to_formal_style
                result = polish_to_formal_style(self.report, text)
                if result["applied"]:
                    self._log(f"도우미: 다듬었어요 → {result['polished_text']}")
                else:
                    # polish_tool.py 자체가 이미 "LLM 빈 응답"을 applied=False로
                    # 명시적으로 구분해서 돌려주고 있는데(작은 로컬 모델에서
                    # 드물지 않게 발생), 여기서 그걸 무시하고 항상 성공 메시지
                    # 형태(f"...다듬었어요 → {빈 문자열}")로 로그를 남기면
                    # "다듬었어요 → " 뒤에 아무것도 없는 채로 찍혀 실제로는
                    # 실패했는데도 성공한 것처럼 보이는 오해를 준다. 도구의
                    # applied 계약을 그대로 반영해 정직하게 실패를 알린다.
                    self._log("도우미: 다듬기에 실패했어요 (응답이 비어있었습니다). 다시 시도해주세요.")
            else:
                # PRD 13-4 "애매하면 되묻기": 실패로 끝내지 않고 다음 입력에서
                # 원문과 합쳐 재판단하도록 원문을 기억해둔다.
                self._pending_clarification = text
                self._log(
                    "도우미: 무슨 뜻인지 잘 모르겠어요. 숫자 검증을 원하시면 "
                    "'검증'이라고, 문장을 다듬고 싶으시면 '공문서체'라고 "
                    "한 번 더 말씀해주시겠어요?"
                )
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

    # 5) F12: 새로 추가된 polish_to_formal_style 도구도 키워드로 잡히는지 확인
    polish_choice = route_intent("이 문장 공문서체로 다듬어줘")
    assert polish_choice == "polish_to_formal_style", polish_choice
    print("route_intent 통과 (공문서체 변환):", polish_choice)


if __name__ == "__main__":
    import sys
    if "--selftest" in sys.argv:
        _selftest_route_intent()
    else:
        app = ChatAssistant()
        app.mainloop()
