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

# (2026-09-03, 실사용 디자인 피드백) 채팅창을 CTkTextbox 로그 한 줄이 아니라
# VS Code Copilot Chat처럼 역할별 말풍선(정렬/색이 다른 CTkFrame)으로 그린다.
#
# (2026-09-03, 다크모드 지원 추가) 처음엔 appearance_mode를 "light"로
# 고정했었다(목업의 밝은 톤을 다크모드 환경에서도 똑같이 재현하려던 것).
# 이번에 다크모드 지원을 요청받아 "system"으로 되돌리고, 모든 색상값을
# 단일 hex 문자열 대신 (라이트, 다크) 튜플로 바꿨다 — CTk는 위젯마다
# color 계열 인자(fg_color/text_color/border_color/hover_color)에 튜플을
# 주면 현재 appearance_mode에 맞는 쪽을 자동으로 골라 쓴다(공식 동작,
# set_appearance_mode("system")이면 OS 다크모드 전환에도 실시간으로
# 따라간다). 다크모드 배색은 밝은 배색과 같은 색상(파랑/초록/빨강 계열)을
# 유지하되, CDS 다크모드 관례대로 "제일 어두운 배경/제일 밝은 글자" 관계를
# 뒤집었다(예: 라이트에서 흰 배경+진한 파란 글씨였던 걸, 다크에서 짙은
# 배경+밝은 파란 글씨로).
_BUBBLE_STYLE = {
    "user": {"align": "e", "bg": ("#DCEAFB", "#1B3A5C"), "fg": ("#0C447C", "#B5D4F4")},
    "assistant": {"align": "w", "bg": ("#F1EFE8", "#2C2C2A"), "fg": ("#2C2C2A", "#E6E6E6")},
    "success": {"align": "w", "bg": ("#F1EFE8", "#2C2C2A"), "fg": ("#3B6D11", "#97C459")},
    "error": {"align": "w", "bg": ("#F1EFE8", "#2C2C2A"), "fg": ("#A32D2D", "#F09595")},
}
_WINDOW_BG = ("#FAFAF8", "#1E1E1C")

# (2026-09-03, 네 번째 디자인 피드백) "하얀 바탕에 글씨는 연파랑" 요청 반영 —
# 스타일 선택 버튼(표/번호서식)을 흰 배경 + 파란 글씨의 아웃라인 버튼으로
# 바꾼다. CTk 기본 테마 버튼(흰 글씨 + 진한 파란 채우기)은 이 카드의 흰/베이지
# 배경과 대비가 너무 강해 튀어 보인다는 지적을 받아, report_button/attach_button과
# 같은 계열(연한 아웃라인)로 통일했다.
_PICKER_BUTTON = {
    "fg_color": ("#FFFFFF", "#2C2C2A"), "text_color": ("#0C447C", "#B5D4F4"), "border_width": 1,
    "border_color": ("#B5D4F4", "#185FA5"), "hover_color": ("#E6F1FB", "#042C53"),
}
_PICKER_BUTTON_CHOSEN = {
    "fg_color": ("#E6F1FB", "#042C53"), "text_color": ("#0C447C", "#B5D4F4"), "border_width": 2,
    "border_color": "#378ADD", "hover_color": ("#E6F1FB", "#042C53"),
}
_PICKER_BUTTON_UNCHOSEN = {
    "fg_color": ("#FFFFFF", "#2C2C2A"), "text_color": ("#B4B2A9", "#5F5E5A"), "border_width": 1,
    "border_color": ("#D3D1C7", "#444441"),
}

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
    {
        "type": "function",
        "function": {
            "name": "insert_table",
            "description": "첨부된 원본자료(엑셀)를 표로 변환해 지금 열려있는 한글 문서의 커서 위치에 삽입한다",
            "parameters": {"type": "object", "properties": {}, "required": []},
        },
    },
    {
        "type": "function",
        "function": {
            "name": "insert_numbering",
            "description": "번호서식(1. / 1) / (1) / ① 등)을 미리보기에서 고른 뒤 커서 위치에 삽입한다",
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

_TABLE_KEYWORDS = ["표", "테이블", "표로", "표 만들어"]

_NUMBERING_KEYWORDS = ["번호", "번호매겨", "번호 매겨", "번호서식", "순번"]

# 오탐(false positive) 방지용 최소 안전장치. bare substring 매칭이라 검증과 무관한
# 문장에도 우연히 걸릴 수 있음이 리뷰에서 실측 확인됨:
#   - "체크카드로 결제했어요" → "체크"가 "체크카드"의 일부로 걸림
#   - "파일 선택 확인했어" → "확인"이 "그냥 선택을 확인했다"는 무관한 진술에 걸림
# 완전한 자연어 이해 없이 이 두 사례만 막기 위해, 매칭 전에 이런 무관한 복합어/구절을
# 먼저 지워버리는 최소 denylist를 둔다("체크"는 살려둬서 "더블체크해줘" 같은 정상
# 요청은 계속 잡히게 함). 이 정도가 "가벼운 안전망" 단계에 맞는 절충이라고 판단함 —
# 완벽한 정확도가 필요해지면(도구가 여러 개로 늘어나는 다음 라운드) 재설계 대상.
_FALSE_POSITIVE_DENYLIST = ["체크카드", "선택 확인", "확인서"]


def _route_by_keywords(user_message: str) -> str | None:
    """키워드 안전망만으로 도구를 판단한다(로컬 LLM 호출 없음, 순수 함수).

    route_intent()의 도구호출이 실패했을 때 쓰이는 바로 그 로직이지만, 이
    함수 자체는 ollama를 전혀 부르지 않는다 — 그래서 결정적(deterministic)이다.

    (F12 2단계, 사용자 요청으로 안정화) 원래는 이 로직이 route_intent() 안에
    있었고, self-test(_selftest_route_intent)도 route_intent()를 통해서만
    검증했다. 그런데 route_intent()는 키워드 판단 전에 먼저 실제로 로컬
    LLM(qwen3.5:2b)의 도구호출을 시도하고, 그게 성공하면(즉 LLM이 스스로
    맞든 틀리든 뭔가 하나를 골라내면) 아래 키워드 로직을 아예 거치지 않고
    그 결과를 그대로 반환해버린다. 그래서 "표에 번호 매겨줘" 같은 tie-break
    케이스를 route_intent()로 테스트하면, 실행할 때마다 LLM이 그 문장을
    스스로 맞히는지 여부에 따라 결과가 달라지는 게 실측 확인됐다(F12 2단계
    Task5 최종검증에서 재현, 그리고 사용자가 직접 실행했을 때도 같은 현상
    재현됨) — 키워드 우선순위 코드 자체는 정확히 구현돼 있는데도 테스트
    결과만 흔들리는 상황이었다. 키워드 로직을 이 별도 함수로 뽑아내고
    self-test가 이 함수를 직접 부르도록 바꾸면, LLM의 성공/실패 여부와
    무관하게 "키워드 안전망 코드 자체가 맞는 우선순위로 짜여있는지"만
    순수하게 검증할 수 있다 — 부수 효과로 self-test 전체가 ollama.chat()을
    한 번도 안 부르게 되어, 이 세션에서 반복 재현됐던 "연속 호출 시 멈춤"
    현상의 영향도 받지 않는다.

    키워드 매칭은 bare substring 매칭이라 오탐 가능성이 남아있다. 실측된
    오탐 두 건은 `_FALSE_POSITIVE_DENYLIST`로 막는다(F11에서 이미 검증됨).
    """
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
    if any(keyword in cleaned_message for keyword in _TABLE_KEYWORDS):
        return "insert_table"
    # (F12 2단계 Task4 코드품질 검토에서 발견) "표"와 "번호" 두 키워드가 동시에
    # 걸리는 문장(예: "표에 번호 매겨줘" — 실제로는 번호서식을 원하는데 "표"라는
    # 단어가 스쳐 지나가듯 들어있는 경우)에서는 위 순서상 insert_table이 항상
    # 이긴다. 이 우선순위를 verify_numbers/polish_to_formal_style 사이의
    # tie-break(F12 1단계 Task8)처럼 의도적으로 설계한 건 아니고, 코드에 먼저
    # 등장한 순서가 그대로 우선순위가 된 것 — 다만 두 팝업 모두 스타일 버튼을
    # 누르기 전까지는 문서를 전혀 건드리지 않으므로(비파괴적), 어느 쪽이 먼저
    # 떠도 사용자가 원치 않는 팝업임을 보고 창을 닫은 뒤 다시 명확히 말하면
    # 복구 가능하다 — 이 정도 무해함을 근거로 순서를 그대로 두고 회귀
    # 테스트로만 고정해둔다(_selftest_route_intent 8번 참고). "표"가 실제로는
    # 그냥 스쳐가는 명사이고 "번호"가 진짜 동사인 경우가 흔할 수 있어 이
    # 우선순위가 항상 최선은 아니다 — 도구가 더 늘어나는 다음 라운드에서
    # 재검토 대상.
    if any(keyword in cleaned_message for keyword in _NUMBERING_KEYWORDS):
        return "insert_numbering"
    return None


def route_intent(user_message: str) -> str | None:
    """사용자의 자연어 입력이 어느 도구(verify_numbers/polish_to_formal_style/
    insert_table/insert_numbering)를 원하는지 판단한다. "뜻을 이해"하는
    역할은 로컬 LLM(qwen3.5:2b)의 도구호출이 담당하고, `_route_by_keywords()`는
    그 도구호출이 실패했을 때의 안전망이다(F11에서 실측된 도구호출 성공률
    약 33% — 코드만으로 완전한 자유 이해를 만들 수는 없고, 이는 결국 로컬
    모델 성능/하드웨어에 달린 문제).

    이 함수 자체는 매 호출마다 ollama.chat()을 실제로 부르므로 결정적이지
    않다(LLM 응답에 따라 같은 입력도 다른 결과가 나올 수 있음) — 그래서
    self-test는 이 함수가 아니라 `_route_by_keywords()`를 직접 검증한다
    (아래 _selftest_route_intent 참고). 이 함수는 실사용 흐름(chat_assistant
    실행 후 채팅 입력)에서 쓰인다."""
    response = ollama.chat(
        model="qwen3.5:2b",
        messages=[{"role": "user", "content": user_message}],
        tools=_TOOLS,
    )
    tool_calls = response.get("message", {}).get("tool_calls") or []
    if tool_calls:
        return tool_calls[0]["function"]["name"]
    return _route_by_keywords(user_message)


class ChatAssistant(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("보고서 도우미")
        self.geometry("320x480")
        self.attributes("-topmost", True)
        self.configure(fg_color=_WINDOW_BG)

        self.report = None  # HwpReport | None — F12: 문서를 한 번만 열고 계속 재사용
        self.source_paths = []  # list[str] — "+"로 첨부된 파일/폴더 경로 목록 (Task 7에서 실제 채워짐, 이 태스크에선 아직 빈 리스트로만 둠)
        self._busy = False
        self._pending_clarification = None  # str | None — 되묻기 대상이었던 원문

        # (2026-09-03, 세 번째 디자인 피드백) 이 버튼은 항상 떠 있는 상시
        # UI라, 채팅 안의 스타일 선택 버튼(그 순간 골라야 하는 것)과 같은
        # 진한 파란색을 쓰면 오히려 이 버튼이 더 튀어 보인다는 지적을 받아
        # 테두리만 있는 연한 버튼으로 바꿨다 — "강조색은 한 화면에 하나만"
        # 원칙(진짜 선택해야 하는 스타일 버튼 쪽에 몰아줌).
        self.report_button = ctk.CTkButton(
            self, text="보고서 파일 선택", command=self._choose_report,
            fg_color="transparent", border_width=1, border_color=("#B4B2A9", "#5F5E5A"),
            text_color=("#2C2C2A", "#E6E6E6"), hover_color=("#F1EFE8", "#2C2C2A"),
        )
        self.report_button.pack(pady=(10, 4), padx=10, fill="x")

        # (2026-09-03, 실사용 디자인 피드백) CTkTextbox 한 줄짜리 로그 대신
        # CTkScrollableFrame 안에 메시지마다 말풍선(CTkFrame)을 쌓는다 — 표/번호
        # 스타일 선택 카드도 별도 팝업 없이 이 안에 그대로 얹힌다(_log 참고).
        self.chat_scroll = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self.chat_scroll.pack(pady=10, padx=6, fill="both", expand=True)

        self.input_row = ctk.CTkFrame(self, fg_color="transparent")
        self.input_row.pack(pady=(0, 10), padx=10, fill="x")

        # report_button과 같은 이유로 연하게 — 입력창 옆의 보조 버튼일 뿐이라
        # 채팅 속 스타일 선택 버튼보다 튀면 안 된다.
        self.attach_button = ctk.CTkButton(
            self.input_row, text="+", width=32, command=self._attach_source,
            fg_color="transparent", border_width=1, border_color=("#B4B2A9", "#5F5E5A"),
            text_color=("#2C2C2A", "#E6E6E6"), hover_color=("#F1EFE8", "#2C2C2A"),
        )
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
            self._log(f"보고서를 열 수 없습니다 - {e}", role="error")
            return
        self.report = new_report
        self._log(f"보고서 선택됨: {path}")
        position_windows(self.report.get_window_handle(), self)

    def _attach_source(self):
        """"+" 버튼 클릭 시 원본자료(파일 여러 개 또는 폴더)를 채팅 카드에서
        바로 고르게 한다. tkinter의 파일 대화상자는 "파일이든 폴더든 한
        화면에서 고르기"를 지원하지 않아, 이 카드로 두 경로를 하나의 "+"
        진입점으로 통합한다.

        (2026-09-03, 다섯 번째 디자인 피드백 — "+" 팝업도 채팅에 통합) 원래는
        별도 CTkToplevel 팝업(_make_topmost_popup, topmost 가려짐 버그를
        after(50, lift/focus_force)로 방어했었음)이었다 — 표/번호서식 스타일
        카드와 같은 이유로 채팅 말풍선 안에 임베드하는 걸로 바꿨다. 팝업 자체가
        없어지니 topmost 가려짐 문제도 함께 사라진다(더는 CTkToplevel을 만들지
        않으므로). 파일탐색기 대화상자(filedialog)는 OS 표준 창이라 그대로
        별도로 뜬다 — 임베드 대상은 "파일이냐 폴더냐"를 고르는 버튼 두 개뿐이다.

        사용자가 파일탐색기에서 취소하면 아무것도 첨부되지 않는데, 이 경우
        버튼을 "골랐다"고 표시하면 실제로는 아무 일도 안 일어났는데 그런 것
        처럼 보여 오해를 준다 — 그래서 _pick_source_files/_pick_source_folder가
        실제로 첨부에 성공했는지(bool)를 반환하도록 바꾸고, 성공했을 때만
        체크마크 강조를 적용한다. 취소했을 땐 버튼을 그대로 살려둬서 다시
        시도할 수 있게 한다."""
        card = self._log("원본자료를 첨부합니다. 파일 또는 폴더를 선택하세요:", role="assistant")
        buttons = []

        def choose(pick_fn, button):
            attached = pick_fn()
            if not attached:
                return
            for b in buttons:
                if b is button:
                    b.configure(text=f"✓ {b.cget('text')}", **_PICKER_BUTTON_CHOSEN)
                else:
                    b.configure(**_PICKER_BUTTON_UNCHOSEN)
                b.configure(state="disabled")

        file_button = ctk.CTkButton(card, text="파일 선택 (여러 개 가능)", **_PICKER_BUTTON)
        file_button.configure(command=lambda b=file_button: choose(self._pick_source_files, b))
        file_button.pack(pady=3, padx=8, fill="x")
        buttons.append(file_button)

        folder_button = ctk.CTkButton(card, text="폴더 선택", **_PICKER_BUTTON)
        folder_button.configure(command=lambda b=folder_button: choose(self._pick_source_folder, b))
        folder_button.pack(pady=3, padx=8, fill="x")
        buttons.append(folder_button)

    def _run_tool_safely(self, fn, *args):
        """미리보기 팝업의 choose() 콜백에서 실제 도구 함수(예:
        insert_table_from_source/insert_numbering_prefix)를 실행한다. 예외가
        나면 채팅창에 정직하게 실패를 알리고 None을 반환한다 —
        _show_table_style_picker/_show_numbering_style_picker가 각자 갖고
        있던 동일한 try/except를 Task 4 코드품질 검토에서 뽑아냈다.

        (코드품질 검토에서 발견) 이 콜백들은 _on_submit()의 try/except 범위
        밖에서(버튼 클릭 시점에 별도로) 실행된다 — _show_*_picker() 자체는
        팝업만 띄우고 곧바로 리턴하므로, _on_submit()의 try 블록은 실제 도구
        실행이 일어나기 전에 이미 끝나 있다. 이 프로젝트에서 이미 여러 번
        실측된 pyhwpx COM 자동화의 간헐적 불안정성(pywintypes.com_error 등)이
        여기서 터지면, 이 가드가 없을 경우 verify_numbers/polish_to_formal_style과
        달리 채팅 로그에 아무 메시지도 안 남고 조용히 실패해 사용자가 원인을
        알 수 없다 — 다른 도구들과 동일하게 정직하게 실패를 알린다."""
        try:
            return fn(*args)
        except Exception as e:
            self._log(f"오류가 발생했습니다 - {e}", role="error")
            return None

    def _draw_table_style_rows(self, parent, on_choose):
        """표 스타일 3종(TABLE_STYLES)을 미니 표 미리보기+버튼 행으로 그려
        parent 안에 넣는다. _show_table_style_picker와, 선택된 텍스트가
        있을 때의 _show_numbering_style_picker가 공유하는 그리기 로직이다
        (2026-09-03, 실사용 피드백 반영 — "번호 매겨줘"도 선택이 있으면
        표를 만드는 동작으로 바뀌면서, 두 팝업이 똑같이 표 스타일을 보여줄
        필요가 생겨 여기로 뽑아냈다). on_choose(style_key)는 버튼 클릭 시
        호출되며, 실제 삽입은 호출자가 담당한다(이 함수는 UI만 그린다).

        (2026-09-03, 두 번째 디자인 피드백) parent가 더는 별도 CTkToplevel
        팝업이 아니라 채팅 말풍선(self._log가 반환한 CTkFrame)이다 — 카드가
        팝업처럼 사라지지 않고 대화 기록에 그대로 남아있어야 하므로, 버튼을
        누르면 그 버튼에 파란 테두리를 표시하고 나머지 버튼은 비활성화해
        "이걸 골랐다"는 게 카드 자체에 남도록 했다(목업 스크린샷의 파란
        테두리 강조와 동일한 의도). 폭도 팝업(360px) 기준에서 채팅창 실제
        폭(모니터에 따라 더 좁을 수 있음, window_layout.py 참고)에 맞춰
        줄였다(라벨 50→38px, 폰트 축소)."""
        from table_tool import TABLE_STYLES

        buttons = []

        def handle_click(style_key, button):
            # 고른 버튼엔 체크마크 + 연파랑 배경 강조, 나머지는 흐린 회색으로
            # 눌러서 "이걸 골랐다"가 한눈에 보이도록 했다(_PICKER_BUTTON_* 참고).
            for b in buttons:
                if b is button:
                    b.configure(text=f"✓ {b.cget('text')}", **_PICKER_BUTTON_CHOSEN)
                else:
                    b.configure(**_PICKER_BUTTON_UNCHOSEN)
                b.configure(state="disabled")
            on_choose(style_key)

        for style_key, style_info in TABLE_STYLES.items():
            row = ctk.CTkFrame(parent, fg_color="transparent")
            row.pack(pady=4, padx=8, fill="x")

            # 축소 모형: 2열짜리 미니 표를 Label 격자로 직접 그린다. 헤더 행에만
            # header_fill 색을 적용해 실제 표 삽입 결과와 시각적으로 대응시킨다.
            #
            # (사용자 요청으로 수정) "기본형(테두리만)" 스타일은 실제 삽입 시
            # header_fill=None이라 table_tool.py가 cell_fill()을 아예 호출하지
            # 않는다 — 즉 실제 결과는 배경색이 전혀 안 입혀진 흰 헤더다. fg_color를
            # 흰색으로 고정해(부모가 이제 팝업이 아니라 다른 톤의 말풍선이라
            # "부모와 같은 배경 물려받기" 방식이 더는 안전하지 않음) "배경색 없음"이
            # 실제로 흰 배경으로 보이게 했다 — 미리보기와 실제 삽입 결과를 일치시켰다.
            # (2026-09-03, "표 모양을 더 다양하게" 요청 반영) 데이터 행을 1개가
            # 아니라 2개(홀/짝) 그려서 "striped" 스타일의 줄무늬가 미리보기에도
            # 보이게 했다 — 데이터 행이 하나뿐이면 header_fill=None인 "기본형"과
            # "striped"가 미리보기에서 똑같아 보여 구분이 안 됐다.
            preview = ctk.CTkFrame(row, fg_color="#FFFFFF")
            preview.pack(side="left", padx=(0, 8))
            header_color = style_info["header_fill"]
            stripe_color = style_info.get("stripe_fill")
            header_kwargs = {"fg_color": "#{:02x}{:02x}{:02x}".format(*header_color)} if header_color else {"fg_color": "#FFFFFF"}
            for col, text in enumerate(["항목", "금액"]):
                ctk.CTkLabel(preview, text=text, width=38, height=18, font=ctk.CTkFont(size=10), **header_kwargs).grid(row=0, column=col, padx=1, pady=1)
            data_rows = [["인건비", "1,000,000"], ["운영비", "500,000"]]
            for row_idx, values in enumerate(data_rows):
                # build_table()의 i % 2 == 0 (0번째, 즉 첫 데이터 행)에 줄무늬를
                # 칠하는 것과 동일한 짝을 맞춘다.
                row_fill = "#{:02x}{:02x}{:02x}".format(*stripe_color) if stripe_color and row_idx % 2 == 0 else "#FFFFFF"
                for col, text in enumerate(values):
                    ctk.CTkLabel(preview, text=text, width=38, height=18, font=ctk.CTkFont(size=9), fg_color=row_fill).grid(row=row_idx + 1, column=col, padx=1, pady=1)

            button = ctk.CTkButton(row, text=style_info["label"], height=28, **_PICKER_BUTTON)
            button.configure(command=lambda k=style_key, b=button: handle_click(k, b))
            button.pack(side="left", fill="x", expand=True)
            buttons.append(button)

    def _show_table_style_picker(self):
        """"표 만들어줘" 요청 시 채팅창 안에 스타일 선택 카드를 바로 추가한다.
        실제 한글 문서를 미리 그리지 않고, CustomTkinter 위젯으로 각 스타일의
        축소 모형을 직접 그려서 보여준다(PRD 14-2 — 위젯 모형으로 "직접 보고
        선택"이라는 목표를 달성). 스타일을 고르면 즉시 insert_table_from_source를
        실행한다 — 카드가 뜨는 시점까지는 문서를 건드리지 않는다.

        (2026-09-03, 실사용 피드백 반영) insert_table_from_source() 자체가
        이제 선택 여부로 내부 분기한다(문서에서 텍스트가 선택돼 있으면 그
        선택 내용을, 없으면 첨부된 엑셀 데이터를 표로 만든다) — 그래서 이
        카드의 UI/호출 방식은 바뀌지 않는다, 어느 쪽이든 같은 함수를 그대로
        호출하면 된다.

        (2026-09-03, 두 번째 디자인 피드백) 별도 CTkToplevel 팝업 대신
        self._log()가 반환하는 말풍선(CTkFrame) 안에 스타일 버튼을 직접
        그린다 — 사용자가 "VS Code+Copilot Chat처럼 대화 흐름 안에서
        선택하고 싶다"고 요청한 것을 반영. 팝업이 사라지는 대신 카드가
        대화 기록에 그대로 남고, 고른 스타일은 파란 테두리로 표시된다
        (_draw_table_style_rows 참고)."""
        from table_tool import TABLE_STYLES, insert_table_from_source

        has_selection = self.report is not None and self.report.hwp.SelectionMode != 0
        label_text = (
            "선택한 텍스트를 표로 삽입합니다. 스타일을 선택하세요:"
            if has_selection else
            "원본자료를 표로 삽입합니다. 스타일을 선택하세요:"
        )
        card = self._log(label_text, role="assistant")

        def choose(style_key):
            result = self._run_tool_safely(insert_table_from_source, self.report, self.source_paths, style_key)
            if result is None:
                return
            if result["inserted"]:
                self._log(f"표를 삽입했어요 ({TABLE_STYLES[style_key]['label']})", role="success")
            else:
                self._log(f"표를 삽입하지 못했어요 - {result.get('reason', '알 수 없는 이유')}", role="error")

        self._draw_table_style_rows(card, choose)

    def _show_numbering_style_picker(self):
        """"번호 매겨줘" 요청 시 채팅창 안에 뜨는 카드 — 두 가지로 갈린다.

        (2026-09-03, 실사용 피드백 반영) 사용자가 실제로 써보고 준 피드백:
        "번호 매겨줘"로 원한 건 커서에 "1." 하나만 넣는 게 아니라, 선택한
        목록 전체를 번호(1·2·3·4) 붙은 표로 만드는 것이었다. 그래서:

        - 문서에 텍스트가 선택돼 있으면: 표 스타일 카드와 똑같은 UI를 보여주고
          (_draw_table_style_rows 공유), 고른 스타일로
          insert_numbered_table_from_selection()을 실행해 "번호/내용" 2열
          표를 만든다.
        - 선택이 없으면: 기존 그대로 4종 서식(1./1)/(1)/①) 버튼을 보여주고,
          고른 프리픽스를 커서 위치에 삽입한다(insert_numbering_prefix).

        (2026-09-03, 두 번째 디자인 피드백) 두 경우 모두 별도 CTkToplevel
        팝업 대신 self._log()가 반환하는 말풍선 안에 버튼을 직접 그린다 —
        _show_table_style_picker와 같은 이유. 콜백 예외처리 누락은
        _run_tool_safely()로 여전히 공통 처리된다."""
        has_selection = self.report is not None and self.report.hwp.SelectionMode != 0

        if has_selection:
            from numbering_tool import insert_numbered_table_from_selection
            from table_tool import TABLE_STYLES

            card = self._log("선택한 목록을 번호 붙은 표로 삽입합니다. 스타일을 선택하세요:", role="assistant")

            def choose_table(style_key):
                result = self._run_tool_safely(
                    insert_numbered_table_from_selection, self.report, style_key
                )
                if result is None:
                    return
                if result["inserted"]:
                    self._log(f"번호 붙은 표를 삽입했어요 ({TABLE_STYLES[style_key]['label']}, {result['row_count']}행)", role="success")
                else:
                    self._log(f"표를 삽입하지 못했어요 - {result.get('reason', '알 수 없는 이유')}", role="error")

            self._draw_table_style_rows(card, choose_table)
            return

        from numbering_tool import NUMBERING_STYLES, insert_numbering_prefix

        card = self._log("커서 위치에 번호를 삽입합니다. 서식을 선택하세요:", role="assistant")
        buttons = []

        def choose(style_key, button):
            for b in buttons:
                if b is button:
                    b.configure(text=f"✓ {b.cget('text')}", **_PICKER_BUTTON_CHOSEN)
                else:
                    b.configure(**_PICKER_BUTTON_UNCHOSEN)
                b.configure(state="disabled")
            result = self._run_tool_safely(insert_numbering_prefix, self.report, style_key)
            if result is None:
                return
            if result["inserted"]:
                self._log(f"번호를 삽입했어요 ({NUMBERING_STYLES[style_key]['label']})", role="success")
            else:
                self._log("번호 삽입에 실패했어요.", role="error")

        for style_key, style_info in NUMBERING_STYLES.items():
            example = f"{style_info['prefix']}예시 항목입니다"
            button = ctk.CTkButton(card, text=example, **_PICKER_BUTTON)
            button.configure(command=lambda k=style_key, b=button: choose(k, b))
            button.pack(pady=3, padx=8, fill="x")
            buttons.append(button)

    def _pick_source_files(self) -> bool:
        """파일탐색기에서 실제로 파일을 골랐으면 True, 취소했으면 False를
        반환한다 — _attach_source()가 이 값으로 버튼에 체크마크를 표시할지
        판단한다(2026-09-03, "+" 팝업 채팅 통합).

        (2026-09-04, 실사용 피드백) 한때 .hwp/.hwpx를 이 목록에서 뺐었다 —
        source_reader.py의 _READERS가 당시엔 .hwp/.hwpx를 지원하지 않아서
        (pyhwpx가 같은 프로세스 안의 다른 Hwp 인스턴스까지 깨뜨리는 라이브러리
        한계 때문에 F12 1단계에서 제외됨), .hwp를 원본자료로 첨부하면
        "첨부됨" 메시지는 뜨지만 실제로는 읽히는 데이터가 하나도 없어서,
        숫자검증을 돌리면 문서의 모든 값이 "대조불가(초록)"로만 나오는
        혼란스러운 결과로 이어졌다(사용자가 실제로 겪고 보고함). 이후
        사용자 제안으로 별도 프로세스 격리 방식(read_hwp_source_isolated,
        source_reader.py 참고)을 구현해 이 문제를 해결했으므로, .hwp/.hwpx도
        다시 선택 가능하게 되돌린다."""
        paths = filedialog.askopenfilenames(
            filetypes=[("원본자료", "*.xlsx *.xls *.hwp *.hwpx *.pdf")]
        )
        if not paths:
            return False
        self.source_paths = list(paths)
        names = ", ".join(os.path.basename(p) for p in self.source_paths)
        self._log(f"📎 원본자료 {len(self.source_paths)}개 첨부됨: {names}")
        return True

    def _pick_source_folder(self) -> bool:
        """_pick_source_files와 동일한 이유로 True/False를 반환한다."""
        path = filedialog.askdirectory()
        if not path:
            return False
        self.source_paths = [path]
        self._log(f"📎 원본자료(폴더) 첨부됨: {path}")
        return True

    def _log(self, message: str, role: str = "assistant"):
        """대화 내용을 말풍선 하나로 chat_scroll에 추가한다. role은
        user(오른쪽 파란)/assistant(왼쪽 기본)/success(왼쪽 초록 글자)/
        error(왼쪽 빨간 글자) 중 하나 — 예전의 "나: "/"도우미: ✅/❌" 같은
        텍스트 접두사 대신, 정렬과 글자색으로 역할과 결과를 구분한다.

        표/번호 스타일 선택 카드(_show_table_style_picker 등)는 이 함수가
        반환하는 CTkFrame(말풍선 본체) 안에 버튼을 직접 그려 넣는 방식으로
        재사용한다 — 팝업 없이 대화 흐름 안에 그대로 남기기 위함
        (2026-09-03, 두 번째 디자인 피드백)."""
        style = _BUBBLE_STYLE[role]
        row = ctk.CTkFrame(self.chat_scroll, fg_color="transparent")
        row.pack(fill="x", pady=3)
        bubble = ctk.CTkFrame(row, fg_color=style["bg"], corner_radius=10)
        bubble.pack(anchor=style["align"], padx=4)
        ctk.CTkLabel(
            bubble, text=message, text_color=style["fg"], font=ctk.CTkFont(size=13),
            wraplength=200, justify="left", anchor="w",
        ).pack(padx=10, pady=6)
        self.update_idletasks()
        # CTkScrollableFrame에는 CTkTextbox의 see("end") 같은 공개 스크롤 API가
        # 없어, 내부 캔버스(_parent_canvas)를 직접 끝까지 스크롤한다 — 공식
        # 문서엔 없지만 customtkinter 커뮤니티에서 통용되는 방법(소스로 직접 확인).
        self.chat_scroll._parent_canvas.yview_moveto(1.0)
        return bubble

    def _on_submit(self, event):
        if self._busy:
            self._log("아직 이전 요청을 처리 중이에요. 잠시만 기다려주세요.")
            return

        text = self.input_box.get()
        self.input_box.delete(0, "end")
        self._log(text, role="user")

        if not self.report:
            self._log("먼저 보고서 파일을 선택해주세요.")
            return
        # (2026-09-03, 실사용 피드백 반영) 원래는 여기서 source_paths도 필수로
        # 요구했다 — 그런데 "표 만들어줘"/"번호 매겨줘"를 문서에서 선택한
        # 텍스트로 실행하는 새 경로는 원본자료(첨부 엑셀)가 전혀 필요 없다
        # (polish_to_formal_style도 원래부터 필요 없었음). 원본자료가 실제로
        # 필요한 도구는 verify_numbers뿐이라, 그 검사는 아래 verify_numbers
        # 분기 안으로 옮겼다 — 여기서 일괄로 막으면 원본자료를 첨부 안 한
        # 사용자가 선택 기반 표/번호 기능조차 못 쓰게 되는 문제가 있었다.

        self._busy = True
        self.input_box.configure(state="disabled")
        try:
            self._log("확인 중입니다... (시간이 좀 걸릴 수 있어요)")
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
                if not self.source_paths:
                    self._log("숫자 검증을 하려면 먼저 원본자료를 '+'로 첨부해주세요.")
                else:
                    from verify_tool import run_verification
                    result = run_verification(self.report, self.source_paths, default_year=2026)
                    # (2026-09-04, 실사용 피드백) 문서에 색으로 표시만 해서는
                    # "빨강/파랑/초록이 각각 무슨 뜻인지" 알 수 없다는 지적을
                    # 받아, 검증 결과와 함께 매번 색 범례를 같이 보여준다.
                    self._log(
                        "빨강: 원본과 다름(오류) / 파랑: 원본과 일치(정상) / "
                        "초록: 원본에 항목 자체가 없어 대조불가",
                        role="assistant",
                    )
                    self._log(result["summary"], role="success")
            elif tool_name == "polish_to_formal_style":
                from polish_tool import polish_to_formal_style
                result = polish_to_formal_style(self.report, text)
                if result["applied"]:
                    self._log(f"다듬었어요 → {result['polished_text']}", role="success")
                else:
                    # polish_tool.py 자체가 이미 "LLM 빈 응답"을 applied=False로
                    # 명시적으로 구분해서 돌려주고 있는데(작은 로컬 모델에서
                    # 드물지 않게 발생), 여기서 그걸 무시하고 항상 성공 메시지
                    # 형태(f"...다듬었어요 → {빈 문자열}")로 로그를 남기면
                    # "다듬었어요 → " 뒤에 아무것도 없는 채로 찍혀 실제로는
                    # 실패했는데도 성공한 것처럼 보이는 오해를 준다. 도구의
                    # applied 계약을 그대로 반영해 정직하게 실패를 알린다.
                    self._log("다듬기에 실패했어요 (응답이 비어있었습니다). 다시 시도해주세요.", role="error")
            elif tool_name == "insert_table":
                self._show_table_style_picker()
            elif tool_name == "insert_numbering":
                self._show_numbering_style_picker()
            else:
                # PRD 13-4 "애매하면 되묻기": 실패로 끝내지 않고 다음 입력에서
                # 원문과 합쳐 재판단하도록 원문을 기억해둔다.
                self._pending_clarification = text
                self._log(
                    "무슨 뜻인지 잘 모르겠어요. 숫자 검증을 원하시면 "
                    "'검증'이라고, 문장을 다듬고 싶으시면 '공문서체'라고 "
                    "한 번 더 말씀해주시겠어요?"
                )
        except Exception as e:
            self._log(f"오류가 발생했습니다 - {e}", role="error")
        finally:
            self._busy = False
            self.input_box.configure(state="normal")


def _selftest_route_intent():
    """키워드 안전망(_route_by_keywords)만 검증한다 — ollama.chat()을 전혀
    부르지 않으므로 결정적이고, 로컬 LLM 서버가 안 떠 있어도 실행 가능하다
    (F12 2단계, 사용자 요청으로 안정화: route_intent()를 통해 검증하면 LLM이
    스스로 도구호출에 성공/실패하는지에 따라 결과가 흔들리는 게 실측
    확인됐음 — _route_by_keywords 함수 docstring에 상세 경위 기록).
    route_intent() 자체(LLM 통합 포함)는 이 self-test 범위가 아니고, 실제
    `python chat_assistant.py` 실행 후 채팅 입력으로 확인한다."""
    # 1) 검증 요청 -> verify_numbers
    tool_called = _route_by_keywords("숫자 검증해줘")
    assert tool_called == "verify_numbers", tool_called
    print("route_intent 통과 (검증 요청):", tool_called)

    # 2) 무관한 요청 -> None (안전망이 과도하게 넓지 않은지 확인)
    unrelated = _route_by_keywords("오늘 날씨 어때")
    assert unrelated is None, unrelated
    print("route_intent 통과 (무관한 요청):", unrelated)

    # 3) 2026-08-31 Task14 리뷰 반영: 새로 추가된 키워드(검토/점검/오류/검사)도
    # 안전망에서 잡히는지 확인 (PRD.md가 "검토"를 이런 요청에 20회 넘게 쓰는데
    # 정작 원래 키워드 목록에는 빠져 있었던 커버리지 공백에 대한 회귀 테스트)
    new_keywords = _route_by_keywords("이거 검토 점검하고 오류 있는지 검사해줘")
    assert new_keywords == "verify_numbers", new_keywords
    print("route_intent 통과 (검토/점검/오류/검사):", new_keywords)

    # 4) 2026-08-31 Task14 리뷰 반영: 실측된 오탐 두 건이 denylist로 막히는지 확인
    fp_choice = _route_by_keywords("파일 선택 확인했어")
    assert fp_choice is None, fp_choice
    print("route_intent 통과 (오탐 방지: 파일 선택 확인했어):", fp_choice)

    fp_card = _route_by_keywords("체크카드로 결제했어요")
    assert fp_card is None, fp_card
    print("route_intent 통과 (오탐 방지: 체크카드로 결제했어요):", fp_card)

    # 5) F12: 새로 추가된 polish_to_formal_style 도구도 키워드로 잡히는지 확인
    polish_choice = _route_by_keywords("이 문장 공문서체로 다듬어줘")
    assert polish_choice == "polish_to_formal_style", polish_choice
    print("route_intent 통과 (공문서체 변환):", polish_choice)

    # 6) F12 2단계: 새로 추가된 insert_table 도구가 키워드로 잡히는지 확인
    table_choice = _route_by_keywords("이 데이터로 표 만들어줘")
    assert table_choice == "insert_table", table_choice
    print("route_intent 통과 (표 삽입):", table_choice)

    # 7) F12 2단계: 새로 추가된 insert_numbering 도구가 키워드로 잡히는지 확인
    numbering_choice = _route_by_keywords("이 목록에 번호 매겨줘")
    assert numbering_choice == "insert_numbering", numbering_choice
    print("route_intent 통과 (번호서식):", numbering_choice)

    # 8) F12 2단계 Task4 코드품질 검토 반영: "표"와 "번호"가 한 문장에 동시에
    # 걸리는 경우(예: "표에 번호 매겨줘") insert_table이 이긴다는 현재 우선순위를
    # 회귀 테스트로 고정해둔다 — Task8에서 verify_numbers/polish_to_formal_style
    # tie-break를 테스트로 고정한 것과 같은 맥락. _route_by_keywords를 직접
    # 부르므로 이제 LLM이 이 문장을 스스로 맞히든 말든 결과가 항상 같다.
    table_numbering_tie = _route_by_keywords("표에 번호 매겨줘")
    assert table_numbering_tie == "insert_table", table_numbering_tie
    print("route_intent 통과 (표/번호 동시 등장 시 표 우선):", table_numbering_tie)


if __name__ == "__main__":
    import sys
    if "--selftest" in sys.argv:
        _selftest_route_intent()
    else:
        app = ChatAssistant()
        app.mainloop()
