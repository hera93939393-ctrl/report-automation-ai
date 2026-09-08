"""chat_assistant.py — 작고 예쁜 채팅창 + Ollama Qwen3 도구호출 + verify_tool 실행.
항상 위에 떠 있고 드래그 가능한 CustomTkinter 창."""
import os
import re

import customtkinter as ctk
from tkinter import filedialog, messagebox
from PIL import Image

from attachment_preview import generate_hwp_preview_isolated, generate_text_preview
from hwp_report import HwpReport
from ignore_list import record_ignored_value
from privacy_guard import detect_pii_patterns
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
    {
        "type": "function",
        "function": {
            "name": "merge_weekly_reports",
            "description": "여러 사람이 첨부한 주간업무보고 문서에서 파란색으로 쓴 내용만 뽑아, 지금 열려있는 대상 문서의 이번주/다음주 칸으로 옮겨 붙인다",
            "parameters": {"type": "object", "properties": {}, "required": []},
        },
    },
    {
        "type": "function",
        "function": {
            "name": "fit_to_one_page",
            "description": "지금 열려있는 문서를 행간/자간/글자크기를 조금씩 줄여 1페이지에 맞춘다",
            "parameters": {"type": "object", "properties": {}, "required": []},
        },
    },
    {
        "type": "function",
        "function": {
            "name": "read_small_text_file",
            "description": "첨부된 작은 텍스트 파일(.txt/.md)을 읽어 내용을 돌려준다 - 파일 관련 질문에 답하기 전에 먼저 호출해야 한다",
            "parameters": {
                "type": "object",
                "properties": {"path": {"type": "string", "description": "읽을 파일 이름(첨부된 파일 이름을 그대로)"}},
                "required": ["path"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "read_fixture_code",
            "description": "coding_fixture/split_cost.py의 현재 코드와 지정 테스트 결과를 읽는다 - 코드를 고치기 전에 먼저 호출해야 한다",
            "parameters": {"type": "object", "properties": {}, "required": []},
        },
    },
    {
        "type": "function",
        "function": {
            "name": "propose_code_fix",
            "description": "coding_fixture/split_cost.py에 적용할 새 코드 전체를 제안한다(파일에 바로 쓰지 않음, 사용자 승인 후에만 적용됨)",
            "parameters": {
                "type": "object",
                "properties": {"new_code": {"type": "string", "description": "적용을 제안하는 새 파이썬 코드 전체"}},
                "required": ["new_code"],
            },
        },
    },
]


# (2026-09-08) 이 도구들은 이미 정확한 구조화된 결과를 채팅창에 보여준다
# (_handle_tool_result 참고) - 그 위에 모델이 다시 요약한 final_text까지
# 덧붙이면, 작은 로컬 모델이 구조화된 결과를 잘못 요약해 전달할 위험만
# 늘어난다(예: 검증에서 불일치가 있는데 "모두 정상입니다"라고 잘못 요약).
# 그래서 이 도구들이 실행됐을 때는 모델의 마무리 문장을 생략한다.
_AUTHORITATIVE_RESULT_TOOLS = {
    "verify_numbers", "polish_to_formal_style", "merge_weekly_reports",
    "fit_to_one_page", "insert_table", "insert_numbering", "propose_code_fix",
}

_VERIFY_KEYWORDS = [
    "검증", "확인", "대조", "체크", "맞는지", "틀린",
    # 2026-08-31 Task14 리뷰 반영: 이 프로젝트의 PRD.md 자체가 "검토"를 이런 요청의
    # 자연어 표현으로 20회 넘게 쓰고 있는데도 원래 목록에 빠져 있었음(예: "형식검토",
    # "최종 검토"). "점검"/"검사"/"오류"도 비개발자가 같은 요청을 할 법한 표현이라 추가.
    "검토", "점검", "검사", "오류",
]

# (2026-09-08, 하네스 재설계) "고쳐줘"는 _POLISH_KEYWORDS에도 있어 "코드
# 고쳐줘"류 문장이 둘 다에 걸릴 수 있다 - "코드"/"버그"/"결함"/"테스트 실패"는
# _POLISH_KEYWORDS에는 없는 훨씬 더 구체적인 신호라, 아래 _route_by_keywords에서
# _POLISH_KEYWORDS보다 먼저 확인해 우선권을 준다.
_CODE_FIX_KEYWORDS = ["코드", "버그", "결함", "테스트 실패", "실패하는 테스트"]

_POLISH_KEYWORDS = [
    "공문서", "다듬어", "정리해", "써줘", "작성해", "바꿔줘", "고쳐줘",
    "손봐줘", "매끄럽게", "격식있게",
]

_TABLE_KEYWORDS = ["표", "테이블", "표로", "표 만들어"]

_NUMBERING_KEYWORDS = ["번호", "번호매겨", "번호 매겨", "번호서식", "순번"]

_WEEKLY_MERGE_KEYWORDS = ["옆 한글파일로", "주간보고 취합", "주간업무보고 취합", "취합해"]

_FIT_TO_PAGE_KEYWORDS = ["한 페이지에 맞춰", "한페이지에 맞춰", "한 장에 맞춰", "페이지 맞춤", "쪽맞춤"]

# 오탐(false positive) 방지용 최소 안전장치. bare substring 매칭이라 검증과 무관한
# 문장에도 우연히 걸릴 수 있음이 리뷰에서 실측 확인됨:
#   - "체크카드로 결제했어요" → "체크"가 "체크카드"의 일부로 걸림
#   - "파일 선택 확인했어" → "확인"이 "그냥 선택을 확인했다"는 무관한 진술에 걸림
# 완전한 자연어 이해 없이 이 두 사례만 막기 위해, 매칭 전에 이런 무관한 복합어/구절을
# 먼저 지워버리는 최소 denylist를 둔다("체크"는 살려둬서 "더블체크해줘" 같은 정상
# 요청은 계속 잡히게 함). 이 정도가 "가벼운 안전망" 단계에 맞는 절충이라고 판단함 —
# 완벽한 정확도가 필요해지면(도구가 여러 개로 늘어나는 다음 라운드) 재설계 대상.
_FALSE_POSITIVE_DENYLIST = ["체크카드", "선택 확인", "확인서"]


def _preview_kind_for(ext: str) -> str | None:
    """확장자별로 어떤 미리보기를 만들지 판단한다. ext는 점(.) 포함
    소문자(예: ".hwp"). HWP류는 이미지 미리보기, 엑셀/PDF는 텍스트
    미리보기, 그 외(이미지 자체를 첨부한 경우 등)는 미리보기를 만들지
    않는다(None)."""
    if ext in (".hwp", ".hwpx"):
        return "image"
    if ext in (".xlsx", ".xls", ".pdf"):
        return "text"
    return None


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
    # (2026-09-08) 코드 관련 요청은 실제 수정 코드까지 모델이 생성해야 해서
    # (propose_code_fix의 new_code는 키워드만으로 만들 수 없음) 여기서는
    # read_fixture_code까지만 결정론적으로 대신한다 - 현재 결함/테스트 상태를
    # 보여주면 사용자가 다시 구체적으로 요청할 실마리가 된다.
    if any(keyword in cleaned_message for keyword in _CODE_FIX_KEYWORDS):
        return "read_fixture_code"
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
    if any(keyword in cleaned_message for keyword in _WEEKLY_MERGE_KEYWORDS):
        return "merge_weekly_reports"
    if any(keyword in cleaned_message for keyword in _FIT_TO_PAGE_KEYWORDS):
        return "fit_to_one_page"
    return None


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
        self._last_mismatch_items: list[str] = []  # list[str] - 마지막 숫자검증에서 빨갛게 표시된 항목들(span 순서)
        self._pending_code_approval: str | None = None  # str | None - propose_code_fix가 만든 approval_id(승인/거절 대기 중)

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
            filetypes=[("원본자료", "*.xlsx *.xls *.hwp *.hwpx *.pdf *.txt *.md")]
        )
        if not paths:
            return False
        self.source_paths = list(paths)
        names = ", ".join(os.path.basename(p) for p in self.source_paths)
        self._log(f"📎 원본자료 {len(self.source_paths)}개 첨부됨: {names}")
        for path in self.source_paths:
            self._show_attachment_preview(path)
        return True

    def _show_attachment_preview(self, path: str):
        """첨부 직후 파일 하나의 작은 미리보기를 채팅창에 바로 보여준다.
        HWP/HWPX는 이미지(별도 프로세스에서 1페이지만 렌더링, 같은
        프로세스 안의 다른 Hwp 인스턴스가 이미 열려있는 보고서의 COM
        연결을 깨뜨리는 pyhwpx 한계 때문 - attachment_preview.py 참고),
        엑셀/PDF는 텍스트(상위 5줄)를 보여준다. 실패해도 조용히
        건너뛴다(미리보기는 부가기능이라 실패가 첨부 자체를 막으면 안 됨)."""
        ext = os.path.splitext(path)[1].lower()
        kind = _preview_kind_for(ext)
        if kind is None:
            return
        try:
            if kind == "image":
                preview_path = generate_hwp_preview_isolated(path)
                if preview_path is not None:
                    card = self._log(f"미리보기: {os.path.basename(path)}", role="assistant")
                    image = ctk.CTkImage(light_image=Image.open(preview_path),
                                          dark_image=Image.open(preview_path), size=(160, 220))
                    ctk.CTkLabel(card, text="", image=image).pack(padx=8, pady=(0, 6))
            else:
                text_preview = generate_text_preview(path)
                if text_preview:
                    self._log(f"미리보기: {os.path.basename(path)}\n{text_preview}", role="assistant")
        except Exception:
            pass  # 미리보기 실패는 첨부 자체를 막지 않음 - 부가기능

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

        # (2026-09-08, 하네스 재설계) 코드 변경 승인/거절은 self.report나
        # source_paths와 무관하고, 아래 PII/보고서선택 검사를 거칠 이유가
        # 없어 가장 먼저 확인한다. 승인 여부는 반드시 사용자의 실제 메시지로만
        # 결정한다 - 모델에게 도구로 맡기면 모델이 스스로 승인해버릴 수 있어
        # code_fix_tool.apply_fix_and_test를 의도적으로 모델 도구 목록에서
        # 뺐다(INTERFACES.md 참고). 이 결정이 이번 메시지의 유일한 목적이므로
        # 다른 처리와 섞지 않고 여기서 바로 끝낸다.
        if self._pending_code_approval is not None:
            approved = parse_approval_response(text)
            if approved is not None:
                from code_fix_tool import apply_fix_and_test
                approval_id = self._pending_code_approval
                self._pending_code_approval = None
                result = apply_fix_and_test(approval_id, approved=approved)
                if not result.get("ok", False):
                    self._log(f"처리하지 못했어요 - {result['error']['message']}", role="error")
                elif not result["applied"]:
                    self._log("적용하지 않았어요(거절). 파일은 그대로예요.", role="assistant")
                else:
                    status = "통과" if result["test_result"] == "PASS" else "실패"
                    self._log(f"적용했어요. 테스트 결과: {status}\n{result['test_output']}",
                               role="success" if result["test_result"] == "PASS" else "error")
                return
            # 승인/거절 의사가 불명확하면 대기 상태를 유지하고 계속 진행 —
            # 이번 메시지는 무관한 다른 요청일 수 있다(예: 딴 얘기를 꺼냄).

        # (F13) 개인정보로 보이는 패턴이 있으면 처리 전에 확인만 구한다 —
        # 정규식 기반 감지라 오탐 가능(privacy_guard.py 참고)하므로 그대로
        # 차단하지 않고 사용자 판단에 맡긴다.
        pii_found = detect_pii_patterns(text)
        if pii_found:
            proceed = messagebox.askyesno(
                "개인정보 확인",
                f"개인정보로 보이는 패턴({', '.join(pii_found)})이 있어요. "
                "패턴이 비슷할 뿐 실제 개인정보가 아닐 수도 있어요. "
                "이대로 진행할까요?",
            )
            if not proceed:
                self._log("입력을 취소했어요.")
                return

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

        # (F14) "N번째는 무시해"도 parse_goto_index()와 같은 이유로
        # route_intent()의 LLM 라우팅을 거치지 않고 결정론적으로 처리한다.
        # parse_goto_index()보다 반드시 먼저 확인해야 한다 - 그 정규식이
        # "2번째는 무시해"에도 매치되기 때문(parse_ignore_index docstring 참고).
        ignore_index = parse_ignore_index(text)
        if ignore_index is not None:
            if not self._last_mismatch_items:
                self._log("먼저 숫자 검증을 실행해주세요.")
            elif 1 <= ignore_index <= len(self._last_mismatch_items):
                target = self._last_mismatch_items[ignore_index - 1]
                record_ignored_value(target)
                self._log(f"{ignore_index}번째 항목('{target}')을 앞으로 무시할게요.", role="assistant")
            else:
                self._log(f"총 {len(self._last_mismatch_items)}건 중 {ignore_index}번째는 없어요.")
            return

        # (F13) "N번째로 가줘"는 결정론적 패턴이라 route_intent()의 LLM
        # 라우팅을 거치지 않고 여기서 바로 처리한다 - 빠른 동작이라
        # self._busy 가드(느린 Ollama/HWP 호출용)를 씌우지 않는다.
        goto_index = parse_goto_index(text)
        if goto_index is not None:
            if not self._last_mismatch_items:
                self._log("먼저 숫자 검증을 실행해주세요.")
            elif 1 <= goto_index <= len(self._last_mismatch_items):
                target = self._last_mismatch_items[goto_index - 1]
                self.report.hwp.MoveDocBegin()
                self.report.hwp.find(target, direction="AllDoc")
                self._log(f"{goto_index}번째 항목으로 이동했어요 - '{target}'", role="assistant")
            else:
                self._log(f"총 {len(self._last_mismatch_items)}건 중 {goto_index}번째는 없어요.")
            return

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
                loop_message = f"{self._pending_clarification} {text}"
                self._pending_clarification = None
            else:
                loop_message = text

            # (2026-09-08, 하네스 재설계) route_intent()의 "이름 하나만 고르고
            # 끝나는" 1회성 호출을 agent_loop.run_agent_loop()로 바꾼다 -
            # 모델이 도구를 부르면 실제로 실행한 결과를 다시 모델에게 보여주고
            # 다음 판단을 잇는 반복 구조다(agent_loop.py 참고). dispatch는
            # 이번 메시지의 첨부/문서 상태를 클로저로 묶어 매번 새로 만든다.
            from agent_loop import run_agent_loop
            dispatch = self._build_tool_dispatch(loop_message)
            result = run_agent_loop(
                loop_message, tools_schema=_TOOLS, dispatch=dispatch,
                on_step=lambda step: (self._log(step, role="assistant"), self.update()),
            )

            if result["trace"]:
                # 모델이 도구를 하나 이상 실제로 불러 실행까지 이어진 정상
                # 경로 - 각 실행 결과를 도구별 형식으로 보여준다.
                for step in result["trace"]:
                    self._handle_tool_result(step["tool"], step["result"])
                ran_authoritative = any(
                    step["tool"] in _AUTHORITATIVE_RESULT_TOOLS for step in result["trace"]
                )
                # verify_numbers 등은 이미 정확한 구조화된 결과를 보여줬으므로,
                # 그 위에 모델이 다시 요약한 문장을 덧붙이면 모델이 구조화된
                # 결과를 잘못 요약할 위험(작은 로컬 모델에서 드물지 않음)만
                # 늘린다 - 이런 도구는 모델의 마무리 문장을 생략한다.
                if result["status"] == "completed" and result["final_text"] and not ran_authoritative:
                    self._log(result["final_text"], role="assistant")
            else:
                # 모델이 도구를 하나도 안 불렀다(답만 하거나, provider·반복
                # 한도 문제) - 키워드 안전망으로 재시도한다(F11에서 실측된
                # 도구호출 성공률 약 33% - 로컬 모델 성능/하드웨어 한계라
                # 코드만으로 완전히 없앨 수 없음). _route_by_keywords는
                # ollama.chat()을 안 부르므로 provider 실패 상황에서도 동작한다.
                tool_name = _route_by_keywords(loop_message)
                text_sources = [
                    path for path in self.source_paths
                    if os.path.splitext(path)[1].lower() in {".txt", ".md"}
                ]
                if tool_name is None and len(text_sources) == 1:
                    # 텍스트 파일이 하나만 첨부돼 있으면 별다른 키워드가 없어도
                    # "이 파일에 대한 질문"으로 본다 - file_answer.answer_file_question은
                    # 읽기+질문답변을 한 번에(Ollama→vLLM 폴백 포함) 처리하므로,
                    # 도구 호출 없이도 정확한 답을 준다(agent_loop를 다시 돌
                    # 필요 없음).
                    from file_answer import answer_file_question
                    self._log("파일 읽는 중...", role="assistant")
                    self.update()
                    answer = answer_file_question(text_sources[0], loop_message)
                    self._log(answer, role="success")
                elif tool_name is not None:
                    step_result = dispatch[tool_name]()
                    self._handle_tool_result(tool_name, step_result)
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

    def _build_tool_dispatch(self, user_message: str) -> dict:
        """agent_loop가 호출할 "도구 이름 → 인자없는 실행 함수" 매핑을 이번
        메시지 기준으로 새로 만든다. self.report/self.source_paths(그 순간의
        첨부·문서 상태)와 user_message(공문서체 변환 등이 필요로 함)를
        클로저로 묶는다 - 매 메시지마다 상태가 바뀔 수 있어(새 첨부, 새 문서)
        한 번만 만들어 재사용하지 않는다.

        각 함수는 agent_loop의 도구 계약대로 항상 {"ok": bool, ...}를
        반환한다 - 예외를 던지는 경우는 agent_loop.run_agent_loop가 공통으로
        잡아 TOOL_EXECUTION_ERROR로 바꾼다(여기서 각자 try/except할 필요 없음).
        """
        def verify_numbers_tool():
            if not self.source_paths:
                return {"ok": False, "error": {
                    "code": "NO_SOURCE_ATTACHED", "message": "원본자료를 먼저 '+'로 첨부해주세요.",
                }}
            from verify_tool import run_verification
            return {"ok": True, **run_verification(self.report, self.source_paths, default_year=2026)}

        def polish_tool():
            from polish_tool import polish_to_formal_style
            from speed_tracker import estimate_seconds
            # (F13) 입력 글자수 × 1.5를 예상 출력 토큰수로 대략 추정해서,
            # 지금까지 기록된 실제 속도로 예상 소요시간을 미리 보여준다.
            # 기록이 아직 없으면(첫 실행) estimate_seconds가 None을
            # 반환하므로 이 메시지 자체를 생략한다.
            expected_tokens = int(len(user_message) * 1.5)
            estimate = estimate_seconds(expected_tokens)
            if estimate is not None:
                lo, hi = estimate
                self._log(f"공문서체로 다듬는 중이에요... (예상 소요시간 약 {lo:.0f}~{hi:.0f}초)", role="assistant")
                self.update()
            return {"ok": True, **polish_to_formal_style(self.report, user_message)}

        def merge_weekly_tool():
            if not self.source_paths:
                return {"ok": False, "error": {
                    "code": "NO_SOURCE_ATTACHED", "message": "취합할 주간업무보고 문서들을 먼저 '+'로 첨부해주세요.",
                }}
            from weekly_report_tool import merge_weekly_reports
            return {"ok": True, **merge_weekly_reports(self.report, self.source_paths)}

        def fit_to_page_tool():
            from fit_to_page_tool import fit_to_one_page
            return {"ok": True, **fit_to_one_page(self.report)}

        def insert_table_tool():
            # 실제 삽입은 스타일 카드에서 사용자가 버튼을 눌러야 이어진다 -
            # 이 함수는 카드를 보여주는 데까지만 책임진다(팝업 자체가
            # "행동이 실행됐다"는 사용자 확인 장치, D06/기존 UX 유지).
            self._show_table_style_picker()
            return {"ok": True, "note": "스타일 선택 카드를 보여드렸어요"}

        def insert_numbering_tool():
            self._show_numbering_style_picker()
            return {"ok": True, "note": "스타일 선택 카드를 보여드렸어요"}

        def read_small_text_file_tool(path: str = ""):
            from file_answer import read_small_text_file
            text_sources = [
                p for p in self.source_paths if os.path.splitext(p)[1].lower() in {".txt", ".md"}
            ]
            if not text_sources:
                return {"ok": False, "error": {
                    "code": "NO_SOURCE_ATTACHED", "message": ".txt/.md 파일을 먼저 '+'로 첨부해주세요.",
                }}
            # 모델이 준 path는 무시하고 실제 첨부된 파일을 읽는다 - 작은
            # 로컬 모델이 파일명을 정확히 그대로 되돌려준다는 보장이 없어
            # (도구호출 성공률 약 33%, _route_by_keywords 주석 참고), 이미
            # UI로 확정된 첨부 상태를 신뢰하는 쪽을 택했다(기존 5개 도구와
            # 같은 판단 - verify_numbers 등도 모델에게 경로를 안 받는다).
            try:
                return {"ok": True, **read_small_text_file(text_sources[0])}
            except (ValueError, FileNotFoundError) as error:
                return {"ok": False, "error": {"code": "FILE_READ_ERROR", "message": str(error)}}

        from code_fix_tool import propose_code_fix, read_fixture_code

        return {
            "verify_numbers": verify_numbers_tool,
            "polish_to_formal_style": polish_tool,
            "merge_weekly_reports": merge_weekly_tool,
            "fit_to_one_page": fit_to_page_tool,
            "insert_table": insert_table_tool,
            "insert_numbering": insert_numbering_tool,
            "read_small_text_file": read_small_text_file_tool,
            "read_fixture_code": read_fixture_code,
            "propose_code_fix": propose_code_fix,
        }

    def _handle_tool_result(self, tool_name: str, result: dict):
        """agent_loop(또는 키워드 안전망의 직접 실행)가 실행한 도구 하나의
        결과를 채팅창에 사람이 읽을 메시지로 바꾸고, 필요한 세션 상태(마지막
        검증 결과, 승인 대기 등)를 갱신한다. 도구별 반환 형태가 달라
        INTERFACES.md의 도구별 계약과 1:1로 대응해 분기한다.

        read_small_text_file/read_fixture_code는 여기서 별도로 로그를
        남기지 않는다 - 전자는 agent_loop 반복에서 모델이 파일 내용을 보고
        다음 턴에 만드는 final_text가 실제 답이고(_on_submit 참고), 후자는
        모델이 이어서 propose_code_fix를 부르거나 final_text로 설명하므로
        원본 코드를 여기서 또 통째로 찍으면 중복이다."""
        if not result.get("ok", False):
            error = result.get("error", {})
            self._log(f"실패했어요 - {error.get('message', error.get('code', '알 수 없는 오류'))}", role="error")
            return

        if tool_name == "verify_numbers":
            self._last_mismatch_items = result["mismatch_items"]  # (F13) "N번째로 가줘" 이동용
            # (2026-09-04, 실사용 피드백) 문서에 색으로 표시만 해서는
            # "빨강/파랑/초록/회색이 각각 무슨 뜻인지" 알 수 없다는 지적을
            # 받아, 검증 결과와 함께 매번 색 범례를 같이 보여준다(회색은
            # 2026-09-08 문맥 연결 판정 추가로 새로 생긴 4번째 색).
            self._log(
                "빨강: 원본과 다름(오류) / 파랑: 원본과 일치(정상) / "
                "초록: 원본에 항목 자체가 없어 대조불가 / 회색: 문맥상 어느 항목인지 불명확(확인 필요)",
                role="assistant",
            )
            self._log(result["summary"], role="success")
        elif tool_name == "polish_to_formal_style":
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
        elif tool_name == "merge_weekly_reports":
            merged_names = ", ".join(os.path.basename(p) for p in result["merged_files"])
            lines = []
            if result["merged_files"]:
                lines.append(f"{len(result['merged_files'])}건 취합했어요: {merged_names}")
            if result["no_content_files"]:
                names = ", ".join(os.path.basename(p) for p in result["no_content_files"])
                lines.append(f"파란색 내용이 없어 건너뜀: {names}")
            if result["skipped_files"]:
                names = ", ".join(os.path.basename(p) for p in result["skipped_files"])
                lines.append(f"한글 문서가 아니라 건너뜀: {names}")
            self._log("\n".join(lines) if lines else "취합할 내용이 없었어요.", role="success")
        elif tool_name == "fit_to_one_page":
            if result["fitted"]:
                method_label = {
                    "already_one_page": "이미 1페이지였어요",
                    "linespacing": "행간을 줄여서",
                    "spacing": "자간까지 줄여서",
                    "font_size": "글자크기까지 줄여서",
                }[result["method"]]
                self._log(f"{method_label} 1페이지로 맞췄어요.", role="success")
            else:
                self._log("행간/자간/글자크기를 다 줄여봐도 1페이지에 안 들어가요. 내용을 좀 줄여주세요.", role="error")
        elif tool_name == "propose_code_fix":
            self._pending_code_approval = result["approval_id"]
            self._log(
                f"변경 전:\n{result['current_code']}\n\n변경 후(제안):\n{result['proposed_code']}\n\n"
                "적용할까요? '승인' 또는 '거절'로 답해주세요.",
                role="assistant",
            )
        elif tool_name == "read_fixture_code":
            status = "실패(FAIL)" if result["test_result"] == "FAIL" else "통과(PASS)"
            self._log(f"현재 코드:\n{result['code']}\n\n지정 테스트 결과: {status}", role="assistant")
        # insert_table/insert_numbering/read_small_text_file은 위 docstring
        # 참고 - 의도적으로 여기서 아무것도 로그하지 않는다.


def _selftest_build_preview_message_routes_by_extension():
    """확장자에 따라 어떤 종류의 미리보기를 만들지 올바르게 판단하는지
    확인한다(HWP류는 이미지, 엑셀/PDF는 텍스트, 그 외는 미리보기 없음).
    이 함수는 실제 pyhwpx/openpyxl을 부르지 않고 순수하게 분기만
    검증한다 - 실제 생성은 attachment_preview.py 쪽 self-test가 담당."""
    assert _preview_kind_for(".hwp") == "image"
    assert _preview_kind_for(".hwpx") == "image"
    assert _preview_kind_for(".xlsx") == "text"
    assert _preview_kind_for(".pdf") == "text"
    assert _preview_kind_for(".txt") is None
    print("_preview_kind_for 통과")


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

    # 9) F14: 주간업무보고 취합 도구가 키워드로 잡히는지 확인
    weekly_choice = _route_by_keywords("옆 한글파일로 옮겨줘")
    assert weekly_choice == "merge_weekly_reports", weekly_choice
    print("route_intent 통과 (주간보고 취합):", weekly_choice)

    # 10) F14: 한 페이지 맞춤 도구가 키워드로 잡히는지 확인
    fit_choice = _route_by_keywords("이거 한 페이지에 맞춰줘")
    assert fit_choice == "fit_to_one_page", fit_choice
    print("route_intent 통과 (한 페이지 맞춤):", fit_choice)

    # 11) 2026-09-08 하네스 재설계: 코드 관련 요청이 read_fixture_code로 잡히는지 확인
    code_choice = _route_by_keywords("코드에 버그가 있는 것 같아")
    assert code_choice == "read_fixture_code", code_choice
    print("route_intent 통과 (코드 결함 조회):", code_choice)

    # 12) "고쳐줘"만으로는 여전히 polish - _CODE_FIX_KEYWORDS가 없는 문장은
    # 기존 우선순위(_POLISH_KEYWORDS)를 그대로 따라야 회귀가 아니다.
    polish_still_wins = _route_by_keywords("이 문장 고쳐줘")
    assert polish_still_wins == "polish_to_formal_style", polish_still_wins
    print("route_intent 통과 (코드 키워드 없는 '고쳐줘'는 여전히 공문서체):", polish_still_wins)

    # 13) "코드"+"고쳐줘"가 같이 있으면 read_fixture_code가 이긴다(더 구체적인 신호 우선)
    code_over_polish = _route_by_keywords("코드 고쳐줘")
    assert code_over_polish == "read_fixture_code", code_over_polish
    print("route_intent 통과 (코드+고쳐줘 동시 등장 시 코드 우선):", code_over_polish)


def parse_goto_index(text: str) -> int | None:
    """"3번째로 가줘"류 입력에서 순서 번호(1-based)를 뽑는다. 매치 안 되면
    None(숫자검증/공문서체 등 다른 요청과 혼동하지 않기 위해, route_intent()의
    LLM 라우팅을 거치지 않고 chat_assistant.py가 이 결정론적 정규식으로
    직접 판단한다)."""
    m = re.search(r'(\d+)번째', text)
    return int(m.group(1)) if m else None


def _selftest_parse_goto_index():
    assert parse_goto_index("3번째로 가줘") == 3
    assert parse_goto_index("10번째 항목 보여줘") == 10
    assert parse_goto_index("숫자 검증해줘") is None
    print("parse_goto_index 통과")


def parse_ignore_index(text: str) -> int | None:
    """"2번째는 무시해"류 입력에서 순서 번호(1-based)를 뽑는다.
    parse_goto_index()와 같은 결정론적 정규식 방식이지만, "무시"/"괜찮"
    키워드가 함께 있어야만 매치된다 - 이 조건이 없으면 "3번째로 가줘"
    (순수 이동 요청)까지 무시 요청으로 잘못 인식하게 된다. _on_submit()은
    반드시 이 함수를 parse_goto_index()보다 먼저 확인해야 한다 -
    parse_goto_index()의 정규식(r'(\d+)번째')은 "2번째는 무시해"에도
    매치되므로, 순서를 바꾸면 무시 요청이 이동 요청으로 잘못 처리된다."""
    m = re.search(r'(\d+)번째.*(?:무시|괜찮)', text)
    return int(m.group(1)) if m else None


def _selftest_parse_ignore_index():
    assert parse_ignore_index("2번째는 무시해") == 2
    assert parse_ignore_index("3번째는 그냥 괜찮아") == 3
    assert parse_ignore_index("3번째로 가줘") is None  # 무시/괜찮 키워드 없음 - goto와 구분
    assert parse_ignore_index("숫자 검증해줘") is None
    print("parse_ignore_index 통과")


def parse_approval_response(text: str) -> bool | None:
    """propose_code_fix가 만든 변경 제안이 승인 대기 중일 때, 사용자의 다음
    메시지가 승인인지 거절인지 판단한다. True/False/None(의사 불명확) 중
    하나를 반환한다 - 모델에게 판단을 맡기지 않는다(code_fix_tool.py의
    apply_fix_and_test docstring 참고, R05 승인 자기결정 방지).

    거절 키워드를 승인 키워드보다 먼저 확인한다 - "거절"에는 "절" 안에
    다른 승인 키워드가 우연히 겹칠 걱정은 없지만, 모호한 표현(예: "아니
    이거 말고 승인해줘")에서는 안전 쪽(거절)으로 치우치는 게 이 프로젝트의
    일관된 원칙(예: verify_numbers의 ambiguous 회색 판정)과 같은 방향이다."""
    if any(keyword in text for keyword in ("거절", "취소", "안 할래", "하지마", "하지 마")):
        return False
    if any(keyword in text for keyword in ("승인", "적용해", "적용", "네", "예", "좋아")):
        return True
    return None


def _selftest_parse_approval_response():
    assert parse_approval_response("승인") is True
    assert parse_approval_response("적용해줘") is True
    assert parse_approval_response("네") is True
    assert parse_approval_response("거절") is False
    assert parse_approval_response("아니 취소할래") is False
    assert parse_approval_response("오늘 날씨 어때") is None
    # 거절 키워드가 있으면 승인 키워드가 같이 있어도 안전 쪽(거절)으로 처리
    assert parse_approval_response("거절할래, 승인 아니야") is False
    print("parse_approval_response 통과")


if __name__ == "__main__":
    import sys
    if "--selftest" in sys.argv:
        _selftest_route_intent()
        _selftest_parse_goto_index()
        _selftest_parse_ignore_index()
        _selftest_parse_approval_response()
        _selftest_build_preview_message_routes_by_extension()
    else:
        app = ChatAssistant()
        app.mainloop()
