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

    def _make_topmost_popup(self, title: str, geometry: str) -> ctk.CTkToplevel:
        """스타일 미리보기 팝업(CTkToplevel)을 만들어 메인 창 뒤에 가려지지
        않도록 앞으로 띄운다. _show_table_style_picker/_show_numbering_style_picker가
        공유하는 보일러플레이트 — Task 4 코드품질 검토에서 두 메서드가 이
        부분을 그대로 복붙하고 있는 걸 발견해 뽑아냈다.

        (스펙 준수 검토 중 재현/수정, Task 3) 메인 ChatAssistant 창도
        -topmost=True라서, 단순히 CTkToplevel을 만들기만 하면 팝업이 메인 창
        뒤에 가려질 수 있음이 win32gui로 실측 재현됐다(포그라운드 창이 메인
        창으로 계속 남음) — 팝업이 안 보이면 사용자에게는 무반응처럼 보여
        F12 1단계의 "무반응보다 되묻는 게 낫다" 원칙에 반한다.

        lift()/focus_force()를 여기서 바로(동기적으로) 호출하는 것만으로는
        고쳐지지 않는다: CTkToplevel.__init__ 내부(Windows용 다크 타이틀바
        처리, customtkinter/windows/ctk_toplevel.py의 _windows_set_titlebar_color)가
        생성자 안에서 self.withdraw()로 창을 일단 숨겨두고, self.after(5, ...)로
        5ms 뒤에야 deiconify()로 다시 보이게 만든다. 그래서 CTkToplevel(self)
        생성 직후 곧바로 lift()/focus_force()를 호출하면 그 시점엔 창이 아직
        (혹은 다시) 숨겨진 상태라 효과가 없다 — 실측으로 확인(win32gui로
        foreground를 찍어보면 lift/focus_force 직후 호출로는 여전히 메인
        창이 foreground로 남아있음). 그 5ms 창보다 더 뒤에 실행되도록
        after(50, ...)로 예약해야 실제로 맨 앞에 나타난다."""
        popup = ctk.CTkToplevel(self)
        popup.title(title)
        popup.geometry(geometry)
        popup.attributes("-topmost", True)
        popup.after(50, lambda: (popup.lift(), popup.focus_force()))
        return popup

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
            self._log(f"도우미: 오류가 발생했습니다 - {e}")
            return None

    def _show_table_style_picker(self):
        """"표 만들어줘" 요청 시 뜨는 스타일 미리보기 팝업. 실제 한글 문서를
        미리 그리지 않고, CustomTkinter 위젯으로 각 스타일의 축소 모형을 직접
        그려서 보여준다(PRD 14-2 — 위젯 모형으로 "직접 보고 선택"이라는
        목표를 달성). 스타일을 고르면 즉시 insert_table_from_source를 실행하고
        팝업을 닫는다 — 팝업이 뜨기 전까지는 문서를 건드리지 않는다."""
        from table_tool import TABLE_STYLES, insert_table_from_source

        picker = self._make_topmost_popup("표 스타일 선택", "360x260")

        def choose(style_key):
            picker.destroy()
            result = self._run_tool_safely(insert_table_from_source, self.report, self.source_paths, style_key)
            if result is None:
                return
            if result["inserted"]:
                self._log(f"도우미: 표를 삽입했어요 ({TABLE_STYLES[style_key]['label']})")
            else:
                self._log(f"도우미: 표를 삽입하지 못했어요 - {result.get('reason', '알 수 없는 이유')}")

        for style_key, style_info in TABLE_STYLES.items():
            row = ctk.CTkFrame(picker)
            row.pack(pady=6, padx=10, fill="x")

            # 축소 모형: 2열짜리 미니 표를 Label 격자로 직접 그린다. 헤더 행에만
            # header_fill 색을 적용해 실제 표 삽입 결과와 시각적으로 대응시킨다.
            preview = ctk.CTkFrame(row)
            preview.pack(side="left", padx=(0, 10))
            header_color = style_info["header_fill"]
            header_hex = "#{:02x}{:02x}{:02x}".format(*header_color) if header_color else "#3a3a3a"
            for col, text in enumerate(["항목", "금액"]):
                ctk.CTkLabel(preview, text=text, fg_color=header_hex, width=50, height=20).grid(row=0, column=col, padx=1, pady=1)
            for col, text in enumerate(["인건비", "1,000,000"]):
                ctk.CTkLabel(preview, text=text, width=50, height=20).grid(row=1, column=col, padx=1, pady=1)

            ctk.CTkButton(row, text=style_info["label"], command=lambda k=style_key: choose(k)).pack(side="left", fill="x", expand=True)

    def _show_numbering_style_picker(self):
        """"번호 매겨줘" 요청 시 뜨는 서식 미리보기 팝업. 각 서식이 실제로
        어떻게 보이는지 예시 문장으로 직접 보여준 뒤, 고른 프리픽스를
        커서 위치에 삽입한다.

        Task 3(_show_table_style_picker)에서 발견된 topmost 가려짐 버그와
        콜백 예외처리 누락은 _make_topmost_popup()/_run_tool_safely()로 이미
        공통 처리된다 — 두 헬퍼의 docstring에 상세 근거가 있다."""
        from numbering_tool import NUMBERING_STYLES, insert_numbering_prefix

        picker = self._make_topmost_popup("번호서식 선택", "280x200")

        def choose(style_key):
            picker.destroy()
            result = self._run_tool_safely(insert_numbering_prefix, self.report, style_key)
            if result is None:
                return
            if result["inserted"]:
                self._log(f"도우미: 번호를 삽입했어요 ({NUMBERING_STYLES[style_key]['label']})")
            else:
                self._log("도우미: 번호 삽입에 실패했어요.")

        for style_key, style_info in NUMBERING_STYLES.items():
            example = f"{style_info['prefix']}예시 항목입니다"
            ctk.CTkButton(
                picker, text=example, command=lambda k=style_key: choose(k)
            ).pack(pady=4, padx=10, fill="x")

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
            elif tool_name == "insert_table":
                self._show_table_style_picker()
            elif tool_name == "insert_numbering":
                self._show_numbering_style_picker()
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

    # 6) F12 2단계: 새로 추가된 insert_table 도구가 키워드로 잡히는지 확인
    table_choice = route_intent("이 데이터로 표 만들어줘")
    assert table_choice == "insert_table", table_choice
    print("route_intent 통과 (표 삽입):", table_choice)

    # 7) F12 2단계: 새로 추가된 insert_numbering 도구가 키워드로 잡히는지 확인
    numbering_choice = route_intent("이 목록에 번호 매겨줘")
    assert numbering_choice == "insert_numbering", numbering_choice
    print("route_intent 통과 (번호서식):", numbering_choice)

    # 8) F12 2단계 Task4 코드품질 검토 반영: "표"와 "번호"가 한 문장에 동시에
    # 걸리는 경우(예: "표에 번호 매겨줘") insert_table이 이긴다는 현재 우선순위를
    # 회귀 테스트로 고정해둔다 — Task8에서 verify_numbers/polish_to_formal_style
    # tie-break를 테스트로 고정한 것과 같은 맥락(route_intent의 if/elif 순서를
    # 코드 리딩만으로 유추해야 하는 상태를 남겨두지 않기 위함).
    table_numbering_tie = route_intent("표에 번호 매겨줘")
    assert table_numbering_tie == "insert_table", table_numbering_tie
    print("route_intent 통과 (표/번호 동시 등장 시 표 우선):", table_numbering_tie)


if __name__ == "__main__":
    import sys
    if "--selftest" in sys.argv:
        _selftest_route_intent()
    else:
        app = ChatAssistant()
        app.mainloop()
