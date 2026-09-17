"""chat_assistant.py — 작고 예쁜 채팅창 + Ollama Qwen3 도구호출 + verify_tool 실행.
항상 위에 떠 있고 드래그 가능한 CustomTkinter 창."""
import os
import re

import customtkinter as ctk
from tkinter import filedialog, messagebox
from PIL import Image

from attachment_preview import (generate_excel_preview_image, generate_hwp_preview_isolated,
                                 generate_text_preview)
from hwp_report import HwpReport
from ignore_list import record_ignored_value
from privacy_guard import detect_pii_patterns
from routing_graph import route_intent_verbose, _route_by_keywords
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


def _preview_kind_for(ext: str) -> str | None:
    """확장자별로 어떤 미리보기를 만들지 판단한다. ext는 점(.) 포함
    소문자(예: ".hwp"). HWP류/엑셀류는 이미지 미리보기(엑셀은 2026-09-17
    추가 — "엑셀 켜서 보는 번거로움을 줄이자"는 사용자 요청, 앞쪽 몇
    행만 표로 그림), PDF는 텍스트 미리보기, 그 외(이미지 자체를 첨부한
    경우 등)는 미리보기를 만들지 않는다(None)."""
    if ext in (".hwp", ".hwpx", ".xlsx", ".xls"):
        return "image"
    if ext in (".pdf",):
        return "text"
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
        self._hover_popup = None  # ctk.CTkToplevel | None — 미리보기 확대창(클릭으로 열고 닫음)
        self._last_mismatch_items: list[str] = []  # list[str] - 마지막 숫자검증에서 빨갛게 표시된 항목들(span 순서)

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
        for path in self.source_paths:
            self._show_attachment_preview(path)
        return True

    def _show_attachment_preview(self, path: str):
        """첨부 직후 파일 하나의 작은 미리보기를 채팅창에 바로 보여준다.
        HWP/HWPX는 문서 이미지(별도 프로세스에서 1페이지만 렌더링, 같은
        프로세스 안의 다른 Hwp 인스턴스가 이미 열려있는 보고서의 COM
        연결을 깨뜨리는 pyhwpx 한계 때문 - attachment_preview.py 참고),
        엑셀은 표 이미지(앞쪽 몇 행만, 2026-09-17 추가), PDF는 텍스트(상위
        5줄)를 보여준다. 실패해도 조용히 건너뛴다(미리보기는 부가기능이라
        실패가 첨부 자체를 막으면 안 됨)."""
        ext = os.path.splitext(path)[1].lower()
        kind = _preview_kind_for(ext)
        if kind is None:
            return
        try:
            if kind == "image":
                if ext in (".hwp", ".hwpx"):
                    preview_path = generate_hwp_preview_isolated(path)
                else:
                    preview_path = generate_excel_preview_image(path)
                if preview_path is not None:
                    card = self._log(f"미리보기: {os.path.basename(path)} (클릭하면 크게 보여요)",
                                     role="assistant")
                    pil_image = Image.open(preview_path)
                    width, height = pil_image.size
                    # 엑셀 표 이미지는 HWP 페이지와 달리 열이 많으면 옆으로
                    # 넓어질 수 있다(고정 비율이 아님) - 채팅창 폭(320)보다
                    # 넓으면 비율 유지하며 줄여서 옆으로 잘려 보이지 않게 한다.
                    # 이렇게 줄이면 특히 엑셀 표의 작은 글씨(숫자)가 안 읽힐
                    # 수 있어 클릭하면 확대해서 보여주는 기능을 같이 둔다.
                    #
                    # (2026-09-17, 실사용 재현) 원래 "마우스를 올리면
                    # 커지고 떼면 사라지는" 방식으로 만들었으나, CTkLabel/
                    # CTkFrame 내부 구조 때문인지 Enter/Leave 이벤트 자체가
                    # 불안정해서(뜨자마자 닫힘, 마우스가 올라가 있는데도
                    # 닫힘, 수초 지연 등 여러 증상이 실제로 재현됨) 근본
                    # 원인을 못 찾고 사용자가 직접 "그냥 클릭으로 하자"고
                    # 결정함 — 클릭은 이벤트가 명확해서 이런 문제가 없다.
                    max_width = 280
                    if width > max_width:
                        height = int(height * max_width / width)
                        width = max_width
                    image = ctk.CTkImage(light_image=pil_image, dark_image=pil_image,
                                          size=(width, height))
                    thumb = ctk.CTkLabel(card, text="", image=image, cursor="hand2")
                    thumb.pack(padx=8, pady=(0, 6))
                    thumb.bind("<Button-1>",
                              lambda _e, p=preview_path: self._toggle_click_preview(p))
            else:
                text_preview = generate_text_preview(path)
                if text_preview:
                    self._log(f"미리보기: {os.path.basename(path)}\n{text_preview}", role="assistant")
        except Exception:
            pass  # 미리보기 실패는 첨부 자체를 막지 않음 - 부가기능

    def _toggle_click_preview(self, preview_path: str):
        """썸네일을 클릭하면, 채팅창 옆에 좀 더 큰(하지만 원본 그대로는
        아닌 — "원본 크기보다는 옆에서 알아볼 정도로만" 사용자 요청)
        미리보기를 테두리 없는 팝업으로 띄운다. 이미 떠있는 상태에서
        다시 클릭하면(썸네일이든 팝업 자신이든) 닫는다 — 마우스
        올림/뗌 이벤트 기반보다 클릭 기반이 훨씬 명확하고 안정적이다
        (2026-09-17, 실사용 재현 — 마우스 올림/뗌 방식은 CTkLabel/
        CTkFrame 내부 구조 때문인지 이벤트 자체가 불안정해서 뜨자마자
        닫히거나, 마우스가 위에 있는데도 닫히거나, 수초씩 지연되는
        증상이 있었고 근본 원인을 못 찾아 사용자가 직접 클릭 방식으로
        바꾸자고 결정함).

        (실사용 재현) HWP 썸네일은 원래 여기서 클릭할 때마다 더 높은
        해상도로 다시 렌더링했었는데, 매번 한글 COM을 새로 여는 데만
        8~9초가 걸려 클릭이 느리게 느껴졌다("한글은 너무 느리다"). 지금은
        attachment_preview.generate_hwp_preview_isolated()가 처음 첨부할
        때부터 이미 150dpi로 만들어두므로(_show_attachment_preview 참고),
        여기서는 항상 preview_path를 그대로 쓰기만 하면 되고 클릭은
        즉시 뜬다.

        (실사용 재현) 채팅창은 원래 position_windows()가 화면 오른쪽
        25% 자리에 놓는 창이라(window_layout.py), "채팅창 오른쪽에 더"
        띄우면 화면 밖으로 나가버려 안 보이는 게 실제로 재현됐다 —
        화면 폭을 확인해서 오른쪽에 공간이 없으면 왼쪽(한글 창이 있는
        방향)에 띄운다."""
        if self._hover_popup is not None:
            self._hide_hover_preview()
            return
        try:
            pil_image = Image.open(preview_path)
        except Exception:
            return
        width, height = pil_image.size
        # (2026-09-17, 실사용 재현) HWP 미리보기는 용량을 아끼려고 원래
        # 저해상도(60dpi)로 만들어서(attachment_preview.py 참고), 이전
        # 코드처럼 "원본이 목표보다 클 때만 줄이기"만 하면 원본이 이미
        # 작은 HWP 미리보기는 확대창에서도 그대로 작게 떠버린다(실사용
        # 재현 — "한글은 붙임처럼 너무 작게 나온다") — 작든 크든 항상
        # 목표 크기에 맞춰(세로가 긴 문서 페이지, 가로가 긴 표 둘 다
        # 화면 안에 들어오게 폭/높이 둘 다 제한) 키우거나 줄인다.
        max_width, max_height = 780, 850
        scale = min(max_width / width, max_height / height)
        width, height = int(width * scale), int(height * scale)

        screen_width = self.winfo_screenwidth()
        x_right = self.winfo_x() + self.winfo_width() + 8
        if x_right + width <= screen_width:
            x = x_right  # 오른쪽에 공간이 있으면 오른쪽에
        else:
            x = max(0, self.winfo_x() - width - 8)  # 없으면 왼쪽에
        # (2026-09-17, 실사용 피드백) 채팅창 맨 위와 나란히 뜨면 너무
        # 높이 뜬다는 의견이 있어, 조금 아래로 내려서 띄운다.
        y = self.winfo_y() + 60

        # (실사용 재현) overrideredirect(True)부터 먼저 부르고 그 다음에
        # geometry()로 위치를 준 순서로 짰더니, 실제로는 원본 썸네일과
        # 거의 같은 자리에 겹쳐서 떴다(Tk가 위치 지정을 무시하는 것처럼
        # 동작 — 직접 재현해서 확인함). geometry()를 overrideredirect()
        # 보다 먼저 한 번 주고, 창을 뒤늦게 건드리는 overrideredirect/
        # attributes 이후에 같은 geometry()를 한 번 더 줘야 실제
        # 화면에도 의도한 위치로 뜬다.
        popup = ctk.CTkToplevel(self)
        popup.geometry(f"{width}x{height}+{x}+{y}")
        popup.update_idletasks()
        popup.overrideredirect(True)  # 제목표시줄 없는 팝업 느낌
        popup.attributes("-topmost", True)
        popup.geometry(f"{width}x{height}+{x}+{y}")
        image = ctk.CTkImage(light_image=pil_image, dark_image=pil_image, size=(width, height))
        popup_label = ctk.CTkLabel(popup, text="", image=image, cursor="hand2")
        popup_label.pack(fill="both", expand=True)
        # 라벨과 팝업 창 둘 다에 바인딩 - 이미지가 라벨 전체를 채우긴
        # 하지만, 혹시 못 채우는 가장자리를 클릭해도 안 닫히는 일이
        #없게 이중으로 걸어둔다(실사용 재현 — "클릭해도 안 없어져").
        popup_label.bind("<Button-1>", lambda _e: self._hide_hover_preview())
        popup.bind("<Button-1>", lambda _e: self._hide_hover_preview())
        self._hover_popup = popup

    def _hide_hover_preview(self):
        """뜬 확대 미리보기가 있으면 닫는다. 없으면 아무 일도 안 한다."""
        if self._hover_popup is not None:
            self._hover_popup.destroy()
            self._hover_popup = None

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
                combined = f"{self._pending_clarification} {text}"
                tool_name, server_reachable = route_intent_verbose(combined)
                self._pending_clarification = None
            else:
                tool_name, server_reachable = route_intent_verbose(text)

            if not server_reachable:
                # (2026-09-17) 홈서버는 필요할 때만 켜는 걸 전제로 하므로
                # (personal-ai-server-roadmap 참고), 꺼져있으면 조용히 키워드
                # 방식으로만 판단하지 않고 매번 알려준다 — 사용자가 지금
                # AI가 얼마나 똑똑하게 판단했는지 신뢰 수준을 알 수 있어야 함.
                self._log("(서버에 연결할 수 없어서 간단한 키워드 방식으로 판단했어요 — 서버가 켜져 있는지 확인해보세요)")

            if tool_name == "verify_numbers":
                if not self.source_paths:
                    self._log("숫자 검증을 하려면 먼저 원본자료를 '+'로 첨부해주세요.")
                else:
                    from verify_tool import run_verification
                    result = run_verification(self.report, self.source_paths, default_year=2026)
                    self._last_mismatch_items = result["mismatch_items"]  # (F13) "N번째로 가줘" 이동용
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
                from speed_tracker import estimate_seconds
                # (F13) 입력 글자수 × 1.5를 예상 출력 토큰수로 대략 추정해서,
                # 지금까지 기록된 실제 속도로 예상 소요시간을 미리 보여준다.
                # 기록이 아직 없으면(첫 실행) estimate_seconds가 None을
                # 반환하므로 이 메시지 자체를 생략한다.
                expected_tokens = int(len(text) * 1.5)
                estimate = estimate_seconds(expected_tokens)
                if estimate is not None:
                    lo, hi = estimate
                    self._log(
                        f"공문서체로 다듬는 중이에요... (예상 소요시간 약 {lo:.0f}~{hi:.0f}초)",
                        role="assistant",
                    )
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
            elif tool_name == "merge_weekly_reports":
                if not self.source_paths:
                    self._log("먼저 취합할 주간업무보고 문서들을 '+'로 첨부해주세요.")
                else:
                    from weekly_report_tool import merge_weekly_reports
                    result = self._run_tool_safely(merge_weekly_reports, self.report, self.source_paths)
                    if result is not None:
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
                from fit_to_page_tool import fit_to_one_page
                result = self._run_tool_safely(fit_to_one_page, self.report)
                if result is not None:
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


def _selftest_build_preview_message_routes_by_extension():
    """확장자에 따라 어떤 종류의 미리보기를 만들지 올바르게 판단하는지
    확인한다(HWP류/엑셀류는 이미지, PDF는 텍스트, 그 외는 미리보기 없음 —
    엑셀이 텍스트에서 이미지로 바뀐 건 2026-09-17, 표를 그려서 보여주는
    쪽으로 변경). 이 함수는 실제 pyhwpx/openpyxl을 부르지 않고 순수하게
    분기만 검증한다 - 실제 생성은 attachment_preview.py 쪽 self-test가
    담당."""
    assert _preview_kind_for(".hwp") == "image"
    assert _preview_kind_for(".hwpx") == "image"
    assert _preview_kind_for(".xlsx") == "image"
    assert _preview_kind_for(".xls") == "image"
    assert _preview_kind_for(".pdf") == "text"
    assert _preview_kind_for(".txt") is None
    print("_preview_kind_for 통과")


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


if __name__ == "__main__":
    import sys
    if "--selftest" in sys.argv:
        _selftest_parse_goto_index()
        _selftest_parse_ignore_index()
        _selftest_build_preview_message_routes_by_extension()
    else:
        app = ChatAssistant()
        app.mainloop()
