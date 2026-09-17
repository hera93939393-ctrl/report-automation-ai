"""routing_graph.py — 사용자 입력을 F-도구 이름으로 분류하는 라우팅 계층.

이전에는 chat_assistant.py 안에 route_intent()가 직접 ollama.chat()을
불러 로컬 qwen3.5:2b로 도구호출을 시도했다(F11 실측 성공률 약 33% —
qwen3.5:2b는 정식 도구호출을 지원하지 않아 결과가 불안정했다). 이제
홈서버의 qwen3.5:9b(정식 tool-calling 지원)로 바꾸면서, 판단 과정 자체를
랭그래프(LangGraph) StateGraph로 재작성한다.

_route_by_keywords()의 우선순위/주석은 chat_assistant.py에서 그대로
옮겨온 것이다 — 로직을 바꾼 게 아니라 위치만 옮겼다(사용자가 "안전망은
그대로 유지"를 명시적으로 선택함, 2026-09-17)."""
from typing import Optional

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
]


_VERIFY_KEYWORDS = [
    "검증", "확인", "대조", "체크", "맞는지", "틀린",
    "검토", "점검", "검사", "오류",
]

_POLISH_KEYWORDS = [
    "공문서", "다듬어", "정리해", "써줘", "작성해", "바꿔줘", "고쳐줘",
    "손봐줘", "매끄럽게", "격식있게",
]

_TABLE_KEYWORDS = ["표", "테이블", "표로", "표 만들어"]

_NUMBERING_KEYWORDS = ["번호", "번호매겨", "번호 매겨", "번호서식", "순번"]

_WEEKLY_MERGE_KEYWORDS = ["옆 한글파일로", "주간보고 취합", "주간업무보고 취합", "취합해"]

_FIT_TO_PAGE_KEYWORDS = ["한 페이지에 맞춰", "한페이지에 맞춰", "한 장에 맞춰", "페이지 맞춤", "쪽맞춤"]

_FALSE_POSITIVE_DENYLIST = ["체크카드", "선택 확인", "확인서"]


def _route_by_keywords(user_message: str) -> Optional[str]:
    """키워드 안전망만으로 도구를 판단한다(LLM 호출 없음, 순수 함수, 결정적).
    chat_assistant.py에 있던 동일 함수를 그대로 옮긴 것 — 우선순위(검증 >
    공문서체 > 표 > 번호서식 > 주간보고취합 > 페이지맞춤)와 오탐 방지
    denylist 처리는 변경 없음(검증이 항상 이기는 이유: 미탐지가 오탐지보다
    위험하다는 판단, F11 12-6 성공기준)."""
    cleaned_message = user_message
    for phrase in _FALSE_POSITIVE_DENYLIST:
        cleaned_message = cleaned_message.replace(phrase, "")

    if any(keyword in cleaned_message for keyword in _VERIFY_KEYWORDS):
        return "verify_numbers"
    if any(keyword in cleaned_message for keyword in _POLISH_KEYWORDS):
        return "polish_to_formal_style"
    if any(keyword in cleaned_message for keyword in _TABLE_KEYWORDS):
        return "insert_table"
    if any(keyword in cleaned_message for keyword in _NUMBERING_KEYWORDS):
        return "insert_numbering"
    if any(keyword in cleaned_message for keyword in _WEEKLY_MERGE_KEYWORDS):
        return "merge_weekly_reports"
    if any(keyword in cleaned_message for keyword in _FIT_TO_PAGE_KEYWORDS):
        return "fit_to_one_page"
    return None


def _selftest_route_by_keywords():
    """키워드 안전망만 검증한다(네트워크 호출 없음, 결정적). 원래
    chat_assistant.py의 _selftest_route_intent를 그대로 옮긴 것 — 케이스/
    단언은 동일하고, 검증 대상 함수 이름에 맞춰 self-test 이름만 바꿨다."""
    tool_called = _route_by_keywords("숫자 검증해줘")
    assert tool_called == "verify_numbers", tool_called
    print("route_by_keywords 통과 (검증 요청):", tool_called)

    unrelated = _route_by_keywords("오늘 날씨 어때")
    assert unrelated is None, unrelated
    print("route_by_keywords 통과 (무관한 요청):", unrelated)

    new_keywords = _route_by_keywords("이거 검토 점검하고 오류 있는지 검사해줘")
    assert new_keywords == "verify_numbers", new_keywords
    print("route_by_keywords 통과 (검토/점검/오류/검사):", new_keywords)

    fp_choice = _route_by_keywords("파일 선택 확인했어")
    assert fp_choice is None, fp_choice
    print("route_by_keywords 통과 (오탐 방지: 파일 선택 확인했어):", fp_choice)

    fp_card = _route_by_keywords("체크카드로 결제했어요")
    assert fp_card is None, fp_card
    print("route_by_keywords 통과 (오탐 방지: 체크카드로 결제했어요):", fp_card)

    polish_choice = _route_by_keywords("이 문장 공문서체로 다듬어줘")
    assert polish_choice == "polish_to_formal_style", polish_choice
    print("route_by_keywords 통과 (공문서체 변환):", polish_choice)

    table_choice = _route_by_keywords("이 데이터로 표 만들어줘")
    assert table_choice == "insert_table", table_choice
    print("route_by_keywords 통과 (표 삽입):", table_choice)

    numbering_choice = _route_by_keywords("이 목록에 번호 매겨줘")
    assert numbering_choice == "insert_numbering", numbering_choice
    print("route_by_keywords 통과 (번호서식):", numbering_choice)

    table_numbering_tie = _route_by_keywords("표에 번호 매겨줘")
    assert table_numbering_tie == "insert_table", table_numbering_tie
    print("route_by_keywords 통과 (표/번호 동시 등장 시 표 우선):", table_numbering_tie)

    weekly_choice = _route_by_keywords("옆 한글파일로 옮겨줘")
    assert weekly_choice == "merge_weekly_reports", weekly_choice
    print("route_by_keywords 통과 (주간보고 취합):", weekly_choice)

    fit_choice = _route_by_keywords("이거 한 페이지에 맞춰줘")
    assert fit_choice == "fit_to_one_page", fit_choice
    print("route_by_keywords 통과 (한 페이지 맞춤):", fit_choice)


if __name__ == "__main__":
    _selftest_route_by_keywords()
