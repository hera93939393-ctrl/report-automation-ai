"""polish_tool.py — 선택된 문장을 공문서체로 다듬거나, 채팅 입력을 새 공문서체
문장으로 만들어 커서 위치에 삽입하는 도구. verify_tool.py와 나란한 두 번째 도구."""
import os
import tempfile
import ollama

from hwp_report import HwpReport
from speed_tracker import record_call_speed


def _generate_formal_style(source_text: str, max_attempts: int = 3) -> str:
    """source_text(선택된 원문 또는 채팅 입력)를 로컬 LLM에 보내 공문서체
    문장으로 바꾼 결과를 반환한다. route_intent()와 달리 도구 호출(tools=)이
    아니라 순수 텍스트 생성이므로 tools 파라미터 없이 호출한다.

    (F12 2단계, 사용자 요청으로 완화) 작은 로컬 모델(qwen3.5:2b)은 드물지
    않게 빈 응답을 내놓는다는 게 이 프로젝트에서 이미 여러 번 실측 확인된
    특성이다 — 원래 이 재시도는 `_selftest_polish_inserts_new_text_when_nothing_selected`
    테스트 코드 안에만 있었고, 실사용 경로(polish_to_formal_style)는 첫 응답이
    비면 바로 applied=False로 실패했다. 사용자가 매번 직접 "다시 시도해주세요"를
    눌러야 했던 부분을 줄이려고, 재시도를 이 함수 안으로 옮겨 실사용에서도
    똑같이 최대 3회까지 자동으로 다시 시도하게 했다.

    근본 해결은 아니다 — 모델 자체가 매번 빈 응답만 내는 상황(3회 전부
    실패)이면 여전히 빈 문자열을 반환하고, 호출자(polish_to_formal_style)가
    이를 applied=False로 정직하게 알린다. 응답 품질(자연스러움) 자체의
    한계는 이 재시도로 해결되지 않는다 — 그건 로컬 모델 성능/하드웨어의
    문제로, 노트북 관련 논의에서 이미 별도로 다룬 사안이다."""
    for attempt in range(1, max_attempts + 1):
        response = ollama.chat(
            model="qwen3.5:2b",
            messages=[{
                "role": "user",
                "content": (
                    "다음 문장을 대한민국 공공기관 공문서에 어울리는 격식있는 "
                    "문체로 다듬어줘. 다듬은 문장만 출력하고 다른 설명은 붙이지 마:\n\n"
                    f"{source_text}"
                ),
            }],
        )
        # (F13) 이번 호출의 실제 처리속도를 기록해, 다음 호출 전 예상
        # 소요시간을 계산할 수 있게 한다(eval_count/eval_duration이 0
        # 이하인 비정상 응답은 record_call_speed가 자체적으로 걸러낸다).
        record_call_speed(response.get("eval_count", 0), response.get("eval_duration", 0))
        polished = response["message"]["content"].strip()
        if polished:
            return polished
    return ""


def polish_to_formal_style(report: HwpReport, chat_input: str) -> dict:
    """선택 여부에 따라 두 가지로 동작하는 공문서체 변환 도구.

    - 선택 있음(SelectionMode != 0): 선택된 원문을 다듬어서 그 자리에 교체
    - 선택 없음(SelectionMode == 0): chat_input 자체를 새 문장으로 만들어
      커서 위치에 삽입

    get_selected_text()는 선택이 없을 때 "현재 단어"를 자동으로 선택해버리는
    부작용이 있어(pyhwpx 소스 확인됨), 반드시 SelectionMode를 먼저 확인해야
    한다. insert_text()는 선택된 구간이 있으면 그 구간을 덮어써서 교체하므로
    (직접 프로브로 실증), 두 경우 모두 최종 적용은 insert_text() 하나로 통일된다.

    LLM이 빈 응답을 준 경우(작은 로컬 모델에서 드물지 않게 발생) 내부적으로
    최대 3회까지 자동 재시도한다(`_generate_formal_style` 참고). 그래도
    계속 비어있으면 applied=False를 반환하고 문서를 건드리지 않는다.

    (F12 2단계, 사용자 요청으로 재시도를 추가한 뒤 실측 재현된 버그를 수정)
    선택된 원문을 다듬는 경우, `_generate_formal_style()`의 재시도가 늘면서
    LLM 왕복(최대 3회)에 걸리는 시간이 늘어났는데, 그 사이 문서의 선택
    상태가 그대로 유지된다는 보장이 없다는 게 실측으로 드러났다 — 재현된
    증상은 `insert_text()`가 선택 전체를 교체하지 않고, 원문 중간의 한
    지점에만 다듬은 텍스트를 끼워넣어 "원문 앞부분 + 다듬은 문장 + 원문
    뒷부분"처럼 뒤섞인 결과를 냈다(즉 그 시점엔 이미 "전체 선택"이 아니라
    "커서 하나"만 남아있었던 것으로 보임). 2회 독립 재현 스크립트로 확인함.
    이 프로젝트의 hwp_report.py(mark_red/get_char_color_at)가 이미 쓰는
    패턴과 동일하게, insert_text() 직전에 원문(source_text)을 다시 find()로
    찾아 확실하게 재선택해서 이 문제를 막는다 — "시간이 얼마나 걸렸든, 지금
    이 순간 정말로 그 원문이 선택된 상태에서" 교체가 일어나도록 보장한다.
    알려진 한계: 문서 안에 완전히 똑같은 문장이 여러 번 나오면 find()가 첫
    occurrence를 찾으므로 사용자가 실제로 선택했던 것과 다른 occurrence가
    교체될 수 있다 — 실사용에서 동일 문장이 중복되는 경우는 드물다고 보고
    이번엔 감수하는 트레이드오프로 남겨둔다.
    """
    was_selection = report.hwp.SelectionMode != 0
    if was_selection:
        source_text = report.hwp.get_selected_text(keep_select=True)
    else:
        source_text = chat_input

    polished = _generate_formal_style(source_text)
    if not polished.strip():
        # 빈 문자열로 insert_text()를 호출하면 아무 것도 삽입되지 않는
        # 무의미한 no-op이면서도 "성공"처럼 보일 수 있다 — 특히 원래 문서가
        # 비어 있던 경우 get_text()가 계속 빈 결과를 돌려주게 되어 실제
        # 원인(LLM 빈 응답)이 가려진다. 여기서 명시적으로 실패를 알린다.
        return {"applied": False, "polished_text": ""}

    if was_selection:
        report.hwp.Cancel()
        if not report.hwp.find(source_text, direction="AllDoc"):
            # 재시도 도중 사용자가 문서를 직접 편집해 원문이 사라졌거나
            # (드문 경쟁 상태) 다른 알 수 없는 이유로 원문을 다시 못 찾으면,
            # 엉뚱한 위치에 삽입해버리는 것보다 정직하게 실패를 알리는 게
            # 낫다 — verify_tool.py/table_tool.py도 같은 원칙을 따른다.
            return {"applied": False, "polished_text": polished}

    applied = report.hwp.insert_text(polished)

    return {"applied": bool(applied), "polished_text": polished}


def _selftest_polish_replaces_selected_text():
    """문서에서 텍스트를 선택한 상태로 호출하면, 그 선택 부분만 다듬어진
    문장으로 교체되는지 확인한다 (채팅 입력이 아니라 선택된 원문이 LLM에
    들어가야 함)."""
    from pyhwpx import Hwp
    test_path = os.path.join(tempfile.gettempdir(), "_test_공문서_선택.hwp")
    setup = Hwp(visible=False, new=True)
    setup.insert_text("이번 달 예산 다 썼고 다음달에 더 필요함")
    setup.save_as(test_path)
    setup.quit()

    report = None
    try:
        report = HwpReport(test_path)
        found = report.hwp.find("이번 달 예산 다 썼고 다음달에 더 필요함", direction="AllDoc")
        assert found, "테스트 문장을 못 찾음"
        assert report.hwp.SelectionMode != 0, "선택이 안 된 상태로 테스트가 시작됨"

        result = polish_to_formal_style(report, chat_input="공문서체로 바꿔줘")
        assert result["applied"] is True, result

        final_text = report.get_text()
        assert "이번 달 예산 다 썼고 다음달에 더 필요함" not in final_text, final_text
        assert len(final_text.strip()) > 0, "문서가 비어버림(다듬은 결과가 빈 문자열)"
        print("polish_to_formal_style(선택 있음) 통과:", repr(final_text))
    finally:
        if report is not None:
            report.close(save=False)
        os.remove(test_path)


def _selftest_polish_inserts_new_text_when_nothing_selected():
    """아무것도 선택 안 한 상태로 호출하면, 채팅 입력 자체가 LLM에 들어가서
    새 문장이 커서 위치(빈 문서라 문서 시작)에 삽입되는지 확인한다.

    (F12 2단계, 사용자 요청으로 완화 후 갱신) 이 테스트는 원래 빈 응답 대응
    재시도를 테스트 코드 자체에 갖고 있었으나, 이제 `_generate_formal_style()`이
    내부적으로 최대 3회까지 자동 재시도하므로(실사용 경로도 동일한 이득을
    보도록 옮김) 이 테스트는 한 번만 호출한다 — 여기서 또 외부 재시도를
    두면 최악의 경우 3×3=9회까지 ollama.chat()을 부르게 되어 과도하다.
    그래도 3회 전부 빈 응답이면 `applied`가 명확히 `False`로 돌아오므로,
    그 경우는 "실패"로 바로 드러난다(모호한 AttributeError 대신)."""
    from pyhwpx import Hwp
    test_path = os.path.join(tempfile.gettempdir(), "_test_공문서_삽입.hwp")
    setup = Hwp(visible=False, new=True)
    setup.save_as(test_path)  # 빈 문서
    setup.quit()

    report = None
    try:
        report = HwpReport(test_path)
        assert report.hwp.SelectionMode == 0, "빈 문서인데 선택 상태로 시작됨"

        result = polish_to_formal_style(
            report, chat_input="이번 달 예산 다 썼고 다음달에 더 필요하다고 좀 써줘"
        )
        assert result["applied"] is True, (
            "3회 내부 재시도에도 로컬 모델이 계속 빈 응답을 반환함 "
            f"(polish_to_formal_style 로직 문제가 아니라 모델 자체의 불안정성으로 보임): {result}"
        )

        final_text = report.get_text()
        assert len(final_text.strip()) > 0, "빈 문서에 아무것도 안 들어감"
        print("polish_to_formal_style(선택 없음) 통과:", repr(final_text))
    finally:
        if report is not None:
            report.close(save=False)
        os.remove(test_path)


if __name__ == "__main__":
    _selftest_polish_replaces_selected_text()
    _selftest_polish_inserts_new_text_when_nothing_selected()
