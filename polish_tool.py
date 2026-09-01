"""polish_tool.py — 선택된 문장을 공문서체로 다듬거나, 채팅 입력을 새 공문서체
문장으로 만들어 커서 위치에 삽입하는 도구. verify_tool.py와 나란한 두 번째 도구."""
import os
import tempfile
import ollama

from hwp_report import HwpReport


def _generate_formal_style(source_text: str) -> str:
    """source_text(선택된 원문 또는 채팅 입력)를 로컬 LLM에 보내 공문서체
    문장으로 바꾼 결과를 반환한다. route_intent()와 달리 도구 호출(tools=)이
    아니라 순수 텍스트 생성이므로 tools 파라미터 없이 호출한다."""
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
    return response["message"]["content"].strip()


def polish_to_formal_style(report: HwpReport, chat_input: str) -> dict:
    """선택 여부에 따라 두 가지로 동작하는 공문서체 변환 도구.

    - 선택 있음(SelectionMode != 0): 선택된 원문을 다듬어서 그 자리에 교체
    - 선택 없음(SelectionMode == 0): chat_input 자체를 새 문장으로 만들어
      커서 위치에 삽입

    get_selected_text()는 선택이 없을 때 "현재 단어"를 자동으로 선택해버리는
    부작용이 있어(pyhwpx 소스 확인됨), 반드시 SelectionMode를 먼저 확인해야
    한다. insert_text()는 선택된 구간이 있으면 그 구간을 덮어써서 교체하므로
    (직접 프로브로 실증), 두 경우 모두 최종 적용은 insert_text() 하나로 통일된다.

    LLM이 빈 응답을 준 경우(작은 로컬 모델에서 드물지 않게 발생) applied=False를
    반환하고 문서를 건드리지 않는다.
    """
    if report.hwp.SelectionMode != 0:
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

    작은 로컬 모델(qwen3.5:2b)은 드물지 않게 빈 응답을 내놓을 수 있다 —
    이는 이 함수 자체의 버그가 아니라 폴리시 대상 모델의 알려진 특성이므로
    (Fix 2로 이제 그런 경우 applied=False가 명확히 돌아온다), 이 테스트는
    빈 응답 한 번으로 바로 실패 처리하지 않고 최대 3회까지 재시도해서
    "실제 생성이 성공했을 때 정말로 문서에 삽입되는지"를 검증한다. 3회
    모두 빈 응답이면 그때는 명확한 메시지로 실패시킨다(모호한 AttributeError
    대신)."""
    from pyhwpx import Hwp
    test_path = os.path.join(tempfile.gettempdir(), "_test_공문서_삽입.hwp")
    setup = Hwp(visible=False, new=True)
    setup.save_as(test_path)  # 빈 문서
    setup.quit()

    report = None
    try:
        report = HwpReport(test_path)
        assert report.hwp.SelectionMode == 0, "빈 문서인데 선택 상태로 시작됨"

        max_attempts = 3
        result = None
        for attempt in range(1, max_attempts + 1):
            result = polish_to_formal_style(
                report, chat_input="이번 달 예산 다 썼고 다음달에 더 필요하다고 좀 써줘"
            )
            if result["applied"] is True:
                break
            print(f"  (시도 {attempt}/{max_attempts}) LLM이 빈 응답을 줘서 재시도함")
        else:
            raise AssertionError(
                f"작은 로컬 모델이 {max_attempts}회 연속 빈 응답을 반환함 "
                "(polish_to_formal_style 로직 문제가 아니라 모델 자체의 불안정성으로 보임)"
            )

        assert result["applied"] is True, result

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
