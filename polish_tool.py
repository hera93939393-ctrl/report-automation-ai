"""polish_tool.py — 선택된 문장을 공문서체로 다듬거나, 채팅 입력을 새 공문서체
문장으로 만들어 커서 위치에 삽입하는 도구. verify_tool.py와 나란한 두 번째 도구.

(2026-10-02 T5, 부품 6 재포장) PRD 16-2의 부품 6 스펙("편람 규칙 + 유사 문단
예시 동봉")에 맞게 프롬프트를 확장했다: ① 개조식 규칙(_STYLE_RULES — 레거시
ai_writer.INSTRUCTION의 실전 규칙을 흡수) ② 과거 보고서 아카이브에서 찾은
유사 문단을 문체 예시로 동봉(_similar_archive_examples, 로컬 BM25라 서버
불필요) ③ 루프/호출자가 주는 힌트(hint) 전달. 프롬프트가 길어져 num_ctx를
명시한다(attachment_qa_tool과 같은 이유). 예시 동봉의 품질(소형 모델이
예시를 베끼지 않는지)은 서버를 켠 뒤 실측 확인 대상."""
import os
import tempfile

from hwp_report import HwpReport
from ollama_client import GENERATION_MODEL, get_client
from speed_tracker import record_call_speed

# 공문서 개조식 핵심 규칙 — 레거시 ai_writer.INSTRUCTION(실사용 검증된 F1
# 프롬프트)에서 흡수. 편람 세부 조항 전체가 아니라 소형 모델이 지킬 수 있는
# 핵심만 추린다(컨텍스트 예산·지시 과다 방지).
_STYLE_RULES = (
    "공문서 개조식 규칙: "
    "① 완전한 문장 종결어미(-습니다, -합니다)를 쓰지 말고 명사형이나 '-함', '-음', '-예정'으로 끝낼 것 "
    "② 하나의 목적/이유에 속한 짧은 요소들은 항목을 나누지 말고 '및'으로 이어 한 줄로 쓸 것 "
    "③ 일정/장소/대상처럼 서로 다른 항목일 때만 '- '로 줄을 나눌 것 "
    "④ 마크다운·특수기호를 쓰지 말 것 "
    "⑤ 원문에 없는 사실·수치를 지어내지 말 것"
)


def _similar_archive_examples(text: str, k: int = 2) -> list[str]:
    """과거 보고서 아카이브(로컬 BM25 색인)에서 비슷한 문단을 찾아 문체 예시로
    쓴다 — 부품 6 스펙의 "유사 문단 예시 동봉"(T5). 색인을 새로 만들지는
    않는다(reindex는 파일당 수 초라 다듬기 도구 안에서 돌리면 안 됨) — 이미
    만들어진 DB만 검색하고, 폴더 미등록·DB 없음·오류는 전부 빈 리스트
    (예시 없이도 다듬기는 그대로 동작한다)."""
    try:
        from archive_index import DB_PATH, search_archive
        if not os.path.exists(DB_PATH):
            return []
        return [r["text"][:300] for r in search_archive(text, k=k)]
    except Exception:  # noqa: BLE001 — 예시는 부가 기능, 실패가 다듬기를 막지 않게
        return []


def _build_polish_prompt(source_text: str, hint: str = "", examples: list = None) -> str:
    """규칙 + (있으면) 유사 문단 예시 + (있으면) 힌트 + 원문으로 프롬프트를
    조립한다. 예시에는 "문체만 참고" 경고를 반드시 붙인다 — 소형 모델이 예시
    문장을 그대로 베끼는 현상을 막기 위한 지시(실측 확인은 서버 켠 뒤)."""
    parts = [_STYLE_RULES, ""]
    if examples:
        parts.append("[문체 예시 — 과거 보고서에서 가져옴. 말투와 형식만 참고하고, 예시의 내용·수치·기관명은 절대 결과에 옮기지 말 것]")
        parts += [f"예시 {i}: {ex}" for i, ex in enumerate(examples, start=1)]
        parts.append("")
    if hint:
        parts += [f"[요청 사항]\n{hint}", ""]
    parts += ["다음 문장을 위 규칙에 맞는 공문서 개조식 문체로 다듬어줘. 다듬은 문장만 출력하고 다른 설명은 붙이지 마:", "", source_text]
    return "\n".join(parts)


def _generate_formal_style(source_text: str, max_attempts: int = 3, client=None,
                           hint: str = "", examples: list = None) -> str | None:
    """source_text(선택된 원문 또는 채팅 입력)를 홈서버 LLM에 보내 공문서체
    문장으로 바꾼 결과를 반환한다. 서버에 연결할 수 없으면 None을 반환한다
    (빈 응답 ""과 구분 — 호출자가 "서버 꺼짐"과 "모델이 빈 답"을 다르게
    안내할 수 있어야 함, routing_graph._classify_via_llm과 같은 관례).

    (2026-10-01, 서버 전환) 원래 노트북 로컬 Ollama의 qwen3.5:2b를
    ollama.chat()으로 직접 불렀는데, 노트북에서 Ollama가 제거된 뒤(2026-09-17
    홈서버 전환) 이 도구만 옛 경로에 남아 실사용에서 실패하던 것을 고침 —
    이제 ollama_client.get_client()로 서버의 GENERATION_MODEL(qwen3.5:9b)을
    쓴다. qwen3.5:9b는 사고형 모델이라 think=False를 주지 않으면 생각만 하다
    빈 답으로 끝나는 게 실측 확인된 특성(llm_disambiguator 참고)이므로 끈다.
    client 인자는 테스트에서 가짜 클라이언트를 끼우기 위한 것(기본 None이면
    get_client()). route_intent()와 달리 도구 호출(tools=)이
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
    if client is None:
        client = get_client()
    prompt = _build_polish_prompt(source_text, hint=hint, examples=examples)
    for attempt in range(1, max_attempts + 1):
        try:
            response = client.chat(
                model=GENERATION_MODEL,
                messages=[{"role": "user", "content": prompt}],
                think=False,
                # 규칙+예시로 프롬프트가 길어질 수 있어 명시(attachment_qa_tool과
                # 같은 이유 — Ollama 기본 num_ctx가 더 작으면 앞부분이 잘린다).
                options={"num_ctx": 8192},
            )
        except Exception:
            # 서버 꺼짐/네트워크 단절 — 재시도해도 같은 결과라 바로 알린다.
            return None
        # (F13) 이번 호출의 실제 처리속도를 기록해, 다음 호출 전 예상
        # 소요시간을 계산할 수 있게 한다(eval_count/eval_duration이 0
        # 이하인 비정상 응답은 record_call_speed가 자체적으로 걸러낸다).
        record_call_speed(response.get("eval_count", 0), response.get("eval_duration", 0))
        polished = response["message"]["content"].strip()
        if polished:
            return polished
    return ""


def polish_to_formal_style(report: HwpReport, chat_input: str, client=None) -> dict:
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

    # (T5) 아카이브 색인이 있으면 유사 문단을 문체 예시로 동봉(없으면 빈 리스트)
    polished = _generate_formal_style(source_text, client=client,
                                      examples=_similar_archive_examples(source_text))
    if polished is None:
        return {"applied": False, "polished_text": "", "error": "server_unreachable"}
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


class _FakeClient:
    """서버 없이 _generate_formal_style의 재시도/실패 경로를 검증하기 위한
    가짜 Ollama 클라이언트. replies의 항목을 차례로 돌려주고, 항목이
    Exception이면 그걸 던진다(서버 꺼짐 흉내)."""

    def __init__(self, replies):
        self.replies = list(replies)
        self.calls = []

    def chat(self, **kwargs):
        self.calls.append(kwargs)
        reply = self.replies.pop(0)
        if isinstance(reply, Exception):
            raise reply
        return {"message": {"content": reply}, "eval_count": 10, "eval_duration": 10**9}


class _FakeReportNoSelection:
    """선택 없는 문서 흉내 — polish_to_formal_style이 LLM 결과를 어떻게
    보고하는지만 볼 때 실제 한글(COM)을 띄우지 않기 위한 것."""

    class _Hwp:
        SelectionMode = 0

        def insert_text(self, text):
            self.inserted = text
            return True

    def __init__(self):
        self.hwp = self._Hwp()


def _selftest_generate_retries_empty_then_succeeds():
    """빈 응답 두 번 뒤 정상 응답이 오면 그 결과를 쓰고, 서버 모델/think=False가
    실제로 전달되는지 확인한다(서버 불필요)."""
    fake = _FakeClient(["", "  ", "예산 집행 완료 및 추가 예산 필요"])
    result = _generate_formal_style("예산 다 썼음", client=fake)
    assert result == "예산 집행 완료 및 추가 예산 필요", result
    assert len(fake.calls) == 3, len(fake.calls)
    assert fake.calls[0]["model"] == GENERATION_MODEL, fake.calls[0]["model"]
    assert fake.calls[0]["think"] is False, fake.calls[0]
    print("_generate_formal_style(재시도 후 성공) 통과")


def _selftest_generate_returns_none_when_server_unreachable():
    """연결 예외가 나면 재시도하지 않고 None을 돌려주는지(빈 문자열과 구분)."""
    fake = _FakeClient([ConnectionError("서버 꺼짐")])
    result = _generate_formal_style("아무 문장", client=fake)
    assert result is None, result
    assert len(fake.calls) == 1, len(fake.calls)
    print("_generate_formal_style(서버 미접속 → None) 통과")


def _selftest_polish_reports_server_unreachable_without_touching_document():
    """서버가 꺼져 있으면 문서를 건드리지 않고 error=server_unreachable로 알리는지."""
    fake = _FakeClient([ConnectionError("서버 꺼짐")])
    report = _FakeReportNoSelection()
    result = polish_to_formal_style(report, "써줘", client=fake)
    assert result == {"applied": False, "polished_text": "", "error": "server_unreachable"}, result
    assert not hasattr(report.hwp, "inserted"), "서버 미접속인데 문서에 삽입함"
    print("polish_to_formal_style(서버 미접속 보고) 통과")


def _selftest_polish_all_empty_reports_failure_without_error_key():
    """3회 모두 빈 응답이면 applied=False이되 error 키는 없어야 한다
    (chat_assistant가 "빈 응답"과 "서버 꺼짐"을 다른 문구로 안내함)."""
    fake = _FakeClient(["", "", ""])
    report = _FakeReportNoSelection()
    result = polish_to_formal_style(report, "써줘", client=fake)
    assert result == {"applied": False, "polished_text": ""}, result
    print("polish_to_formal_style(3회 빈 응답) 통과")


def _selftest_build_polish_prompt():
    """(T5) 프롬프트 조립: 규칙은 항상, 예시는 "문체만 참고" 경고와 함께,
    힌트는 요청 사항으로 들어가는지. 예시·힌트가 없으면 그 구획 자체가 없는지."""
    p = _build_polish_prompt("예산 다 썼음")
    assert _STYLE_RULES in p and "예산 다 썼음" in p, p
    assert "문체 예시" not in p and "요청 사항" not in p, p
    p = _build_polish_prompt("예산 다 썼음", hint="한 문장으로", examples=["○ 예산 집행 완료 및 추가 확보 추진"])
    assert "예시 1: ○ 예산 집행 완료" in p and "절대 결과에 옮기지 말 것" in p, p
    assert "[요청 사항]\n한 문장으로" in p, p
    assert p.index(_STYLE_RULES) < p.index("예시 1") < p.index("예산 다 썼음"), "순서: 규칙 → 예시 → 원문"
    print("_build_polish_prompt(규칙·예시 경고·힌트) 통과")


def _selftest_generate_passes_prompt_and_num_ctx():
    """(T5) 조립된 프롬프트와 num_ctx가 실제 chat 호출에 들어가는지."""
    fake = _FakeClient(["다듬은 결과"])
    result = _generate_formal_style("원문", client=fake, hint="짧게", examples=["예시문"])
    assert result == "다듬은 결과", result
    sent = fake.calls[0]["messages"][0]["content"]
    assert _STYLE_RULES in sent and "예시 1: 예시문" in sent and "[요청 사항]\n짧게" in sent, sent
    assert fake.calls[0]["options"]["num_ctx"] == 8192, fake.calls[0]
    print("_generate_formal_style(프롬프트 조립·num_ctx 전달) 통과")


def _selftest_similar_archive_examples_safe_without_db():
    """(T5) 색인 DB가 없으면 빈 리스트 — 예시 기능이 다듬기를 막지 않는다.
    _similar_archive_examples는 DB_PATH를 호출 시점에 from-import로 읽으므로
    모듈 속성을 잠시 없는 경로로 바꿔 "DB 없음" 분기를 확정적으로 태운다."""
    import archive_index
    saved = archive_index.DB_PATH
    try:
        archive_index.DB_PATH = os.path.join(tempfile.gettempdir(), "_없는_색인.db")
        assert _similar_archive_examples("아무 문장") == []
    finally:
        archive_index.DB_PATH = saved
    print("_similar_archive_examples(DB 없음 안전) 통과")


if __name__ == "__main__":
    # 서버·한글 없이 도는 것부터
    _selftest_build_polish_prompt()
    _selftest_generate_passes_prompt_and_num_ctx()
    _selftest_similar_archive_examples_safe_without_db()
    _selftest_generate_retries_empty_then_succeeds()
    _selftest_generate_returns_none_when_server_unreachable()
    _selftest_polish_reports_server_unreachable_without_touching_document()
    _selftest_polish_all_empty_reports_failure_without_error_key()
    # 아래 둘은 홈서버가 켜져 있고 한글 COM을 쓸 수 있을 때만(다른 한글 창이
    # 떠 있으면 실행하지 말 것 — hwp-report-tool 메모의 tasklist 확인 규칙)
    _selftest_polish_replaces_selected_text()
    _selftest_polish_inserts_new_text_when_nothing_selected()
