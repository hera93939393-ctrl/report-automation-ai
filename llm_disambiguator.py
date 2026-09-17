# -*- coding: utf-8 -*-
"""llm_disambiguator.py — categorize_values()가 내린 결정론적 판정 중
"값은 원본 어딘가에 분명히 있는데 엉뚱한 행에 연결된" 건만 골라내
로컬 LLM에게 의미로 다시 물어보는 보정 단계(F14 후속).

왜 별도 파일인가
----------------
verify_numbers.py는 docstring에 "외부 의존성 없는 순수 규칙 기반 모듈이라
LLM 의미 비교를 쓰지 않는다"고 스스로 못박아 둔 파일이다. 그 성질(테스트가
모델 없이 언제나 같은 결과를 내는 것)은 이 프로젝트에서 실제로 값을 해왔기
때문에 깨지 않는다 — 그래서 LLM 호출은 전부 이 파일에 모으고,
categorize_values()의 판정 로직 자체는 한 글자도 건드리지 않는다. 이 단계는
categorize_values()가 네 갈래를 다 만든 "뒤에" 돌아가는, 꺼도 그만인
후처리다(run_verification(llm_reconsider=False)면 예전 동작 그대로).

무엇을 고치려고 만들었나 (사용자 실제 문서에서 재현)
----------------------------------------------------
보고서 문장:
    * 식품 및 계약관련 법률 위반업체 제재 : ('24) 83개소 → ('25) 101 (63.2%↑)
원본 엑셀(핵심추이 요약) 9행:
    행정처분 정보연계 제재업체 / 2024=83 / 2025=101
        비고: 전년대비 63.2%↑ (식품·계약 관련 법률 위반업체)

83·101은 9행에 글자 그대로 있는데도 결정론 판정은 "빨강(오류)"이었다.
_label_matches_context()는 라벨을 공백으로 쪼갠 조각의 부분문자열 일치로
행을 찾는데, 9행의 조각들("행정처분","정보연계","제재업체")은 보고서 문장의
"위반업체 제재"와 글자가 하나도 안 겹친다(사람에겐 같은 말, 글자로는 남남).
반대로 엉뚱한 행이 우연히 걸렸다 — 실측해보니 "평가시스템 실적" 시트의
합계 행 라벨이 딱 한 글자 "계"인데, 문장의 "계약관련"에 "계"가 들어 있어서
통짜 부분문자열 비교를 그냥 통과했다. 그래서 83은 그 행의 값(109/120/494)과
비교돼 "확신에 찬 오답(빨강)"이 됐고, 진짜 출처인 9행은 후보로 검토조차
되지 않았다.

즉 이건 "못 찾는" 문제가 아니라 "엉뚱한 걸 찾아놓고 확신하는" 문제라
더 위험하다. 글자 겹침으로는 원리적으로 못 푸는 부분(같은 뜻, 다른 단어)이라
의미 판단이 필요하고, 그 자리에만 LLM을 부른다.

호출 조건(공짜가 아니므로 아주 좁게)
------------------------------------
빨강(mismatches)·회색(ambiguous)으로 떨어진 값에 대해서만, 그리고 그 값과
"글자 그대로 똑같은 값을 가진 진짜 데이터 칸"이 원본에 실제로 있을 때만
부른다. 그런 대안이 하나도 없으면 LLM을 불러봐야 바꿀 수 있는 게 없으므로
아예 부르지 않는다(대부분의 값이 여기서 걸러진다 — 실제 문서 32건 중 3건만
LLM까지 갔다).

"진짜 데이터 칸"의 뜻은 is_named_data_cell() 참고 — 비고(메모) 칸의 문장
속에서 긁어낸 숫자는 대안으로 쳐주지 않는다. 이게 그냥 깐깐한 게 아니라
이 문서에서 실제로 중요했다: 같은 문장의 "63.2%↑"는 진짜 보고서 오류인데
(83→101은 +21.7%이지 63.2%가 아니다 — 63.2%는 다른 문장의 19→31 것이다),
원본 엑셀 9행 비고에 "전년대비 63.2%↑"라는 문구가 그대로 옮겨 적혀 있다.
그 비고를 대안으로 인정하면 LLM이 "9행 얘기가 맞다"고 옳게 골라주는 순간
63.2%가 파랑(정상)으로 바뀌어, 진짜 오류가 조용히 사라진다. 비고는 독립된
근거가 아니라 보고서를 옮겨 적은 말이라(이 엑셀은 A2에 "출처: …관리결과.hwp"
라고 스스로 밝히고 있다) 보고서가 보고서로 자기를 증명하는 꼴이 된다.
그래서 숫자 칸만 대안으로 인정한다 — 그 결과 63.2%는 LLM을 아예 거치지
않고 빨강 그대로 남는다(실측 확인).

두 번 묻는 이유
---------------
exaone3.5:7.8b로 실측해보니 후보를 나열해 고르게 하면 목록 순서에 흔들렸다
(정답을 맨 뒤에 두면 맞히고 맨 앞에 두면 "해당 없음"이라고 답하는 현상).
그래서 (1) 고르게 한 뒤 (2) 고른 항목 하나만 놓고 "이 문장이 이 항목
얘기가 맞냐"를 예/아니오로 다시 물어, 둘 다 통과할 때만 판정을 바꾼다.
(2)는 순서와 무관한 내용 질문이라 (1)의 순서 편향을 걸러낸다 — 실측에서
모호한 케이스 두 건을 (2)가 "아니오"로 정확히 잡아냈다.

실패하면 어떻게 되나
--------------------
Ollama가 꺼져 있든, 느려서 타임아웃이 나든, 모델이 엉뚱한 말을 지어내든,
번호가 목록 범위를 벗어나든 — 전부 "판정을 못 바꿨다"로만 끝난다. 원래
빨강이던 건 빨강 그대로, 회색이던 건 회색 그대로다. 표를 못 읽어도 검증은
끝까지 진행된다는 이 프로젝트의 원칙을 LLM 단계에도 그대로 적용한 것으로,
검증이 LLM 때문에 중단되거나 예외로 죽는 경로는 없다.
"""
import re

from ollama_client import ROUTING_MODEL, get_client
from verify_numbers import (_matching_context_text, _label_matches_context,
                            filter_candidates_by_context)

# 한 번의 검증에서 LLM까지 갈 수 있는 항목 수 상한. 이 도구는 CPU에서 도는
# 로컬 모델을 쓰고(실측: 이 PC에서 exaone3.5:7.8b가 100% CPU로 로드됨) 항목당
# 두 번 부르므로, 이상한 문서 하나가 검증을 몇십 분씩 붙들지 않도록 막아둔다.
# 넘치면 앞에서부터 이만큼만 다시 보고 나머지는 결정론 판정 그대로 둔다.
MAX_RECONSIDERED_ITEMS = 12

# LLM에게 보여줄 후보 개수 상한(결정론이 고른 것 + 값이 같은 대안).
MAX_SHORTLIST = 6
_MAX_DETERMINISTIC_SHOWN = 2
_MAX_ALTERNATIVES_SHOWN = 4

# 호출 한 번의 제한 시간(초). 2026-09-17부터 로컬이 아니라 홈서버(Tailscale
# 경유, qwen3.5:9b)를 부른다 — 서버가 막 켜져 모델을 새로 올려야 하면 1분
# 안팎 걸릴 수 있어(실측), 여유 있게 잡는다.
DEFAULT_TIMEOUT_SEC = 120

_SELECTION_HEAD = (
    "보고서 문장에 나온 숫자가 원본 자료의 어느 '항목'을 말하는 것인지 고르는 문제다.\n"
    "숫자 값이 맞는지는 묻지 않는다. 오직 '어느 항목에 대한 문장인가'만 판단한다.\n\n"
)
_SELECTION_RULES = (
    "규칙:\n"
    "- 문장이 말하는 주제와 같은 항목의 번호를 고른다.\n"
    "- 뜻이 통하는 항목이 하나도 없으면 0을 고른다.\n"
    "- 설명 없이 번호 숫자 하나만 출력한다.\n"
)
_CONFIRM_TEMPLATE = (
    "아래 보고서 문장이 원본 자료의 '{label}' 항목에 관한 문장인지 판단하시오.\n\n"
    "[보고서 문장] {context} {raw}\n"
    "[원본 항목] {label}\n\n"
    "같은 주제를 말하고 있으면 '예', 서로 다른 항목이면 '아니오'라고만 답하시오.\n"
    "설명 없이 예 또는 아니오만 출력한다.\n\n"
    "답:"
)


def is_named_data_cell(candidate: dict) -> bool:
    """이 후보가 "이름 붙은 진짜 데이터 칸"인가 — LLM에게 '이 행이 출처일지도
    모른다'고 제시할 자격이 있는지의 기준.

    두 가지를 요구한다.

    1) 숫자 칸에서 온 값일 것(from_text가 아닐 것). source_reader는 엑셀
       문자열 칸도 읽어서 그 문장 속 숫자를 정답 풀에 넣는데(비고/메모 칸이
       대표적), 그렇게 긁어낸 숫자는 독립된 근거가 아니라 보고서 문장을
       옮겨 적은 말인 경우가 많다. 그걸 근거로 빨강을 파랑으로 뒤집으면
       보고서가 보고서로 자기를 증명하게 된다(이 파일 맨 위 63.2% 사례).
    2) 항목명(label)이 있을 것. LLM에게 묻는 질문 자체가 "어느 항목이냐"라서
       이름이 없는 후보는 판단할 거리가 없다. 열 합계처럼 파생으로 만들어진
       항목(compute_column_sums)은 label이 아예 없어 여기서 자연히 빠진다.
    """
    if candidate.get("from_text"):
        return False
    return bool((candidate.get("label") or "").strip())


def narrowed_candidates(rv: dict, answer_pool: list) -> list:
    """categorize_values()가 이 값을 판정할 때 실제로 쓴 후보 풀.

    같은 타입으로 거르고, 값이 보고서 "표 칸"에서 나왔으면 그 칸의 행/열
    머리말로 한 번 더 좁힌다 — categorize_values()의 좁히기 순서를 그대로
    따라가며, 쓰는 함수도 verify_numbers의 바로 그 함수라 규칙이 몰래
    갈라질 수 없다.

    (이 좁히기를 빠뜨렸다가 기존 회귀 테스트
    _selftest_run_verification_table_cell_wrong_year_is_flagged에 잡혔다.
    표의 "'24년" 칸에 2025년 값을 잘못 적은 오류는, 행 이름은 멀쩡히 맞고
    연도 열만 틀린 경우다. 그런데 연도를 무시하고 "같은 값이 원본에 있다"만
    보면 2025년 열의 그 값이 대안으로 잡히고, LLM에게 "이 문장이 어느
    항목이냐"고 물으면 행 이름이 맞으니 당연히 그 행을 골라서 빨강이
    파랑으로 뒤집혀 버린다 — 이 브랜치가 만든 표 맥락 기능을 정면으로
    무력화하는 셈이다. LLM에게는 "어느 항목이냐"만 묻고 "어느 연도냐"는
    묻지 않으므로, 연도 판정은 지금까지처럼 전적으로 결정론 규칙
    (column_identity_agrees)에 맡기고 그 결과를 존중해야 한다.)
    """
    candidates = [a for a in answer_pool if a.get("type") == rv.get("type")]
    table_context = rv.get("table_context")
    if table_context:
        return filter_candidates_by_context(table_context, candidates)
    return candidates


def find_value_matching_alternatives(rv: dict, answer_pool: list) -> list:
    """보고서 값과 타입·값이 똑같은 "이름 붙은 진짜 데이터 칸" 후보들.

    이게 비어 있으면 LLM을 불러도 판정을 바꿀 재료가 없다는 뜻이라 호출을
    건너뛴다(대부분의 값이 여기서 걸러진다)."""
    return [c for c in narrowed_candidates(rv, answer_pool)
            if c.get("normalized") == rv.get("normalized")
            and is_named_data_cell(c)]


def deterministic_candidates(rv: dict, answer_pool: list) -> list:
    """categorize_values()가 이 값에 대해 "라벨이 문맥에 걸린다"고 본 후보들.

    판정을 다시 계산하는 게 아니라 "결정론이 무엇을 보고 그렇게 판단했는지"를
    LLM에게 같이 보여주려고 읽기만 한다."""
    context = _matching_context_text(rv)
    return [c for c in narrowed_candidates(rv, answer_pool)
            if _label_matches_context(c.get("label"), context)]


def _dedupe_by_label(candidates: list) -> list:
    """항목명이 같은 후보는 하나만 남긴다. 실측에서 똑같은 라벨("계")이 세 번
    나열되자 모델이 눈에 띄게 흔들렸다 — 어차피 우리가 묻는 건 "어느 항목
    이냐"라서 같은 이름이 여러 번 나올 이유가 없다."""
    seen, result = set(), []
    for c in candidates:
        label = (c.get("label") or "").strip()
        if label in seen:
            continue
        seen.add(label)
        result.append(c)
    return result


def build_shortlist(rv: dict, answer_pool: list) -> list:
    """LLM에게 보여줄 후보 목록. 결정론이 고른 것을 앞에, 값이 같은 대안을
    뒤에 둔다.

    순서를 이렇게 고정한 데는 실측 근거가 있다 — 이 모델은 목록의 뒤쪽
    항목을 고르는 쪽으로 기울었고(정답을 맨 앞에 두면 "해당 없음"이라고
    답해버리는 현상을 재현), 우리가 새로 알려주려는 정보는 "값이 같은 다른
    행이 있다"는 쪽이다. 이 편향에 기대는 배치라는 점이 께름칙해서, 고른
    결과는 순서와 무관한 예/아니오 확인 질문으로 한 번 더 검증한다."""
    alternatives = _dedupe_by_label(find_value_matching_alternatives(rv, answer_pool))
    chosen = _dedupe_by_label(deterministic_candidates(rv, answer_pool))
    alternative_labels = {(c.get("label") or "").strip() for c in alternatives}
    chosen = [c for c in chosen if (c.get("label") or "").strip() not in alternative_labels]
    shortlist = (chosen[:_MAX_DETERMINISTIC_SHOWN]
                 + alternatives[:_MAX_ALTERNATIVES_SHOWN])
    return shortlist[:MAX_SHORTLIST]


def build_selection_prompt(context: str, raw: str, shortlist: list) -> str:
    lines = [f"{i}. {(c.get('label') or '').strip()}" for i, c in enumerate(shortlist, 1)]
    return (_SELECTION_HEAD
            + f"[보고서 문장] {context} {raw}\n\n"
            + "[원본 항목 목록]\n" + "\n".join(lines) + "\n0. 해당하는 항목 없음\n\n"
            + _SELECTION_RULES + "\n번호:")


def build_confirm_prompt(context: str, raw: str, label: str) -> str:
    return _CONFIRM_TEMPLATE.format(label=label, context=context, raw=raw)


def parse_index(response: str, option_count: int):
    """모델 답에서 번호를 뽑는다. 0은 "해당 없음", None은 "못 알아들었다".

    자유 텍스트라 방어적으로 읽는다 — 첫 번째 정수만 보고, 목록 범위를
    벗어나면(모델이 없는 번호를 지어내면) None으로 버린다. 판정을 바꾸는
    쪽으로는 절대 추측하지 않는다."""
    if not response:
        return None
    match = re.search(r"\d+", response)
    if not match:
        return None
    value = int(match.group())
    if value < 0 or value > option_count:
        return None
    return value


def parse_yes(response: str):
    """'예'/'아니오' 답을 읽는다. True/False, 못 알아들으면 None.

    "아니오"가 "예"를 부분문자열로 갖고 있지 않아 순서 함정은 없지만,
    모델이 "예."·"예, 같은 항목입니다"처럼 덧붙이는 경우가 있어 앞부분만 본다."""
    if not response:
        return None
    text = response.strip()
    if text.startswith("아니"):
        return False
    if text.startswith("예") or text.startswith("네"):
        return True
    if "아니오" in text or "아닙니다" in text:
        return False
    if "예" in text[:10]:
        return True
    return None


def ask_ollama(prompt: str, timeout: float = DEFAULT_TIMEOUT_SEC):
    """홈서버의 qwen3.5:9b에 평문 완성으로 한 번 묻는다(도구호출 아님 —
    번호/예-아니오 하나만 받으면 되므로 도구호출 스키마 자체가 필요 없다).

    think=False로 사고 과정(thinking)을 꺼서 곧바로 짧은 답만 받는다 —
    qwen3.5:9b는 기본이 사고형 모델이라, 이걸 안 끄면 아래처럼 num_predict를
    짧게 잡았을 때 "생각하는 중"에 예산을 다 쓰고 정작 답(response)은 빈
    문자열로 끝나버리는 게 실측 확인됐다.

    어떤 이유로든 실패하면 예외를 올리지 않고 None을 돌려준다 — 호출부는
    None을 "판정을 못 바꿨다"로만 취급하므로, 서버가 꺼져 있거나 응답이
    늦어도 숫자검증 자체는 예전과 똑같이 끝까지 진행된다."""
    try:
        response = get_client(timeout=timeout).generate(
            model=ROUTING_MODEL,
            prompt=prompt,
            # 같은 문서를 두 번 검증하면 같은 답이 나와야 하므로 온도 0.
            # 번호 하나만 받으면 되니 생성 길이도 짧게 잡아 추론 시간을 아낀다.
            options={"temperature": 0, "num_predict": 8},
            # 한 번의 검증에서 여러 항목을 연달아 묻게 되므로, 그 사이에
            # 모델이 서버 메모리에서 내려가 매번 재로딩되지 않도록 유지시간을
            # 넉넉히 준다.
            keep_alive="10m",
            think=False,
        )
        return response.response or ""
    except Exception:
        return None


def reconsider_item(rv: dict, answer_pool: list, ask=ask_ollama,
                    timeout: float = DEFAULT_TIMEOUT_SEC):
    """값 하나를 LLM에게 다시 물어본다.

    반환: (판정, 고른 후보) — 판정은 "match"/"mismatch"/"ambiguous" 중 하나.
    고른 후보는 None일 수 있다(모델이 못 고름).
    """
    context = _matching_context_text(rv)
    if not context.strip():
        # 값 앞에 문맥이 한 글자도 없으면 의미로 판단할 거리가 없다 —
        # 물어봐야 찍는 답만 돌아오므로 부르지 않는다.
        return None, None
    shortlist = build_shortlist(rv, answer_pool)
    if not shortlist:
        return None, None

    raw = str(rv.get("raw") or rv.get("normalized") or "")
    answer = ask(build_selection_prompt(context, raw, shortlist), timeout)
    index = parse_index(answer, len(shortlist))
    if not index:
        # 0(해당 없음)이거나 못 알아들은 응답 — 어느 쪽이든 특정 실패다.
        return "ambiguous", None

    picked = shortlist[index - 1]
    label = (picked.get("label") or "").strip()
    confirm = ask(build_confirm_prompt(context, raw, label), timeout)
    if parse_yes(confirm) is not True:
        # 고르긴 했는데 내용 확인에서 통과 못 했다 — 순서에 흔들린 답일
        # 가능성이 높으므로 "확인 필요"로 넘긴다.
        return "ambiguous", None

    verdict = "match" if picked.get("normalized") == rv.get("normalized") else "mismatch"
    return verdict, picked


def reconsider(categorized: dict, answer_pool: list, skip_raws=frozenset(),
               ask=ask_ollama, timeout: float = DEFAULT_TIMEOUT_SEC,
               max_items: int = MAX_RECONSIDERED_ITEMS) -> dict:
    """categorize_values()의 결과를 받아, 빨강·회색 중 다시 볼 값만 골라
    LLM에게 물어본 뒤 갈래를 재배치한 새 dict를 돌려준다.

    원본 dict와 값(rv) 객체는 그대로 두고 갈래만 새로 만든다 — 다만 판정이
    LLM으로 바뀐 값에는 출처를 남긴다(resolved_by/resolved_location/
    resolved_label). 이 프로젝트에는 항목별 출처를 남기는 기존 관례가 없어
    새로 만든 것이라, 이름을 눈에 띄게 하고 키를 세 개로만 제한했다.

    반환 dict에는 네 갈래 외에 llm_calls(실제 호출 횟수)와 resolutions(바뀐
    항목 설명)를 같이 담는다 — 호출부가 요약에 쓰고, 테스트가 "안 불렀어야
    할 때 정말 안 불렀는지"를 확인하는 데 쓴다."""
    matches = list(categorized["matches"])
    unverifiable = list(categorized["unverifiable"])
    mismatches, ambiguous = [], []
    calls = {"n": 0}
    resolutions = []

    def counting_ask(prompt, call_timeout):
        calls["n"] += 1
        return ask(prompt, call_timeout)

    reconsidered = 0
    for bucket_name, bucket in (("mismatch", categorized["mismatches"]),
                                ("ambiguous", categorized["ambiguous"])):
        for rv in bucket:
            target = mismatches if bucket_name == "mismatch" else ambiguous
            # 이 값과 똑같은 값을 가진 진짜 데이터 칸이 원본에 없으면 LLM이
            # 바꿀 수 있는 게 없다 — 부르지 않고 결정론 판정 그대로 둔다.
            # (정상 문서의 거의 모든 값이 여기서 끝난다.)
            if (str(rv.get("raw")) in skip_raws
                    or not find_value_matching_alternatives(rv, answer_pool)
                    or reconsidered >= max_items):
                target.append(rv)
                continue
            reconsidered += 1
            verdict, picked = reconsider_item(rv, answer_pool, ask=counting_ask,
                                              timeout=timeout)
            if verdict is None:
                target.append(rv)
                continue
            if verdict == "ambiguous":
                # 판정을 못 바꿨다. 빨강이었던 값도 회색으로 내린다 — 값이
                # 같은 다른 행이 분명히 있는데 어느 쪽인지 못 가린 상태에서
                # "틀렸다"고 단정하는 건 이 문서에서 실제로 났던 사고(확신에
                # 찬 오답)와 같은 종류다. 이 프로젝트의 정직한 기본값은
                # "확인 필요"다.
                ambiguous.append(rv)
                continue
            rv["resolved_by"] = "llm"
            rv["resolved_location"] = picked.get("location")
            rv["resolved_label"] = (picked.get("label") or "").strip()
            (matches if verdict == "match" else mismatches).append(rv)
            resolutions.append({
                "raw": rv.get("raw"), "from": bucket_name, "to": verdict,
                "location": picked.get("location"),
                "label": (picked.get("label") or "").strip(),
            })

    return {"matches": matches, "mismatches": mismatches,
            "unverifiable": unverifiable, "ambiguous": ambiguous,
            "llm_calls": calls["n"], "resolutions": resolutions}
