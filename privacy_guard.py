"""privacy_guard.py — 개인정보로 보이는 패턴(주민등록번호/휴대폰번호/
계좌번호)을 감지하는 순수 함수. 외부 의존성 없음(re만 사용).

이 모듈은 개인정보를 "확실히 판별"하지 않는다 — 정규식 패턴 매칭이므로
숫자 형태가 비슷하면 실제 개인정보가 아니어도 감지될 수 있다(오탐 가능).
그래서 chat_assistant.py 쪽 연동은 감지되면 곧바로 차단하지 않고 사용자에게
"계속 진행할지"만 확인한다.

(2026-10-02 T0, PRD 16-5 ⑥ 준비) 쓰임새가 둘로 늘었다:
1. 채팅 입력 확인(기존) — detect_pii_patterns()가 유형 이름만 돌려준다.
2. 서버 전송 전 마스킹(신규) — ⑥의 임베딩·사업 카드 생성은 문서 전문이
   홈서버 LLM 프롬프트로 나가므로, 나가기 전에 mask_pii()로 가린다(원문이
   서버로 가기 전이 유일한 안전 지점). 위치가 필요해 find_pii_spans()가
   매치 구간(span)을 돌려준다.
정규식의 한계(이름·이메일·주소는 못 잡음)는 그대로다 — 패턴 기반 최소
방어이며, 민감 문서를 색인 폴더에서 빼는 책임은 사용자에게 있다."""
import re

# 주민등록번호: 생년월일 6자리 + 하이픈 + 성별코드(1~4) + 나머지 6자리.
# 성별코드를 1~4로 제한해 일반 전화/계좌번호와 우연히 겹칠 가능성을 줄인다.
_RRN_PATTERN = re.compile(r'\d{6}-[1-4]\d{6}')

# 휴대폰번호: 010/011/016/017/018/019로 시작하는 표준 이동전화 형식.
# 하이픈은 있어도 없어도 매치되게 해서 "01012345678" 같은 붙여쓰기도 잡는다.
_MOBILE_PATTERN = re.compile(r'01[016789]-?\d{3,4}-?\d{4}')

# 지역번호가 있는 유선전화(예: 031-1234-5678) — verify_numbers.py의
# _PHONE_PATTERN과 동일한 모양. 이 모양은 "전화번호"이지 "계좌번호"가
# 아니므로, 아래 계좌번호 패턴이 이 구간과 겹치면 계좌번호로 오인하지
# 않도록 제외구간으로 쓴다.
_LANDLINE_PATTERN = re.compile(r'\d{2,3}-\d{3,4}-\d{4}')

# 계좌번호로 보이는 패턴: 은행마다 자릿수가 제각각이라 완벽한 정규식은
# 없다 — 하이픈으로 구분된 세 자리 숫자 그룹이라는 느슨한 모양만 잡는다.
_ACCOUNT_PATTERN = re.compile(r'\d{2,6}-\d{2,6}-\d{2,8}')

# "2026-09-07"류 날짜 — 계좌번호 패턴(하이픈 3-그룹)과 모양이 겹쳐 그동안
# "계좌번호"로 오탐되던 알려진 한계였다(기존 docstring에 기록). 채팅 한 줄이면
# 확인 한 번으로 끝나지만, ⑥의 문서 전문 처리에서는 날짜가 수십 개라 오탐이
# 폭주하므로 제외구간으로 승격(2026-10-02 T0). 월 1~12, 일 1~31만 날짜로 본다
# — "123-45-67890" 같은 진짜 계좌 모양은 월/일 범위를 벗어나 계속 잡힌다.
_DATE_PATTERN = re.compile(r'(?<!\d)\d{4}-(?:0?[1-9]|1[0-2])-(?:0?[1-9]|[12]\d|3[01])(?!\d)')

_TYPE_ORDER = ["주민등록번호", "휴대폰번호", "계좌번호"]


def find_pii_spans(text: str) -> list[dict]:
    """개인정보로 보이는 매치를 위치와 함께 돌려준다(start 순 정렬).
    반환: [{"type": 유형, "start": int, "end": int, "text": 매치 원문}]

    계좌번호 판정의 제외구간: 휴대폰·유선전화(전화는 전화지 계좌가 아님)와
    날짜(_DATE_PATTERN 주석 참고). 겹치면 계좌번호로 세지 않는다."""
    spans = []
    for m in _RRN_PATTERN.finditer(text):
        spans.append({"type": "주민등록번호", "start": m.start(), "end": m.end(), "text": m.group()})
    for m in _MOBILE_PATTERN.finditer(text):
        spans.append({"type": "휴대폰번호", "start": m.start(), "end": m.end(), "text": m.group()})

    excluded_spans = (
        [m.span() for m in _MOBILE_PATTERN.finditer(text)]
        + [m.span() for m in _LANDLINE_PATTERN.finditer(text)]
        + [m.span() for m in _DATE_PATTERN.finditer(text)]
    )

    def _overlaps_excluded(span):
        a_start, a_end = span
        return any(a_start < e and s < a_end for s, e in excluded_spans)

    for m in _ACCOUNT_PATTERN.finditer(text):
        if not _overlaps_excluded(m.span()):
            spans.append({"type": "계좌번호", "start": m.start(), "end": m.end(), "text": m.group()})

    spans.sort(key=lambda d: (d["start"], d["end"]))
    return spans


def detect_pii_patterns(text: str) -> list[str]:
    """text에서 개인정보로 보이는 패턴을 감지해 발견된 패턴 유형 이름
    리스트를 반환한다(주민등록번호 → 휴대폰번호 → 계좌번호 순, 유형당 한 번).
    아무 것도 감지 안 되면 빈 리스트. 위치까지 필요하면 find_pii_spans()를,
    가린 텍스트가 필요하면 mask_pii()를 쓴다."""
    types = {d["type"] for d in find_pii_spans(text)}
    return [t for t in _TYPE_ORDER if t in types]


def mask_pii(text: str) -> tuple[str, list[str]]:
    """개인정보로 보이는 구간을 "[유형 삭제]"로 치환한 텍스트와, 발견된 유형
    목록을 돌려준다. ⑥ 서버 색인·사업 카드 생성처럼 문서 전문이 홈서버 LLM
    프롬프트로 나가기 직전 단계에서 쓴다(모듈 docstring). 감지된 게 없으면
    원문 그대로 반환. 겹치는 매치는 앞선 것만 가린다(이중 치환 방지)."""
    spans = find_pii_spans(text)
    if not spans:
        return text, []
    pieces, last, types = [], 0, []
    for d in spans:
        if d["start"] < last:  # 앞 매치와 겹침 — 이미 가려진 구간
            continue
        pieces.append(text[last:d["start"]])
        pieces.append(f"[{d['type']} 삭제]")
        last = d["end"]
        if d["type"] not in types:
            types.append(d["type"])
    pieces.append(text[last:])
    return "".join(pieces), [t for t in _TYPE_ORDER if t in types]


def _selftest_detect_pii_patterns_rrn():
    result = detect_pii_patterns("담당자 주민번호는 900101-1234567 입니다")
    assert result == ["주민등록번호"], result
    print("detect_pii_patterns 통과(주민등록번호):", result)


def _selftest_detect_pii_patterns_mobile():
    result = detect_pii_patterns("연락처는 010-1234-5678로 부탁드립니다")
    assert result == ["휴대폰번호"], result
    print("detect_pii_patterns 통과(휴대폰번호):", result)


def _selftest_detect_pii_patterns_no_false_positive():
    """일반 문장, 그리고 이 프로젝트 다른 self-test(verify_numbers.py)가
    이미 쓰고 있는 유선전화 형식("031-1234-5678")도 아무 패턴이 감지되면
    안 된다 - 특히 유선전화가 계좌번호로 오탐되지 않는지가 핵심(계좌번호
    패턴이 하이픈 3-그룹 숫자라는 느슨한 모양이라, 지역번호 유선전화와
    모양이 겹치기 쉽다)."""
    assert detect_pii_patterns("숫자 검증해줘") == []
    result = detect_pii_patterns("담당자: 김주무관 (031-1234-5678), 예산 1,850,000원")
    assert result == [], result
    print("detect_pii_patterns 통과(오탐 없음, 유선전화 포함):", result)


def _selftest_date_not_account():
    """(T0) 기존의 알려진 한계였던 날짜("2026-09-07") 계좌 오탐이 해소됐는지,
    그리고 월/일 범위를 벗어나는 진짜 계좌 모양은 계속 잡히는지."""
    assert detect_pii_patterns("회의는 2026-09-07에 열렸고 2026-10-1로 연기됨") == []
    result = detect_pii_patterns("입금계좌 123-45-67890 로 보내주세요")
    assert result == ["계좌번호"], result
    print("detect_pii_patterns 통과(날짜 제외·계좌 유지)")


def _selftest_find_pii_spans_positions():
    text = "담당 김주무관 010-1234-5678, 주민번호 900101-1234567"
    spans = find_pii_spans(text)
    assert [d["type"] for d in spans] == ["휴대폰번호", "주민등록번호"], spans
    for d in spans:
        assert text[d["start"]:d["end"]] == d["text"], d
    print("find_pii_spans 통과(위치·원문 일치):", [d["text"] for d in spans])


def _selftest_mask_pii():
    """(T0) ⑥ 서버 전송 전 마스킹: 개인정보만 가리고 업무 정보(예산·유선전화·
    날짜)는 남긴다. 감지 없음이면 원문 그대로."""
    text = "김주무관(010-1234-5678, 900101-1234567)이 2026-09-07에 예산 1,850,000원을 계좌 123-45-67890로 집행"
    masked, types = mask_pii(text)
    assert types == ["주민등록번호", "휴대폰번호", "계좌번호"], types
    assert "010-1234-5678" not in masked and "900101-1234567" not in masked and "123-45-67890" not in masked, masked
    assert "[휴대폰번호 삭제]" in masked and "[주민등록번호 삭제]" in masked and "[계좌번호 삭제]" in masked, masked
    assert "2026-09-07" in masked and "1,850,000원" in masked, masked  # 업무 정보는 보존
    clean = "예산 1,850,000원, 문의 031-1234-5678"
    assert mask_pii(clean) == (clean, []), mask_pii(clean)
    print("mask_pii 통과(개인정보만 가림·업무 정보 보존)")


if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8")
    _selftest_detect_pii_patterns_rrn()
    _selftest_detect_pii_patterns_mobile()
    _selftest_detect_pii_patterns_no_false_positive()
    _selftest_date_not_account()
    _selftest_find_pii_spans_positions()
    _selftest_mask_pii()
