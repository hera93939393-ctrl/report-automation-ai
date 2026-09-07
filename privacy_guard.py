"""privacy_guard.py — 채팅 입력에 개인정보로 보이는 패턴(주민등록번호/
휴대폰번호/계좌번호)이 있는지 감지하는 순수 함수. 외부 의존성 없음(re만 사용).

이 모듈은 개인정보를 "확실히 판별"하지 않는다 — 정규식 패턴 매칭이므로
숫자 형태가 비슷하면 실제 개인정보가 아니어도 감지될 수 있다(오탐 가능).
그래서 chat_assistant.py 쪽 연동은 감지되면 곧바로 차단하지 않고 사용자에게
"계속 진행할지"만 확인한다."""
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


def detect_pii_patterns(text: str) -> list[str]:
    """text에서 개인정보로 보이는 패턴을 감지해 발견된 패턴 유형 이름
    리스트를 반환한다. 아무 것도 감지 안 되면 빈 리스트.

    검사 순서: 주민등록번호 → 휴대폰번호 → 계좌번호. 계좌번호 패턴은
    휴대폰번호/유선전화 패턴과 모양이 겹치므로(하이픈 구분 숫자 그룹),
    두 전화 패턴이 이미 매치한 구간과 겹치는 매치는 계좌번호로 세지 않는다
    (안 그러면 "031-1234-5678" 같은 흔한 유선전화가 매번 "계좌번호"로
    오탐된다).

    알려진 한계: 계좌번호 패턴은 "2026-09-07"처럼 하이픈 두 개로 나뉜
    날짜와도 모양이 겹친다 — 전화번호만 제외구간으로 처리하므로 날짜는
    걸러지지 않는다. 채팅에 날짜가 있으면 "계좌번호로 보이는 패턴"으로
    오탐될 수 있다는 뜻이며, 이번 라운드에서는 감수하는 트레이드오프로
    남겨둔다(사용자에게 확인만 구하고 차단하지 않는 설계라 피해가 제한적
    — 다음 라운드에서 날짜 패턴도 제외구간에 넣는 개선 여지가 있음).
    """
    found = []
    if _RRN_PATTERN.search(text):
        found.append("주민등록번호")
    if _MOBILE_PATTERN.search(text):
        found.append("휴대폰번호")

    excluded_spans = (
        [m.span() for m in _MOBILE_PATTERN.finditer(text)]
        + [m.span() for m in _LANDLINE_PATTERN.finditer(text)]
    )

    def _overlaps_excluded(span):
        a_start, a_end = span
        return any(a_start < e and s < a_end for s, e in excluded_spans)

    for m in _ACCOUNT_PATTERN.finditer(text):
        if not _overlaps_excluded(m.span()):
            found.append("계좌번호")
            break

    return found


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


if __name__ == "__main__":
    _selftest_detect_pii_patterns_rrn()
    _selftest_detect_pii_patterns_mobile()
    _selftest_detect_pii_patterns_no_false_positive()
