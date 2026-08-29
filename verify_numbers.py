"""verify_numbers.py — 원본데이터 대비 보고서 숫자검증 핵심 로직 (순수 함수, 외부 의존성 없음)"""
import re


_AMOUNT_PATTERN = re.compile(r'(\d+(?:,\d{3})*(?:\.\d+)?)(?:\s*(만원|천원|원|%))?')
_UNIT_MULTIPLIER = {"만원": 10000, "천원": 1000, "원": 1, "%": 1, None: 1}


def extract_amounts(text: str) -> list[dict]:
    """텍스트에서 금액/일반숫자를 뽑아 정규화된 값과 원문을 반환한다.

    반환 형식: [{"type": "amount", "normalized": "1850000", "raw": "1,850,000원",
                 "span": (start, end)}, ...]
    normalized는 단위환산이 적용된 순수 숫자 문자열이다(비교 시 float으로 변환해 사용).
    """
    results = []
    for m in _AMOUNT_PATTERN.finditer(text):
        digits, unit = m.group(1), m.group(2)
        value = float(digits.replace(",", "")) * _UNIT_MULTIPLIER[unit]
        # 정수면 소수점 없이, 아니면 소수점 유지 (185.5만원 같은 케이스 대비)
        normalized = str(int(value)) if value == int(value) else str(value)
        results.append({
            "type": "amount",
            "normalized": normalized,
            "raw": m.group(0),
            "span": m.span(),
        })
    return results


_DATE_ISO = re.compile(r'(\d{4})-(\d{1,2})-(\d{1,2})')
_DATE_KOR = re.compile(r'(\d{1,2})월\s*(\d{1,2})일')
_DATE_ABBR = re.compile(r"'(\d{2})\.(\d{1,2})\.(\d{1,2})(?:\([월화수목금토일]\))?")


def extract_dates(text: str, default_year: int) -> list[dict]:
    """ISO(2026-09-07) / 한국어(9월 7일) / 공문서축약형('26.9.7(화)) 3종을
    모두 YYYY-MM-DD 문자열로 정규화해서 반환한다.
    한국어 형식은 연도 정보가 없어 default_year를 사용한다.
    """
    results = []
    for m in _DATE_ISO.finditer(text):
        y, mo, d = m.groups()
        results.append({"type": "date", "normalized": f"{int(y):04d}-{int(mo):02d}-{int(d):02d}",
                         "raw": m.group(0), "span": m.span()})
    for m in _DATE_KOR.finditer(text):
        mo, d = m.groups()
        results.append({"type": "date", "normalized": f"{default_year:04d}-{int(mo):02d}-{int(d):02d}",
                         "raw": m.group(0), "span": m.span()})
    for m in _DATE_ABBR.finditer(text):
        yy, mo, d = m.groups()
        results.append({"type": "date", "normalized": f"{2000 + int(yy):04d}-{int(mo):02d}-{int(d):02d}",
                         "raw": m.group(0), "span": m.span()})
    return results


def _selftest_extract_amounts():
    result = extract_amounts("예산 1,850,000원이고 참가인원 342명, 증가율 12%")
    normalized = [r["normalized"] for r in result]
    assert normalized == ["1850000", "342", "12"], normalized
    print("extract_amounts 통과:", result)


def _selftest_unit_multipliers():
    # 만원(x10000)/천원(x1000) 배율 변환이 실제로 적용되는지 검증
    result = extract_amounts("185만원")
    assert len(result) == 1 and result[0]["normalized"] == "1850000", result
    assert result[0]["raw"] == "185만원", result

    result = extract_amounts("3천원")
    assert len(result) == 1 and result[0]["normalized"] == "3000", result
    assert result[0]["raw"] == "3천원", result
    print("_selftest_unit_multipliers 통과:", result)


def _selftest_no_unit_no_trailing_space():
    # 단위 없는 숫자 뒤 공백이 raw/span에 섞이면 안 됨 (mark_red의 exact-text find 대비)
    result = extract_amounts("회의는 2026-09-07 14:00~16:00")
    matches = [r for r in result if r["raw"].startswith("07")]
    assert matches, result
    for r in matches:
        assert r["raw"] == "07", repr(r["raw"])
        assert not r["raw"].endswith(" "), repr(r["raw"])
    print("_selftest_no_unit_no_trailing_space 통과:", result)


def _selftest_extract_dates():
    text = "회의는 2026-09-07, 또는 9월 7일, 혹은 '26.9.7(화)에 진행"
    result = extract_dates(text, default_year=2026)
    normalized = [r["normalized"] for r in result]
    assert normalized == ["2026-09-07", "2026-09-07", "2026-09-07"], normalized
    print("extract_dates 통과:", result)


_TIME_COLON = re.compile(r'(\d{1,2}):(\d{2})\s*[~-]\s*(\d{1,2}):(\d{2})')
_TIME_FULL_SI = re.compile(r'(\d{1,2})시\s*[~-]\s*(\d{1,2})시')
_TIME_SHORT_SI = re.compile(r'(\d{1,2})\s*[~-]\s*(\d{1,2})시')
_COLON_DIGIT_RUN = re.compile(r':\d+')


def extract_times(text: str) -> list[dict]:
    """14:00~16:00 / 14시~16시 / 14~16시 3종을 "HH:MM~HH:MM"(24시간制)로
    정규화해서 반환한다. 콜론형은 실제 분(分)을 그대로 유지하고, 시(時) 단위로만
    표현되는 두 형식(전체시/축약시)은 분을 00으로 간주한다 — 그래야
    "14:30~16:45"(원본과 분 단위로 다른 값) 같은 불일치를 놓치지 않는다.

    콜론(:) 바로 뒤에 오는 숫자 구간(분에 해당, 자릿수 무관)은 전체시/축약시
    패턴이 시작 위치로 삼지 못하도록 배제한다 — 부정 후방탐색은 고정 길이만
    막을 수 있어 일반화가 안 되므로, `:\\d+` 구간의 위치를 직접 찾아 그 구간
    안에서 시작하는 매치를 전부 걸러내는 방식을 쓴다.
    """
    results = []
    for m in _TIME_COLON.finditer(text):
        h1, m1, h2, m2 = m.groups()
        results.append({"type": "time", "normalized": f"{int(h1):02d}:{m1}~{int(h2):02d}:{m2}",
                         "raw": m.group(0), "span": m.span()})

    colon_digit_spans = [m.span() for m in _COLON_DIGIT_RUN.finditer(text)]

    def _starts_inside_colon_digits(pos):
        return any(s < pos < e for s, e in colon_digit_spans)

    for m in _TIME_FULL_SI.finditer(text):
        if _starts_inside_colon_digits(m.start()):
            continue
        h1, h2 = m.groups()
        results.append({"type": "time", "normalized": f"{int(h1):02d}:00~{int(h2):02d}:00",
                         "raw": m.group(0), "span": m.span()})
    for m in _TIME_SHORT_SI.finditer(text):
        if _starts_inside_colon_digits(m.start()):
            continue
        h1, h2 = m.groups()
        results.append({"type": "time", "normalized": f"{int(h1):02d}:00~{int(h2):02d}:00",
                         "raw": m.group(0), "span": m.span()})
    return results


def _selftest_extract_times():
    text = "회의 시간은 14:00~16:00, 또는 14시~16시, 혹은 14~16시"
    result = extract_times(text)
    normalized = [r["normalized"] for r in result]
    assert normalized == ["14:00~16:00", "14:00~16:00", "14:00~16:00"], normalized
    print("extract_times 통과:", result)


def _selftest_extract_times_minute_precision():
    """콜론형은 분(分)까지 정확히 비교해야 하므로, 분이 다르면 다른 값으로 나와야 한다."""
    result = extract_times("14:30~16:45")
    assert result[0]["normalized"] == "14:30~16:45", result
    print("_selftest_extract_times_minute_precision 통과:", result)


def _selftest_extract_times_mixed_notation_no_false_match():
    """콜론형 한쪽만 있는 손상/혼합 표기에서, 분 숫자를 시로 잘못 읽어 매치하면 안 된다
    (예: "14:00~16시"에서 "00~16시"를 "0시~16시"로 오인하는 과거 버그 재발 방지)."""
    result = extract_times("14:00~16시")
    assert result == [], result
    print("_selftest_extract_times_mixed_notation_no_false_match 통과:", result)


def _selftest_extract_times_long_digit_run_no_false_match():
    """콜론 뒤 숫자가 3자리 이상(오타 등)이어도 시로 오독하지 않아야 한다."""
    assert extract_times("14:100~16시") == []
    assert extract_times("14:009~16시") == []
    assert extract_times("09:001~10시") == []
    assert extract_times("14:100시~16시") == []
    print("_selftest_extract_times_long_digit_run_no_false_match 통과")


_PHONE_PATTERN = re.compile(r'\d{2,3}-\d{3,4}-\d{4}')


def extract_phones(text: str) -> list[dict]:
    return [{"type": "phone", "normalized": m.group(0), "raw": m.group(0), "span": m.span()}
            for m in _PHONE_PATTERN.finditer(text)]


def extract_values(text: str, default_year: int) -> list[dict]:
    """텍스트에서 금액/날짜/시간/전화번호 4종을 전부 뽑아 하나의 리스트로 반환한다.
    날짜·시간·전화번호를 먼저 뽑고, 그 글자 범위는 금액 추출 대상에서 제외한다
    (전화번호의 하이픈숫자, 날짜의 연도, 시간의 숫자가 금액으로 오인되는 것을 방지).
    """
    dates = extract_dates(text, default_year)
    times = extract_times(text)
    phones = extract_phones(text)
    excluded_spans = [r["span"] for r in dates + times + phones]

    def _in_excluded(span):
        return any(s <= span[0] < e for s, e in excluded_spans)

    amounts = [r for r in extract_amounts(text) if not _in_excluded(r["span"])]
    return amounts + dates + times + phones


def _selftest_extract_values():
    text = "담당자 031-1234-5678, 예산 1,850,000원, 2026-09-07 14:00~16:00 진행"
    result = extract_values(text, default_year=2026)
    types = sorted(r["type"] for r in result)
    assert types == ["amount", "date", "phone", "time"], types
    phone = [r for r in result if r["type"] == "phone"][0]
    assert phone["normalized"] == "031-1234-5678", phone
    print("extract_values 통과:", result)


if __name__ == "__main__":
    _selftest_extract_amounts()
    _selftest_unit_multipliers()
    _selftest_no_unit_no_trailing_space()
    _selftest_extract_dates()
    _selftest_extract_times()
    _selftest_extract_times_minute_precision()
    _selftest_extract_times_mixed_notation_no_false_match()
    _selftest_extract_times_long_digit_run_no_false_match()
    _selftest_extract_values()
