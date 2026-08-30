"""verify_numbers.py — 원본데이터 대비 보고서 숫자검증 핵심 로직 (순수 함수, 외부 의존성 없음)"""
import re
from decimal import Decimal


_AMOUNT_PATTERN = re.compile(r'(\d+(?:,\d{3})*(?:\.\d+)?)(?:\s*(만원|천원|원|%))?')
_UNIT_MULTIPLIER = {"만원": 10000, "천원": 1000, "원": 1, "%": 1, None: 1}


def _decimal_to_normalized_str(value: Decimal) -> str:
    """Decimal 값을 정규화된 문자열로 바꾼다. 정수면 소수점 없이, 아니면
    고정소수점 표기로 — 어느 분기든 항상 format(x, 'f')를 써서 과학적 표기법이
    나오지 않게 강제한다. 소수 분기는 아주 작은 값("1E-7")에서, 정수 분기는
    Decimal의 기본 28자리 정밀도를 넘는 아주 큰 값(예: 1e28)에서 str()이
    과학적 표기법을 낼 수 있어 둘 다 format(x, 'f')가 필요하다.
    """
    integral = value.to_integral_value()
    if value == integral:
        return format(integral, 'f')
    return format(value.normalize(), 'f')


def extract_amounts(text: str) -> list[dict]:
    """텍스트에서 금액/일반숫자를 뽑아 정규화된 값과 원문을 반환한다.

    반환 형식: [{"type": "amount", "normalized": "1850000", "raw": "1,850,000원",
                 "span": (start, end)}, ...]
    normalized는 단위환산이 적용된 순수 숫자 문자열이다. float이 아니라 Decimal로
    계산하는 이유: "1.005천원"처럼 소수×배율을 float으로 계산하면 이진 부동소수점
    표현 오차(예: 1004.9999999999999)가 생겨, 같은 실제 값을 다르게 표기한
    "1,005원"과 문자열이 달라져 Task 5의 정확일치 비교에서 오탐이 난다. Decimal은
    십진수를 오차 없이 그대로 표현하므로 이 문제가 생기지 않는다.
    """
    results = []
    for m in _AMOUNT_PATTERN.finditer(text):
        digits, unit = m.group(1), m.group(2)
        value = Decimal(digits.replace(",", "")) * _UNIT_MULTIPLIER[unit]
        normalized = _decimal_to_normalized_str(value)
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


def _selftest_extract_amounts_decimal_unit_no_float_noise():
    """소수+단위 조합이 float 오차 없이, 같은 값을 다르게 쓴 것과 정확히 같은
    문자열로 정규화되어야 한다 (Task 5의 정확일치 비교가 성립하는 전제)."""
    a = extract_amounts("예산은 1.005천원입니다")
    b = extract_amounts("예산은 1,005원입니다")
    assert a[0]["normalized"] == b[0]["normalized"] == "1005", (a, b)
    print("_selftest_extract_amounts_decimal_unit_no_float_noise 통과:", a, b)


def _selftest_extract_amounts_no_scientific_notation():
    """아주 작은 소수라도 과학적 표기법("1E-7")이 아니라 고정소수점("0.0000001")으로
    나와야 한다 (그래야 정확일치 비교가 깨지지 않는다)."""
    result = extract_amounts("증감률은 0.0000001%입니다")
    assert result[0]["normalized"] == "0.0000001", result
    print("_selftest_extract_amounts_no_scientific_notation 통과:", result)


def _selftest_decimal_to_normalized_str_large_integral_no_scientific_notation():
    """Decimal의 기본 28자리 정밀도를 넘는 아주 큰 정수값도 str()이 과학적 표기법
    ("1E+28")을 내지 않고 고정소수점 전체 자릿수로 나와야 한다."""
    result = _decimal_to_normalized_str(Decimal("1E+28"))
    assert result == "10000000000000000000000000000", result
    assert "E" not in result and "e" not in result, result

    result2 = _decimal_to_normalized_str(Decimal("2E+30"))
    assert result2 == "2000000000000000000000000000000", result2
    assert "E" not in result2 and "e" not in result2, result2

    print("_selftest_decimal_to_normalized_str_large_integral_no_scientific_notation 통과:",
          result, result2)


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
    """전화번호(031-1234-5678 등 표준 하이픈 형식)를 추출한다."""
    return [{"type": "phone", "normalized": m.group(0), "raw": m.group(0), "span": m.span()}
            for m in _PHONE_PATTERN.finditer(text)]


def extract_values(text: str, default_year: int) -> list[dict]:
    """텍스트에서 금액/날짜/시간/전화번호 4종을 전부 뽑아 하나의 리스트로 반환한다.
    날짜·시간·전화번호를 먼저 뽑고, 그 글자 범위와 조금이라도 겹치는 금액 매치는
    제외한다 (전화번호의 하이픈숫자, 날짜의 연도, 시간의 숫자가 금액으로 오인되는
    것을 방지 — 시작 위치만 보면 안 되고 구간 겹침 전체를 봐야 한다. 예:
    "1"+날짜가 공백 없이 붙은 "12026-09-07" 같은 입력에서 금액 매치("12026")가
    제외구간보다 먼저 시작하면서 겹치는 경우까지 잡아야 함).
    """
    dates = extract_dates(text, default_year)
    times = extract_times(text)
    phones = extract_phones(text)
    excluded_spans = [r["span"] for r in dates + times + phones]

    def _overlaps_excluded(span):
        a_start, a_end = span
        return any(a_start < e and s < a_end for s, e in excluded_spans)

    amounts = [r for r in extract_amounts(text) if not _overlaps_excluded(r["span"])]
    return amounts + dates + times + phones


def _selftest_extract_values():
    text = "담당자 031-1234-5678, 예산 1,850,000원, 2026-09-07 14:00~16:00 진행"
    result = extract_values(text, default_year=2026)
    types = sorted(r["type"] for r in result)
    assert types == ["amount", "date", "phone", "time"], types
    phone = [r for r in result if r["type"] == "phone"][0]
    assert phone["normalized"] == "031-1234-5678", phone
    print("extract_values 통과:", result)


def _selftest_extract_values_adjacent_no_separator():
    """구분자 없이 붙어있어 금액 매치가 제외구간보다 먼저 시작하며 겹치는 경우도
    걸러내야 한다 (시작 위치만 보는 검사로는 놓치는 케이스)."""
    result = extract_values("1" + "2026-09-07 회의", default_year=2026)
    amounts = [r for r in result if r["type"] == "amount"]
    assert amounts == [], amounts
    print("_selftest_extract_values_adjacent_no_separator 통과:", result)


def _selftest_extract_values_reverse_direction_overlap():
    """금액 매치가 제외구간 안에서 시작해 밖으로 걸치는 반대 방향도 걸러야 한다."""
    result = extract_values("연락처 031-1234-56789999 입니다", default_year=2026)
    amounts = [r for r in result if r["type"] == "amount"]
    assert amounts == [], amounts
    print("_selftest_extract_values_reverse_direction_overlap 통과:", result)


def _selftest_extract_values_straddles_two_adjacent_excluded_spans():
    """제외구간(날짜+시간) 두 개가 붙어있고, 그 경계를 금액이 걸치는 경우도 걸러야 한다."""
    result = extract_values("2026-09-0714:00~16:0099900", default_year=2026)
    amounts = [r for r in result if r["type"] == "amount"]
    assert amounts == [], amounts
    print("_selftest_extract_values_straddles_two_adjacent_excluded_spans 통과:", result)


def compare_values(report_values: list[dict], answer_pool: list[dict]) -> list[dict]:
    """report_values 각각을 answer_pool과 대조해, 원본에서 확인 안 되는 항목만 반환한다.
    4종(금액/날짜/시간/전화번호) 전부 normalized 값이 정확히 일치해야 통과한다.
    금액에 별도 오차 허용을 두지 않는 이유: normalized는 이미 단위환산까지 끝난
    깨끗한 정수 문자열이고, 이 도구가 다루는 금액 규모(최대 수조 원)는 배정밀도
    float의 정수 정확 표현 한계(2^53)에 전혀 못 미쳐 부동소수점 잡음이 생기지
    않는다. 상대오차(%) 허용은 금액이 클수록 허용되는 절대 오차도 커져서,
    큰 금액에서 실제 오타를 놓치는 근본적인 결함이 있었다(임계값을 아무리
    좁혀도 해결 안 됨) — 그래서 정확 일치로 되돌린다.
    """
    mismatches = []
    for rv in report_values:
        candidates = [a for a in answer_pool if a["type"] == rv["type"]]
        found = any(a["normalized"] == rv["normalized"] for a in candidates)
        if not found:
            mismatches.append(rv)
    return mismatches


def _selftest_compare_values():
    answer_pool = [
        {"type": "amount", "normalized": "1850000", "raw": "1,850,000", "source_file": "원본.xlsx", "location": "Sheet1!C15"},
        {"type": "date", "normalized": "2026-09-07", "raw": "2026-09-07", "source_file": "원본.xlsx", "location": "Sheet1!D2"},
    ]
    report_values = [
        {"type": "amount", "normalized": "1850000", "raw": "185만원", "span": (0, 4)},   # 정상(단위환산 일치)
        {"type": "amount", "normalized": "9999999", "raw": "999만9900원", "span": (10, 20)},  # 오탐(원본에 없음)
        {"type": "date", "normalized": "2026-09-07", "raw": "'26.9.7(화)", "span": (30, 40)},  # 정상
    ]
    mismatches = compare_values(report_values, answer_pool)
    assert len(mismatches) == 1, mismatches
    assert mismatches[0]["raw"] == "999만9900원", mismatches
    print("compare_values 통과:", mismatches)


def _selftest_compare_values_catches_digit_transposition_typo():
    """0.5% 오차 허용에서는 놓쳤던, 자릿수를 바꿔 쓴 전형적인 오타를 다시 잡아야 한다."""
    answer_pool = [{"type": "amount", "normalized": "12345000", "raw": "12,345,000",
                     "source_file": "원본.xlsx", "location": "Sheet1!C1"}]
    report_values = [{"type": "amount", "normalized": "12354000", "raw": "12,354,000", "span": (0, 10)}]
    mismatches = compare_values(report_values, answer_pool)
    assert len(mismatches) == 1, mismatches
    print("_selftest_compare_values_catches_digit_transposition_typo 통과:", mismatches)


def _selftest_compare_values_catches_typo_in_large_amount():
    """상대오차 방식의 근본 결함이었던 케이스 — 10억 단위에서 마지막 자리 오타도 잡혀야 한다."""
    answer_pool = [{"type": "amount", "normalized": "1000000000", "raw": "10억",
                     "source_file": "원본.xlsx", "location": "Sheet1!C2"}]
    report_values = [{"type": "amount", "normalized": "1000000001", "raw": "1,000,000,001", "span": (0, 10)}]
    mismatches = compare_values(report_values, answer_pool)
    assert len(mismatches) == 1, mismatches
    print("_selftest_compare_values_catches_typo_in_large_amount 통과:", mismatches)


def compute_column_sums(rows: list[dict], source_file: str, sheet: str) -> list[dict]:
    """엑셀에서 읽은 행(딕셔너리 리스트)에서, 숫자로만 이루어진 각 컬럼의 단순 합계를
    미리 계산해 정답 풀에 추가할 항목으로 반환한다 (PRD 12-2: 합계만, 증감률/평균은 제외).
    float을 직접 sum()하지 않고 Decimal로 변환해 더한다 — 소수를 포함하는 컬럼(단가 등)의
    합계에서 이진 부동소수점 오차가 생겨 Task 5의 정확일치 비교와 충돌하는 것을 방지한다.
    bool은 int의 서브클래스라 isinstance(v, (int, float))를 그냥 두면 True/False가
    섞인 컬럼(완료여부 등)에서 Decimal("True") 파싱 오류가 나므로 명시적으로 제외한다.
    """
    if not rows:
        return []
    results = []
    for col in rows[0].keys():
        values = [r[col] for r in rows
                  if isinstance(r.get(col), (int, float)) and not isinstance(r.get(col), bool)]
        if len(values) == len(rows) and values:
            total = sum(Decimal(str(v)) for v in values)
            normalized = _decimal_to_normalized_str(total)
            results.append({
                "type": "amount", "normalized": normalized, "raw": f"{col} 합계",
                "source_file": source_file, "location": f"{sheet}!{col}(합계)",
            })
    return results


def _selftest_compute_column_sums():
    rows = [{"예산": 500000, "인원": 10}, {"예산": 300000, "인원": 20}]
    sums = compute_column_sums(rows, source_file="원본.xlsx", sheet="Sheet1")
    normalized = sorted(r["normalized"] for r in sums)
    assert normalized == ["30", "800000"], normalized
    print("compute_column_sums 통과:", sums)


def _selftest_compute_column_sums_decimal_no_float_noise():
    """소수를 포함하는 컬럼의 합계도 float 오차 없이 정확해야 한다."""
    rows = [{"단가": 0.1}, {"단가": 0.2}]
    sums = compute_column_sums(rows, source_file="원본.xlsx", sheet="Sheet1")
    assert sums[0]["normalized"] == "0.3", sums
    print("_selftest_compute_column_sums_decimal_no_float_noise 통과:", sums)


def _selftest_compute_column_sums_ignores_boolean_column():
    """완료여부 같은 True/False 컬럼이 섞여 있어도 크래시 없이 무시해야 한다."""
    rows = [{"예산": 100, "완료": True}, {"예산": 200, "완료": False}]
    sums = compute_column_sums(rows, source_file="원본.xlsx", sheet="Sheet1")
    normalized = sorted(r["normalized"] for r in sums)
    assert normalized == ["300"], normalized  # "완료" 컬럼은 합계 대상에서 제외됨
    print("_selftest_compute_column_sums_ignores_boolean_column 통과:", sums)


if __name__ == "__main__":
    _selftest_extract_amounts()
    _selftest_extract_amounts_decimal_unit_no_float_noise()
    _selftest_extract_amounts_no_scientific_notation()
    _selftest_decimal_to_normalized_str_large_integral_no_scientific_notation()
    _selftest_unit_multipliers()
    _selftest_no_unit_no_trailing_space()
    _selftest_extract_dates()
    _selftest_extract_times()
    _selftest_extract_times_minute_precision()
    _selftest_extract_times_mixed_notation_no_false_match()
    _selftest_extract_times_long_digit_run_no_false_match()
    _selftest_extract_values()
    _selftest_extract_values_adjacent_no_separator()
    _selftest_extract_values_reverse_direction_overlap()
    _selftest_extract_values_straddles_two_adjacent_excluded_spans()
    _selftest_compare_values()
    _selftest_compare_values_catches_digit_transposition_typo()
    _selftest_compare_values_catches_typo_in_large_amount()
    _selftest_compute_column_sums()
    _selftest_compute_column_sums_decimal_no_float_noise()
    _selftest_compute_column_sums_ignores_boolean_column()
