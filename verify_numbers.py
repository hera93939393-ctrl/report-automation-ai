"""verify_numbers.py — 원본데이터 대비 보고서 숫자검증 핵심 로직 (순수 함수, 외부 의존성 없음)"""
import datetime
import re
from decimal import Decimal


_AMOUNT_PATTERN = re.compile(r'(\d+(?:,\d{3})*(?:\.\d+)?)(?:\s*(만원|천원|원|%))?')
_UNIT_MULTIPLIER = {"만원": 10000, "천원": 1000, "원": 1, "%": 1, None: 1}

# (2026-09-04, 실사용 피드백으로 발견) 공공기관 보고서에서 흔한 "('23)","'24년"
# 같은 연도 약칭 표기 — 작은따옴표(또는 스마트따옴표) 바로 뒤에 두 자리 숫자가
# 오는 패턴. 이 두 자리 숫자("23","24")가 금액으로 뽑혀버리면, 원본자료의
# 진짜 데이터 값과는 아무 관계도 없는데 "오류(빨강)"로 잘못 표시된다(실사용
# 문서 "정기점검 실적 : ('23) 1,595개소 → ('24) 1,694 → ('25) 1,827"에서
# 실제 재현됨 — 1,595/1,694/1,827은 정상인데 23/24/25가 전부 오류로 잘못
# 표시됨). extract_values()가 날짜/시간/전화번호처럼 이 패턴도 "제외 구간"으로
# 다뤄서, 겹치는 금액 매치를 애초에 뽑지 않게 한다.
_YEAR_ABBREVIATION_PATTERN = re.compile(r"['’‘]\d{2}(?!\d)")

# (2026-09-07, 실사용 피드백으로 발견) "1." "2)"처럼 목차·개요·붙임 번호 등
# 문서 어디서나 등장하는 번호매기기 표기 — numbering_tool.py가 실제로 쓰는
# 표준 번호서식("1. "/"1) ") 관례와 정확히 일치한다. 이런 숫자는 순서를
# 나타내는 라벨일 뿐 원본자료와 대조할 데이터가 아니다. 소수점 금액
# ("97.1")은 마침표 뒤에 또 숫자가 오므로 (?!\d)로 걸러 오배제하지 않는다.
_LIST_MARKER_PATTERN = re.compile(r'(\d+)(?:\.(?!\d)|\))')

# (2026-09-07) 연속된 1~12가 월(月) 표시로 흔히 쓰인다("1 2 3 ... 12" 월별
# 헤더) — extract_values()가 후보 amounts 중에서 이 조건에 맞는 런을 찾아
# 검증 대상에서 제외한다(_find_month_sequence_spans 참고).


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

# (F13) "'26.9.7(화)"처럼 날짜 뒤 괄호에 요일이 적힌 경우만 잡는다.
# _DATE_ABBR과 같은 날짜 모양에 괄호+요일을 필수로 요구하는 점만 다르다.
_DATE_WEEKDAY_PATTERN = re.compile(r"'(\d{2})\.(\d{1,2})\.(\d{1,2})\(([월화수목금토일])\)")
_WEEKDAY_NAMES = ["월", "화", "수", "목", "금", "토", "일"]


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


def check_weekday_consistency(text: str) -> list[dict]:
    """"'26.9.7(화)"처럼 날짜 뒤 괄호에 적힌 요일이, 그 날짜의 실제 요일과
    맞는지 검증한다. 표기된 요일과 실제 요일이 다른 것만 반환한다(일치하면
    결과에 안 넣음). 요일 표기가 아예 없는 날짜("2026-09-07")는 이 정규식
    자체가 매치하지 않으므로 건드리지 않는다.
    """
    mismatches = []
    for m in _DATE_WEEKDAY_PATTERN.finditer(text):
        yy, mo, d, written_weekday = m.groups()
        actual_date = datetime.date(2000 + int(yy), int(mo), int(d))
        actual_weekday = _WEEKDAY_NAMES[actual_date.weekday()]
        if written_weekday != actual_weekday:
            mismatches.append({
                "raw": m.group(0),
                "written_weekday": written_weekday,
                "actual_weekday": actual_weekday,
                "span": m.span(),
            })
    return mismatches


def _selftest_check_weekday_consistency_detects_mismatch():
    """2026-09-07의 실제 요일을 파이썬으로 미리 확인:
    datetime.date(2026,9,7).weekday() == 0(월요일). 표기를 일부러 틀리게
    ("화") 써서 불일치가 잡히는지 확인한다."""
    assert datetime.date(2026, 9, 7).weekday() == 0, "전제 확인: 2026-09-07은 월요일이어야 함"
    text = "회의는 '26.9.7(화)에 진행"
    result = check_weekday_consistency(text)
    assert len(result) == 1, result
    assert result[0]["written_weekday"] == "화", result
    assert result[0]["actual_weekday"] == "월", result
    print("check_weekday_consistency 통과(불일치 감지):", result)


def _selftest_check_weekday_consistency_matches_when_correct():
    """실제 요일과 맞게 쓰면 빈 리스트여야 한다."""
    text = "회의는 '26.9.7(월)에 진행"
    result = check_weekday_consistency(text)
    assert result == [], result
    print("check_weekday_consistency 통과(일치 시 빈 리스트):", result)


def _selftest_check_weekday_consistency_ignores_date_without_weekday():
    """요일 표기가 아예 없는 날짜는 이 함수가 건드리지 않아야 한다."""
    result = check_weekday_consistency("회의는 2026-09-07에 진행")
    assert result == [], result
    print("check_weekday_consistency 통과(요일 표기 없으면 무시):", result)


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


def _find_month_sequence_spans(amounts: list[dict]) -> set[tuple[int, int]]:
    """amounts(문서 내 위치 순으로 정렬되지 않았을 수 있음) 중에서 "1부터
    12까지 정확히 순서대로 연속 등장"하는 런을 찾아, 그 12개 항목의 span을
    반환한다(월별 헤더로 흔히 쓰이는 형태 - 실제 데이터가 아니라 월을
    나타내는 라벨). 소수점/단위 없이 순수 정수로 정규화된 값만 후보로 보고,
    문서 내 위치(span) 순서대로 정렬한 뒤 "바로 다음 값 = 직전 값 + 1"이
    끊기지 않고 1부터 12까지 이어지는 구간만 채택한다. 1에서 시작하지
    않거나 12에서 끝나지 않는 부분적인 연속(예: "3 4 5 6 7")은 실제 데이터일
    가능성을 배제할 수 없어 제외하지 않는다."""
    candidates = sorted(
        (a for a in amounts if a["normalized"].isdigit() and 1 <= int(a["normalized"]) <= 12),
        key=lambda a: a["span"][0],
    )
    excluded = set()
    i = 0
    while i < len(candidates):
        run = [candidates[i]]
        j = i + 1
        while j < len(candidates) and int(candidates[j]["normalized"]) == int(run[-1]["normalized"]) + 1:
            run.append(candidates[j])
            j += 1
        if len(run) == 12 and int(run[0]["normalized"]) == 1:
            excluded.update(a["span"] for a in run)
        i = j if j > i + 1 else i + 1
    return excluded


def _selftest_find_month_sequence_spans_detects_full_run():
    amounts = [{"normalized": str(v), "raw": str(v), "span": (i * 3, i * 3 + 1)}
               for i, v in enumerate(range(1, 13))]
    spans = _find_month_sequence_spans(amounts)
    assert spans == {a["span"] for a in amounts}, spans
    print("_find_month_sequence_spans 통과(1~12 전체 런 감지)")


def _selftest_find_month_sequence_spans_ignores_partial_run():
    """1에서 시작하지 않거나 12에서 끝나지 않는 부분 연속은 실제 데이터일 수
    있으므로 제외하면 안 된다."""
    amounts = [{"normalized": str(v), "raw": str(v), "span": (i * 3, i * 3 + 1)}
               for i, v in enumerate(range(3, 8))]  # 3,4,5,6,7 (부분 연속)
    spans = _find_month_sequence_spans(amounts)
    assert spans == set(), spans
    print("_find_month_sequence_spans 통과(부분 연속은 무시)")


def _selftest_find_month_sequence_spans_ignores_non_amount_type():
    """정수가 아니거나 1~12 범위를 벗어난 값은 애초에 후보에서 제외된다."""
    amounts = [
        {"normalized": "1", "raw": "1", "span": (0, 1)},
        {"normalized": "1850000", "raw": "1,850,000", "span": (5, 14)},  # 범위 밖 - 후보 아님
    ]
    spans = _find_month_sequence_spans(amounts)
    assert spans == set(), spans
    print("_find_month_sequence_spans 통과(런이 안 되면 제외 없음)")


def extract_values(text: str, default_year: int) -> list[dict]:
    """텍스트에서 금액/날짜/시간/전화번호 4종을 전부 뽑아 하나의 리스트로 반환한다.
    날짜·시간·전화번호를 먼저 뽑고, 그 글자 범위와 조금이라도 겹치는 금액 매치는
    제외한다 (전화번호의 하이픈숫자, 날짜의 연도, 시간의 숫자가 금액으로 오인되는
    것을 방지 — 시작 위치만 보면 안 되고 구간 겹침 전체를 봐야 한다. 예:
    "1"+날짜가 공백 없이 붙은 "12026-09-07" 같은 입력에서 금액 매치("12026")가
    제외구간보다 먼저 시작하면서 겹치는 경우까지 잡아야 함).

    (2026-09-04 추가) "('23)","'24년"처럼 작은따옴표+두자리 숫자로 된 연도
    약칭도 같은 방식으로 제외 구간에 포함한다 — 이런 숫자는 실제 데이터 값이
    아니라 연도를 줄여 쓴 것뿐이라, 원본자료와 대조할 대상이 아니다.

    (2026-09-07 추가, 실사용 피드백) "1." "2)" 같은 목차/개요 번호매기기와,
    "1 2 3 ... 12"처럼 연속된 월(月) 표시도 같은 이유로 검증 대상에서
    뺀다 - 전자는 _LIST_MARKER_PATTERN으로 제외구간에 포함하고, 후자는
    금액 후보를 다 뽑은 뒤 _find_month_sequence_spans()로 별도 제거한다
    (달 전체가 연속으로 등장해야 확정되는 패턴이라 다른 제외구간과 달리
    금액끼리 서로 비교해야 하므로 별도 단계로 처리).
    """
    dates = extract_dates(text, default_year)
    times = extract_times(text)
    phones = extract_phones(text)
    year_abbreviations = [m.span() for m in _YEAR_ABBREVIATION_PATTERN.finditer(text)]
    list_markers = [m.span(1) for m in _LIST_MARKER_PATTERN.finditer(text)]
    excluded_spans = [r["span"] for r in dates + times + phones] + year_abbreviations + list_markers

    def _overlaps_excluded(span):
        a_start, a_end = span
        return any(a_start < e and s < a_end for s, e in excluded_spans)

    amounts = [r for r in extract_amounts(text) if not _overlaps_excluded(r["span"])]
    month_spans = _find_month_sequence_spans(amounts)
    amounts = [r for r in amounts if r["span"] not in month_spans]
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


def _selftest_extract_values_excludes_year_abbreviation():
    """(2026-09-04, 실사용 피드백으로 발견한 실제 버그 재현) "('23) 1,595개소"
    처럼 연도 약칭 바로 뒤에 진짜 금액이 붙어 있을 때, 연도 약칭("23")은
    금액으로 뽑히면 안 되고 진짜 금액("1,595")만 뽑혀야 한다. 사용자의
    실제 문서 문장을 그대로 재현한다."""
    text = "정기점검 실적 : ('23) 1,595개소 → ('24) 1,694 → ('25) 1,827"
    result = extract_values(text, default_year=2026)
    amounts = sorted(r["normalized"] for r in result if r["type"] == "amount")
    assert amounts == ["1595", "1694", "1827"], amounts
    print("_selftest_extract_values_excludes_year_abbreviation 통과:", result)


def _selftest_extract_values_excludes_list_marker_dot():
    """(2026-09-07, 실사용 피드백) "1. 등록심사"처럼 목차/개요 번호매기기로
    쓰인 "1."은 검증 대상에서 빠져야 하고, 그 뒤 실제 데이터("1,827")는
    영향받지 않아야 한다."""
    text = "1. 등록심사 실적은 1,827건이다"
    result = extract_values(text, default_year=2026)
    amounts = sorted(r["normalized"] for r in result if r["type"] == "amount")
    assert amounts == ["1827"], amounts
    print("_selftest_extract_values_excludes_list_marker_dot 통과:", result)


def _selftest_extract_values_excludes_list_marker_paren():
    """"2) 세부내용"처럼 괄호 번호매기기("1) "/numbering_tool.py의
    arabic_paren 스타일)로 쓰인 숫자도 같은 이유로 제외돼야 한다."""
    text = "2) 세부내용 : 예산 1,850,000원"
    result = extract_values(text, default_year=2026)
    amounts = sorted(r["normalized"] for r in result if r["type"] == "amount")
    assert amounts == ["1850000"], amounts
    print("_selftest_extract_values_excludes_list_marker_paren 통과:", result)


def _selftest_extract_values_list_marker_does_not_exclude_decimal_amount():
    """"97.1%"처럼 마침표가 소수점으로 쓰인 진짜 금액은 목차 번호로
    오인해서 제외하면 안 된다(마침표 뒤에 숫자가 더 오는 경우)."""
    text = "IP 안전지수는 97.1%이다"
    result = extract_values(text, default_year=2026)
    amounts = sorted(r["normalized"] for r in result if r["type"] == "amount")
    assert amounts == ["97.1"], amounts
    print("_selftest_extract_values_list_marker_does_not_exclude_decimal_amount 통과:", result)


def _selftest_extract_values_excludes_month_sequence():
    """(2026-09-07, 실사용 피드백으로 발견한 실제 버그 재현) "1 2 3 ... 12"처럼
    연속된 월 표시는 검증 대상에서 빠지고, 그 사이에 낀 진짜 데이터는
    영향받지 않아야 한다."""
    months = " ".join(str(i) for i in range(1, 13))
    text = f"월별 실적: {months} 합계는 1,827건"
    result = extract_values(text, default_year=2026)
    amounts = sorted(r["normalized"] for r in result if r["type"] == "amount")
    assert amounts == ["1827"], amounts
    print("_selftest_extract_values_excludes_month_sequence 통과:", result)


def categorize_values(report_values: list[dict], answer_pool: list[dict]) -> dict:
    """report_values 각각을 answer_pool과 대조해 세 갈래로 나눈다:
    - matches: 원본과 정확히 일치(정상) → 문서에 파란색으로 표시할 대상
    - mismatches: 원본에 같은 타입 값이 있지만 정확히 일치하는 게 없음(오류)
      → 빨간색 표시 대상. 기존 compare_values()가 반환하던 것과 동일한 판정.
    - unverifiable: answer_pool에 그 타입 자체가 하나도 없어서 애초에 대조가
      불가능함 → 초록색 표시 대상. "원본에 이 종류의 데이터 자체가 없어서
      확인 못 했다"는 뜻이지, 값이 틀렸다는 뜻이 아니다.

    (2026-09-04, 실사용 피드백) 원래 compare_values()는 mismatches만
    반환했다 — "문서가 길어지니 전부 확인한 건지 못 알아보겠다"는 사용자
    피드백에 따라, 확인은 됐지만 정상인 값(파랑)과 애초에 대조가 불가능한
    값(초록)도 구분해 색으로 표시하게 됐다. 4종(금액/날짜/시간/전화번호)
    전부 normalized 값이 정확히 일치해야 match로 판정한다. 금액에 별도
    오차 허용을 두지 않는 이유: normalized는 이미 단위환산까지 끝난 깨끗한
    정수 문자열이고, 이 도구가 다루는 금액 규모(최대 수조 원)는 배정밀도
    float의 정수 정확 표현 한계(2^53)에 전혀 못 미쳐 부동소수점 잡음이
    생기지 않는다. 상대오차(%) 허용은 금액이 클수록 허용되는 절대 오차도
    커져서, 큰 금액에서 실제 오타를 놓치는 근본적인 결함이 있었다(임계값을
    아무리 좁혀도 해결 안 됨) — 그래서 정확 일치로 되돌린다.
    """
    matches, mismatches, unverifiable = [], [], []
    for rv in report_values:
        candidates = [a for a in answer_pool if a["type"] == rv["type"]]
        if not candidates:
            unverifiable.append(rv)  # 원본에 이 타입 자체가 없음 → 대조 불가
            continue
        if any(a["normalized"] == rv["normalized"] for a in candidates):
            matches.append(rv)
        else:
            mismatches.append(rv)
    return {"matches": matches, "mismatches": mismatches, "unverifiable": unverifiable}


def compare_values(report_values: list[dict], answer_pool: list[dict]) -> list[dict]:
    """(기존 동작 그대로 유지) report_values 중 원본과 불일치하는 것만
    반환한다. categorize_values()의 "mismatches"만 뽑아 쓰는 얇은 래퍼 —
    기존 호출자(아래 self-test 여러 개, verify_tool.py)와의 하위호환을
    위해 이름과 반환 형태를 그대로 둔다. 새로 파랑/초록 분류까지 필요한
    호출자는 categorize_values()를 직접 쓴다(verify_tool.py run_verification 참고)."""
    return categorize_values(report_values, answer_pool)["mismatches"]


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


def _selftest_compare_values_skips_type_absent_from_source():
    """(2026-08-31 최종 검토 반영 회귀테스트) 원본 정답 풀에 phone 타입이 아예
    없으면, 보고서의 (완전히 정상적인) 전화번호를 오탐으로 찍으면 안 된다."""
    answer_pool = [{"type": "amount", "normalized": "1850000", "raw": "1,850,000",
                     "source_file": "원본.xlsx", "location": "Sheet1!C15"}]
    report_values = extract_values(
        "담당자: 김주무관 (031-1234-5678), 예산 1,850,000원", default_year=2026,
    )
    mismatches = compare_values(report_values, answer_pool)
    assert mismatches == [], mismatches
    print("_selftest_compare_values_skips_type_absent_from_source 통과:", mismatches)


def _selftest_compare_values_still_flags_when_type_present_but_no_match():
    """타입 자체는 원본에 있지만(phone 항목 존재), 이 값과는 하나도 일치하지
    않는 경우는 여전히 mismatch로 잡아야 한다 — 위 스킵 로직이 진짜 오타
    탐지까지 죽여버리면 안 된다."""
    answer_pool = [
        {"type": "phone", "normalized": "031-9999-0000", "raw": "031-9999-0000",
         "source_file": "원본.xlsx", "location": "Sheet1!B3"},
    ]
    report_values = [{"type": "phone", "normalized": "031-1234-5678",
                       "raw": "031-1234-5678", "span": (0, 13)}]
    mismatches = compare_values(report_values, answer_pool)
    assert len(mismatches) == 1 and mismatches[0]["raw"] == "031-1234-5678", mismatches
    print("_selftest_compare_values_still_flags_when_type_present_but_no_match 통과:", mismatches)


def _selftest_compare_values_per_type_skip_does_not_hide_other_mismatches():
    """스킵 로직은 "이 타입이 원본에 아예 없을 때"에만 적용돼야지, 다른 타입에
    영향을 주거나 같은 타입 안에서 진짜 오타를 덮어버리면 안 된다. phone은
    원본에 없어 스킵되지만, amount는 원본에 있고 그중 하나는 진짜 오타라
    여전히 잡혀야 한다(같은 타입이 다른 곳(정상 항목)에도 등장하는 상황 포함)."""
    answer_pool = [
        {"type": "amount", "normalized": "1850000", "raw": "1,850,000",
         "source_file": "원본.xlsx", "location": "Sheet1!C15"},
    ]
    report_values = [
        {"type": "phone", "normalized": "031-1234-5678", "raw": "031-1234-5678", "span": (0, 13)},  # 스킵돼야 함(원본에 phone 없음)
        {"type": "amount", "normalized": "1850000", "raw": "185만원", "span": (20, 24)},  # 정상(일치)
        {"type": "amount", "normalized": "9999999", "raw": "999만9900원", "span": (30, 40)},  # 진짜 오타(여전히 잡혀야 함)
    ]
    mismatches = compare_values(report_values, answer_pool)
    assert len(mismatches) == 1, mismatches
    assert mismatches[0]["raw"] == "999만9900원", mismatches
    print("_selftest_compare_values_per_type_skip_does_not_hide_other_mismatches 통과:", mismatches)


def _selftest_categorize_values_three_buckets():
    """(2026-09-04, 실사용 피드백) categorize_values()가 matches(파랑)/
    mismatches(빨강)/unverifiable(초록) 세 갈래를 정확히 나누는지 확인한다.
    _selftest_compare_values_per_type_skip_does_not_hide_other_mismatches와
    같은 픽스처(phone 스킵 + amount 정상/오타 혼재)를 재사용해, compare_values()
    가 이미 검증한 mismatches 판정과 categorize_values()의 mismatches가
    여전히 같은 결과인지도 함께 확인한다."""
    answer_pool = [
        {"type": "amount", "normalized": "1850000", "raw": "1,850,000",
         "source_file": "원본.xlsx", "location": "Sheet1!C15"},
    ]
    report_values = [
        {"type": "phone", "normalized": "031-1234-5678", "raw": "031-1234-5678", "span": (0, 13)},  # unverifiable(원본에 phone 없음)
        {"type": "amount", "normalized": "1850000", "raw": "185만원", "span": (20, 24)},  # match(일치)
        {"type": "amount", "normalized": "9999999", "raw": "999만9900원", "span": (30, 40)},  # mismatch(오타)
    ]
    result = categorize_values(report_values, answer_pool)
    assert [m["raw"] for m in result["matches"]] == ["185만원"], result["matches"]
    assert [m["raw"] for m in result["mismatches"]] == ["999만9900원"], result["mismatches"]
    assert [m["raw"] for m in result["unverifiable"]] == ["031-1234-5678"], result["unverifiable"]
    print("_selftest_categorize_values_three_buckets 통과:", result)


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


def compute_column_growth_rate(rows: list[dict], source_file: str, sheet: str,
                                from_col: str, to_col: str) -> list[dict]:
    """지정한 두 컬럼(from_col 합계 대비 to_col 합계)의 증감률(%)을 계산해
    정답 풀에 추가할 항목 하나로 반환한다. compute_column_sums()와 같은
    반환 dict 형태(type/normalized/raw/source_file/location)를 따르며, 그
    함수가 이미 계산한 컬럼별 합계를 그대로 재사용한다(같은 "숫자 컬럼
    판별" 로직을 중복 구현하지 않기 위함). from_col/to_col이 숫자 컬럼이
    아니거나(rows[0]에 없거나 일부 행에 비숫자 값이 섞여 있으면
    compute_column_sums가 애초에 그 컬럼을 대상에서 뺀다), from_col 합계가
    0이면(0으로 나누기 방지) 빈 리스트를 반환한다."""
    if not rows:
        return []
    sums = {r["raw"].removesuffix(" 합계"): Decimal(r["normalized"])
            for r in compute_column_sums(rows, source_file, sheet)}
    if from_col not in sums or to_col not in sums or sums[from_col] == 0:
        return []
    growth = (sums[to_col] - sums[from_col]) / sums[from_col] * 100
    normalized = _decimal_to_normalized_str(growth.quantize(Decimal("0.01")))
    return [{
        "type": "amount", "normalized": normalized,
        "raw": f"{from_col}→{to_col} 증감률",
        "source_file": source_file, "location": f"{sheet}!{from_col}->{to_col}(증감률)",
    }]


def compute_column_average(rows: list[dict], source_file: str, sheet: str, col: str) -> list[dict]:
    """지정한 컬럼의 평균을 계산해 정답 풀에 추가할 항목 하나로 반환한다.
    compute_column_sums()와 같은 반환 dict 형태를 따르며, col이 숫자
    컬럼이 아니면(compute_column_sums가 대상에서 뺐으면) 빈 리스트를
    반환한다."""
    if not rows:
        return []
    sums = compute_column_sums(rows, source_file, sheet)
    matching = [r for r in sums if r["raw"] == f"{col} 합계"]
    if not matching:
        return []
    total = Decimal(matching[0]["normalized"])
    average = total / len(rows)
    normalized = _decimal_to_normalized_str(average.quantize(Decimal("0.01")))
    return [{
        "type": "amount", "normalized": normalized, "raw": f"{col} 평균",
        "source_file": source_file, "location": f"{sheet}!{col}(평균)",
    }]


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


def _selftest_compute_column_growth_rate():
    rows = [{"작년": 100, "올해": 150}, {"작년": 200, "올해": 250}]
    result = compute_column_growth_rate(
        rows, source_file="원본.xlsx", sheet="Sheet1", from_col="작년", to_col="올해"
    )
    assert len(result) == 1, result
    assert result[0]["normalized"] == "33.33", result
    assert result[0]["type"] == "amount", result
    print("compute_column_growth_rate 통과:", result)


def _selftest_compute_column_growth_rate_zero_base_returns_empty():
    """분모(from_col 합계)가 0이면 0으로 나누기를 피해 빈 리스트를 반환해야 한다."""
    rows = [{"작년": 0, "올해": 150}]
    result = compute_column_growth_rate(
        rows, source_file="원본.xlsx", sheet="Sheet1", from_col="작년", to_col="올해"
    )
    assert result == [], result
    print("compute_column_growth_rate(분모 0) 통과: 빈 리스트")


def _selftest_compute_column_average():
    rows = [{"예산": 100}, {"예산": 200}, {"예산": 300}]
    result = compute_column_average(rows, source_file="원본.xlsx", sheet="Sheet1", col="예산")
    assert len(result) == 1, result
    assert result[0]["normalized"] == "200", result
    print("compute_column_average 통과:", result)


def _selftest_compute_column_average_rounds_repeating_decimal():
    """나눗셈이 딱 떨어지지 않아도(100/3 형태) 소수점 둘째자리로 반올림돼야
    한다 - 사람이 읽는 보고서 문장과 대조하기 위함(끝없는 소수는 무의미함)."""
    rows = [{"값": 30}, {"값": 30}, {"값": 40}]
    result = compute_column_average(rows, source_file="원본.xlsx", sheet="Sheet1", col="값")
    assert result[0]["normalized"] == "33.33", result
    print("compute_column_average(반복소수 반올림) 통과:", result)


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
    _selftest_extract_values_excludes_year_abbreviation()
    _selftest_extract_values_excludes_list_marker_dot()
    _selftest_extract_values_excludes_list_marker_paren()
    _selftest_extract_values_list_marker_does_not_exclude_decimal_amount()
    _selftest_extract_values_excludes_month_sequence()
    _selftest_find_month_sequence_spans_detects_full_run()
    _selftest_find_month_sequence_spans_ignores_partial_run()
    _selftest_find_month_sequence_spans_ignores_non_amount_type()
    _selftest_compare_values()
    _selftest_compare_values_catches_digit_transposition_typo()
    _selftest_compare_values_catches_typo_in_large_amount()
    _selftest_compare_values_skips_type_absent_from_source()
    _selftest_compare_values_still_flags_when_type_present_but_no_match()
    _selftest_compare_values_per_type_skip_does_not_hide_other_mismatches()
    _selftest_categorize_values_three_buckets()
    _selftest_compute_column_sums()
    _selftest_compute_column_sums_decimal_no_float_noise()
    _selftest_compute_column_sums_ignores_boolean_column()
    _selftest_compute_column_growth_rate()
    _selftest_compute_column_growth_rate_zero_base_returns_empty()
    _selftest_compute_column_average()
    _selftest_compute_column_average_rounds_repeating_decimal()
    _selftest_check_weekday_consistency_detects_mismatch()
    _selftest_check_weekday_consistency_matches_when_correct()
    _selftest_check_weekday_consistency_ignores_date_without_weekday()
