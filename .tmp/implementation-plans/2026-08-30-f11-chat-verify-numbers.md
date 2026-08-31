# F11 채팅 인터페이스 + 숫자검증 Implementation Plan

> For the agent worker: required sub skill. Implement task by task with `subagent-driven-development` (recommended) or `executing-plans`. Track steps with checkbox (`- [ ]`) syntax.

Goal: 채팅창에 자연어로 "숫자 검증해줘"라고 입력하면, 로컬 LLM(Qwen3)이 verify_numbers 도구를 호출해 원본데이터(엑셀/한글/디지털PDF)와 지금 채팅 도구가 직접 열어둔 한글 보고서를 대조하고, 원본에 없는 금액·날짜·시간·전화번호를 문서 안에 빨간색으로 표시한다.

Architecture: (1) 순수 함수 계층 — 정규식 추출/정규화/대조 로직(verify_numbers.py, 외부 의존성 없음, 가장 먼저 만들고 가장 많이 테스트) → (2) 원본 읽기 계층 — 엑셀/한글/PDF에서 값을 뽑아 (1)의 형식으로 변환(source_reader.py) → (3) 문서 조작 계층 — pyhwpx로 보고서를 열고 빨간색 표시(hwp_report.py) → (4) UI+라우팅 계층 — CustomTkinter 채팅창 + Ollama Qwen3 도구호출(chat_assistant.py, 마지막에 나머지를 다 연결). 이 순서로 만들면 매 단계 콘솔에서 바로 테스트 가능하고, 가장 복잡한 UI/LLM 연동은 안정된 하위 계층 위에서 마지막에 붙인다.

Tech stack: Python 3.13, 정규식(re) + 표준 datetime, openpyxl(엑셀 읽기), pyhwpx(한글 자동화, 기존 F1~F10과 동일), pdfplumber(PDF 텍스트/표 추출), pywin32/win32com(엑셀 COM, 셀 이동용), customtkinter(채팅창 UI), ollama(파이썬 공식 클라이언트, Qwen3 도구호출)

---

## 사전 준비 (Phase 0 — 코드 작성 전 1회성 설치)

이 단계는 TDD 대상이 아니라 환경 설정이라 태스크로 세지 않지만, Task 1을 시작하기 전에 반드시 완료해야 한다.

```bash
ollama pull qwen3:8b
pip install customtkinter pdfplumber ollama openpyxl
```

기존 F1~F10이 이미 pywin32, pyhwpx를 설치해뒀으므로 이번엔 위 4개만 추가로 설치하면 된다. `pip install -r requirements.txt`(기존 파일)에 이 4개를 추가하는 건 Task 9(설치 목록 갱신)에서 처리한다.

---

## 파일 구조

| 파일 | 책임 |
|---|---|
| `verify_numbers.py` (신규) | 정규식 추출, 정규화, 대조 — 외부 파일/네트워크 의존성 전혀 없는 순수 함수 계층 |
| `source_reader.py` (신규) | 엑셀/한글/PDF 파일 또는 폴더를 읽어 `verify_numbers.py`가 쓰는 정답 항목 리스트로 변환. 원본 파일 간 충돌 탐지 포함 |
| `hwp_report.py` (신규) | pyhwpx로 보고서 문서를 열고, 텍스트를 읽고, 특정 텍스트를 빨간색으로 바꾸는 얇은 래퍼 (기존 `fill_core.py`는 서식 채우기용이라 건드리지 않고 별도 파일로 분리) |
| `chat_assistant.py` (신규) | CustomTkinter 채팅창 UI + Ollama Qwen3 도구호출 라우팅 + 위 세 모듈을 엮는 진입점 |
| `requirements.txt` (수정) | 신규 의존성 4개 추가 |
| `README.md` (수정) | F11 실행 방법 추가 |

기존 프로젝트 관례(각 파일을 직접 실행하면 `if __name__ == "__main__":` 블록에서 자체 테스트가 도는 방식, 예: `ai_writer.py`, `format_checker.py`)를 그대로 따른다 — 별도 `tests/` 폴더나 pytest는 도입하지 않는다.

---

### Task 1: 금액 추출·정규화 ✅ 완료 (커밋 8349bd6, 리뷰 반영 741f430)

Files

- Create: `verify_numbers.py`

- [x] Step 1: 실패하는 테스트 작성

`verify_numbers.py` 맨 아래에 아래 코드를 작성한다 (아직 `extract_amounts` 함수는 없으므로 실행하면 실패한다):

```python
"""verify_numbers.py — 원본데이터 대비 보고서 숫자검증 핵심 로직 (순수 함수, 외부 의존성 없음)"""
import re


def _selftest_extract_amounts():
    result = extract_amounts("예산 1,850,000원이고 참가인원 342명, 증가율 12%")
    normalized = [r["normalized"] for r in result]
    assert normalized == ["1850000", "342", "12"], normalized
    print("extract_amounts 통과:", result)


if __name__ == "__main__":
    _selftest_extract_amounts()
```

- [x] Step 2: 실패 확인

Run: `python verify_numbers.py`
Expected: `NameError: name 'extract_amounts' is not defined`

- [x] Step 3: 최소 구현

`_selftest_extract_amounts` 함수 **위에** 아래 코드를 추가한다:

```python
_AMOUNT_PATTERN = re.compile(r'(\d+(?:,\d{3})*(?:\.\d+)?)\s*(만원|천원|원|%)?')
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
```

- [x] Step 4: 통과 확인

Run: `python verify_numbers.py`
Expected: `extract_amounts 통과: [...]` 출력, 에러 없음

- [x] Step 5: 커밋

```bash
git add verify_numbers.py
git commit -m "F11: 금액/일반숫자 추출·단위환산 정규화"
```

(리뷰 반영 커밋 `741f430`: 단위 없는 숫자 뒤 공백이 raw/span에 섞이는 정규식 버그 수정, 만원/천원 배율 검증 테스트 추가)

**(2026-08-30 Task 5 코드품질 검토 중 재발견 — 중요, Task 1로 소급 수정)**: Task 5에서 "오차 허용을 없애고 정확일치로 비교해도 된다"고 결정한 근거가 "normalized는 항상 깨끗한 정수 문자열"이었는데, 이 가정 자체가 틀렸음이 실제로 드러났다. `float(digits) * _UNIT_MULTIPLIER[unit]`처럼 **소수(예: "1.005") × 배율(1000)을 float으로 계산하면 이진 부동소수점 표현 오차**가 생긴다 — 실제로 `extract_amounts("예산은 1.005천원입니다")`를 돌려보면 `normalized`가 `"1004.9999999999999"`로 나온다(정확히는 1005여야 함). 반면 `extract_amounts("예산은 1,005원입니다")`(같은 실제 값을 다르게 표기)는 정확히 `"1005"`가 나와서, Task 5의 정확일치 비교에서 **같은 값인데 다르다고 오탐**하게 된다.

**근본 수정**: float 대신 `decimal.Decimal`로 계산한다. `Decimal`은 십진수를 있는 그대로 정확히 표현하므로(이진 부동소수점처럼 표현 오차가 생기지 않음), "1.005천원" 같은 케이스도 오차 없이 정확히 "1005"가 된다.

**(2026-08-30 코드품질 재검토 후 2차 수정)**: 이 Decimal 변환 로직이 `extract_amounts`와 `compute_column_sums`(Task 6) 양쪽에 똑같이 필요해서, 중복을 없애고 한 곳에서만 고치면 되도록 공용 헬퍼 `_decimal_to_normalized_str`로 뽑는다. 또한 `str(value.normalize())`가 아주 작은 소수(예: 0.0000001)에서 **과학적 표기법("1E-7")을 만들어낼 수 있어**(실제로 `extract_amounts("증감률은 0.0000001%입니다")`로 재현됨), 이러면 Task 5의 정확일치 비교가 "1E-7" vs "0.0000001"처럼 같은 값인데 다른 문자열이 되어 깨진다. `format(value, 'f')`로 바꿔서 항상 고정소수점 표기만 나오게 한다.

```python
from decimal import Decimal

_AMOUNT_PATTERN = re.compile(r'(\d+(?:,\d{3})*(?:\.\d+)?)\s*(만원|천원|원|%)?')
_UNIT_MULTIPLIER = {"만원": 10000, "천원": 1000, "원": 1, "%": 1, None: 1}


def _decimal_to_normalized_str(value: Decimal) -> str:
    """Decimal 값을 정규화된 문자열로 바꾼다. 정수면 소수점 없이, 아니면
    고정소수점 표기로 (과학적 표기법 "1E-7" 같은 건 정확일치 비교를 깨뜨리므로
    항상 고정소수점만 나오게 강제한다).

    (2026-08-30 코드품질 검토 후 3차 수정) 정수 분기(`str(integral)`)도
    안전하지 않았다 — 아주 큰 정수(예: 1e28 이상, Decimal 기본 정밀도 28자리를
    넘는 값)에서는 str()도 과학적 표기법("1E+28")을 낼 수 있음이 실제로
    재현됐다. 두 분기 모두 format(..., 'f')로 통일해서 이 문제를 완전히 막는다.
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
```

추가 회귀 테스트 (`_selftest_extract_amounts` 아래에 추가, 실제 `extract_amounts`를 통해 검증 — 손으로 만든 dict가 아니라 이 함수 자체의 출력으로 확인해야 함):

```python
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
    """아주 큰 정수(Decimal 기본정밀도 28자리를 넘는 값)도 정수 분기에서
    과학적 표기법 없이 고정소수점으로 나와야 한다."""
    assert _decimal_to_normalized_str(Decimal("1E+28")) == "10000000000000000000000000000", \
        _decimal_to_normalized_str(Decimal("1E+28"))
    assert _decimal_to_normalized_str(Decimal("2E+30")) == "2000000000000000000000000000000", \
        _decimal_to_normalized_str(Decimal("2E+30"))
    print("_selftest_decimal_to_normalized_str_large_integral_no_scientific_notation 통과")
```

`__main__`에 `_selftest_extract_amounts_decimal_unit_no_float_noise()`와 `_selftest_extract_amounts_no_scientific_notation()` 둘 다 추가 (기존 `_selftest_extract_amounts()` 호출 바로 다음 자리에).

---

### Task 2: 날짜 3종 추출·정규화 ✅ 완료 (커밋 a4797d8, 리뷰 권고 2건은 Task 4에서 반영 예정: span assert 추가, 2자리연도→20xx 가정 주석)

Files

- Modify: `verify_numbers.py`

- [ ] Step 1: 실패하는 테스트 작성

`_selftest_extract_amounts` 아래에 추가:

```python
def _selftest_extract_dates():
    text = "회의는 2026-09-07, 또는 9월 7일, 혹은 '26.9.7(화)에 진행"
    result = extract_dates(text, default_year=2026)
    normalized = [r["normalized"] for r in result]
    assert normalized == ["2026-09-07", "2026-09-07", "2026-09-07"], normalized
    print("extract_dates 통과:", result)
```

그리고 `if __name__ == "__main__":` 블록에 `_selftest_extract_dates()` 호출을 추가한다:

```python
if __name__ == "__main__":
    _selftest_extract_amounts()
    _selftest_extract_dates()
```

- [ ] Step 2: 실패 확인

Run: `python verify_numbers.py`
Expected: `NameError: name 'extract_dates' is not defined`

- [ ] Step 3: 최소 구현

`extract_amounts` 함수 아래에 추가:

```python
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
```

- [ ] Step 4: 통과 확인

Run: `python verify_numbers.py`
Expected: `extract_dates 통과: [...]` 출력, 에러 없음

- [ ] Step 5: 커밋

```bash
git add verify_numbers.py
git commit -m "F11: 날짜 3종(ISO/한국어/공문서축약형) 추출·정규화"
```

---

### Task 3: 시간 3종 추출·정규화 ✅ 완료 (3라운드 수정 거침 — 8349bd6계열 위 5b51d67→a508761→787c657→7b90107. 최종: 콜론 뒤 숫자구간 span 제외 방식, 분단위 정밀비교)

Files

- Modify: `verify_numbers.py`

- [x] Step 1: 실패하는 테스트 작성

**(2026-08-30 코드품질 검토 후 수정됨)**: 최초 버전은 분(分)을 버리고 시(時)만 비교했으나, 이러면 "14:30~16:45"(분 단위 오차)를 원본과 다른데도 같다고 오판할 위험이 있어 사용자 확인 후 분까지 비교하도록 강화함. 또한 검토자가 `_overlaps`/`consumed_spans` 메커니즘이 세 패턴 구조상 실제로는 결코 트리거되지 않는 죽은 코드임을 직접 증명했고, 대신 콜론형의 분(分) 자릿수를 뒤의 두 패턴이 시(時)로 잘못 읽어버리는 실제 버그(예: "14:00~16시" 같은 혼합/손상 표기에서 "00"을 시작 시각으로 오인)를 발견했다. 이번 버전은 그 죽은 코드를 제거하고, 대신 정규식 자체에 `(?<!:)` 부정 후방탐색으로 콜론 직후 숫자에서 매치가 시작되지 못하게 막는다.

```python
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
```

`__main__` 블록에 `_selftest_extract_times()`, `_selftest_extract_times_minute_precision()`, `_selftest_extract_times_mixed_notation_no_false_match()` 3개 전부 추가.

- [ ] Step 2: 실패 확인

Run: `python verify_numbers.py`
Expected: `NameError: name 'extract_times' is not defined`

- [ ] Step 3: 최소 구현

**(2026-08-30 재수정)**: 최초의 `(?<!:)` 가드 하나만으로는 부족함이 구현 중 실제로 드러났다 — 분(分)이 2자리(예: "00")일 때, 그 **두 번째 자릿수**에서 매치가 시작되는 경우("14:00~16시"에서 "0~16시"로 오인)는 `(?<!:)`가 막지 못한다(그 위치 바로 앞 글자는 ':'가 아니라 숫자이므로). 콜론 뒤 숫자 1개 또는 2개 전 위치 모두를 막도록 `(?<!:)(?<!:\d)` 두 개의 부정 후방탐색을 겹쳐 쓴다.

**(2026-08-30 3차 수정 — 최종)**: 이중 부정 후방탐색도 여전히 불완전함이 드러났다 — Python 정규식은 **가변 길이 후방탐색을 지원하지 않아서**, `(?<!:)(?<!:\d)`는 콜론 뒤 정확히 1~2자리까지만 막고 3자리 이상(예: "14:100~16시" 같은 오타)은 여전히 오독한다. 후방탐색을 계속 겹쳐 쌓는 방식은 근본적으로 일반화될 수 없으므로, **콜론 뒤 숫자 구간의 위치를 정규식으로 직접 찾아, 그 구간 안에서 시작하는 매치는 길이에 상관없이 전부 제외**하는 방식으로 바꾼다.

```python
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
```

추가 회귀 테스트(3자리 이상 분 구간도 막히는지 확인, 기존 `_selftest_extract_times_mixed_notation_no_false_match` 아래에 추가):

```python
def _selftest_extract_times_long_digit_run_no_false_match():
    """콜론 뒤 숫자가 3자리 이상(오타 등)이어도 시로 오독하지 않아야 한다."""
    assert extract_times("14:100~16시") == []
    assert extract_times("14:009~16시") == []
    assert extract_times("09:001~10시") == []
    assert extract_times("14:100시~16시") == []
    print("_selftest_extract_times_long_digit_run_no_false_match 통과")
```

`__main__`에 `_selftest_extract_times_long_digit_run_no_false_match()`도 추가.

- [x] Step 4: 통과 확인 (최종 커밋 7b90107 기준, 8개 셀프테스트 전부 통과)

- [x] Step 5: 커밋 (최종 7b90107, 중간 경과: 5b51d67 → a508761 → 787c657 → 7b90107)

---

### Task 4: 전화번호 추출 + 통합 extract_values 함수 ✅ 완료 (c0ea3d6 → 30c2377(겹침버그수정) → 7fabc11(테스트보강), 12개 셀프테스트 전부 통과)

Files

- Modify: `verify_numbers.py`

- [ ] Step 1: 실패하는 테스트 작성

```python
def _selftest_extract_values():
    text = "담당자 031-1234-5678, 예산 1,850,000원, 2026-09-07 14:00~16:00 진행"
    result = extract_values(text, default_year=2026)
    types = sorted(r["type"] for r in result)
    assert types == ["amount", "date", "phone", "time"], types
    phone = [r for r in result if r["type"] == "phone"][0]
    assert phone["normalized"] == "031-1234-5678", phone
    print("extract_values 통과:", result)
```

`__main__` 블록에 `_selftest_extract_values()` 추가.

- [ ] Step 2: 실패 확인

Run: `python verify_numbers.py`
Expected: `NameError: name 'extract_values' is not defined`

- [ ] Step 3: 최소 구현

**중요**: 전화번호·날짜·시간을 먼저 뽑고, 그 범위(span)를 제외한 나머지에서만 금액을 뽑아야 한다 (과거 F7이 날짜의 하이픈 숫자를 금액으로 잘못 인식했던 버그 재발 방지 — PRD 12-4 참고).

**(2026-08-30 코드품질 검토 후 수정됨)**: 최초 버전의 `_in_excluded`는 금액 매치의 **시작 위치만** 제외 구간 안에 있는지 확인했다(`s <= span[0] < e`). 그런데 구분자 없이 붙어있는 경우(예: "12026-09-07" — 항목번호 "1" 바로 뒤에 날짜가 공백 없이 붙음), 금액 매치("12026")가 제외 구간보다 **먼저 시작해서 겹치기만** 하면 시작 위치 검사로는 못 잡아서 날짜의 "2026"이 그대로 금액에 섞여 들어간다(같은 자릿수가 두 번 카운트됨 — 이게 바로 Task 4가 막으려던 버그 그 자체). **시작 위치가 아니라 구간이 조금이라도 겹치는지(interval overlap)** 검사하도록 바꾼다.

```python
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
```

추가 회귀 테스트 (구분자 없이 붙은 경우도 새지 않는지 확인, `_selftest_extract_values` 아래에 추가):

```python
def _selftest_extract_values_adjacent_no_separator():
    """구분자 없이 붙어있어 금액 매치가 제외구간보다 먼저 시작하며 겹치는 경우도
    걸러내야 한다 (시작 위치만 보는 검사로는 놓치는 케이스)."""
    result = extract_values("1" + "2026-09-07 회의", default_year=2026)
    amounts = [r for r in result if r["type"] == "amount"]
    assert amounts == [], amounts
    print("_selftest_extract_values_adjacent_no_separator 통과:", result)
```

`__main__`에 `_selftest_extract_values_adjacent_no_separator()`도 추가.

**(2026-08-30 재검토 후 테스트 보강)**: 위 테스트는 "금액이 제외구간보다 먼저 시작해서 겹치는" 방향만 검증한다. 반대 방향(금액이 제외구간 **안에서 시작해서 밖으로 걸치는** 경우)과 "제외구간 두 개가 붙어있고 그 경계를 금액이 걸치는" 경우도 추가한다:

```python
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
```

`__main__`에 두 함수 호출 모두 추가.

- [ ] Step 4: 통과 확인

Run: `python verify_numbers.py`
Expected: `extract_values 통과: [...]` 출력, 에러 없음

- [ ] Step 5: 커밋

```bash
git add verify_numbers.py
git commit -m "F11: 전화번호 추출 + 4종 통합 extract_values (날짜/시간/전화 우선 제외 후 금액 추출)"
```

---

### Task 5: 대조 로직 (오차·컬럼합계 포함) ✅ 완료 (7d4ec85 → ea97ddf → 23289ec(오차허용 전면제거, 정확일치로 전환) → 9dfc2ae → 74608dc(불리언크래시+과학적표기법 수정), 21개 셀프테스트 전부 통과. Task 6도 함께 완료됨)

Files

- Modify: `verify_numbers.py`

- [ ] Step 1: 실패하는 테스트 작성

```python
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
```

`__main__` 블록에 `_selftest_compare_values()` 추가.

- [ ] Step 2: 실패 확인

Run: `python verify_numbers.py`
Expected: `NameError: name 'compare_values' is not defined`

- [ ] Step 3: 최소 구현

**(2026-08-30 코드품질 검토 후 수정됨 — 중요, 2차)**: 1차 수정(±0.5%→±0.0001%)도 여전히 근본 결함이 있었다 — **상대오차(%) 방식 자체가 문제**다. 오차 허용을 아무리 좁혀도, 금액이 커질수록(수억~수십억원) 그 비율에 해당하는 절대 액수는 커지므로, 큰 금액에서는 여전히 수백~수천원 단위 오타가 허용범위 안에 숨을 수 있다 (예: 10억원 중 마지막 한 자리 오타 999→998은 상대차이가 1e-10 수준으로, 0.0001%보다도 훨씬 작음). 임계값을 아무리 조정해도 "비율" 방식으로는 이 문제를 근본적으로 못 없앤다.

다시 생각해보면, **오차 허용 자체가 애초에 필요 없다.** 정규화된 금액(`normalized`)은 전부 단위환산까지 끝난 깨끗한 정수 문자열이고(Task 1 참고: `str(int(value))`), 이 도구가 다루는 범위(정부 예산 규모, 최대 수조 원 수준)는 배정밀도 float가 정수를 오차 없이 표현하는 한계(2^53 ≈ 9천조)에 전혀 못 미친다. 즉 **부동소수점 계산 잡음이 생길 여지 자체가 실질적으로 없다.** 컬럼 합계(Task 6)도 정수만 더하므로 결과 역시 정확한 정수다. 따라서 오차 허용 메커니즘을 통째로 없애고, **금액도 다른 3종과 똑같이 정확히 일치하는지만** 본다 — 이러면 어떤 크기의 금액에서도 오타가 숨을 곳이 없다.

```python
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
```

추가 회귀 테스트(자릿수 교체형 오타가, 특히 아주 큰 금액에서도 잡히는지 확인, `_selftest_compare_values` 아래에 추가):

```python
def _selftest_compare_values_catches_digit_transposition_typo():
    """자릿수를 바꿔 쓴 전형적인 오타는 어떤 크기의 금액에서도 잡혀야 한다."""
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
```

`__main__`에 `_selftest_compare_values_catches_digit_transposition_typo()`와 `_selftest_compare_values_catches_typo_in_large_amount()` 둘 다 추가.

- [ ] Step 4: 통과 확인

Run: `python verify_numbers.py`
Expected: `compare_values 통과: [...]` 출력, 에러 없음

- [ ] Step 5: 커밋

```bash
git add verify_numbers.py
git commit -m "F11: 원본 정답풀 대조 로직 (금액 오차허용, 날짜/시간/전화 정확일치)"
```

---

### Task 6: 컬럼 합계 파생값 계산 ✅ 완료 (Task 5와 같은 커밋 계열에서 완료, 7d20e95→9dfc2ae→74608dc)

Files

- Modify: `verify_numbers.py`

- [ ] Step 1: 실패하는 테스트 작성

```python
def _selftest_compute_column_sums():
    rows = [{"예산": 500000, "인원": 10}, {"예산": 300000, "인원": 20}]
    sums = compute_column_sums(rows, source_file="원본.xlsx", sheet="Sheet1")
    normalized = sorted(r["normalized"] for r in sums)
    assert normalized == ["30", "800000"], normalized
    print("compute_column_sums 통과:", sums)
```

`__main__` 블록에 `_selftest_compute_column_sums()` 추가.

**(2026-08-30 미리 반영)** 소수 컬럼 합계가 float 오차 없이 정확한지 확인하는 회귀 테스트도 같이 추가한다:

```python
def _selftest_compute_column_sums_decimal_no_float_noise():
    """소수를 포함하는 컬럼의 합계도 float 오차 없이 정확해야 한다."""
    rows = [{"단가": 0.1}, {"단가": 0.2}]
    sums = compute_column_sums(rows, source_file="원본.xlsx", sheet="Sheet1")
    assert sums[0]["normalized"] == "0.3", sums
    print("_selftest_compute_column_sums_decimal_no_float_noise 통과:", sums)
```

`__main__`에 `_selftest_compute_column_sums_decimal_no_float_noise()`도 추가.

- [ ] Step 2: 실패 확인

Run: `python verify_numbers.py`
Expected: `NameError: name 'compute_column_sums' is not defined`

- [ ] Step 3: 최소 구현

**(2026-08-30, Task 5 검토에서 미리 반영)**: Task 1을 Decimal 기반으로 고친 것과 같은 이유로, 여기서도 `sum()`을 순수 float 리스트에 바로 쓰면 안 된다 — 예를 들어 단가·비율처럼 소수를 포함하는 컬럼(`[0.1, 0.2]`)을 그냥 `sum()`하면 `0.30000000000000004`처럼 이진 부동소수점 오차가 생겨, 보고서의 "0.3원"과 정확일치 비교에서 오탐이 난다. 각 값을 `Decimal(str(v))`로 변환해서 더한다.

**(2026-08-30 코드품질 검토 후 2차 수정 — 중요)**: 실제 돌려보니 **불리언(True/False) 컬럼에서 크래시**가 났다 — 파이썬에서 `bool`은 `int`의 서브클래스라 `isinstance(v, (int, float))` 필터를 그대로 통과하는데, `Decimal(str(True))` = `Decimal("True")`는 파싱 불가능해 `decimal.InvalidOperation` 예외가 난다. "완료여부"/"참석여부" 같은 예/아니오 컬럼은 실제 공공기관 엑셀에 흔하므로, **불리언은 명시적으로 제외**해야 한다. 또한 Task 1과 같은 이유로 정규화도 공용 헬퍼 `_decimal_to_normalized_str`를 재사용한다(과학적 표기법 방지, 중복 제거).

```python
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
```

추가 회귀 테스트(불리언 컬럼이 섞여도 크래시 없이 정상 동작하는지 확인, `_selftest_compute_column_sums_decimal_no_float_noise` 아래에 추가):

```python
def _selftest_compute_column_sums_ignores_boolean_column():
    """완료여부 같은 True/False 컬럼이 섞여 있어도 크래시 없이 무시해야 한다."""
    rows = [{"예산": 100, "완료": True}, {"예산": 200, "완료": False}]
    sums = compute_column_sums(rows, source_file="원본.xlsx", sheet="Sheet1")
    normalized = sorted(r["normalized"] for r in sums)
    assert normalized == ["300"], normalized  # "완료" 컬럼은 합계 대상에서 제외됨
    print("_selftest_compute_column_sums_ignores_boolean_column 통과:", sums)
```

`__main__`에 `_selftest_compute_column_sums_ignores_boolean_column()`도 추가.

- [ ] Step 4: 통과 확인

Run: `python verify_numbers.py`
Expected: `compute_column_sums 통과: [...]` 출력, 에러 없음

- [ ] Step 5: 커밋

```bash
git add verify_numbers.py
git commit -m "F11: 엑셀 컬럼 단순합계 파생값 계산 (증감률/평균은 비목표로 제외)"
```

`verify_numbers.py`는 이걸로 완성이다 — 이 파일은 순수 함수만 담고 있어 엑셀/한글/PDF/UI 어느 것도 몰라야 한다. 다음 태스크부터 실제 파일을 읽는 계층을 만든다.

---

### Task 7: 엑셀 원본 읽기 ✅ 완료 (58c8299 → 7db9d84(날짜셀 누락 핵심버그 수정+손상파일 예외처리) → 681c232(시간단독셀 한계 문서화))

Files

- Create: `source_reader.py`

- [ ] Step 1: 실패하는 테스트 작성 (샘플 엑셀 파일을 코드로 직접 생성해서 테스트)

```python
"""source_reader.py — 엑셀/한글/PDF 원본데이터를 읽어 verify_numbers의 정답 풀 형식으로 변환"""
import os
import openpyxl
from verify_numbers import extract_values, compute_column_sums


def _selftest_read_excel_source():
    test_path = "_test_원본.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "실적"
    ws.append(["예산", "인원"])
    ws.append([500000, 10])
    ws.append([300000, 20])
    wb.save(test_path)
    try:
        result = read_excel_source(test_path, default_year=2026)
        amounts = sorted(r["normalized"] for r in result if r["type"] == "amount")
        assert "500000" in amounts and "300000" in amounts and "800000" in amounts, amounts
        assert all(r["source_file"] == test_path for r in result)
        print("read_excel_source 통과:", len(result), "건")
    finally:
        os.remove(test_path)


if __name__ == "__main__":
    _selftest_read_excel_source()
```

- [ ] Step 2: 실패 확인

Run: `python source_reader.py`
Expected: `NameError: name 'read_excel_source' is not defined`

- [ ] Step 3: 최소 구현

`_selftest_read_excel_source` **위에** 추가:

**(2026-08-30 코드품질 검토 후 수정됨 — 중요)**: 최초 버전은 엑셀의 **날짜 타입 셀**(openpyxl이 `datetime.date`/`datetime.datetime`으로 돌려주는 값)을 `int/float`도 `str`도 아니라서 **조용히 누락**시켰다. 날짜는 이 도구가 검증하는 4종 중 하나인데, 원본에 날짜가 실제로 있어도 정답 풀에 안 들어가서 "확인 안 됨"으로 오탐이 나는 심각한 문제다. `datetime.date`/`datetime.datetime` 분기를 추가해서 `extract_dates`와 같은 형식(`YYYY-MM-DD`)으로 정규화한다.

또한 `openpyxl.load_workbook`이 손상된 파일에서 예외를 던지면 전체가 죽는데, 이미 만든 `read_pdf_source`(Task 9)는 "읽을 수 없으면 조용히 건너뜀"으로 설계했으므로 여기도 맞춰서 try/except로 감싼다 (Task 10 폴더스캔이 이 "건너뜀" 동작에 의존하게 됨).

```python
import datetime


def read_excel_source(path: str, default_year: int) -> list[dict]:
    """엑셀 파일의 모든 시트, 모든 셀 값을 정답 풀 항목으로 변환한다.
    셀 값이 순수 숫자면 amount로, 날짜 타입이면 date로, 문자열이면 extract_values로
    4종을 뽑는다. 추가로 각 시트의 숫자 컬럼(1행이 헤더라고 가정) 합계도 파생값으로
    포함한다. 손상되었거나 읽을 수 없는 파일은 예외 없이 빈 리스트를 반환한다
    (PRD 12-5, read_pdf_source와 동일한 관례).

    알려진 한계(문서화만 하고 이번엔 해결하지 않음): 수식 셀은 data_only=True로
    열어도 실제 Excel에서 한 번도 저장된 적 없으면 캐시된 값이 없어 None으로
    읽힐 수 있다(합계 등 파생값이 원본에 있어도 못 읽는 경우 발생 가능). 병합된
    제목 행이 실제 헤더보다 위에 있는 레이아웃도 1행=헤더 가정과 어긋날 수 있다.
    """
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception:
        return []
    results = []
    for ws in wb.worksheets:
        rows_as_dicts = []
        header = None
        for row in ws.iter_rows(values_only=False):
            if header is None:
                header = [c.value for c in row]
                continue
            row_dict = {}
            for col_name, cell in zip(header, row):
                value = cell.value
                if value is None:
                    continue
                if isinstance(value, (datetime.date, datetime.datetime)):
                    # (2026-08-30 재검토 후 추가) 엑셀 날짜 타입 셀 — extract_dates와
                    # 같은 형식(YYYY-MM-DD)으로 정규화해야 정답 풀에서 날짜로 인식된다.
                    results.append({
                        "type": "date", "normalized": value.strftime("%Y-%m-%d"),
                        "raw": str(value), "source_file": path, "location": f"{ws.title}!{cell.coordinate}",
                    })
                elif isinstance(value, (int, float)) and not isinstance(value, bool):
                    # (2026-08-30 Task 5/6 검토에서 미리 반영) bool은 int의 서브클래스라
                    # 별도 제외 필요. 정규화도 float 대신 Decimal 기반 공용 헬퍼를 써서
                    # 부동소수점 오차·과학적 표기법 문제를 처음부터 피한다.
                    results.append({
                        "type": "amount", "normalized": _decimal_to_normalized_str(Decimal(str(value))),
                        "raw": str(value), "source_file": path, "location": f"{ws.title}!{cell.coordinate}",
                    })
                    row_dict[col_name] = value
                elif isinstance(value, str):
                    for v in extract_values(value, default_year):
                        v["source_file"] = path
                        v["location"] = f"{ws.title}!{cell.coordinate}"
                        del v["span"]
                        results.append(v)
            if row_dict:
                rows_as_dicts.append(row_dict)
        results.extend(compute_column_sums(rows_as_dicts, source_file=path, sheet=ws.title))
    return results
```

추가 회귀 테스트(엑셀 날짜 타입 셀이 누락 없이 인식되는지, `_selftest_read_excel_source` 아래에 추가):

```python
def _selftest_read_excel_source_date_cell():
    """엑셀의 날짜 타입 셀(datetime.date)이 date 항목으로 정확히 인식되어야 한다."""
    test_path = "_test_원본_날짜.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "일정"
    ws.append(["착수일", "예산"])
    ws.append([datetime.date(2026, 9, 7), 500000])
    wb.save(test_path)
    try:
        result = read_excel_source(test_path, default_year=2026)
        dates = [r for r in result if r["type"] == "date"]
        assert len(dates) == 1 and dates[0]["normalized"] == "2026-09-07", dates
        print("read_excel_source_date_cell 통과:", dates)
    finally:
        os.remove(test_path)


def _selftest_read_excel_source_corrupted_file_no_crash():
    """손상된/엑셀이 아닌 파일도 예외 없이 빈 리스트를 반환해야 한다."""
    test_path = "_test_손상됨.xlsx"
    with open(test_path, "w") as f:
        f.write("이건 엑셀 파일이 아님")
    try:
        result = read_excel_source(test_path, default_year=2026)
        assert result == [], result
        print("read_excel_source_corrupted_file_no_crash 통과:", result)
    finally:
        os.remove(test_path)
```

`__main__`에 `_selftest_read_excel_source_date_cell()`과 `_selftest_read_excel_source_corrupted_file_no_crash()` 둘 다 추가.

- [ ] Step 4: 통과 확인

Run: `python source_reader.py`
Expected: `read_excel_source 통과: N 건` 출력, 에러 없음

- [ ] Step 5: 커밋

```bash
git add source_reader.py
git commit -m "F11: 엑셀 원본 읽기 (셀값+컬럼합계를 정답풀 형식으로 변환)"
```

---

### Task 8: 한글(HWP) 원본 읽기 ✅ 완료 (62de8e8 → 1f21977(프로세스누수 수정+실패시 빈리스트 통일+위치스니펫 매치표시), 6개 셀프테스트 전부 통과)

Files

- Modify: `source_reader.py`

- [ ] Step 1: 실패하는 테스트 작성

이 테스트는 **실제 한컴오피스가 설치된 이 PC에서만** 통과한다 (pyhwpx가 한글 프로그램을 실제로 구동하기 때문). 기존 F1~F10과 동일한 전제다.

```python
def _selftest_read_hwp_source():
    from pyhwpx import Hwp
    test_path = os.path.abspath("_test_원본.hwp")
    hwp = Hwp(visible=False)
    hwp.insert_text("예산은 1,850,000원이며 회의는 2026-09-07 14:00~16:00 진행")
    hwp.save_as(test_path)
    hwp.quit()
    try:
        result = read_hwp_source(test_path, default_year=2026)
        types = sorted(r["type"] for r in result)
        assert types == ["amount", "date", "time"], types
        print("read_hwp_source 통과:", result)
    finally:
        os.remove(test_path)
```

`__main__` 블록에 추가:
```python
if __name__ == "__main__":
    _selftest_read_excel_source()
    _selftest_read_hwp_source()
```

- [ ] Step 2: 실패 확인

Run: `python source_reader.py`
Expected: `NameError: name 'read_hwp_source' is not defined`

- [ ] Step 3: 최소 구현

**(2026-08-30 코드품질 검토 후 수정됨 — 중요)**: 최초 버전은 `hwp.quit()`이 `try/finally` 없이 호출돼서, `hwp.open`/`GetTextFile` 사이에 예외가 나면(파일 없음, 손상 등) **보이지 않는 Hwp.exe 프로세스가 그대로 남는 문제**가 실제로 재현됐다 (이 태스크 개발 중에도 한 번 실제로 이 문제 때문에 혼란스러운 COM 에러가 났었음). 또한 실패 시 예외를 그대로 던져서, `read_excel_source`/`read_pdf_source`가 따르는 "읽을 수 없으면 조용히 빈 리스트 반환"(PRD 12-5) 관례와도 어긋났다. `try/finally`로 감싸 항상 `quit()`이 호출되게 하고, 실패 시 빈 리스트를 반환하도록 통일한다. 위치 스니펫도 여러 값이 가까이 붙어 있을 때 어떤 숫자가 실제로 걸린 건지 구분이 안 되는 문제가 있어, 매치된 부분을 `**`로 표시해 구분되게 한다.

```python
def read_hwp_source(path: str, default_year: int) -> list[dict]:
    """한글 문서를 안 보이게(visible=False) 열어서 전체 텍스트를 읽고,
    4종 값을 뽑아 정답 풀 항목으로 변환한다. 원본 참고용으로만 열기 때문에
    화면에 띄우지 않는다 (사용자가 실제로 편집 중인 보고서와는 별개의 인스턴스).
    손상되었거나 열 수 없는 파일은 예외 없이 빈 리스트를 반환한다 (PRD 12-5,
    read_excel_source/read_pdf_source와 동일한 관례). hwp.quit()은 어떤 예외가
    나든 반드시 호출되도록 try/finally로 감싼다 — 그렇지 않으면 보이지 않는
    Hwp.exe 프로세스가 누적되어 남는 문제가 실제로 재현된 적 있다.
    """
    from pyhwpx import Hwp
    hwp = None
    try:
        hwp = Hwp(visible=False)
        if not hwp.open(path):
            return []
        text = hwp.GetTextFile("TEXT", "")
        if not text:
            return []
    except Exception:
        return []
    finally:
        if hwp is not None:
            hwp.quit()

    results = extract_values(text, default_year)
    for r in results:
        start, end = r["span"]
        before = text[max(0, start - 10):start]
        matched = text[start:end]
        after = text[end:end + 10]
        r["source_file"] = path
        r["location"] = f"{before}**{matched}**{after}"
        del r["span"]
    return results
```

추가 회귀 테스트(무매치 문서, 존재하지 않는 파일 — 둘 다 크래시·프로세스 누수 없이 빈 리스트를 반환하는지 확인, `_selftest_read_hwp_source` 아래에 추가):

```python
def _selftest_read_hwp_source_no_match():
    """숫자/날짜/시간/전화번호가 전혀 없는 문서는 빈 리스트를 반환해야 한다."""
    from pyhwpx import Hwp
    test_path = os.path.abspath("_test_원본_빈값.hwp")
    hwp = Hwp(visible=False)
    hwp.insert_text("이 문서에는 특별한 값이 없습니다")
    hwp.save_as(test_path)
    hwp.quit()
    try:
        result = read_hwp_source(test_path, default_year=2026)
        assert result == [], result
        print("read_hwp_source_no_match 통과:", result)
    finally:
        os.remove(test_path)


def _selftest_read_hwp_source_missing_file_no_crash():
    """존재하지 않는 파일은 예외 없이 빈 리스트를 반환해야 한다 (프로세스 누수 방지 포함)."""
    result = read_hwp_source(os.path.abspath("_존재하지_않는_파일.hwp"), default_year=2026)
    assert result == [], result
    print("read_hwp_source_missing_file_no_crash 통과:", result)
```

`__main__`에 `_selftest_read_hwp_source_no_match()`와 `_selftest_read_hwp_source_missing_file_no_crash()` 둘 다 추가.

- [ ] Step 4: 통과 확인

Run: `python source_reader.py`
Expected: `read_hwp_source 통과: [...]` 출력, 에러 없음

- [ ] Step 5: 커밋

```bash
git add source_reader.py
git commit -m "F11: 한글(hwp) 원본 읽기, visible=False로 백그라운드에서 텍스트만 추출"
```

---

### Task 9: 디지털 PDF 원본 읽기 ✅ 완료 (db7cbae → 6d5acfe(docstring 정정+fpdf2 회귀테스트 자동화), 9개 셀프테스트 전부 통과)

Files

- Modify: `source_reader.py`

- [ ] Step 1: 실패하는 테스트 작성

이 테스트는 pdfplumber만으로 PDF를 생성할 수 없으므로, reportlab 없이 **텍스트 추출 실패 케이스(스캔 PDF 흉내)**와 **정상 텍스트 PDF는 별도 준비된 파일로 수동 검증**을 나눈다. 자동화 가능한 부분만 먼저 작성한다.

```python
def _selftest_read_pdf_source_missing_file():
    result = read_pdf_source("존재하지_않는_파일.pdf", default_year=2026)
    assert result == [], result
    print("read_pdf_source(없는 파일) 통과: 빈 리스트, 오류 없이 건너뜀")
```

`__main__` 블록에 추가.

- [ ] Step 2: 실패 확인

Run: `python source_reader.py`
Expected: `NameError: name 'read_pdf_source' is not defined`

- [ ] Step 3: 최소 구현

**(2026-08-30 코드품질 검토 후 수정됨)**: 최초 버전의 docstring이 "텍스트·표를 추출"이라고 과장돼 있었는데, 실제 구현은 `extract_text()`만 쓰고 `extract_tables()`는 안 쓴다. 실제로 여러 표 케이스로 테스트해보니 흔한 케이스(칸이 잘 구분된 표)는 텍스트 추출만으로도 값을 잘 잡아내서, 이번엔 표 전용 파싱 기능을 새로 추가하지 않고 **docstring을 사실대로 정정 + 한계로 문서화**하는 것으로 정리한다. 또한 실제로는 `fpdf2`가 이미 설치돼 있어 합성 PDF를 만들어 자동 테스트할 수 있었는데, 계획서가 "pdfplumber만으로는 PDF를 못 만든다"고 잘못 가정해서 수동검증으로 미뤄뒀던 부분도 자동 테스트로 승격한다 (여러 페이지 중 빈 페이지가 껴 있어도 나머지가 정상 처리되는지, 표 형태 데이터도 잡히는지).

```python
def read_pdf_source(path: str, default_year: int) -> list[dict]:
    """디지털 PDF에서 텍스트를 추출해 정답 풀 항목으로 변환한다. 표 형태 데이터는
    별도 표 파싱(extract_tables) 없이 흐르는 텍스트로 추출하는데, 칸이 잘
    구분된 단순한 표는 이 방식으로도 값이 잘 잡히지만, 여러 줄에 걸친 셀이나
    칸 구분이 애매한 복잡한 표는 놓칠 수 있다(알려진 한계, 이번엔 해결 안 함).
    파일이 없거나, 텍스트를 전혀 추출할 수 없는 스캔 이미지 PDF인 경우
    예외를 던지지 않고 빈 리스트를 반환한다 (PRD 12-2: OCR 비목표, 조용히 건너뜀).
    """
    import pdfplumber
    if not os.path.exists(path):
        return []
    results = []
    try:
        with pdfplumber.open(path) as pdf:
            for page_num, page in enumerate(pdf.pages, start=1):
                text = page.extract_text() or ""
                if not text.strip():
                    continue  # 텍스트 없음 = 스캔 이미지로 추정, 건너뜀
                for v in extract_values(text, default_year):
                    v["source_file"] = path
                    v["location"] = f"{page_num}페이지"
                    del v["span"]
                    results.append(v)
    except Exception:
        return []  # 손상되었거나 읽을 수 없는 PDF는 조용히 건너뜀 (PRD 12-5)
    return results
```

추가 회귀 테스트(다중 페이지 중 빈 페이지 스킵, 단순 표 인식 — `fpdf2`로 합성 PDF를 만들어 자동 검증. `_selftest_read_pdf_source_missing_file` 아래에 추가):

```python
def _selftest_read_pdf_source_multipage_skips_blank():
    """여러 페이지 중 빈 페이지(스캔 이미지 가정)가 껴 있어도, 나머지 페이지는
    정상적으로 처리되어야 한다."""
    from fpdf import FPDF
    test_path = "_test_원본_다중페이지.pdf"
    pdf = FPDF()
    pdf.add_page(); pdf.set_font("Helvetica", size=12)
    pdf.cell(0, 10, "예산은 1,850,000원입니다")
    pdf.add_page()  # 빈 페이지 (텍스트 없음)
    pdf.add_page(); pdf.set_font("Helvetica", size=12)
    pdf.cell(0, 10, "인원은 342명입니다")
    pdf.output(test_path)
    try:
        result = read_pdf_source(test_path, default_year=2026)
        locations = sorted(r["location"] for r in result)
        assert locations == ["1페이지", "3페이지"], locations  # 2페이지(빈 페이지)는 건너뜀
        print("read_pdf_source_multipage_skips_blank 통과:", result)
    finally:
        os.remove(test_path)


def _selftest_read_pdf_source_simple_table():
    """칸이 잘 구분된 단순 표 형태 데이터도 텍스트 추출로 값이 잡혀야 한다."""
    from fpdf import FPDF
    test_path = "_test_원본_표.pdf"
    pdf = FPDF()
    pdf.add_page(); pdf.set_font("Helvetica", size=12)
    pdf.cell(60, 10, "항목", border=1); pdf.cell(60, 10, "금액", border=1); pdf.ln()
    pdf.cell(60, 10, "예산", border=1); pdf.cell(60, 10, "1850000", border=1); pdf.ln()
    pdf.output(test_path)
    try:
        result = read_pdf_source(test_path, default_year=2026)
        amounts = [r["normalized"] for r in result if r["type"] == "amount"]
        assert "1850000" in amounts, amounts
        print("read_pdf_source_simple_table 통과:", result)
    finally:
        os.remove(test_path)
```

`__main__`에 `_selftest_read_pdf_source_multipage_skips_blank()`와 `_selftest_read_pdf_source_simple_table()` 둘 다 추가.

- [ ] Step 4: 통과 확인

Run: `python source_reader.py`
Expected: `read_pdf_source(없는 파일) 통과: ...` 출력, 에러 없음

- [ ] Step 6: 커밋

```bash
git add source_reader.py
git commit -m "F11: 디지털 PDF 원본 읽기 (스캔 PDF·손상 파일은 예외 없이 빈 리스트)"
```

---

### Task 10: 폴더 스캔 + 원본 파일 간 충돌 탐지 ✅ 완료 (ee93123 → 5c187e0(location+type 튜플로 그룹핑, 셀 내 다중타입 충돌 오탐 수정) → 080e52b(코드품질 리뷰 반영 - 충돌 결과에 type 포함, docstring에 location 비고유 명시), 2개 셀프테스트 전부 통과)

Files

- Modify: `source_reader.py`

- [ ] Step 1: 실패하는 테스트 작성

```python
def _selftest_read_source_folder_conflict():
    os.makedirs("_test_원본폴더", exist_ok=True)
    wb1 = openpyxl.Workbook(); ws1 = wb1.active; ws1.title = "Sheet1"
    ws1["A1"] = "예산"; ws1["A2"] = 1800000
    wb1.save("_test_원본폴더/초안.xlsx")
    wb2 = openpyxl.Workbook(); ws2 = wb2.active; ws2.title = "Sheet1"
    ws2["A1"] = "예산"; ws2["A2"] = 1850000
    wb2.save("_test_원본폴더/최종.xlsx")
    with open("_test_원본폴더/무관한파일.txt", "w") as f:
        f.write("이건 읽으면 안 되는 파일")
    try:
        pool, conflicts = read_source_folder("_test_원본폴더", default_year=2026)
        assert len(conflicts) == 1, conflicts
        assert conflicts[0]["location"] == "Sheet1!A2", conflicts
        print("read_source_folder 통과: 충돌", conflicts)
    finally:
        import shutil
        shutil.rmtree("_test_원본폴더")
```

`__main__` 블록에 추가.

- [ ] Step 2: 실패 확인

Run: `python source_reader.py`
Expected: `NameError: name 'read_source_folder' is not defined`

- [ ] Step 3: 최소 구현

```python
_READERS = {".xlsx": read_excel_source, ".xls": read_excel_source,
            ".hwp": read_hwp_source, ".hwpx": read_hwp_source,
            ".pdf": read_pdf_source}


def read_source_folder(folder_path: str, default_year: int) -> tuple[list[dict], list[dict]]:
    """폴더 안 모든 파일을 읽되, 처리 가능한 확장자만 실제로 읽고 나머지는 조용히
    건너뛴다. 서로 다른 엑셀 파일의 "같은 시트!같은 셀 주소"에 다른 값이 있으면
    충돌로 간주해 별도 리스트로 반환한다 (PRD 12-5).

    반환: (정답_풀, 충돌_목록)
    충돌_목록 항목 형식: {"location": "Sheet1!A2", "type": "amount", "values": [
        {"file": "초안.xlsx", "normalized": "1800000"},
        {"file": "최종.xlsx", "normalized": "1850000"}]}

    (2026-08-31 코드품질 검토 후 명시) 충돌은 (location, type) 단위로 탐지하므로,
    같은 location이 서로 다른 타입의 충돌 항목으로 두 번 이상 나타날 수 있다 —
    예를 들어 한 셀에 금액도 날짜도 잘못된 값이 있으면 "Sheet1!A2"가 amount
    충돌 하나, date 충돌 하나로 각각 별도 항목이 된다. 즉 location만으로는
    충돌_목록 항목을 유일하게 식별할 수 없다.
    """
    pool = []
    for name in os.listdir(folder_path):
        ext = os.path.splitext(name)[1].lower()
        reader = _READERS.get(ext)
        if reader is None:
            continue  # 이미지, 워드, 알 수 없는 형식 등 → 조용히 건너뜀
        full_path = os.path.join(folder_path, name)
        pool.extend(reader(full_path, default_year))

    by_location: dict[tuple[str, str], list[dict]] = {}
    for item in pool:
        location = item.get("location", "")
        if location.count("!") != 1:  # 엑셀 셀 주소 형식만 충돌 탐지 대상
            continue
        _, _, cell_ref = location.partition("!")
        if not _CELL_ADDRESS_PATTERN.match(cell_ref):
            # (2026-08-30 구현 중 발견) "Sheet1!예산(합계)" 같은 compute_column_sums의
            # 파생값 라벨도 "!"가 정확히 1개라 이 체크를 그냥 통과해버림 — 진짜 셀 주소
            # 형식(A2, XFD1048576 등)인지 정규식으로 한 번 더 확인해야 한다.
            continue
        # (2026-08-30 코드품질 검토 후 수정됨 — 중요) 키를 location만으로 잡으면,
        # 같은 셀에서 금액+날짜처럼 여러 타입이 같이 나올 때(예: "1,850,000원
        # 2026-09-07"이 든 셀) 파일별로 나중 타입이 앞 타입을 덮어써서, 날짜와
        # 금액을 서로 비교하는 말도 안 되는 "충돌"이 나오거나 진짜 금액 불일치가
        # 가려지는 문제가 있었다. (location, type) 튜플로 키를 잡아 타입별로
        # 따로 그룹핑해야 한다.
        by_location.setdefault((location, item.get("type", "")), []).append(item)

    conflicts = []
    for (location, value_type), items in by_location.items():
        distinct_files = {i["source_file"]: i["normalized"] for i in items}
        distinct_values = set(distinct_files.values())
        if len(distinct_values) > 1:
            conflicts.append({
                "location": location,
                "type": value_type,
                "values": [{"file": f, "normalized": v} for f, v in distinct_files.items()],
            })
    return pool, conflicts
```

**(2026-08-31 코드품질 리뷰 반영)**: 위 코드에서 `conflicts` 항목이 `type`을 버리고 있었다 — 같은 `location`에 amount 충돌과 date 충돌이 각각 따로 있어도 두 항목이 겉보기엔 똑같아서 어느 게 어떤 타입인지 구분이 안 됐다. `_value_type` → `value_type`으로 바꾸고 `"type": value_type`을 충돌 dict에 추가했다(위 코드에 이미 반영됨). 이에 따라 docstring도 "같은 location이 타입별로 여러 번 나타날 수 있다"는 점을 명시하도록 갱신했다(위 docstring 예시에 반영됨). 실제 커밋은 080e52b.

**모델링 결정 사항 (계획 단계에서 확정)**: "원본 파일 간 값 불일치"는 PRD에서 추상적으로만 정의돼 있어, 이 계획에서 구체적인 규칙으로 확정한다 — **서로 다른 엑셀 파일에서 같은 시트이름+같은 셀 주소("Sheet1!A2" 형식)에 다른 값이 있을 때만** 충돌로 판단한다. 이건 같은 서식의 초안/최종본처럼 실제로 흔한 케이스를 다루기 위함이며, 구조가 다른 파일끼리의 의미적 충돌 탐지(예: 한글 문서와 엑셀의 같은 항목 비교)는 이번 범위에 포함하지 않는다.

추가 회귀 테스트(같은 셀에 금액+날짜 등 여러 타입이 같이 있어도 타입별로 올바르게 비교되는지, `_selftest_read_source_folder_conflict` 아래에 추가):

```python
def _selftest_read_source_folder_multi_type_cell_no_cross_type_conflict():
    """같은 셀에서 금액+날짜 등 여러 타입이 같이 나올 때, 타입을 섞어서 엉뚱하게
    비교하면 안 된다 — 실제 금액 불일치만 정확히 잡아야 한다."""
    os.makedirs("_test_원본폴더2", exist_ok=True)
    wb1 = openpyxl.Workbook(); ws1 = wb1.active; ws1.title = "Sheet1"
    ws1["A1"] = "메모"; ws1["A2"] = "예산은 1,850,000원이며 회의는 2026-09-07 진행"
    wb1.save("_test_원본폴더2/초안.xlsx")
    wb2 = openpyxl.Workbook(); ws2 = wb2.active; ws2.title = "Sheet1"
    ws2["A1"] = "메모"; ws2["A2"] = 1900000  # 금액만 다름(진짜 불일치), 날짜는 없음
    wb2.save("_test_원본폴더2/최종.xlsx")
    try:
        pool, conflicts = read_source_folder("_test_원본폴더2", default_year=2026)
        amount_conflicts = [c for c in conflicts if any(
            v["normalized"] in ("1850000", "1900000") for v in c["values"])]
        assert len(amount_conflicts) == 1, conflicts
        values = sorted(v["normalized"] for v in amount_conflicts[0]["values"])
        assert values == ["1850000", "1900000"], values  # 날짜와 뒤섞이지 않아야 함
        print("read_source_folder_multi_type_cell_no_cross_type_conflict 통과:", conflicts)
    finally:
        import shutil
        shutil.rmtree("_test_원본폴더2")
```

`__main__`에 `_selftest_read_source_folder_multi_type_cell_no_cross_type_conflict()`도 추가.

- [ ] Step 4: 통과 확인

Run: `python source_reader.py`
Expected: `read_source_folder 통과: 충돌 [...]` 출력, 에러 없음

- [ ] Step 5: 커밋

```bash
git add source_reader.py
git commit -m "F11: 폴더 스캔 + 엑셀 동일셀 값 불일치 충돌 탐지 (무관한 파일은 조용히 건너뜀)"
```

---

### Task 11: 한글 보고서 문서 열기 + 텍스트 읽기 + 빨간색 표시

Files

- Create: `hwp_report.py`

- [ ] Step 1: 실패하는 테스트 작성

```python
"""hwp_report.py — 채팅 도구가 직접 여는 한글 보고서 문서를 다루는 얇은 pyhwpx 래퍼"""
import os


def _selftest_open_and_mark_red():
    from pyhwpx import Hwp
    test_path = os.path.abspath("_test_보고서.hwp")
    setup = Hwp(visible=False)
    setup.insert_text("예산은 9999999원이며 정상입니다")
    setup.save_as(test_path)
    setup.quit()
    try:
        report = HwpReport(test_path)
        text = report.get_text()
        assert "9999999" in text, text
        report.mark_red("9999999")
        color_at_match = report.get_char_color_at("9999999")
        assert color_at_match == (255, 0, 0), color_at_match
        report.close(save=False)
        print("HwpReport 통과: 텍스트 읽기 + 빨간색 표시 확인")
    finally:
        os.remove(test_path)


if __name__ == "__main__":
    _selftest_open_and_mark_red()
```

- [ ] Step 2: 실패 확인

Run: `python hwp_report.py`
Expected: `NameError: name 'HwpReport' is not defined`

- [ ] Step 3: 최소 구현

```python
from pyhwpx import Hwp


class HwpReport:
    """채팅 도구가 pyhwpx로 직접 여는 보고서 문서. 사용자가 한글 아이콘으로
    따로 열지 않게 해서, F9에서 확인된 COM 재연결 제약(PRD 12-3)을 피한다.
    """

    def __init__(self, path: str):
        self.hwp = Hwp(visible=True)  # 사용자가 직접 봐야 하므로 visible=True
        self.hwp.open(path)
        self.path = path

    def get_text(self) -> str:
        return self.hwp.GetTextFile("TEXT", "")

    def mark_red(self, target_text: str) -> bool:
        """문서 안에서 target_text를 찾아 글자색을 빨간색(255,0,0)으로 바꾼다.
        찾지 못하면 False를 반환한다. 자동저장은 하지 않는다.
        """
        found = self.hwp.find(target_text)
        if not found:
            return False
        self.hwp.set_font(TextColor=self.hwp.RGBColor(255, 0, 0))
        return True

    def get_char_color_at(self, target_text: str):
        """target_text 위치의 현재 글자색을 (R,G,B) 튜플로 반환한다 (테스트 검증용)."""
        self.hwp.find(target_text)
        color_value = self.hwp.CharShape.Item("TextColor")
        return (color_value & 0xFF, (color_value >> 8) & 0xFF, (color_value >> 16) & 0xFF)

    def close(self, save: bool):
        self.hwp.quit(save=save)
```

- [ ] Step 4: 통과 확인

Run: `python hwp_report.py`
Expected: `HwpReport 통과: ...` 출력, 에러 없음. **한글 창이 화면에 실제로 뜨는 것도 눈으로 확인**(visible=True이므로).

- [ ] Step 5: 커밋

```bash
git add hwp_report.py
git commit -m "F11: HwpReport 래퍼 - 채팅도구가 직접 열기+텍스트읽기+빨간색표시, 자동저장 없음"
```

**✅ 완료 (커밋 `7d161e5`)**. 실제 버그 발견 및 수정: `mark_red()` 직후 `get_char_color_at()`가 다시 `find()`를 호출할 때, 이미 선택된 상태에서 `RepeatFind`를 또 실행하면 HWP가 색상 변경 *이전* 글자모양을 보고하는 COM 동작이 확인됨(재현 100%) — `get_char_color_at()` 내부에서 `find()` 전에 `self.hwp.Cancel()`로 선택을 먼저 해제하도록 수정해 해결. 화면 스크린샷으로 실제 문서에 빨간색이 적용된 것도 확인됨(사람이 직접 봄, "9999999"만 빨간색·나머지는 검정). 부수 개선: `Hwp(visible=True)` → `Hwp(visible=True, new=True)`로 변경(이미 실행 중인 HWP COM 인스턴스를 조용히 재사용하지 않도록).

**(2026-08-31 스펙검토 중 발견 — 환경 관련, 코드 결함 아님)**: 이 워크트리 폴더(`.worktrees\f11-chat-verify-numbers`) 안에서 `save_as()`/`open()`으로 .hwp 파일을 쓰거나 열 때, 간헐적으로 한글의 네이티브 "파일 접근 허용" 보안 대화상자가 떠서 사람이 직접 닫아주기 전까지 자동화 스크립트가 멈추는 현상이 재현됨(스펙 검토자가 3회 반복 재현, 매번 orphan된 `Hwp.exe` 프로세스가 남아 수동으로 종료해야 했음). `pyhwpx`의 `FilePathCheckerModule.dll`이 이 경고를 억제하도록 레지스트리(`HKCU\Software\HNC\HwpAutomation\Modules\FilePathCheckerModule`)에 정상 등록되어 있음을 직접 확인했으나(DLL 파일도 실제로 존재), 그럼에도 간헐적으로 발생한다 — Task 7~10의 수많은 반복 실행에서는 문제없이 통과했던 걸 보면 100%가 아니라 간헐적 현상으로 보인다(이날 밤 여러 서브에이전트 시도로 쌓인 orphan 프로세스/부하가 원인일 가능성). **완전 무인 자동화가 이 폴더에서 100% 보장되지 않을 수 있다는 걸 알고 있어야 한다** — 실제 배포 시나리오(F11 최종 기능)는 어차피 `visible=True`로 사용자가 화면을 보고 있는 상태를 전제로 하므로 이 대화상자는 사용자가 한 번 눌러주면 그만이라 제품 설계상 치명적이지 않지만, Task 12/15/16의 이후 자동 테스트 실행 중 유사한 멈춤이 재발할 수 있음을 감안해야 한다.

---

### Task 12: verify_numbers 전체 파이프라인 연결 (도구 함수)

Files

- Create: `verify_tool.py`

- [ ] Step 1: 실패하는 테스트 작성

```python
"""verify_tool.py — Task 1~11의 모듈을 엮어 "숫자검증" 도구 하나로 만든다.
chat_assistant.py가 Ollama 도구호출로 실행할 최종 진입점."""
import os
import openpyxl


def _selftest_run_verification():
    os.makedirs("_test_원본", exist_ok=True)
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Sheet1"
    ws["A1"] = "예산"; ws["A2"] = 1850000
    wb.save("_test_원본/원본.xlsx")

    from pyhwpx import Hwp
    report_path = os.path.abspath("_test_보고서.hwp")
    setup = Hwp(visible=False)
    setup.insert_text("예산은 185만원이며, 오타는 9999999원입니다")
    setup.save_as(report_path)
    setup.quit()

    try:
        result = run_verification(report_path=report_path, source_path="_test_원본", default_year=2026)
        assert result["mismatch_count"] == 1, result
        assert "9999999" in result["summary"], result
        print("run_verification 통과:", result["summary"])
    finally:
        import shutil
        shutil.rmtree("_test_원본")
        os.remove(report_path)


if __name__ == "__main__":
    _selftest_run_verification()
```

- [ ] Step 2: 실패 확인

Run: `python verify_tool.py`
Expected: `NameError: name 'run_verification' is not defined`

- [ ] Step 3: 최소 구현

```python
import os
from verify_numbers import extract_values, compare_values
from source_reader import read_excel_source, read_hwp_source, read_pdf_source, read_source_folder
from hwp_report import HwpReport


def run_verification(report_path: str, source_path: str, default_year: int) -> dict:
    """채팅창이 호출하는 최종 도구 함수.
    1) report_path를 pyhwpx로 열어 화면에 띄운다 (사용자가 보는 그 창)
    2) source_path(파일 또는 폴더)에서 정답 풀을 만든다
    3) 보고서 텍스트에서 4종 값을 뽑아 대조한다
    4) 불일치 항목을 빨간색으로 표시한다 (자동저장 안 함)
    5) 채팅창에 보여줄 요약 텍스트와 충돌 목록을 반환한다
    """
    ext = os.path.splitext(source_path)[1].lower()
    if os.path.isdir(source_path):
        answer_pool, conflicts = read_source_folder(source_path, default_year)
    elif ext in (".xlsx", ".xls"):
        answer_pool, conflicts = read_excel_source(source_path, default_year), []
    elif ext in (".hwp", ".hwpx"):
        answer_pool, conflicts = read_hwp_source(source_path, default_year), []
    elif ext == ".pdf":
        answer_pool, conflicts = read_pdf_source(source_path, default_year), []
    else:
        return {"mismatch_count": 0, "summary": f"지원하지 않는 원본 형식: {source_path}", "conflicts": []}

    report = HwpReport(report_path)
    report_text = report.get_text()
    report_values = extract_values(report_text, default_year)
    mismatches = compare_values(report_values, answer_pool)

    for mismatch in mismatches:
        report.mark_red(mismatch["raw"])

    lines = [f"- [{m['type']}] '{m['raw']}' 원본에서 확인 안 됨" for m in mismatches]
    for c in conflicts:
        value_desc = ", ".join(f"{v['file']}={v['normalized']}" for v in c["values"])
        # (2026-08-31) 같은 location에 타입이 다른 충돌이 각각 별도 항목으로 올 수 있으므로
        # (Task 10 참고) type을 같이 표시해야 두 줄이 나와도 사용자가 구분할 수 있다.
        lines.append(f"- ⚠ 원본자료 불일치[{c['type']}]: {c['location']} ({value_desc})")

    summary = f"{len(mismatches)}건 확인 필요\n" + "\n".join(lines) if mismatches or conflicts else "이상 없음, 모두 원본과 일치합니다"

    return {"mismatch_count": len(mismatches), "summary": summary, "conflicts": conflicts,
            "_report_handle": report}  # UI가 문서를 닫지 않고 계속 보여주기 위해 핸들 유지
```

- [ ] Step 4: 통과 확인

Run: `python verify_tool.py`
Expected: `run_verification 통과: ...` 출력. **한글 창이 실제로 뜨고, "9999999" 부분이 빨갛게 표시된 것을 눈으로 확인.**

- [ ] Step 5: 커밋

```bash
git add verify_tool.py
git commit -m "F11: run_verification 전체 파이프라인 (문서열기+원본읽기+대조+빨간색표시+요약)"
```

---

### Task 13: 채팅창 UI 뼈대 (도구 연결 없이, 화면만)

Files

- Create: `chat_assistant.py`

- [ ] Step 1: 최소 구현 (이 태스크는 순수 UI라 자동 테스트 대신 육안 확인으로 대체한다)

```python
"""chat_assistant.py — 작고 예쁜 채팅창 + Ollama Qwen3 도구호출 + verify_tool 실행.
항상 위에 떠 있고 드래그 가능한 CustomTkinter 창."""
import customtkinter as ctk
from tkinter import filedialog

ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")


class ChatAssistant(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("보고서 도우미")
        self.geometry("320x480")
        self.attributes("-topmost", True)

        self.report_path = None
        self.source_path = None

        self.report_button = ctk.CTkButton(self, text="보고서 파일 선택", command=self._choose_report)
        self.report_button.pack(pady=(10, 4), padx=10, fill="x")

        self.source_button = ctk.CTkButton(self, text="원본자료 선택 (파일)", command=self._choose_source_file)
        self.source_button.pack(pady=4, padx=10, fill="x")

        self.source_folder_button = ctk.CTkButton(self, text="원본자료 선택 (폴더)", command=self._choose_source_folder)
        self.source_folder_button.pack(pady=4, padx=10, fill="x")

        self.chat_log = ctk.CTkTextbox(self, height=280)
        self.chat_log.pack(pady=10, padx=10, fill="both", expand=True)
        self.chat_log.configure(state="disabled")

        self.input_box = ctk.CTkEntry(self, placeholder_text="예: 숫자 검증해줘")
        self.input_box.pack(pady=(0, 10), padx=10, fill="x")
        self.input_box.bind("<Return>", self._on_submit)

    def _choose_report(self):
        path = filedialog.askopenfilename(filetypes=[("한글 문서", "*.hwp *.hwpx")])
        if path:
            self.report_path = path
            self._log(f"보고서 선택됨: {path}")

    def _choose_source_file(self):
        path = filedialog.askopenfilename(filetypes=[("원본자료", "*.xlsx *.xls *.hwp *.hwpx *.pdf")])
        if path:
            self.source_path = path
            self._log(f"원본자료(파일) 선택됨: {path}")

    def _choose_source_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.source_path = path
            self._log(f"원본자료(폴더) 선택됨: {path}")

    def _log(self, message: str):
        self.chat_log.configure(state="normal")
        self.chat_log.insert("end", message + "\n")
        self.chat_log.configure(state="disabled")
        self.chat_log.see("end")

    def _on_submit(self, event):
        text = self.input_box.get()
        self.input_box.delete(0, "end")
        self._log(f"나: {text}")
        # Task 15에서 Ollama 도구호출로 교체 예정. 지금은 입력이 화면에 찍히는지만 확인.


if __name__ == "__main__":
    app = ChatAssistant()
    app.mainloop()
```

- [ ] Step 2: 실행 확인 + 스크린샷

Run: `python chat_assistant.py`

Expected: 320x480 크기의 작은 창이 화면 맨 위에 뜨고, 버튼 3개(보고서 선택/원본 파일 선택/원본 폴더 선택), 채팅 로그 창, 하단 입력창이 보인다. 버튼을 눌러 파일 선택 대화상자가 뜨는지, 입력창에 텍스트 치고 Enter 누르면 "나: ..." 로 로그에 찍히는지 확인한다.

**이 스텝은 실제로 실행해서 화면을 캡처해 사용자에게 보여주고 확인받는다** (사용자가 목업 없이 "만들어서 보여주는" 방식을 선택함, 앞선 대화 참고). 확인 없이 다음 태스크로 넘어가지 않는다.

- [ ] Step 3: 커밋

```bash
git add chat_assistant.py
git commit -m "F11: 채팅창 UI 뼈대 (topmost, 파일/폴더 선택 버튼, 로그, 입력창) - 도구 연결 전"
```

---

### Task 14: Ollama Qwen3 도구호출 연결

Files

- Modify: `chat_assistant.py`

- [ ] Step 1: 실패하는 테스트 작성 (UI 없이 라우팅 로직만 분리해서 테스트 가능하게 만든다)

```python
def _selftest_route_intent():
    # Ollama가 실제로 설치되어 있어야 통과한다 (Phase 0 사전준비 완료 전제)
    tool_called = route_intent("숫자 검증해줘")
    assert tool_called == "verify_numbers", tool_called
    print("route_intent 통과:", tool_called)
```

`chat_assistant.py` 맨 아래, `if __name__ == "__main__":` 블록 **위**에 추가하고, 블록 안에 조건부 테스트 실행을 넣는다:

```python
if __name__ == "__main__":
    import sys
    if "--selftest" in sys.argv:
        _selftest_route_intent()
    else:
        app = ChatAssistant()
        app.mainloop()
```

- [ ] Step 2: 실패 확인

Run: `python chat_assistant.py --selftest`
Expected: `NameError: name 'route_intent' is not defined`

- [ ] Step 3: 최소 구현

`ChatAssistant` 클래스 **위에** 추가:

```python
import ollama

_TOOLS = [{
    "type": "function",
    "function": {
        "name": "verify_numbers",
        "description": "지금 열려있는 한글 보고서의 금액/날짜/시간/전화번호를 원본데이터와 대조해서 틀린 부분을 빨간색으로 표시한다",
        "parameters": {"type": "object", "properties": {}, "required": []},
    },
}]


def route_intent(user_message: str) -> str | None:
    """사용자의 자연어 입력을 Qwen3에 보내 어떤 도구를 부를지 판단한다.
    지금은 등록된 도구가 verify_numbers 하나뿐이라 반환값도 그것뿐이거나 None이다.
    """
    response = ollama.chat(
        model="qwen3:8b",
        messages=[{"role": "user", "content": user_message}],
        tools=_TOOLS,
    )
    tool_calls = response.get("message", {}).get("tool_calls") or []
    if not tool_calls:
        return None
    return tool_calls[0]["function"]["name"]
```

- [ ] Step 4: 통과 확인

Run: `python chat_assistant.py --selftest`
Expected: `route_intent 통과: verify_numbers` 출력

- [ ] Step 5: 커밋

```bash
git add chat_assistant.py
git commit -m "F11: Ollama Qwen3 도구호출 라우팅 (route_intent), --selftest 플래그로 UI 없이 검증 가능"
```

---

### Task 15: 채팅창 ↔ verify_tool 실제 연결 (전체 통합)

Files

- Modify: `chat_assistant.py`

- [ ] Step 1: 최소 구현 (전체 조립이라 자동 테스트보다 실사용 시나리오로 검증한다)

`_on_submit` 메서드를 아래로 교체:

```python
    def _on_submit(self, event):
        text = self.input_box.get()
        self.input_box.delete(0, "end")
        self._log(f"나: {text}")

        if not self.report_path or not self.source_path:
            self._log("도우미: 먼저 보고서 파일과 원본자료를 선택해주세요.")
            return

        tool_name = route_intent(text)
        if tool_name == "verify_numbers":
            from verify_tool import run_verification
            result = run_verification(self.report_path, self.source_path, default_year=2026)
            self._log(f"도우미: {result['summary']}")
        else:
            self._log("도우미: 아직 이 요청은 처리할 수 있는 도구가 없어요. '숫자 검증해줘'라고 말씀해보세요.")
```

- [ ] Step 2: 실사용 시나리오로 검증 (자동화된 assert 대신, 실제 파일로 직접 확인)

1. 오타를 하나 심은 테스트용 한글 보고서와 엑셀 원본을 준비한다 (Task 12의 `_test_보고서.hwp`, `_test_원본` 폴더를 재사용하거나 새로 만든다)
2. `python chat_assistant.py` 실행
3. "보고서 파일 선택" → 테스트 보고서 선택
4. "원본자료 선택 (폴더)" → 테스트 원본 폴더 선택
5. 입력창에 "숫자 검증해줘" 입력 후 Enter

Expected: 한글 창에서 오타 부분이 빨갛게 표시되고, 채팅 로그에 "N건 확인 필요: ..." 요약이 뜬다. **이것도 스크린샷으로 캡처해 확인받는다.**

- [ ] Step 3: 커밋

```bash
git add chat_assistant.py
git commit -m "F11: 채팅창-verify_tool 통합 완료, 자연어 입력으로 숫자검증 전체 흐름 동작"
```

---

### Task 16: 의존성 목록 및 README 갱신

Files

- Modify: `requirements.txt`
- Modify: `README.md`

- [ ] Step 1: `requirements.txt`에 추가 (기존 내용 유지, 아래 4줄만 추가)

```
customtkinter
pdfplumber
ollama
openpyxl
```

- [ ] Step 2: `README.md`의 F11 상태 행을 최종 갱신

기존 "📝 PRD 확정, 구현 전" 상태를, "✅ 구현 완료" 로 바꾸고 실행법 섹션에 아래를 추가:

```markdown
### 6. 채팅 인터페이스 + 숫자검증 (F11)

```bash
ollama pull qwen3:8b   # 최초 1회
python chat_assistant.py
```

"보고서 파일 선택" → "원본자료 선택(파일 또는 폴더)" → 채팅창에 "숫자 검증해줘"라고 입력. 원본에서 확인 안 되는 금액·날짜·시간·전화번호가 보고서 안에 빨간색으로 표시됩니다 (자동저장 안 됨, 확인 후 직접 저장).
```

- [ ] Step 3: 커밋

```bash
git add requirements.txt README.md
git commit -m "F11: 의존성 목록·README 실행법 갱신, 구현 완료 표시"
```

---

## Self-Review 결과 (계획 작성자가 직접 확인함)

- **스펙 커버리지**: PRD 12-3(아키텍처)→Task 11·12·15, 12-4(매칭규칙)→Task 1~4, 12-5(원본데이터 처리)→Task 7~10, 12-6(성공기준)→Task 15의 실사용 시나리오, 12-7(비기능)→Task 13(topmost), 12-9(테스트계획 표)의 10개 케이스 중 8개는 Task 1~12의 selftest로 커버, 나머지 2개(원본 간 불일치의 "열어줘" 이동, 폴더 내 무관 파일)는 Task 10에서 무관파일 처리는 커버했으나 **"열어줘" 명령으로 엑셀 셀 이동/hwp 위치 이동하는 기능은 이번 계획에 빠져 있음 — 아래 참고**
- **발견된 누락**: PRD 12-5의 "열어줘"라고 하면 해당 위치로 이동하는 기능(win32com 엑셀 셀 선택, pyhwpx find 이동)이 태스크로 안 만들어졌다. 이건 별도 태스크로 추가하거나, 이번 1차 구현 범위에서 제외하고 다음 라운드로 미루는 두 가지 선택이 있다 — **사용자 확인 필요** (아래 참고)
- **플레이스홀더 검사**: "TBD", "나중에 구현" 등 없음. 모든 스텝에 실제 코드 있음
- **타입 일관성**: `extract_values`, `compare_values`, `run_verification` 등 함수 시그니처가 정의된 태스크와 사용되는 태스크에서 동일하게 유지됨을 확인함

## 확인 필요 사항 (실행 시작 전에 답변 부탁드립니다)

**"열어줘" 기능(원본 파일 간 불일치 시 해당 위치로 직접 이동)을 이번 1차 구현에 포함할까요, 아니면 위 16개 태스크로 먼저 끝내고 다음 라운드에 추가할까요?** PRD엔 있지만 계획을 짜다 보니 이번 범위(채팅+숫자검증 핵심)와 별개로 떼어내도 되는 독립 기능이라, 먼저 핵심을 완성하고 검증한 뒤 추가하는 걸 추천드립니다.
