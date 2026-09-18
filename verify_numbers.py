"""verify_numbers.py — 원본데이터 대비 보고서 숫자검증 핵심 로직 (순수 함수, 외부 의존성 없음)"""
import datetime
import re
import unicodedata
from decimal import Decimal


_AMOUNT_PATTERN = re.compile(r'(\d+(?:,\d{3})*(?:\.\d+)?)(?:\s*(조원|억원|만원|천원|원|%))?')
_UNIT_MULTIPLIER = {"조원": 1000000000000, "억원": 100000000, "만원": 10000, "천원": 1000,
                     "원": 1, "%": 1, None: 1}

# (2026-09-04, 실사용 피드백으로 발견) 공공기관 보고서에서 흔한 "('23)","'24년"
# 같은 연도 약칭 표기 — 작은따옴표(또는 스마트따옴표) 바로 뒤에 두 자리 숫자가
# 오는 패턴. 이 두 자리 숫자("23","24")가 금액으로 뽑혀버리면, 원본자료의
# 진짜 데이터 값과는 아무 관계도 없는데 "오류(빨강)"로 잘못 표시된다(실사용
# 문서 "정기점검 실적 : ('23) 1,595개소 → ('24) 1,694 → ('25) 1,827"에서
# 실제 재현됨 — 1,595/1,694/1,827은 정상인데 23/24/25가 전부 오류로 잘못
# 표시됨). extract_values()가 날짜/시간/전화번호처럼 이 패턴도 "제외 구간"으로
# 다뤄서, 겹치는 금액 매치를 애초에 뽑지 않게 한다.
#
# (2026-09-17 추가, 실사용 재현) 같은 문서의 다른 절에서는 작은따옴표 대신
# 백틱(`)으로 연도를 적은 경우("`24개소" 등, 작성자가 키보드에서 잘못
# 눌렀거나 자동고침 결과로 추정)가 실제로 있었다 — 이 문자는 원래 패턴에
# 없어서 "24"/"25"가 그대로 금액으로 뽑혀 똑같은 오탐이 재현됐다(19→31,
# 63.2%↑ 자체는 정상인데 연도 24/25가 오류로 잘못 표시됨). 백틱도 같은
# 취급을 받도록 문자 집합에 추가한다.
_YEAR_ABBREVIATION_PATTERN = re.compile(r"['’‘`]\d{2}(?!\d)")

# (2026-09-17 추가, 사용자가 직접 지적한 근본적 한계) 위 패턴은 특정 기호를
# 하나씩 미리 나열하는 방식이라, 다른 사람이 다른 기호를 쓰거나 오타를 내면
# (예: 큰따옴표, 그냥 기호 없이 "(23)") 또 놓친다 — "두더지 잡기"의 한계.
# 그런데 사용자 문서들을 보면 진짜 불변의 구조는 "괄호 안에 숫자 2자리"라는
# 형태 자체다("('23)","(`24)" 전부 이 틀을 따름) — 괄호 안 기호가 뭐든, 심지어
# 아예 없어도("(23)") 이 틀로 잡으면 특정 기호 나열에 기대지 않아도 된다.
# 다만 진짜 데이터가 "(23)"처럼 괄호에 숫자 2자리만 단독으로 들어가는 경우는
# 이 프로젝트가 다루는 문서에서 실제로 관찰된 적이 없어(괄호 안 숫자는 보통
# "(1,750)"처럼 콤마 있는 목표값이거나 "(63.2%)"처럼 %가 붙음) 이 일반화가
# 안전하다고 판단한다. 위 _YEAR_ABBREVIATION_PATTERN은 괄호 없이 그냥
# "'23"만 있는 경우까지 잡으므로 그대로 남겨두고, 이 패턴은 추가로 합쳐서 쓴다.
_YEAR_PAREN_MARKER_PATTERN = re.compile(r"\([^\d()]?(\d{2})\)")

# (2026-09-07, 실사용 피드백으로 발견) "1." "2)"처럼 목차·개요·붙임 번호 등
# 문서 어디서나 등장하는 번호매기기 표기 — numbering_tool.py가 실제로 쓰는
# 표준 번호서식("1. "/"1) "/"(1) ")과 정확히 일치한다. 이런 숫자는 순서를
# 나타내는 라벨일 뿐 원본자료와 대조할 데이터가 아니다. 소수점 금액
# ("97.1")은 마침표 뒤에 또 숫자가 오므로 (?!\d)로 걸러 오배제하지 않는다.
#
# (2026-09-11, 실사용 문장으로 재현한 진짜 버그 수정 — 이 패턴은 원래
# 아무 데도 고정돼 있지 않았다: r'(\d+)(?:\.(?!\d)|\))')
# 그래서 "번호매기기"가 아니라 "괄호가 닫히거나 마침표로 끝나는 모든 숫자"를
# 번호로 오인해, 아래 문장들의 진짜 데이터가 검증 대상에서 통째로 빠졌다
# (빨강도 파랑도 아닌 "아예 안 봄" — 사용자가 알아챌 방법이 없는 조용한
# 누락이라 가장 위험한 종류의 오류다):
#   "계약 건수(1,827)는 전년 대비 증가"  -> ")" 앞의 "827"이 번호로 오인돼
#                                          1,827 전체가 제외됨
#   "총 예산 규모는 1850000."            -> 문장 끝 마침표 앞이라 제외됨
#   "점검 실적 1,694) 기준"              -> ")" 앞이라 제외됨
# 괄호 안에 수치를 적는 것은 공공기관 보고서에서 대단히 흔한 표기라 실사용
# 영향이 크다. 기존 self-test가 이걸 못 잡은 이유는 하필 검증 문장의 숫자
# 뒤에 단위("원")가 붙어 있어서 ")"/"."에 닿지 않았기 때문이다.
#
# 고치는 방향: "번호매기기가 실제로 나타나는 자리"에만 고정한다.
#   (가) 줄머리 — 들여쓰기 공백은 건너뛴다. numbering_tool.py가 번호를
#        각 줄 맨 앞에 붙이므로 이 도구가 만든 문서와 정확히 맞는다.
#   (나) 탭 또는 두 칸 이상 띄어쓴 뒤 — "붙임  1. 계획서 1부." 처럼 줄 안에서
#        구조적으로 벌려 쓴 번호를 놓치지 않기 위해서다. 한 칸 띄어쓰기는
#        일부러 제외했다 — 위 "규모는 1850000." 같은 평범한 문장이 다시
#        걸려들기 때문이다.
# 여는 괄호 한 개는 선택적으로 허용한다("(1) " 이중괄호 서식). 번호 자릿수는
# 1~2자리로 제한한다 — 실제 개요 번호는 두 자리를 넘지 않고, 넓게 잡을수록
# 진짜 데이터를 번호로 오인할 위험만 커진다.
_LIST_MARKER_PATTERN = re.compile(
    r'(?:(?<![^\r\n])[ \t 　]*|(?<=\t)|(?<=  ))'  # 줄머리(들여쓰기 포함) 또는 탭/두 칸 이상 뒤
    r'[(（]?'                                              # "(1)" 서식의 여는 괄호(선택)
    r'(\d{1,2})'                                           # 번호 자체 — 이 구간만 제외 대상이다
    r'(?:\.(?!\d)|[)）])'                                  # "1." 또는 "1)" 꼴의 번호 꼬리
)

# (2026-09-07) 연속된 1~12가 월(月) 표시로 흔히 쓰인다("1 2 3 ... 12" 월별
# 헤더) — extract_values()가 후보 amounts 중에서 이 조건에 맞는 런을 찾아
# 검증 대상에서 제외한다(_find_month_sequence_spans 참고).

# (2026-09-17, 사용자가 직접 지적) 목차의 쪽번호("Ⅰ. 공급업체 관리 개요\t 1")는
# 실제 데이터가 아니라 단순 안내용 숫자다 — 원본자료와 대조하면 우연히 겹치는
# 진짜 데이터와 잘못 비교돼 엉뚱하게 정상/오류로 판정될 위험이 있다(실사용
# 문서로 재현 확인: "목   차" 다음에 오는 1/2/3/10/12가 전부 금액으로 잘못
# 뽑혔었음). "목차"는 글자 사이에 공백이 여러 칸 들어갈 수 있어(사용자가
# 직접 지적, "목   차"처럼) \s*로 허용한다.
_TOC_HEADING_PATTERN = re.compile(r'목\s*차')
_TOC_ENTRY_LINE_PATTERN = re.compile(r'^.*\t\s*(\d+)\s*$')


def _find_toc_number_spans(text: str) -> list[tuple[int, int]]:
    """"목차" 표제 뒤에 "제목\t쪽번호" 꼴로 이어지는 목차 항목들을 찾아
    그 줄 전체를 제외구간으로 반환한다. 표제 다음 줄부터 시작해서, 그
    줄이 "텍스트 + 탭 + 숫자"꼴이면 그 줄 전체를 제외구간에 넣고 다음
    줄로 계속 가고, 그 모양이 아닌 줄을 만나면(목차가 끝났다고 보고)
    멈춘다. 빈 줄은 건너뛰고 계속 본다(목차 항목 사이에 빈 줄이 있는
    문서가 실제로 있음).

    (2026-09-17, 실사용 문서로 재현) 처음엔 줄 끝의 쪽번호만 제외했는데,
    "붙 임 1. 기관별 합동점검 실적\t 14"처럼 목차 항목 자체에 번호가 또
    붙는 경우(붙임 자료 번호매기기) 그 앞쪽 번호("1")는 걸러지지
    않았다 — 줄이 목차 항목으로 확정되면 쪽번호뿐 아니라 그 줄 전체를
    제외해야, 줄 안 어디에 다른 숫자가 있어도 안전하다."""
    spans = []
    for heading in _TOC_HEADING_PATTERN.finditer(text):
        first_newline = text.find('\n', heading.end())
        if first_newline == -1:
            continue
        pos = first_newline + 1
        while pos < len(text):
            next_newline = text.find('\n', pos)
            line_end = next_newline if next_newline != -1 else len(text)
            line = text[pos:line_end].rstrip('\r')
            if line.strip() == "":
                pos = line_end + 1
                continue
            m = _TOC_ENTRY_LINE_PATTERN.match(line)
            if not m:
                break
            spans.append((pos, pos + len(line)))
            pos = line_end + 1
    return spans


def _selftest_find_toc_number_spans_handles_spaced_heading():
    """(2026-09-17, 사용자가 직접 지적) "목차"뿐 아니라 "목   차"처럼
    글자 사이 공백이 여러 칸이어도 목차로 인식해야 한다. 이제 줄 전체를
    제외구간으로 반환하므로, 각 span이 해당 줄의 쪽번호를 포함하는지로
    확인한다(정확히 그 줄 하나씩과 대응하는지도 개수로 확인)."""
    text = ("목   차\n\n"
            "Ⅰ. 공급업체 관리 개요\t 1\n"
            "Ⅱ. 사전관리\t 2\n"
            "Ⅲ. 상시관리\t 3\n"
            "Ⅳ. 사후관리\t 10\n"
            "Ⅴ. 향후 개선방향\t 12\n"
            "Ⅵ. 교육청 협조요청\n")
    spans = _find_toc_number_spans(text)
    lines = sorted(text[s:e] for s, e in spans)
    assert lines == [
        "Ⅰ. 공급업체 관리 개요\t 1", "Ⅱ. 사전관리\t 2", "Ⅲ. 상시관리\t 3",
        "Ⅳ. 사후관리\t 10", "Ⅴ. 향후 개선방향\t 12",
    ], lines
    print("_selftest_find_toc_number_spans_handles_spaced_heading 통과:", lines)


_CONTEXT_BOUNDARY_CHARS = ",.\n;、。，"
_DIGIT_GROUPING_COMMA_CHARS = ",，"


def _is_digit_grouping_comma(text: str, pos: int) -> bool:
    """text[pos]가 절 구분자가 아니라 "1,595"처럼 숫자 세 자리마다 넣는
    자릿수 구분 쉼표인지(양옆이 모두 숫자인지) 확인한다."""
    return (0 < pos < len(text) - 1
            and text[pos - 1].isdigit() and text[pos + 1].isdigit())


def _last_clause_boundary(text: str, char: str, start: int) -> int:
    """start 앞에서 char가 "절 구분자로" 쓰인 마지막 위치(없으면 -1).
    자릿수 구분 쉼표는 경계로 세지 않고 건너뛰어 더 앞에서 계속 찾는다
    (이유는 _preceding_context의 docstring 참고)."""
    search_end = start
    while True:
        pos = text.rfind(char, 0, search_end)
        if pos == -1:
            return -1
        if char in _DIGIT_GROUPING_COMMA_CHARS and _is_digit_grouping_comma(text, pos):
            search_end = pos
            continue
        return pos


def _preceding_context(text: str, start: int) -> str:
    """start 위치 바로 앞의 문맥(가장 가까운 문장/절 구분자부터 start까지)을
    반환한다. categorize_values()가 "이 값이 어떤 항목을 가리키는지"를 판단할
    실마리로 쓴다 — LLM 없이 순수 규칙(가장 가까운 쉼표/마침표/줄바꿈 이후
    텍스트)만으로 항목명을 유추하는 가벼운 방법이라, "참여 인원 21명" 같은
    라벨+숫자 근접 표기에는 잘 맞지만 라벨이 문장 앞쪽에 멀리 떨어진 경우는
    잡지 못하는 한계가 있다(예: "이번 사업의 참여 인원은... 총 21명").

    (2026-09-12, 실사용 문서에서 재현) 쉼표는 두 가지 전혀 다른 역할로 쓰인다 —
    "1,595"처럼 숫자 세 자리마다 넣는 자릿수 구분(양옆이 모두 숫자)과,
    "...점검, 2024년도에는..."처럼 절을 나누는 구분자(쉼표 뒤에 공백과 다음
    말이 옴)다. 원래는 이 둘을 구분하지 않고 "가장 가까운 쉼표"이면 무조건
    끊었다 — 그래서 "정기점검 실적 : ('23) 1,595개소 → ('24) 1,694 → ('25)
    1,827"에서 1,694의 문맥을 구하면, 진짜 절 구분 쉼표가 하나도 없는데도
    "1,595"의 자릿수 구분 쉼표에서 끊겨 "595개소 → ('24) "만 남고 정작
    항목명 "정기점검 실적"은 통째로 잘려나갔다(직접 재현해 확인함 — 그래서
    바로 앞 숫자에 자릿수 구분 쉼표가 없는 첫 값만 라벨에 연결되고 나머지는
    전부 회색으로 떨어졌다). 이제 자릿수 구분 쉼표는 경계로 세지 않고 계속
    더 앞의 진짜 경계를 찾는다."""
    boundary = max((_last_clause_boundary(text, char, start)
                    for char in _CONTEXT_BOUNDARY_CHARS), default=-1)
    return text[boundary + 1:start].strip()


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
            "context": _preceding_context(text, m.start()),
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
                         "raw": m.group(0), "span": m.span(),
                         "context": _preceding_context(text, m.start())})
    for m in _DATE_KOR.finditer(text):
        mo, d = m.groups()
        results.append({"type": "date", "normalized": f"{default_year:04d}-{int(mo):02d}-{int(d):02d}",
                         "raw": m.group(0), "span": m.span(),
                         "context": _preceding_context(text, m.start())})
    for m in _DATE_ABBR.finditer(text):
        yy, mo, d = m.groups()
        results.append({"type": "date", "normalized": f"{2000 + int(yy):04d}-{int(mo):02d}-{int(d):02d}",
                         "raw": m.group(0), "span": m.span(),
                         "context": _preceding_context(text, m.start())})
    return results


_YEAR_WORD_FULL = re.compile(r'(\d{4})년')
# (2026-09-17 추가, 실사용 우려사항 반영) 따옴표/백틱 없이 그냥 "24년"만 쓴
# 경우도 잡아야 한다 — "년"이라는 글자 자체가 이미 "이건 연도다"라는 확실한
# 신호라 따옴표 유무와 무관하게 안전하다(2024년을 "24개소"처럼 단위와
# 헷갈릴 일이 없음, 그래서 _YEAR_ABBREVIATION_PATTERN처럼 따옴표를 필수로
# 요구하지 않는다). 다만 "2024년"의 "24"까지 별도로 다시 잡히면 안 되므로
# ((?<!\d)로 앞에 숫자가 이어지지 않을 때만 매치), 4자리 연도의 뒷부분과
# 겹치지 않는다.
_YEAR_WORD_ABBR = re.compile(r"(?:['’‘`]|(?<!\d))(\d{2})년")


def extract_years(text: str) -> list[dict]:
    """"2026년"(4자리)과 "'26년"/"26년"(2자리, 따옴표는 있어도 없어도 됨)
    표기를 모두 4자리 연도 문자열로 정규화해서 반환한다 — D05가 요구하는
    "같은 연도로 정규화"를 실제로 비교 가능한 값으로 만드는 부분. "년"
    없이 그냥 "'23"만 쓴 표기는(예: "('23) 1,595개소") 연도인지 확정할 수
    없어 종전대로 _YEAR_ABBREVIATION_PATTERN이 금액 오인식만 막고, 이
    함수는 다루지 않는다 — "년"이 실제로 붙어있을 때만 이 함수가 연도로
    확정한다."""
    results = []
    for m in _YEAR_WORD_ABBR.finditer(text):
        results.append({"type": "year", "normalized": f"20{m.group(1)}", "raw": m.group(0),
                         "span": m.span(), "context": _preceding_context(text, m.start())})
    for m in _YEAR_WORD_FULL.finditer(text):
        results.append({"type": "year", "normalized": m.group(1), "raw": m.group(0),
                         "span": m.span(), "context": _preceding_context(text, m.start())})
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


def _selftest_extract_amounts_eok_jo_units():
    """(2026-09-17, 사용자 직접 요청) "억"/"조" 단위도 "만원"/"천원"처럼
    배율이 적용돼야 하고, 특히 소수와 결합했을 때(1.5억원 등) 정확해야
    한다 — float 오차 없이, 그리고 다른 표기와 같은 실제 값이면 같은
    문자열로 정규화되어야 한다."""
    a = extract_amounts("예산 1.5억원")
    assert a[0]["normalized"] == "150000000", a
    b = extract_amounts("예산 150,000,000원")
    assert a[0]["normalized"] == b[0]["normalized"], (a, b)  # 같은 값, 다른 표기

    c = extract_amounts("예산 2.3조원")
    assert c[0]["normalized"] == "2300000000000", c
    print("_selftest_extract_amounts_eok_jo_units 통과:", a, c)


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
                         "raw": m.group(0), "span": m.span(),
                         "context": _preceding_context(text, m.start())})

    colon_digit_spans = [m.span() for m in _COLON_DIGIT_RUN.finditer(text)]

    def _starts_inside_colon_digits(pos):
        return any(s < pos < e for s, e in colon_digit_spans)

    for m in _TIME_FULL_SI.finditer(text):
        if _starts_inside_colon_digits(m.start()):
            continue
        h1, h2 = m.groups()
        results.append({"type": "time", "normalized": f"{int(h1):02d}:00~{int(h2):02d}:00",
                         "raw": m.group(0), "span": m.span(),
                         "context": _preceding_context(text, m.start())})
    for m in _TIME_SHORT_SI.finditer(text):
        if _starts_inside_colon_digits(m.start()):
            continue
        h1, h2 = m.groups()
        results.append({"type": "time", "normalized": f"{int(h1):02d}:00~{int(h2):02d}:00",
                         "raw": m.group(0), "span": m.span(),
                         "context": _preceding_context(text, m.start())})
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
    return [{"type": "phone", "normalized": m.group(0), "raw": m.group(0), "span": m.span(),
             "context": _preceding_context(text, m.start())}
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

    (2026-09-08 추가, D05 연도 정규화 요건 재검토) "2026년"/"'26년"은
    extract_years()가 type="year"로 뽑고, 그 글자 범위도 제외구간에 넣어
    금액으로 중복 추출되지 않게 한다 — 이전에는 '26년의 '26만 제외됐고
    2026년의 2026은 그냥 금액으로 잡혀 원본에 없는 값이면 오류(빨강)로
    잘못 표시됐다.
    """
    dates = extract_dates(text, default_year)
    times = extract_times(text)
    phones = extract_phones(text)
    years = extract_years(text)
    year_abbreviations = [m.span() for m in _YEAR_ABBREVIATION_PATTERN.finditer(text)]
    year_paren_markers = [m.span() for m in _YEAR_PAREN_MARKER_PATTERN.finditer(text)]
    list_markers = [m.span(1) for m in _LIST_MARKER_PATTERN.finditer(text)]
    toc_numbers = _find_toc_number_spans(text)
    excluded_spans = ([r["span"] for r in dates + times + phones + years]
                       + year_abbreviations + year_paren_markers + list_markers
                       + toc_numbers)

    def _overlaps_excluded(span):
        a_start, a_end = span
        return any(a_start < e and s < a_end for s, e in excluded_spans)

    amounts = [r for r in extract_amounts(text) if not _overlaps_excluded(r["span"])]
    month_spans = _find_month_sequence_spans(amounts)
    amounts = [r for r in amounts if r["span"] not in month_spans]
    return amounts + dates + times + phones + years


def _selftest_extract_values_excludes_toc_page_numbers():
    """(2026-09-17, 실사용 문서로 재현) 목차의 쪽번호는 실제 데이터가
    아니므로 extract_values()가 금액으로 뽑으면 안 된다 - 목차 뒤에 이어지는
    진짜 본문의 금액은 평소처럼 그대로 뽑혀야 한다."""
    text = ("목   차\n\n"
            "Ⅰ. 공급업체 관리 개요\t 1\n"
            "Ⅱ. 사전관리\t 2\n"
            "Ⅲ. 상시관리\t 3\n\n"
            "Ⅰ. 공급업체 관리 개요\n\n"
            "예산은 1,850,000원입니다.")
    result = extract_values(text, default_year=2026)
    amounts = sorted(r["normalized"] for r in result if r["type"] == "amount")
    assert amounts == ["1850000"], amounts
    print("_selftest_extract_values_excludes_toc_page_numbers 통과:", amounts)


def _selftest_extract_values_excludes_toc_embedded_item_number():
    """(2026-09-17, 실사용 문서로 재현) "붙 임 1. 기관별 합동점검 실적\t 14"
    처럼 목차 항목 자체에 번호(붙임 자료 번호매기기)가 또 붙어있으면,
    줄 끝 쪽번호(14)뿐 아니라 그 앞의 "1"도 데이터로 뽑히면 안 된다 -
    처음엔 줄 끝 쪽번호만 제외해서 이 "1"이 새어나갔던 게 실제 문서로
    확인됨."""
    text = ("목   차\n\n"
            "Ⅰ. 공급업체 관리 개요\t 1\n"
            " 붙 임 1. 기관별 합동점검 실적\t 14\n"
            " 붙 임 2. 정부부처 간 협업 통한 학교급식 안전망 구축 현황\t 15\n\n"
            "Ⅰ. 공급업체 관리 개요\n\n"
            "예산은 1,850,000원입니다.")
    result = extract_values(text, default_year=2026)
    amounts = sorted(r["normalized"] for r in result if r["type"] == "amount")
    assert amounts == ["1850000"], amounts
    print("_selftest_extract_values_excludes_toc_embedded_item_number 통과:", amounts)


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


def _selftest_extract_values_excludes_year_in_parens_any_symbol():
    """(2026-09-17, 사용자가 직접 지적한 근본적 한계 반영) 괄호 안 연도
    표기에 어떤 기호를 쓰든(작은따옴표/백틱/큰따옴표/기호 없음) 특정 기호를
    미리 나열하는 방식으로는 언젠가 놓친다 — "괄호+숫자2자리"라는 틀
    자체로 잡아야 튼튼하다. 동시에 진짜 데이터가 담긴 괄호(목표값처럼
    콤마 있는 4자리, 증감률처럼 %가 붙은 경우)는 계속 정상적으로 뽑혀야
    한다."""
    no_symbol = extract_values(
        "정기점검 실적 : (23) 1,595개소 → (24) 1,694 → (25) 1,827", default_year=2026)
    amounts = sorted(r["normalized"] for r in no_symbol if r["type"] == "amount")
    assert amounts == ["1595", "1694", "1827"], amounts

    unexpected_symbol = extract_values(
        '실적 : ("23) 1,595개소 → ("24) 1,694', default_year=2026)
    amounts2 = sorted(r["normalized"] for r in unexpected_symbol if r["type"] == "amount")
    assert amounts2 == ["1595", "1694"], amounts2

    # 괄호 안 진짜 데이터는 계속 정상적으로 뽑혀야 한다(2자리가 아니므로 안 걸림)
    real_target = extract_values("목표(1,750) 대비 실적", default_year=2026)
    assert [r["normalized"] for r in real_target] == ["1750"], real_target
    real_percent = extract_values("비고(63.2%)", default_year=2026)
    assert [r["normalized"] for r in real_percent] == ["63.2"], real_percent
    print("_selftest_extract_values_excludes_year_in_parens_any_symbol 통과")


def _selftest_extract_values_excludes_year_abbreviation_with_backtick():
    """(2026-09-17, 실사용 재현) 같은 문서의 다른 절에서 연도 약칭에
    작은따옴표 대신 백틱(`)을 쓴 경우 — "동일IP주소 중복투찰 제재 : (`24)
    19개소 → (`25) 31 (63.2%↑)". 백틱도 작은따옴표와 똑같이 연도 약칭으로
    인식해서, "24"/"25"는 안 뽑히고 진짜 데이터(19/31/63.2)만 뽑혀야 한다."""
    text = "동일IP주소 중복투찰 제재 : (`24) 19개소 → (`25) 31 (63.2%↑)"
    result = extract_values(text, default_year=2026)
    amounts = sorted(r["normalized"] for r in result if r["type"] == "amount")
    assert amounts == ["19", "31", "63.2"], amounts
    print("_selftest_extract_values_excludes_year_abbreviation_with_backtick 통과:", result)


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


def _selftest_extract_values_list_marker_does_not_swallow_real_numbers():
    """(2026-09-11, 실사용 문장으로 재현한 진짜 버그의 회귀테스트) 예전
    _LIST_MARKER_PATTERN은 아무 데도 고정돼 있지 않아서, 괄호가 닫히거나
    마침표로 끝나는 숫자를 전부 "번호매기기"로 오인해 검증 대상에서
    조용히 빼버렸다. 셋 다 공공기관 보고서에 매우 흔한 표기다."""
    cases = [
        ("계약 건수(1,827)는 전년 대비 증가", ["1827"]),
        ("총 예산 규모는 1850000.", ["1850000"]),
        ("점검 실적 1,694) 기준", ["1694"]),
    ]
    for text, expected in cases:
        result = extract_values(text, default_year=2026)
        amounts = sorted(r["normalized"] for r in result if r["type"] == "amount")
        assert amounts == expected, (text, amounts, expected)
    print("_selftest_extract_values_list_marker_does_not_swallow_real_numbers 통과")


def _selftest_extract_values_excludes_list_marker_in_real_positions():
    """고정을 걸고도 진짜 번호매기기는 여전히 전부 제외돼야 한다. 이 도구가
    실제로 만들어내는 자리(numbering_tool.py가 각 줄 맨 앞에 붙이는 서식)와
    공문서에서 흔한 "붙임  1." 배치를 모두 확인한다."""
    # 줄머리 "1." / "2)" — 그 줄의 진짜 데이터는 그대로 살아 있어야 한다
    result = extract_values("1. 등록심사 실적은 1,827건이다", default_year=2026)
    assert sorted(r["normalized"] for r in result if r["type"] == "amount") == ["1827"], result
    # 여러 줄 문서에서 줄머리(들여쓰기 포함)와 이중괄호 서식
    text = ("가. 개요\r\n"
            "1. 등록심사 1,827건\r\n"
            "  2) 세부내용 1,694건\r\n"
            "(3) 참고사항 500,000원\r\n"
            "붙임  1. 세부계획 1부.\r\n")
    amounts = sorted(r["normalized"] for r in extract_values(text, default_year=2026)
                      if r["type"] == "amount")
    # 번호(1/2/3/1)는 전부 빠지고, "1부."의 1만 단위 없는 숫자로 남는다
    # (번호매기기가 아니라 수량 표기라 이 패턴의 대상이 아니다).
    assert amounts == ["1", "1694", "1827", "500000"], amounts
    print("_selftest_extract_values_excludes_list_marker_in_real_positions 통과:", amounts)


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


def _selftest_extract_years_full_and_abbreviated():
    result = extract_years("행사 연도는 2026년이며, 별도 통계에는 '26년 기준으로 기재")
    normalized = sorted(r["normalized"] for r in result)
    assert normalized == ["2026", "2026"], normalized
    print("extract_years 통과:", result)


def _selftest_extract_years_bare_abbreviation_without_quote():
    """(2026-09-17, 실사용 우려사항 반영) 따옴표/백틱 없이 그냥 "24년"만
    쓴 경우도 2024로 인식해야 한다 — 동시에 "2024년"의 뒷부분("24년")이
    별도 연도로 중복 추출되면 안 된다."""
    bare = extract_years("24년 실적은 좋았다")
    assert [r["normalized"] for r in bare] == ["2024"], bare

    full = extract_years("2024년 실적은 좋았다")
    assert [r["normalized"] for r in full] == ["2024"], full  # "24년" 중복 없이 하나만

    mixed = extract_years("23개소에서 24년까지 활동")
    assert [r["normalized"] for r in mixed] == ["2024"], mixed  # "23"은 연도 아님(년 없음)
    print("_selftest_extract_years_bare_abbreviation_without_quote 통과:", bare, full, mixed)


def _selftest_extract_values_year_word_excluded_from_amounts():
    """"2026년"의 2026이 금액으로 중복 추출되지 않고 year 타입 하나로만 잡혀야 한다."""
    result = extract_values("행사 연도는 2026년입니다", default_year=2026)
    types = sorted(r["type"] for r in result)
    assert types == ["year"], types
    print("_selftest_extract_values_year_word_excluded_from_amounts 통과:", result)


# ---------------------------------------------------------------------------
# (F14) 표 안에 있는 값의 "맥락"(행 머리말 / 열 머리말)을 이용한 대조 제한
#
# 왜 필요한가 — 실사용에서 확인된 문제: 공공기관 보고서는 KPI/실적 데이터를
# 보통 "표"로 넣는다. 그런데 한글의 GetTextFile("TEXT","")는 표를 읽을 때
# 셀을 그냥 한 줄에 하나씩 위→아래·왼→오른쪽 순서로 늘어놓을 뿐, 어느 셀이
# 몇 행 몇 열인지에 대한 표시를 전혀 남기지 않는다(직접 확인함:
#   "구분\r\n2024년\r\n2025년\r\n정기점검 실적\r\n1,694\r\n1,827\r\n...").
# 그래서 표 안의 숫자는 "이 값이 어느 항목의 몇 년도 값인지"를 알 수 없고,
# 값만 같으면 정답 풀 아무 값하고나 일치를 인정해 버린다. 실제로 이게
# 만드는 최악의 오탐은 이런 것이다:
#   보고서 표: [정기점검 실적] 열 "2024년" 칸에 1,827이 적혀 있음
#   원본 자료: 정기점검 실적의 2024년 값은 1,694, 2025년 값이 1,827
# → 1,827이 정답 풀 어딘가에 있다는 이유만으로 "정상(파랑)"으로 표시된다.
#   2024년 칸에 2025년 값을 잘못 옮겨적은 진짜 오류가 파란색으로 가려진다.
#
# 그래서 표 안의 값에 한해, 그 칸의 행 머리말/열 머리말과 "출처가 실제로
# 맞는" 원본 값하고만 대조하도록 후보를 걸러낸다.
#
# (2026-09-11 중요한 단서) 이 좁히기가 옳으려면 표기를 먼저 흡수해야 한다.
# 같은 열인데 표기만 달라서(예: 보고서는 "'24년", 원본은 "2024") 후보가 전부
# 걸러지면, 진짜 오류가 조용히 사라진다
# (split_column_identity / _find_year 참고).
# ---------------------------------------------------------------------------

# 4자리 연도만 잡는다. 앞뒤로 숫자가 더 붙어있으면(예: "12024", "20245")
# 연도가 아니라 그냥 긴 숫자의 일부이므로 잡으면 안 된다.
_YEAR_PATTERN = re.compile(r'(?<!\d)(\d{4})(?!\d)')

# 연도 뒤에 붙는 꼬리표. 반드시 긴 것부터 확인해야 한다 — "년"을 먼저 떼면
# "2024년도 목표"에서 "년"만 떨어져 "도 목표"라는 엉뚱한 수식어가 남는다.
_YEAR_SUFFIXES = ("년도", "년")

# ---------------------------------------------------------------------------
# (2026-09-11, 두 차례 독립 검토에서 같은 결론으로 발견된 회귀를 고침)
#
# 무엇이 잘못됐었나 — split_column_identity()가 위 _YEAR_PATTERN(4자리 아스키
# 숫자)만 연도로 인정해서, 정작 이 파일이 이미 읽을 줄 아는 다른 연도 표기를
# "열 머리말"에서는 못 읽었다:
#   "'24년", "('24)"  -> 연도 None    (이 파일 맨 위 _YEAR_ABBREVIATION_PATTERN
#                                      주석이 인용한, 사용자가 실제 문서에
#                                      쓰는 바로 그 표기다)
#   "２０２４년"        -> 연도 "２０２４" (전각 숫자라 엑셀 머리말의 "2024"와
#                                      글자가 달라 같은 연도로 안 보인다)
#
# 그래서 생긴 실제 피해(진짜 .hwp+.xlsx로 재현 확인) — 보고서 표의 열 머리말이
# "'24년"이고 원본 엑셀 머리말이 "2024"이면, column_identity_agrees()가 모든
# 후보에 대해 불일치를 내서 후보가 전부 걸러지고, 그 칸의 값이 통째로
# 대조 실패로 떨어졌다. 원본 2024년 값이 1,694인데 보고서에 1,900이라고
# 잘못 적힌 진짜 오류가 빨강으로 잡히던 것이, 이 기능 도입 후 조용히
# 사라졌다(mismatch_items에서도 빠져 F13의 "N번째로 가줘" 대상에서도 없어진다).
#
# 고치는 방향: 새 규칙을 만들지 않는다. 이 파일이 이미 쓰는 것을 그대로
# 재사용한다 — 연도 약칭은 _YEAR_ABBREVIATION_PATTERN, 두 자리→네 자리
# 확장은 extract_dates()/check_weekday_consistency()가 쓰는 2000 + int(yy).
# ---------------------------------------------------------------------------

# 머리말 양 끝에서 떼어내도 의미가 변하지 않는 장식 문자(괄호·따옴표·구분점).
# "('24)"에서 연도를 떼면 "()"만 남는데, 이걸 그대로 수식어로 두면 평범한
# "2024" 머리말의 수식어("")와 달라져 또 엉뚱하게 불일치가 난다. 하이픈·물결·
# 빗금은 일부러 넣지 않았다 — "2024/25"처럼 실제로 의미가 있는 표기를
# 조용히 지워버릴 수 있어서다.
_COLUMN_WRAPPER_CHARS = " \t　()[]{}<>（）［］｛｝〈〉《》【】〔〕「」『』,.·ㆍ、:;'’‘\"“”"


def _normalize_digit_forms(text: str) -> str:
    """전각 숫자("２０２４")·전각 괄호·전각 따옴표 같은 호환 문자를 아스키로
    접는다. NFKC가 정확히 이 일을 한다 — 전각 영숫자/괄호를 대응하는 아스키로
    바꿔주므로 "２０２４년"과 "2024년"을 같은 글자로 볼 수 있게 된다.
    한글 음절(가-힣)은 NFKC에 영향을 받지 않아 라벨 텍스트가 망가지지 않는다.
    """
    return unicodedata.normalize("NFKC", text or "")


def _find_year(text: str):
    """정규화가 끝난 텍스트에서 연도를 찾아 (네자리연도, 시작위치, 끝위치)를
    돌려준다. 못 찾으면 None.

    4자리 표기를 먼저 본다. 이유가 두 가지다 — (가) "'2024년"처럼 따옴표와
    4자리가 같이 있는 표기에서는 4자리 쪽이 정답이고, (나) 이미 4자리로 잘
    읽히던 머리말의 동작이 한 글자도 바뀌지 않는다. 4자리가 없을 때만 연도
    약칭("'24")을 본다.
    """
    m = _YEAR_PATTERN.search(text)
    if m:
        return m.group(1), m.start(), m.end()
    m = _YEAR_ABBREVIATION_PATTERN.search(text)
    if m:
        # 맨 앞 따옴표를 뗀 두 자리를 네 자리로 편다. extract_dates()와
        # check_weekday_consistency()가 쓰는 것과 같은 2000 + int(yy) 규칙.
        return f"{2000 + int(m.group(0)[1:]):04d}", m.start(), m.end()
    return None


def _normalize_ws(text: str) -> str:
    """공백 차이만 있는 표기를 같게 보기 위해 공백을 전부 없앤다
    ("2024년 목표" == "2024년목표")."""
    return re.sub(r'\s+', '', text or '')


def _loose_text_match(a: str, b: str) -> bool:
    """공백을 무시한 양방향 부분문자열 비교. 이 저장소가 라벨 표기 차이에
    이미 관용적인 수준("집행률" vs "예산 집행률")을 그대로 따른다.
    한쪽이 빈 문자열이면 항상 True — "구분할 정보가 없다"는 뜻이므로
    걸러내지 않는다(모르는 것은 막지 않는다는 이 모듈의 기본 원칙)."""
    a, b = _normalize_ws(a), _normalize_ws(b)
    if not a or not b:
        return True
    return a in b or b in a


def split_column_identity(header: str) -> tuple:
    """열 머리말을 (연도, 수식어)로 쪼갠다.
      "2024년"      -> ("2024", "")
      "2024년도"    -> ("2024", "")     # 년/년도 표기 차이는 같은 것으로 본다
      "2024년 목표" -> ("2024", "목표")
      "'24년"       -> ("2024", "")     # 연도 약칭도 같은 연도로 본다
      "('24)"       -> ("2024", "")
      "２０２４년"    -> ("2024", "")     # 전각 숫자도 같은 연도로 본다
      "목표"        -> (None, "목표")   # 연도 정보 없음
      ""            -> (None, "")
    연도가 없으면 첫 값이 None이다 — "연도를 모른다"와 "연도가 다르다"를
    반드시 구분해야 해서 빈 문자열이 아니라 None을 쓴다.

    표기가 달라도 같은 연도면 같은 ("2024", ...)가 나오게 하는 것이 핵심이다.
    원본 엑셀 머리말은 보통 평범한 "2024"인데 보고서 표 머리말은 "'24년"처럼
    줄여 쓰는 일이 흔해서, 여기서 표기를 흡수하지 않으면 같은 열인데도
    "다른 열"로 판정돼 그 칸의 검증이 통째로 무력해진다(위 회귀 주석 참고).
    """
    text = _normalize_digit_forms((header or "").strip())
    found = _find_year(text)
    if not found:
        return None, _normalize_ws(text)
    year, start, end = found
    # 연도를 떼고 남은 부분에서 괄호/따옴표 같은 장식을 먼저 걷어낸다 —
    # 그래야 "('24년)" 같은 표기에서도 그다음의 "년" 꼬리표 제거가 먹는다.
    rest = text[:start] + text[end:]
    rest = rest.strip(_COLUMN_WRAPPER_CHARS)
    for suffix in _YEAR_SUFFIXES:  # 긴 것부터 — 위 _YEAR_SUFFIXES 주석 참고
        if rest.startswith(suffix):
            rest = rest[len(suffix):]
            break
    return year, _normalize_ws(rest).strip(_COLUMN_WRAPPER_CHARS)


def column_identity_agrees(a: str, b: str) -> bool:
    """두 열 머리말이 "같은 출처"를 가리키는지 판정한다. 표 안의 값을 원본
    후보와 대조해도 되는지 거르는 핵심 필터.

    규칙(이 순서가 중요하다):
    1) 양쪽 다 연도가 있으면 → 연도는 반드시 정확히 같아야 한다. 다르면 무조건
       불일치. ("2024년" 칸을 "2025년" 원본값과 대조하면 안 된다 — 이걸 부분
       문자열 비교로 대충 했다가, 2024년 칸에 적힌 2025년 값이 "정상"으로
       표시되는 오탐이 실제로 났다.)
    2) 연도가 같으면 → 수식어("목표"/"계획"/"실적" 등)도 정확히 같아야 한다.
       여기서 부분문자열 비교를 쓰면 안 된다 — "목표"와 "상반기 목표"는
       한쪽이 다른 쪽을 포함한다는 이유로 같은 것으로 취급돼 버리는데,
       실제로는 범위가 다른 별개의 값이다.

       (2026-09-11, 알려진 한계 — 일부러 안 고쳤다) 이 규칙은 "한쪽에만
       수식어가 있는" 경우도 불일치로 본다. 그래서 보고서 머리말이
       "2024년 실적"이고 원본 엑셀 머리말이 그냥 "2024"이면, 사람 눈에는
       같은 열인데 여기서 갈라져 그 칸이 회색(확인 필요)이 된다 — 원래는
       빨강으로 잡히던 오류를 놓치는 실제 사례다(진짜 .hwp/.xlsx로 재현 확인).
       그런데 이걸 "한쪽 수식어가 비었으면 통과"로 풀면 더 나쁜 일이 난다:
       보고서에 "2024년 목표"와 "2024년 실적" 두 열이 있고 원본에는 실적
       열("2024") 하나뿐일 때, 목표 열의 값까지 그 실적 열과 대조돼 멀쩡한
       숫자가 전부 빨강(오류)으로 찍힌다. 놓치는 회색보다 거짓 빨강이 훨씬
       해로워서, 표기가 애매할 때는 회색으로 두는 지금 규칙을 유지한다.
       (연도 "표기" 차이는 이와 다른 문제이고, split_column_identity()가
       전부 흡수하므로 여기까지 오지 않는다.)
    3) 한쪽에만 연도가 있으면 → 같은 열이라고 볼 근거가 없으므로 어긋난 것으로
       본다. 애초에 "열을 모르는" 원본(머리행 없는 세로형 엑셀 등)은 아래 첫
       줄에서 이미 걸러져 여기까지 오지 않는다 — 이 분기는 "원본에 머리행이
       있는데 거기에 연도가 없다"는 드문 경우만 다룬다.

       (2026-09-11 정정) 이 좁힘이 안전한 건 연도를 정말로 알 수 없을
       때뿐이다 — 원본 머리말에 연도가 실제로 안 적혀 있어서 몇 년도 값인지
       판단할 근거가 없는 경우. 그때는 "비교할 같은 출처가 없었다"는 뜻이
       사실과 맞는다. 반대로, 연도가 분명히 적혀 있는데 우리가 그 표기를 못
       읽어서 여기로 떨어지는 것은 안전하지 않다 — 진짜 오류가 조용히
       사라진다(실제로 "'24년"·전각 연도 표기에서 그 회귀가 났다). 그래서
       split_column_identity()가 아는 표기(4자리·연도 약칭·전각)는 전부
       흡수해서 같은 연도로 만든다. 여기 남는 건 "표기를 다 흡수하고도
       한쪽에 연도가 없는" 경우뿐이다.
    4) 양쪽 다 연도가 없으면 → 기존처럼 공백무시 양방향 부분문자열 비교로
       관대하게 본다("집행률" vs "예산 집행률"). 여기서 정확일치를 요구하면
       정당한 표기 차이가 깨진다.
    """
    if not (a or "").strip() or not (b or "").strip():
        return True  # 한쪽이라도 열 정보를 모르면 거르지 않는다
    year_a, mod_a = split_column_identity(a)
    year_b, mod_b = split_column_identity(b)
    if year_a != year_b:
        return False  # 연도가 다르거나, 한쪽에만 연도가 있다
    if year_a is not None:
        return mod_a == mod_b  # 연도가 같으면 수식어도 정확히 같아야 한다
    return _loose_text_match(mod_a, mod_b)


def candidate_allowed_by_context(table_context: dict, candidate: dict) -> bool:
    """표 안 값의 맥락(table_context)과 원본 후보(candidate)의 출처가 서로
    맞는지. 맥락이 없거나(표 밖의 일반 문장) 후보에 출처 정보가 없으면 항상
    True — 즉 기존 동작을 그대로 둔다. 모르는 것은 막지 않는다.
    """
    if not table_context:
        return True
    if not column_identity_agrees(table_context.get("column_header", ""),
                                  candidate.get("column_header", "")):
        return False
    return _loose_text_match(table_context.get("row_header", ""),
                             candidate.get("row_label", ""))


def _has_column_identity(entry: dict) -> bool:
    """이 원본 후보가 "몇 년도 열에서 왔는지"를 알고 있는가."""
    return bool(str((entry or {}).get("column_header", "") or "").strip())


def filter_candidates_by_context(table_context: dict, candidates: list) -> list:
    """표 안 값의 맥락에 맞는 원본 후보만 골라낸다.

    (2026-09-11 추가) 출처 정보가 섞인 후보 풀 처리 — read_source_files()는
    폴더를 통째로 받아 엑셀·한글·PDF 원본을 한 풀에 합친다. 그런데 열 출처
    정보를 붙이는 건 가로형 엑셀뿐이고, read_hwp_source()/read_pdf_source()가
    만든 후보에는 열 정보가 아예 없다. candidate_allowed_by_context()는
    "모르는 것은 막지 않는다"는 원칙에 따라 열 정보가 없는 후보를 무조건
    통과시키는데, 그러면 열 정보 없는 후보가 풀에 단 하나만 섞여 있어도
    그게 아무 값이나 받아줘서, 열을 아는 엑셀 후보들이 "그 연도 열의 값은
    이게 아니다"라고 말하고 있어도 열 구분 기능이 그 값에 대해 통째로
    무력해진다.

    그래서 이 칸이 열을 알고 있고 후보 중에 열을 아는 것이 하나라도 있으면,
    열을 아는 후보끼리만 대조한다. "같은 열인지 비교할 수 있는 원본이 실제로
    있는" 상황이라, 굳이 열을 모르는 후보로 후퇴할 이유가 없다. 열을 아는
    후보가 하나도 없으면 예전처럼 전체 후보로 관대하게 간다(= 기존 동작).
    """
    if not table_context or not str(table_context.get("column_header", "") or "").strip():
        return [c for c in candidates if candidate_allowed_by_context(table_context, c)]
    column_aware = [c for c in candidates if _has_column_identity(c)]
    pool = column_aware if column_aware else candidates
    return [c for c in pool if candidate_allowed_by_context(table_context, c)]


def _matching_context_text(rv: dict) -> str:
    """categorize_values()의 "라벨이 문맥에 들어있는가" 검사가 쓸 문맥 문자열.

    평문 값은 _preceding_context()가 채워둔 rv["context"](값 직전 문장/절
    텍스트) 그대로다 — 지금까지의 동작이 한 글자도 바뀌지 않는다.

    표 안의 값은 여기에 그 칸의 행 머리말을 덧붙인다. 표를 평문으로 펼치면
    셀마다 줄이 나뉘어서 rv["context"]가 거의 항상 빈 문자열이 되는데(값
    바로 앞이 줄바꿈이라 문맥이 잘린다), 정작 사람이 보기에 그 값의 항목명은
    같은 행의 머리말이다. 세로형("항목/값") 원본의 라벨과 연결되려면 이
    행 머리말이 문맥에 들어 있어야 한다. 열 머리말은 일부러 넣지 않았다 —
    열 쪽 판정은 column_identity_agrees()가 연도까지 정확히 비교하는 전용
    규칙으로 이미 처리하므로, 부분일치 문자열 비교에 섞으면 그 엄격함만
    흐려진다.
    """
    parts = [rv.get("context") or ""]
    table_context = rv.get("table_context")
    if table_context:
        parts.append(str(table_context.get("row_header") or ""))
    return " ".join(p for p in parts if p)


def _selftest_split_column_identity():
    assert split_column_identity("2024년") == ("2024", ""), split_column_identity("2024년")
    assert split_column_identity("2024년도") == ("2024", ""), split_column_identity("2024년도")
    # "년도"를 "년"보다 먼저 떼지 않으면 여기서 "도목표"가 나온다(회귀 방지)
    assert split_column_identity("2024년도 목표") == ("2024", "목표"), split_column_identity("2024년도 목표")
    assert split_column_identity("2024년 목표") == ("2024", "목표")
    assert split_column_identity("목표") == (None, "목표")
    assert split_column_identity("") == (None, "")
    # 긴 숫자 안에 우연히 들어있는 4자리는 연도가 아니다
    assert split_column_identity("120245") == (None, "120245"), split_column_identity("120245")
    print("_selftest_split_column_identity 통과")


def _selftest_split_column_identity_year_notations():
    """(2026-09-11, 두 차례 독립 검토에서 같은 결론으로 발견된 회귀의 재현
    테스트) 이 파일이 다른 곳에서는 이미 연도로 알아보는 표기를, 열 머리말
    에서도 똑같이 연도로 읽어야 한다. 못 읽으면 같은 연도인데 "다른 열"로
    판정돼 그 칸의 검증이 통째로 무력화된다."""
    # 연도 약칭 — 이 파일 맨 위 _YEAR_ABBREVIATION_PATTERN 주석이 인용한,
    # 사용자의 실제 문서에 나오는 표기 그대로다.
    assert split_column_identity("'24년") == ("2024", ""), split_column_identity("'24년")
    assert split_column_identity("'24") == ("2024", ""), split_column_identity("'24")
    assert split_column_identity("('24)") == ("2024", ""), split_column_identity("('24)")
    assert split_column_identity("('24년)") == ("2024", ""), split_column_identity("('24년)")
    assert split_column_identity("'25년") == ("2025", ""), split_column_identity("'25년")
    # 스마트 따옴표(한글이 자동 변환해 넣는 경우)도 같은 연도여야 한다
    assert split_column_identity("’24년") == ("2024", ""), split_column_identity("’24년")
    # 약칭 + 수식어
    assert split_column_identity("'24년 목표") == ("2024", "목표"), split_column_identity("'24년 목표")

    # 전각 숫자 — NFKC로 접어서 평범한 "2024"와 같은 연도가 되어야 한다
    assert split_column_identity("２０２４년") == ("2024", ""), split_column_identity("２０２４년")
    assert split_column_identity("２０２４") == ("2024", ""), split_column_identity("２０２４")
    assert split_column_identity("２０２４년 목표") == ("2024", "목표"), \
        split_column_identity("２０２４년 목표")

    # 4자리가 있으면 그쪽이 우선 — 기존 동작이 바뀌지 않는지 확인
    assert split_column_identity("'2024년") == ("2024", ""), split_column_identity("'2024년")
    # 연도 약칭이 아닌 따옴표+숫자(세 자리 이상)는 연도가 아니다
    assert split_column_identity("'243") == (None, "'243"), split_column_identity("'243")
    print("_selftest_split_column_identity_year_notations 통과")


def _selftest_column_identity_agrees_truth_table():
    """열 머리말 판정 진리표 — 이 표가 이 기능의 핵심이라 통째로 못박아 둔다."""
    # 같은 연도 + 같은 수식어 -> 연결
    assert column_identity_agrees("2024년", "2024년") is True
    assert column_identity_agrees("2024년", "2024년도") is True  # 년/년도 표기차이
    assert column_identity_agrees("2024년 목표", "2024년 목표") is True
    # 같은 연도 + 다른 수식어 -> 연결 안 됨
    assert column_identity_agrees("2024년 목표", "2024년 실적") is False
    # 같은 연도 + 한쪽만 수식어 -> 연결 안 됨(부분문자열로 통과시키면 안 됨)
    assert column_identity_agrees("2024년 목표", "2024년") is False
    # (실제로 났던 회귀) 같은 연도에서 "목표"와 "상반기 목표"는 범위가 다른
    # 별개 값인데, 부분문자열 비교를 쓰면 같은 것으로 취급돼 버렸다.
    assert column_identity_agrees("2024년 목표", "2024년 상반기 목표") is False
    # 다른 연도 -> 수식어가 같아도 절대 연결 안 됨
    assert column_identity_agrees("2024년", "2025년") is False
    assert column_identity_agrees("2024년 목표", "2025년 목표") is False
    # 양쪽 다 연도 없음 -> 기존의 관대한 부분문자열 비교 유지
    assert column_identity_agrees("집행률", "예산 집행률") is True
    assert column_identity_agrees("목표", "실적") is False
    # 한쪽만 연도 있음 -> 같은 열로 볼 근거가 없다
    assert column_identity_agrees("2025년", "값") is False
    assert column_identity_agrees("2024년 목표", "목표") is False
    assert column_identity_agrees("목표", "2024년 목표") is False
    # 한쪽이라도 열 정보를 모르면 거르지 않는다
    assert column_identity_agrees("", "2024년") is True
    assert column_identity_agrees("2024년", "") is True
    assert column_identity_agrees("   ", "2024년") is True
    print("_selftest_column_identity_agrees_truth_table 통과")


def _selftest_column_identity_agrees_across_year_notations():
    """(회귀 방지 — 이게 A 브랜치 3차 검토의 핵심 수정이었다) 보고서 표는
    "'24년"·전각으로 쓰고 원본 엑셀 머리말은 평범한 "2024"인, 실사용에서
    가장 흔한 조합에서 같은 열로 인정돼야 한다. 예전에는 전부 불일치가 나서
    그 칸의 후보가 통째로 걸러졌고, 진짜 오류가 조용히 사라졌다."""
    # 약칭 ↔ 평범한 연도
    assert column_identity_agrees("'24년", "2024") is True
    assert column_identity_agrees("'24년", "2024년") is True
    assert column_identity_agrees("('24)", "2024") is True
    assert column_identity_agrees("2024", "'24년") is True
    # 전각 ↔ 평범한 연도
    assert column_identity_agrees("２０２４년", "2024") is True
    assert column_identity_agrees("２０２４년", "2024년") is True
    # 표기만 다를 뿐 연도가 다르면 여전히 절대 연결되면 안 된다(이 기능의 본래 목적)
    assert column_identity_agrees("'24년", "2025") is False
    assert column_identity_agrees("'25년", "2024") is False
    assert column_identity_agrees("２０２４년", "2025") is False
    # 표기가 달라도 수식어 규칙은 그대로 적용된다
    assert column_identity_agrees("'24년 목표", "2024년 목표") is True
    assert column_identity_agrees("'24년 목표", "2024년 실적") is False
    print("_selftest_column_identity_agrees_across_year_notations 통과")


def _selftest_candidate_allowed_by_context_falls_back_when_unknown():
    """맥락이 없거나 후보에 출처 정보가 없으면 무조건 허용(기존 동작 유지)."""
    cand_plain = {"type": "amount", "normalized": "1827"}
    assert candidate_allowed_by_context(None, cand_plain) is True
    assert candidate_allowed_by_context({}, cand_plain) is True
    # 맥락은 있지만 후보엔 출처 정보가 전혀 없음 -> 거르지 않는다
    ctx = {"row_header": "정기점검 실적", "column_header": "2024년"}
    assert candidate_allowed_by_context(ctx, cand_plain) is True
    # 후보에 출처가 있고 연도가 다르면 -> 걸러진다
    cand_2025 = {"type": "amount", "normalized": "1827",
                 "row_label": "정기점검 실적", "column_header": "2025년"}
    assert candidate_allowed_by_context(ctx, cand_2025) is False
    # 행 라벨이 아예 다르면 -> 걸러진다
    cand_other_row = {"type": "amount", "normalized": "1694",
                      "row_label": "예산", "column_header": "2024년"}
    assert candidate_allowed_by_context(ctx, cand_other_row) is False
    print("_selftest_candidate_allowed_by_context_falls_back_when_unknown 통과")


def _selftest_filter_candidates_by_context_mixed_pool():
    """(2026-09-11) 폴더 하나에 가로형 엑셀과 한글/PDF 원본이 섞여 있을 때.
    한글/PDF에서 온 후보에는 열 정보가 아예 없어서 "모르는 건 막지 않는다"
    원칙으로 무조건 통과해 버리는데, 그 후보 하나 때문에 열 구분이 통째로
    무력해지면 안 된다 — 열을 아는 후보가 있으면 그쪽끼리만 대조한다."""
    ctx = {"row_header": "정기점검 실적", "column_header": "2024년"}
    excel_2024 = {"type": "amount", "normalized": "1694",
                  "row_label": "정기점검 실적", "column_header": "2024"}
    excel_2025 = {"type": "amount", "normalized": "1827",
                  "row_label": "정기점검 실적", "column_header": "2025"}
    hwp_blind = {"type": "amount", "normalized": "1827"}  # 열 정보 없음(한글 원본)

    # 열을 아는 후보가 있으면 열 모르는 후보는 빠진다 → 2024년 후보만 남는다
    allowed = filter_candidates_by_context(ctx, [excel_2024, excel_2025, hwp_blind])
    assert allowed == [excel_2024], allowed
    # → 2025년 값(1827)을 2024년 칸에 적은 오류가 hwp_blind 덕에 살아나지 않는다
    assert not any(a["normalized"] == "1827" for a in allowed), allowed

    # 열을 아는 후보가 하나도 없으면 예전처럼 전부 통과(기존 동작 유지)
    only_blind = [hwp_blind, {"type": "amount", "normalized": "1694"}]
    assert filter_candidates_by_context(ctx, only_blind) == only_blind

    # 이 칸이 열을 모르면(표 밖 값·머리말 없는 표) 좁히지 않는다
    ctx_no_col = {"row_header": "", "column_header": ""}
    assert filter_candidates_by_context(ctx_no_col, [excel_2024, excel_2025, hwp_blind]) == \
        [excel_2024, excel_2025, hwp_blind]
    print("_selftest_filter_candidates_by_context_mixed_pool 통과")


_MIN_LABEL_FRAGMENT_LENGTH = 3


def _label_matches_context(label: str, context: str) -> bool:
    """label(엑셀 행의 문자열 셀을 전부 이어붙인 것)이 context(보고서에서 그
    수치 바로 앞부분)와 연결되는가.

    (2026-09-12, 실사용 문서에서 재현) 통짜 부분문자열 비교만으로는 실패하는
    실제 사례: 원본 행이 "지표명"("정기점검 실적")과 "단위"("개소") 두
    문자열 칸을 갖고 있으면 label은 "정기점검 실적 개소"인데, 보고서 문장은
    "정기점검 실적 : ('23) 1,595개소 → ('24) 1,694 → ('25) 1,827"처럼
    "개소"가 숫자 뒤(그 다음 숫자의 앞이 아니라)에만 한 번 붙는다. 그러면
    각 숫자 앞의 context("정기점검 실적 : ('23) " 등)에는 "개소"가 없어서
    통짜 label 전체가 안 걸리고, 지표명과 단위 둘 다 진짜 원본 정보인데도
    전부 회색(확인 필요)으로 떨어진다.

    그래서 통짜 비교가 실패하면 label을 공백으로 쪼갠 조각 중 하나라도
    context에 있으면 연결로 본다. 다만 너무 짧은 조각("개소","건","%" 같은
    단위어)은 이 매칭에서 빼야 한다 — 그런 조각은 보고서 어디에나 흔해서,
    엉뚱한 다른 행과도 걸려버리는 오탐을 만든다(예: 다른 지표도 전부 "개소"
    단위면 그 지표들끼리 서로 라벨이 뒤섞인다). 그래서 길이가
    _MIN_LABEL_FRAGMENT_LENGTH 이상인 조각만 독립 매칭 신호로 인정한다.
    """
    if not label:
        return False
    if label in context:
        return True
    return any(
        len(fragment) >= _MIN_LABEL_FRAGMENT_LENGTH and fragment in context
        for fragment in label.split()
    )


def categorize_values(report_values: list[dict], answer_pool: list[dict]) -> dict:
    """report_values 각각을 answer_pool과 대조해 네 갈래로 나눈다:
    - matches: 원본과 정확히 일치(정상) → 문서에 파란색으로 표시할 대상
    - mismatches: 원본에 같은 타입 값이 있지만 정확히 일치하는 게 없음(오류)
      → 빨간색 표시 대상. 기존 compare_values()가 반환하던 것과 동일한 판정.
    - unverifiable: answer_pool에 그 타입 자체가 하나도 없어서 애초에 대조가
      불가능함 → 초록색 표시 대상. "원본에 이 종류의 데이터 자체가 없어서
      확인 못 했다"는 뜻이지, 값이 틀렸다는 뜻이 아니다.
    - ambiguous: 같은 타입 후보가 원본에 둘 이상 있는데, 이 값이 어느
      후보를 가리키는지 문맥(라벨)으로 특정할 수 없음 → 회색 "확인 필요"
      표시 대상.

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

    (2026-09-08 추가, 실제 fixture 테스트로 발견 — D01/D05가 요구한 "문맥상
    같은 항목으로 연결된 경우에만 비교" 요건이 누락돼 있었다) 이전에는
    "같은 타입 + 같은 값"이면 항목명이 전혀 달라도 그냥 matches/mismatches로
    판정했다 — 그 결과 "참여 인원 21명"(원본과 일치)과 아무 관계 없는
    "회의 횟수도 21회"가 같이 파랑으로 칠해지고, 반대로 "지원 인원 33명"이
    원본과 달라 빨강으로 칠해질 때 무관한 "자원봉사자도 33명"까지 같이
    빨갛게 칠해지는 오탐이 실제 한글 문서 테스트로 재현됐다. 이제는 같은
    타입 후보들의 값이 전부 같을 때만(어느 후보와 비교하든 판정이 똑같을
    때만, 혼동할 여지가 없을 때만) 문맥 없이 바로 비교하고, 서로 다른 값의
    후보가 둘 이상이면 report_values의 "context"(값 직전 문장/절 텍스트,
    extract_amounts 등이 채워둠)에 후보의 "label"(원본
    쪽 항목명, source_reader.read_excel_source가 채워둠)이 실제로 들어있는
    경우에만 그 후보와 비교한다. 문맥에 라벨이 전혀 안 걸리면(예: "회의
    횟수도 21회"의 문맥엔 "참여 인원"도 "지원 인원"도 없음) 어느 후보와도
    연결됐다고 볼 수 없으므로 ambiguous(회색)로 남긴다. "라벨이 문맥
    문자열에 포함되는지"라는 단순 부분일치만 쓰는 이유: 이 파일은 외부
    의존성 없는 순수 규칙 기반 모듈이라 LLM 의미 비교를 쓰지 않는다 —
    한계는 있지만("이번 사업의 참여 인원은 ... 총 21명"처럼 라벨이 멀리
    떨어지면 못 잡음), 적어도 완전히 무관한 항목을 같은 항목으로 오인하는
    사고는 막는다.

    (2026-09-11, 표 맥락 기능과의 통합) 값이 보고서의 "표 칸"에서 나온
    경우에는 rv["table_context"]에 그 칸의 행/열 머리말이 dict로 붙어 있다
    (verify_tool.py가 attach_table_context()로 채운다). 그때는 위의 라벨
    부분일치보다 훨씬 확실한 단서가 있으므로, 먼저 그 출처(연도 열 + 행
    머리말)가 실제로 맞는 원본 후보로 범위를 좁힌다. 순서가 중요하다:

      1) 표 칸이면 filter_candidates_by_context()로 후보를 좁힌다. 좁힌
         결과가 비면 → ambiguous(회색 "확인 필요").
      2) 그다음은 표 칸이든 평문이든 똑같은 규칙을 적용한다 — 남은 후보의
         서로 다른 값이 하나뿐이면 바로 비교하고, 여럿이면 문맥·라벨로
         연결되는 후보하고만 비교한다.

    (1)에서 후보가 안 남을 때 회색인 이유 — 표 맥락 기능을 처음 만들 때는
    이 자리를 초록(unverifiable)으로 뒀었다. 그때는 회색이라는 갈래 자체가
    없었기 때문이다. 지금은 두 색의 뜻이 분명히 다르다: 초록은 "원본에 이
    종류의 데이터가 아예 없어서 대조 자체가 불가능"이고, 회색은 "후보는
    있는데 어느 것과 비교해야 할지 못 좁혔다"이다. 표 칸의 열이 안 맞아
    후보가 사라지는 상황은 정확히 후자다(보통 같은 행의 다른 연도 열
    후보들이 멀쩡히 존재한다). 그래서 회색이 사실에 맞는 표시다.

    (2)를 표 칸에도 그대로 적용하는 이유 — 표 맥락은 "좁히는" 장치일 뿐이라,
    좁힐 근거가 없을 때(원본에 열/행 출처 정보가 아예 없는 세로형 시트 등)
    전부 통과시킨다. 그 상태에서 곧장 값만 비교하면, 표 안의 값만 평문보다
    관대하게 판정되는 모순이 생긴다 — 같은 원본·같은 숫자인데 문장에 쓰면
    회색, 표에 쓰면 파랑/빨강이 된다. 표 맥락으로 좁히지 못한 값은 평문과
    똑같은 라벨 연결 규칙을 거치게 해서 그 구멍을 막는다.
    """
    matches, mismatches, unverifiable, ambiguous = [], [], [], []
    for rv in report_values:
        candidates = [a for a in answer_pool if a["type"] == rv["type"]]
        if not candidates:
            unverifiable.append(rv)  # 원본에 이 타입 자체가 없음 → 대조 불가
            continue
        table_context = rv.get("table_context")
        if table_context:
            # (F14) 이 값이 표 안에 있고 행/열 머리말을 알아낸 경우에만, 출처가
            # 실제로 맞는 원본 값으로 후보를 좁힌다. table_context가 없는 값
            # (= 표 밖의 일반 문장)은 이 블록을 그냥 통과하므로 평문 동작이
            # 한 글자도 바뀌지 않는다.
            allowed = filter_candidates_by_context(table_context, candidates)
            if not allowed:
                # 타입이 같은 원본 값은 있지만, 그중 이 칸과 출처가 맞는 건
                # 하나도 없다 → "후보는 있는데 못 좁혔다"이므로 회색이다
                # (위 docstring의 색 구분 설명 참고).
                ambiguous.append(rv)
                continue
            candidates = allowed
        distinct_values = {c["normalized"] for c in candidates}
        if len(distinct_values) == 1:
            # (2026-09-08 정정) 후보 "개수"가 아니라 후보들의 "서로 다른 값
            # 개수"로 판단해야 한다 — compute_column_sums 등 원본 파생값이
            # 같은 값을 라벨 없이 하나 더 answer_pool에 추가하는 경우(예:
            # 데이터가 한 행뿐인 시트는 "값 합계"가 그 행 값과 같음)가
            # 실제로 흔해서, "후보 개수==1"만 보면 이런 데이터에서 매번
            # 라벨 요구 조건에 걸려 정상 판정까지 전부 회색(확인 필요)으로
            # 잘못 떨어지는 회귀가 실제 self-test로 재현됐다. 후보가 여럿이라도
            # 값이 전부 같으면 어느 후보와 비교하든 결과가 같으므로(어느
            # 항목이든 판정이 달라지지 않으므로) 문맥 연결 없이 바로 비교해도
            # 안전하다.
            (matches if rv["normalized"] in distinct_values else mismatches).append(rv)
            continue
        context = _matching_context_text(rv)
        labeled_candidates = [c for c in candidates if _label_matches_context(c.get("label"), context)]
        if not labeled_candidates:
            ambiguous.append(rv)  # 후보가 여럿인데 문맥으로 특정 못 함 → 확인 필요(회색)
            continue
        if any(c["normalized"] == rv["normalized"] for c in labeled_candidates):
            matches.append(rv)
        else:
            mismatches.append(rv)
    return {"matches": matches, "mismatches": mismatches, "unverifiable": unverifiable,
            "ambiguous": ambiguous}


def _selftest_categorize_values_connects_via_label_when_multiple_candidates():
    """(2026-09-08, 실제 fixture 테스트로 발견한 문제의 재현·수정 확인) 같은
    타입 후보가 둘 이상이어도, 문맥에 그 후보의 라벨이 들어있으면 정확히
    그 후보와만 비교해야 한다 — "참여 인원 21명"은 라벨 "참여 인원"이
    문맥에 있어 21과 비교해 일치(파랑), "지원 인원 33명"은 라벨 "지원
    인원"이 문맥에 있어 40과 비교해 불일치(빨강)로 갈려야 한다."""
    answer_pool = [
        {"type": "amount", "normalized": "21", "raw": "21", "label": "참여 인원",
         "source_file": "원본.xlsx", "location": "실적!B2"},
        {"type": "amount", "normalized": "40", "raw": "40", "label": "지원 인원",
         "source_file": "원본.xlsx", "location": "실적!B3"},
    ]
    text = "참여 인원 21명이며, 지원 인원 33명입니다"
    report_values = extract_values(text, default_year=2026)
    result = categorize_values(report_values, answer_pool)
    assert [r["raw"] for r in result["matches"]] == ["21"], result["matches"]
    assert [r["raw"] for r in result["mismatches"]] == ["33"], result["mismatches"]
    assert result["ambiguous"] == [], result["ambiguous"]
    print("_selftest_categorize_values_connects_via_label_when_multiple_candidates 통과:", result)


def _selftest_categorize_values_marks_ambiguous_when_context_has_no_label():
    """같은 타입 후보가 둘 이상인데 문맥에 어느 후보의 라벨도 없으면, 값이
    우연히 같거나 달라도 matches/mismatches로 단정하지 않고 ambiguous(회색
    "확인 필요")로 남겨야 한다 — 오탐(무관한 항목을 같은 항목으로 오인)
    방지가 핵심."""
    answer_pool = [
        {"type": "amount", "normalized": "21", "raw": "21", "label": "참여 인원",
         "source_file": "원본.xlsx", "location": "실적!B2"},
        {"type": "amount", "normalized": "40", "raw": "40", "label": "지원 인원",
         "source_file": "원본.xlsx", "location": "실적!B3"},
    ]
    text = "참여 인원 21명이며, 회의 횟수도 21회 진행됐고, 자원봉사자도 33명입니다"
    report_values = extract_values(text, default_year=2026)
    result = categorize_values(report_values, answer_pool)
    assert [r["raw"] for r in result["matches"]] == ["21"], result["matches"]  # 첫 "21"만 라벨 연결됨
    assert result["mismatches"] == [], result["mismatches"]  # 무관한 "33"을 빨강으로 단정하지 않음
    ambiguous_raw = sorted(r["raw"] for r in result["ambiguous"])
    assert ambiguous_raw == ["21", "33"], ambiguous_raw  # 문맥 불명확한 두번째 "21"과 "33"
    print("_selftest_categorize_values_marks_ambiguous_when_context_has_no_label 통과:", result)


def _selftest_categorize_values_single_candidate_ignores_missing_label():
    """(회귀 확인) 같은 타입 후보가 하나뿐이면, 그 후보에 label이 없거나
    (기존 answer_pool 형태) 문맥이 안 맞아도 여전히 문맥 없이 바로
    비교해야 한다 — 기존 self-test들이 label 없는 answer_pool을 그대로
    쓰므로 하위호환이 깨지면 안 된다."""
    answer_pool = [{"type": "amount", "normalized": "1850000", "raw": "1,850,000",
                     "source_file": "원본.xlsx", "location": "Sheet1!C15"}]
    report_values = [{"type": "amount", "normalized": "1850000", "raw": "185만원",
                       "span": (0, 4), "context": "전혀 무관한 문맥"}]
    result = categorize_values(report_values, answer_pool)
    assert [r["raw"] for r in result["matches"]] == ["185만원"], result
    print("_selftest_categorize_values_single_candidate_ignores_missing_label 통과:", result)


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


# ---------------------------------------------------------------------------
# (F14) 평문 텍스트 ↔ 표 격자 맞추기
#
# 한글의 GetTextFile("TEXT","")는 표를 "셀 하나당 한 줄"로, 왼→오른쪽·위→아래
# 순서로 늘어놓는다(직접 확인). 그래서 표 격자(table_to_df로 따로 읽은 것)를
# 이 줄 목록에 순서대로 맞춰보면, 어느 줄이 어느 행·열 칸인지 복원할 수 있다.
#
# 병합 셀 주의사항(직접 실험으로 확인): table_to_df는 병합된 칸의 값을 그
# 범위의 모든 칸에 그대로 복사해 넣는다. 그런데 평문 텍스트에는 그 내용이
# 딱 한 번만 나온다. 게다가 원래 글자가 있던 두 칸을 병합하면 병합된 셀의
# 내용 자체가 여러 문단이 된다 — 실제로 "점검"과 "예산"이 든 두 칸을
# 병합했더니 table_to_df는 두 행 모두에 '점검\r\n예산'을 넣었고, 평문에는
# "점검", "예산" 두 줄이 나왔다. 이걸 처리하지 않으면 줄 정렬이 한 줄씩
# 밀려서 표 전체의 맥락이 어긋난다.
#
# 그래서 각 칸마다 "그 자리에서 줄을 직접 소비하는" 시도를 먼저 하고,
# 실패했을 때만 "앞서 같은 값이 소비한 줄을 다시 가리키는"(병합 흔적)
# 처리로 넘어간다. 이 순서가 중요하다 — 순서를 뒤집으면, 서로 다른 두
# 행에 우연히 같은 값이 적혀 있는 정상적인 경우를 병합으로 오해한다.
# ---------------------------------------------------------------------------

_LINE_SPLIT_PATTERN = re.compile(r'\r\n|\r|\n')


def _flatten_cell(value) -> str:
    """표 한 칸의 값을 한 줄짜리 문자열로 만든다. 병합으로 여러 문단이 된
    칸은 공백으로 이어 붙인다. pandas가 빈 칸에 넣는 NaN도 빈 문자열로 본다."""
    if value is None:
        return ""
    text = str(value)
    if text.lower() == "nan":
        return ""
    return " ".join(x.strip() for x in _LINE_SPLIT_PATTERN.split(text) if x.strip())


def _split_lines_with_offsets(text: str) -> list:
    """텍스트를 줄 단위로 나누되 각 줄의 (내용, 시작위치, 끝위치)를 같이 준다.
    끝위치는 줄바꿈 문자를 뺀 위치다."""
    lines = []
    pos = 0
    for m in _LINE_SPLIT_PATTERN.finditer(text):
        lines.append((text[pos:m.start()], pos, m.start()))
        pos = m.end()
    lines.append((text[pos:], pos, len(text)))
    return lines


def _match_table_at(line_texts: list, start: int, grid: list):
    """grid(행×열 격자)를 line_texts의 start번째 줄부터 맞춰본다.
    성공하면 ({줄번호: (행, 열)}, 다음_탐색_시작줄)을, 실패하면 None."""
    pos = start
    line_to_cell = {}
    consumed_by_value = {}  # 이미 소비한 칸 내용 -> 그 칸이 차지했던 줄 번호들
    for r, row in enumerate(grid):
        for c, raw_value in enumerate(row):
            value = "" if raw_value is None else str(raw_value)
            if value.lower() == "nan":
                value = ""
            cell_lines = [x.strip() for x in _LINE_SPLIT_PATTERN.split(value)]
            end = pos + len(cell_lines)
            if end <= len(line_texts) and line_texts[pos:end] == cell_lines:
                # 직접 소비 — 항상 이쪽을 먼저 시도한다(위 주석 참고)
                for i in range(pos, end):
                    line_to_cell[i] = (r, c)
                consumed_by_value[value] = list(range(pos, end))
                pos = end
            elif value in consumed_by_value:
                pass  # 병합된 칸의 복사본 — 새 줄을 소비하지 않는다
            else:
                return None
    return line_to_cell, pos


def build_table_contexts(text: str, tables: list) -> list:
    """평문 텍스트와 표 격자 목록을 맞춰, 표 안 글자 구간마다 행/열 머리말을
    알려주는 목록을 만든다.

    tables는 [격자, 격자, ...]이고 각 격자는 [[머리말행...], [데이터행...], ...]
    형태(0행 = 열 머리말, 0열 = 행 머리말)로, 문서에 나오는 순서여야 한다.

    반환: [{"start", "end", "row_header", "column_header"}, ...]

    맥락은 데이터 칸(1행 1열 이후)에만 붙인다 — 머리말 칸에 들어있는 숫자는
    데이터가 아니라 라벨(예: "2024년"의 2024)이라 대조 대상이 아니고, 그런
    칸까지 건드리면 기존 동작이 예상 밖으로 바뀐다. 표를 평문에 맞추지
    못하면 그 표는 조용히 건너뛴다(맥락 없이 = 기존 동작)."""
    lines = _split_lines_with_offsets(text)
    line_texts = [t.strip() for t, _s, _e in lines]
    contexts = []
    search_from = 0
    for grid in tables:
        if not grid or len(grid) < 2 or not grid[0]:
            continue
        matched = None
        for start in range(search_from, len(line_texts)):
            result = _match_table_at(line_texts, start, grid)
            if result is not None:
                matched = result
                break
        if matched is None:
            continue  # 이 표는 평문과 맞출 수 없었다 → 맥락 없이 진행
        line_to_cell, next_start = matched
        for line_index, (r, c) in line_to_cell.items():
            if r == 0 or c == 0:
                continue  # 머리말 행/열 자체는 대조 대상이 아니다
            row_header = _flatten_cell(grid[r][0]) if grid[r] else ""
            column_header = _flatten_cell(grid[0][c]) if c < len(grid[0]) else ""
            _t, s, e = lines[line_index]
            contexts.append({"start": s, "end": e,
                             "row_header": row_header,
                             "column_header": column_header})
        search_from = next_start
    return contexts


def attach_table_context(values: list, contexts: list) -> list:
    """extract_values()가 뽑은 값들에 표 맥락을 달아준다. 값의 글자 구간이
    어떤 표 칸(줄) 안에 완전히 들어있으면 그 칸의 행/열 머리말을 붙인다.
    표 밖의 값은 아무것도 붙지 않아 기존 동작 그대로 간다.

    (2026-09-11) 표 맥락은 반드시 "table_context"라는 별도 키에 담는다.
    "context"는 _preceding_context()가 채우는 평문 문자열이라 타입도 뜻도
    다르다 — 같은 키를 쓰면 한쪽이 다른 쪽을 조용히 덮어써서, 문자열을
    기대하는 라벨 연결 코드에 dict가 들어가는 사고가 난다. 표 안의 값은
    두 정보를 모두 가진다(평문 문맥 + 표 머리말)."""
    for v in values:
        s, e = v["span"]
        for ctx in contexts:
            if ctx["start"] <= s and e <= ctx["end"]:
                v["table_context"] = {"row_header": ctx["row_header"],
                                      "column_header": ctx["column_header"]}
                break
    return values


# 아래 self-test들이 쓰는 평문 텍스트는 실제 한글이 표를 내보낸 모양 그대로다
# (직접 만든 .hwp에서 GetTextFile("TEXT","")로 받아 확인한 문자열).
_SAMPLE_TABLE_TEXT = (
    "2025년 주요 실적\r\n\r\n구분\r\n2024년\r\n2025년\r\n"
    "정기점검 실적\r\n1,694\r\n1,827\r\n예산\r\n500,000\r\n600,000\r\n\r\n이상입니다"
)
_SAMPLE_TABLE_GRID = [
    ["구분", "2024년", "2025년"],
    ["정기점검 실적", "1,694", "1,827"],
    ["예산", "500,000", "600,000"],
]


def _selftest_build_table_contexts_basic():
    contexts = build_table_contexts(_SAMPLE_TABLE_TEXT, [_SAMPLE_TABLE_GRID])
    got = {(_SAMPLE_TABLE_TEXT[c["start"]:c["end"]], c["row_header"], c["column_header"])
           for c in contexts}
    assert got == {
        ("1,694", "정기점검 실적", "2024년"),
        ("1,827", "정기점검 실적", "2025년"),
        ("500,000", "예산", "2024년"),
        ("600,000", "예산", "2025년"),
    }, got
    print("_selftest_build_table_contexts_basic 통과:", got)


def _selftest_build_table_contexts_merged_cell():
    """(직접 실험으로 확인한 실제 동작 재현) 글자가 들어있던 두 칸을 세로로
    병합하면, table_to_df는 두 행 모두에 '점검\\r\\n예산'을 넣고 평문에는
    "점검","예산" 두 줄이 나온다. 이걸 처리하지 못하면 줄 정렬이 한 줄
    밀려서 이 표의 맥락이 통째로 어긋난다."""
    text = ("머지 실험\r\n\r\n구분\r\n2024년\r\n2025년\r\n"
            "점검\r\n예산\r\n1,694\r\n1,827\r\n500,000\r\n600,000\r\n\r\n")
    grid = [
        ["구분", "2024년", "2025년"],
        ["점검\r\n예산", "1,694", "1,827"],
        ["점검\r\n예산", "500,000", "600,000"],
    ]
    contexts = build_table_contexts(text, [grid])
    got = {(text[c["start"]:c["end"]], c["row_header"], c["column_header"])
           for c in contexts}
    assert got == {
        ("1,694", "점검 예산", "2024년"),
        ("1,827", "점검 예산", "2025년"),
        ("500,000", "점검 예산", "2024년"),
        ("600,000", "점검 예산", "2025년"),
    }, got
    print("_selftest_build_table_contexts_merged_cell 통과:", got)


def _selftest_build_table_contexts_repeated_value_is_not_treated_as_merge():
    """서로 다른 두 행에 우연히 같은 값이 적혀 있는 정상적인 표를, 병합
    흔적으로 오해하면 안 된다("직접 소비 먼저" 순서가 지켜지는지 확인)."""
    text = "구분\r\n2024년\r\n갑\r\n100\r\n을\r\n100\r\n"
    grid = [["구분", "2024년"], ["갑", "100"], ["을", "100"]]
    contexts = build_table_contexts(text, [grid])
    got = sorted((text[c["start"]:c["end"]], c["row_header"]) for c in contexts)
    assert got == [("100", "갑"), ("100", "을")], got
    print("_selftest_build_table_contexts_repeated_value_is_not_treated_as_merge 통과:", got)


def _selftest_build_table_contexts_unmatchable_table_is_skipped():
    """평문과 맞출 수 없는 표(격자와 본문이 어긋남)는 크래시 없이 조용히
    건너뛰고, 맥락 없이 기존 동작으로 가야 한다."""
    contexts = build_table_contexts("전혀 다른 내용입니다", [_SAMPLE_TABLE_GRID])
    assert contexts == [], contexts
    print("_selftest_build_table_contexts_unmatchable_table_is_skipped 통과")


def _selftest_attach_table_context_does_not_clobber_prose_context():
    """(2026-09-11, 두 기능 통합의 핵심 계약) 표 맥락은 "table_context"에만
    담기고, _preceding_context()가 채운 평문 "context" 문자열을 덮어쓰지
    않아야 한다. 덮어쓰면 라벨 연결 코드에 dict가 들어가 터진다."""
    values = extract_values(_SAMPLE_TABLE_TEXT, default_year=2026)
    attach_table_context(values, build_table_contexts(_SAMPLE_TABLE_TEXT,
                                                      [_SAMPLE_TABLE_GRID]))
    in_table = [v for v in values if v.get("table_context")]
    assert in_table, values
    for v in values:
        assert isinstance(v.get("context", ""), str), v  # 평문 문맥은 계속 문자열
        if v.get("table_context"):
            assert isinstance(v["table_context"], dict), v
    cell = [v for v in in_table if v["raw"] == "1,694"][0]
    assert cell["table_context"] == {"row_header": "정기점검 실적",
                                     "column_header": "2024년"}, cell
    print("_selftest_attach_table_context_does_not_clobber_prose_context 통과")


def _selftest_attach_table_context_and_categorize_catches_wrong_year():
    """이 기능이 존재하는 이유 그 자체 — 표의 "2024년" 칸에 2025년 값을
    잘못 옮겨적은 오류를, 표 맥락 없이는 특정할 수 없다.

    맥락을 붙이면 각 칸이 자기 열의 원본 값하고만 대조되므로 2024년 칸은
    빨강(오류), 2025년 칸은 파랑(정상)으로 갈린다."""
    text = ("구분\r\n2024년\r\n2025년\r\n정기점검 실적\r\n1,827\r\n1,827\r\n")
    grid = [["구분", "2024년", "2025년"], ["정기점검 실적", "1,827", "1,827"]]
    answer_pool = [
        {"type": "amount", "normalized": "1694", "raw": "1694",
         "row_label": "정기점검 실적", "column_header": "2024년"},
        {"type": "amount", "normalized": "1827", "raw": "1827",
         "row_label": "정기점검 실적", "column_header": "2025년"},
    ]
    values = extract_values(text, default_year=2026)
    # 머리말 칸의 "2024년"/"2025년"에서도 숫자가 뽑히지만 이 변경과 무관한
    # 기존 동작이라(머리말 칸에는 맥락을 붙이지 않는다) 데이터 값만 본다.
    def _data_cells(result_bucket):
        return [m["raw"] for m in result_bucket if m["raw"] == "1,827"]

    # 맥락을 붙이기 전: 후보 값이 1694/1827 둘로 갈리는데 표 밖 문맥으로는
    # 어느 항목인지 못 좁히므로, 두 칸 다 회색(확인 필요)에 머문다.
    # (2026-09-11 정정: 표 맥락 기능을 처음 만들 때는 이 자리가 "둘 다
    # 파랑"이었다. 그 뒤 master에 들어간 문맥 연결 수정이 "라벨로 연결
    # 못 하면 단정하지 않는다"로 바꿔서, 가짜 파랑은 이미 사라졌다.
    # 표 맥락이 하는 일은 그 회색을 확실한 빨강/파랑으로 끌어올리는 것이다.)
    before = categorize_values([dict(v) for v in values], answer_pool)
    assert _data_cells(before["matches"]) == [], before["matches"]
    assert _data_cells(before["mismatches"]) == [], before["mismatches"]
    assert _data_cells(before["ambiguous"]) == ["1,827", "1,827"], before["ambiguous"]

    attach_table_context(values, build_table_contexts(text, [grid]))
    after = categorize_values(values, answer_pool)
    mismatched = _data_cells(after["mismatches"])
    matched = _data_cells(after["matches"])
    # 2024년 칸의 1,827은 오류로 잡히고, 2025년 칸의 1,827은 그대로 정상
    assert mismatched == ["1,827"], mismatched
    assert matched == ["1,827"], matched
    wrong_cell = [m for m in after["mismatches"] if m["raw"] == "1,827"][0]
    right_cell = [m for m in after["matches"] if m["raw"] == "1,827"][0]
    assert wrong_cell["table_context"]["column_header"] == "2024년", wrong_cell
    assert right_cell["table_context"]["column_header"] == "2025년", right_cell
    print("_selftest_attach_table_context_and_categorize_catches_wrong_year 통과")


def _selftest_categorize_mixed_pool_column_blind_does_not_defeat_filter():
    """열 정보 없는 후보(한글/PDF 원본) 하나 때문에 "2024년 칸에 적힌
    2025년 값"이 다시 파랑이 되면 안 된다."""
    ctx = {"row_header": "정기점검 실적", "column_header": "2024년"}
    answer_pool = [
        {"type": "amount", "normalized": "1694",
         "row_label": "정기점검 실적", "column_header": "2024"},
        {"type": "amount", "normalized": "1827",
         "row_label": "정기점검 실적", "column_header": "2025"},
        {"type": "amount", "normalized": "1827"},  # 한글 원본에서 온 열 모르는 후보
    ]
    rv = {"type": "amount", "normalized": "1827", "raw": "1,827",
          "span": (0, 5), "table_context": ctx}
    result = categorize_values([rv], answer_pool)
    assert [m["raw"] for m in result["mismatches"]] == ["1,827"], result
    assert result["matches"] == [], result
    print("_selftest_categorize_mixed_pool_column_blind_does_not_defeat_filter 통과")


def _selftest_categorize_abbreviated_year_column_still_catches_error():
    """(A 브랜치가 진짜 .hwp+.xlsx로 재현해 고친 회귀의 단위 재현) 보고서 표
    머리말이 "'24년"이고 원본 엑셀 머리말이 "2024"일 때, 잘못된 값은
    빨강(오류)으로, 맞는 값은 파랑(정상)으로 나와야 한다. 고치기 전에는
    둘 다 대조 실패로 떨어져 진짜 오류가 통째로 사라졌다."""
    answer_pool = [
        {"type": "amount", "normalized": "1694",
         "row_label": "정기점검 실적", "column_header": "2024"},
        {"type": "amount", "normalized": "1827",
         "row_label": "정기점검 실적", "column_header": "2025"},
    ]
    wrong = {"type": "amount", "normalized": "1900", "raw": "1,900", "span": (0, 5),
             "table_context": {"row_header": "정기점검 실적", "column_header": "'24년"}}
    right = {"type": "amount", "normalized": "1827", "raw": "1,827", "span": (6, 11),
             "table_context": {"row_header": "정기점검 실적", "column_header": "'25년"}}
    result = categorize_values([wrong, right], answer_pool)
    assert [m["raw"] for m in result["mismatches"]] == ["1,900"], result
    assert [m["raw"] for m in result["matches"]] == ["1,827"], result
    assert result["unverifiable"] == [], result
    assert result["ambiguous"] == [], result
    print("_selftest_categorize_abbreviated_year_column_still_catches_error 통과")


def _selftest_categorize_no_matching_column_falls_back_to_ambiguous():
    """이 칸과 출처가 맞는 원본 값이 하나도 없으면 빨강(오류)이 아니어야
    한다 — 원본에 그 연도 열이 없다는 이유만으로 멀쩡한 값을 오류로 찍으면
    안 된다.

    (2026-09-11 변경) 이 자리는 표 맥락 기능을 처음 만들 때 초록
    (unverifiable)이었다. 그때는 회색이라는 갈래가 아예 없었기 때문이다.
    지금은 초록이 "원본에 이 종류의 데이터 자체가 없음"이라는 더 좁은 뜻을
    갖고, 회색이 "후보는 있는데 못 좁혔다"를 뜻한다. 여기 상황은 amount
    후보가 멀쩡히 존재하는데 열이 안 맞아 못 좁힌 것이므로 회색이 맞다."""
    text = "구분\r\n2030년\r\n예산\r\n777\r\n"
    grid = [["구분", "2030년"], ["예산", "777"]]
    answer_pool = [{"type": "amount", "normalized": "500", "raw": "500",
                    "row_label": "예산", "column_header": "2024년"}]
    values = extract_values(text, default_year=2026)
    attach_table_context(values, build_table_contexts(text, [grid]))
    result = categorize_values(values, answer_pool)
    # 머리말 칸의 "2030년"은 type="year"로 뽑혀 원본에 year 후보가 없으니
    # 초록(대조불가)으로 간다 — 이 변경과 무관한 기존 동작이라(머리말 칸에는
    # 맥락을 안 붙인다) 데이터 칸 값만 골라서 확인한다.
    assert [m["raw"] for m in result["ambiguous"]] == ["777"], result
    assert [m["raw"] for m in result["unverifiable"] if m["raw"] == "777"] == [], result
    assert [m["raw"] for m in result["mismatches"] if m["raw"] == "777"] == [], result
    print("_selftest_categorize_no_matching_column_falls_back_to_ambiguous 통과")


def _selftest_categorize_table_cell_uses_row_header_as_label_context():
    """(2026-09-11, 두 기능을 합치면서 생긴 구멍을 메움) 세로형("항목/값")
    원본은 값마다 label(항목명)을 갖고, 평문 값은 문장 문맥으로 그 라벨과
    연결된다. 그런데 표 안의 값은 평문 문맥이 거의 항상 빈 문자열이다 —
    표를 펼치면 값 바로 앞이 줄바꿈이라 문맥이 잘리기 때문이다. 그래서
    같은 원본·같은 숫자인데 문장에 쓰면 연결되고 표에 쓰면 회색이 되는
    모순이 생긴다. 표 칸에서는 그 행의 머리말을 문맥으로 함께 쓴다."""
    answer_pool = [
        {"type": "amount", "normalized": "21", "raw": "21", "label": "참여 인원"},
        {"type": "amount", "normalized": "40", "raw": "40", "label": "지원 인원"},
    ]
    text = "구분\r\n인원\r\n참여 인원\r\n21\r\n지원 인원\r\n33\r\n"
    grid = [["구분", "인원"], ["참여 인원", "21"], ["지원 인원", "33"]]
    values = extract_values(text, default_year=2026)
    attach_table_context(values, build_table_contexts(text, [grid]))
    result = categorize_values(values, answer_pool)
    assert [m["raw"] for m in result["matches"]] == ["21"], result["matches"]
    assert [m["raw"] for m in result["mismatches"]] == ["33"], result["mismatches"]
    print("_selftest_categorize_table_cell_uses_row_header_as_label_context 통과")


def _selftest_categorize_table_cell_not_more_lenient_than_prose():
    """(2026-09-11, 두 기능 통합에서 가장 중요한 계약) 표 맥락은 후보를
    "좁히는" 장치일 뿐이므로, 좁힐 근거가 없을 때 표 안의 값이 평문보다
    관대하게 판정되면 안 된다. 원본에 열/행 출처 정보가 전혀 없고 라벨도
    비어 있으면, 같은 숫자는 문장에 있든 표에 있든 똑같이 회색이어야 한다.
    (여기서 표만 파랑/빨강이 되면, master에 이미 들어간 "라벨로 연결 못
    하면 단정하지 않는다" 보호가 표 안에서만 조용히 풀린다.)"""
    answer_pool = [
        {"type": "amount", "normalized": "1850000", "raw": "1850000", "label": ""},
        {"type": "amount", "normalized": "342", "raw": "342", "label": ""},
    ]
    prose = extract_values("예산은 1,850,000원이고 기타는 7,777,777원입니다", default_year=2026)
    prose_result = categorize_values(prose, answer_pool)
    assert sorted(m["raw"] for m in prose_result["ambiguous"]) == \
        ["1,850,000원", "7,777,777원"], prose_result

    text = "구분\r\n금액\r\n예산\r\n1,850,000\r\n기타\r\n7,777,777\r\n"
    grid = [["구분", "금액"], ["예산", "1,850,000"], ["기타", "7,777,777"]]
    values = extract_values(text, default_year=2026)
    attach_table_context(values, build_table_contexts(text, [grid]))
    table_result = categorize_values(values, answer_pool)
    assert sorted(m["raw"] for m in table_result["ambiguous"]) == \
        ["1,850,000", "7,777,777"], table_result
    assert table_result["matches"] == [], table_result
    assert table_result["mismatches"] == [], table_result
    print("_selftest_categorize_table_cell_not_more_lenient_than_prose 통과")


def _selftest_label_matches_context_fragment_fix():
    """(2026-09-12, 실사용 문서에서 재현) label이 "지표명 + 단위" 두 문자열
    칸을 이어붙인 통짜("정기점검 실적 개소")면, 단위가 숫자 뒤에만 붙는
    보통 문장에서는 통짜 전체가 절대 안 걸린다. 지표명 조각("정기점검
    실적")만으로도 충분히 연결된다."""
    assert _label_matches_context("정기점검 실적 개소", "정기점검 실적 : ('23) ") is True
    # 너무 짧은 조각("개소")만 겹치고 나머지 실질 조각("우리동네")은 안
    # 겹치면 연결로 보면 안 된다 - 다른 지표도 흔히 같은 단위를 쓰므로,
    # 짧은 단위어만으로 걸리면 엉뚱한 행끼리 서로 라벨이 섞인다.
    assert _label_matches_context("우리동네 개소", "다른동네 개소 관련 문장") is False
    assert _label_matches_context("", "아무 문장") is False
    print("_selftest_label_matches_context_fragment_fix 통과")


def _selftest_preceding_context_ignores_digit_grouping_comma():
    """(2026-09-12, 실사용 문서에서 재현) "정기점검 실적 : ('23) 1,595개소 →
    ('24) 1,694 → ('25) 1,827"에서 1,694의 문맥을 구하면, 진짜 절 구분
    쉼표가 하나도 없는데도 "1,595"의 자릿수 구분 쉼표에서 끊겨 항목명이
    통째로 잘려나갔었다."""
    text = "정기점검 실적 : ('23) 1,595개소 -> ('24) 1,694 -> ('25) 1,827"
    idx = text.index("1,694")
    context = _preceding_context(text, idx)
    assert "정기점검 실적" in context, context
    print("_selftest_preceding_context_ignores_digit_grouping_comma 통과:", context)


def _selftest_categorize_prose_wide_source_real_document_shape():
    """(2026-09-12, 실사용 문서 전체 경로 재현) 위 두 수정이 함께 있어야만
    통과하는 통합 수준 테스트 — 보고서는 순수 평문(표 아님), 원본 엑셀은
    연도 약칭 머리말("'23","'24","'25")을 쓰는 가로형 표. 이 조합이
    실사용에서 그대로 재현됐다: 셋 다 회색으로 떨어졌었다."""
    source_pool = [
        {"type": "amount", "normalized": "1595", "raw": "1595",
         "label": "정기점검 실적 개소", "row_label": "정기점검 실적 개소",
         "column_header": "'23"},
        {"type": "amount", "normalized": "1694", "raw": "1694",
         "label": "정기점검 실적 개소", "row_label": "정기점검 실적 개소",
         "column_header": "'24"},
        {"type": "amount", "normalized": "1827", "raw": "1827",
         "label": "정기점검 실적 개소", "row_label": "정기점검 실적 개소",
         "column_header": "'25"},
    ]
    text = "정기점검 실적 : ('23) 1,595개소 -> ('24) 1,694 -> ('25) 1,827"
    report_values = extract_values(text, default_year=2026)
    result = categorize_values(report_values, source_pool)
    matched_raws = sorted(m["raw"] for m in result["matches"])
    assert matched_raws == ["1,595", "1,694", "1,827"], result
    assert result["mismatches"] == [] and result["ambiguous"] == [], result
    print("_selftest_categorize_prose_wide_source_real_document_shape 통과")


def compute_column_sums(rows: list[dict], source_file: str, sheet: str,
                        wide_columns: set = None) -> list[dict]:
    """엑셀에서 읽은 행(딕셔너리 리스트)에서, 숫자로만 이루어진 각 컬럼의 단순 합계를
    미리 계산해 정답 풀에 추가할 항목으로 반환한다 (PRD 12-2: 합계만, 증감률/평균은 제외).
    float을 직접 sum()하지 않고 Decimal로 변환해 더한다 — 소수를 포함하는 컬럼(단가 등)의
    합계에서 이진 부동소수점 오차가 생겨 Task 5의 정확일치 비교와 충돌하는 것을 방지한다.
    bool은 int의 서브클래스라 isinstance(v, (int, float))를 그냥 두면 True/False가
    섞인 컬럼(완료여부 등)에서 Decimal("True") 파싱 오류가 나므로 명시적으로 제외한다.

    (F14) wide_columns가 주어지면(가로형 표에서 온 열 이름들), 그 열의 합계
    항목에도 열 머리말과 행 라벨("합계")을 붙인다. 안 붙이면 구멍이 생긴다 —
    출처 정보가 없는 후보는 "모르는 건 막지 않는다" 원칙에 따라 열 필터를
    무조건 통과하므로, 데이터 행이 하나뿐이라 합계가 그 행의 값과 같아지는
    흔한 경우에 "2024년 칸에 적힌 2025년 값"이 2025년 열의 합계와 일치해
    다시 정상(파랑)으로 빠져나간다(실제로 이 테스트를 만들다 발견). 합계는
    그 열에 속한 값이 맞으므로 열 머리말을 붙이는 게 사실에 맞고, 행 라벨을
    "합계"로 두면 개별 항목 행과는 자연히 연결되지 않는다(보고서 표에 진짜
    "합계" 행이 있으면 그때는 제대로 연결된다).
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
            entry = {
                "type": "amount", "normalized": normalized, "raw": f"{col} 합계",
                "source_file": source_file, "location": f"{sheet}!{col}(합계)",
            }
            if wide_columns and str(col).strip() in wide_columns:
                entry["column_header"] = str(col).strip()
                entry["row_label"] = "합계"
            results.append(entry)
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
    _selftest_extract_amounts_eok_jo_units()
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
    _selftest_extract_values_excludes_year_abbreviation_with_backtick()
    _selftest_extract_values_excludes_year_in_parens_any_symbol()
    _selftest_find_toc_number_spans_handles_spaced_heading()
    _selftest_extract_values_excludes_toc_page_numbers()
    _selftest_extract_values_excludes_toc_embedded_item_number()
    _selftest_extract_values_excludes_list_marker_dot()
    _selftest_extract_values_excludes_list_marker_paren()
    _selftest_extract_values_list_marker_does_not_exclude_decimal_amount()
    _selftest_extract_values_list_marker_does_not_swallow_real_numbers()
    _selftest_extract_values_excludes_list_marker_in_real_positions()
    _selftest_extract_values_excludes_month_sequence()
    _selftest_extract_years_full_and_abbreviated()
    _selftest_extract_years_bare_abbreviation_without_quote()
    _selftest_extract_values_year_word_excluded_from_amounts()
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
    _selftest_categorize_values_connects_via_label_when_multiple_candidates()
    _selftest_categorize_values_marks_ambiguous_when_context_has_no_label()
    _selftest_categorize_values_single_candidate_ignores_missing_label()
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
    # (F14) 표 맥락 + 열 출처 구분
    _selftest_split_column_identity()
    _selftest_split_column_identity_year_notations()
    _selftest_column_identity_agrees_truth_table()
    _selftest_column_identity_agrees_across_year_notations()
    _selftest_candidate_allowed_by_context_falls_back_when_unknown()
    _selftest_filter_candidates_by_context_mixed_pool()
    _selftest_build_table_contexts_basic()
    _selftest_build_table_contexts_merged_cell()
    _selftest_build_table_contexts_repeated_value_is_not_treated_as_merge()
    _selftest_build_table_contexts_unmatchable_table_is_skipped()
    _selftest_attach_table_context_does_not_clobber_prose_context()
    _selftest_attach_table_context_and_categorize_catches_wrong_year()
    _selftest_categorize_mixed_pool_column_blind_does_not_defeat_filter()
    _selftest_categorize_abbreviated_year_column_still_catches_error()
    _selftest_categorize_no_matching_column_falls_back_to_ambiguous()
    _selftest_categorize_table_cell_uses_row_header_as_label_context()
    _selftest_categorize_table_cell_not_more_lenient_than_prose()
    _selftest_label_matches_context_fragment_fix()
    _selftest_preceding_context_ignores_digit_grouping_comma()
    _selftest_categorize_prose_wide_source_real_document_shape()
