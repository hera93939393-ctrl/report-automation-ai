# -*- coding: utf-8 -*-
"""
최종 검토 체크 (F7) — 규칙 기반, AI/외부 통신 없음
- 항목 번호 형식 체크: 1. -> 가. -> 1) -> 가) 순서가 올바른지
- 계산(합계) 오류 체크: 숫자 목록과 합계가 실제로 맞는지
모든 처리는 로컬에서만 이루어지며 외부로 아무것도 전송하지 않는다.
"""
import re

# 개조식 문서에서 흔히 쓰는 번호 체계 순서 (레벨 0~3)
LEVEL_PATTERNS = [
    re.compile(r"^(\d+)\.\s"),        # 1. 2. 3.
    re.compile(r"^([가-힣])\.\s"),     # 가. 나. 다.
    re.compile(r"^(\d+)\)\s"),         # 1) 2) 3)
    re.compile(r"^([가-힣])\)\s"),     # 가) 나) 다)
]

GANADA = "가나다라마바사아자차카타파하"


def _level_of(line: str):
    for level, pat in enumerate(LEVEL_PATTERNS):
        if pat.match(line.strip()):
            return level
    return None


def check_item_numbering(text: str):
    """줄마다 번호 체계를 확인해, 순서가 어긋난 곳을 오류로 반환한다.
    반환: [{"line": 줄번호, "content": 원문, "issue": 설명}, ...]
    """
    issues = []
    counters = {0: 0, 1: 0, 2: 0, 3: 0}  # 레벨별 마지막으로 본 순번(숫자 또는 가나다 인덱스)

    for i, line in enumerate(text.splitlines(), start=1):
        level = _level_of(line)
        if level is None:
            continue

        # 이 레벨보다 깊은 레벨의 카운터는 새 항목이 시작됐으니 리셋
        for deeper in range(level + 1, 4):
            counters[deeper] = 0

        m = LEVEL_PATTERNS[level].match(line.strip())
        token = m.group(1)

        if level in (0, 2):  # 숫자 레벨
            num = int(token)
            expected = counters[level] + 1
            if num != expected:
                issues.append({
                    "line": i, "content": line.strip(),
                    "issue": f"번호 순서 오류 (기대: {expected}, 실제: {num})"
                })
            counters[level] = num
        else:  # 가나다 레벨
            idx = GANADA.find(token) + 1
            expected = counters[level] + 1
            if idx != expected:
                expected_char = GANADA[expected - 1] if 0 < expected <= len(GANADA) else "?"
                issues.append({
                    "line": i, "content": line.strip(),
                    "issue": f"번호 순서 오류 (기대: {expected_char}, 실제: {token})"
                })
            counters[level] = idx

    return issues


NUMBER_PATTERN = re.compile(r"(?<!\d)[-+]?\d[\d,]*(?:\.\d+)?")

# 합계 계산에 끼면 안 되는 숫자(날짜/전화번호/시간)를 미리 지워내기 위한 패턴들.
# 예: "2026-09-07", "031-1234-5678", "14:00~16:00", "9월 7일" 같은 표기는
# 예산 항목 숫자가 아닌데도 NUMBER_PATTERN에 그대로 걸려 합계를 오염시킬 수 있음.
NON_VALUE_PATTERNS = [
    re.compile(r"\d{2,4}-\d{1,2}-\d{1,2}"),           # 날짜(2026-09-07)
    re.compile(r"\d{2,4}-\d{3,4}-\d{4}"),              # 전화번호(031-1234-5678)
    re.compile(r"\d{1,2}:\d{2}(?:~\d{1,2}:\d{2})?"),   # 시간/시간대(14:00~16:00)
    re.compile(r"\d{1,2}월\s*\d{1,2}일"),               # 한글 날짜(9월 7일)
]


def _to_number(s: str) -> float:
    return float(s.replace(",", ""))


def _strip_non_value_numbers(line: str) -> str:
    for pat in NON_VALUE_PATTERNS:
        line = pat.sub("", line)
    return line


def check_arithmetic(text: str, total_keywords=("합계", "총액", "총 금액", "총계", "소계")):
    """'합계/총액' 등이 포함된 줄을 찾아, 그 앞에 나온 숫자들의 합과 실제로 맞는지 확인한다.
    반환: [{"line": 줄번호, "stated": 표기된 합계, "computed": 계산된 합, "issue": 설명}, ...]
    """
    issues = []
    lines = text.splitlines()
    pending_numbers = []

    for i, line in enumerate(lines, start=1):
        is_total_line = any(kw in line for kw in total_keywords)
        # 번호 표시 줄(1. / 가. / 1) / 가))은 항목 순번일 뿐 값이 아니므로 합산 대상에서 제외
        if _level_of(line) is not None and not is_total_line:
            continue
        numbers_in_line = [_to_number(n) for n in NUMBER_PATTERN.findall(_strip_non_value_numbers(line))]

        if is_total_line:
            if numbers_in_line:
                stated = numbers_in_line[-1]  # 합계 줄의 마지막 숫자를 표기된 합계로 간주
                computed = sum(pending_numbers)
                if pending_numbers and abs(stated - computed) > 0.01:
                    issues.append({
                        "line": i, "content": line.strip(),
                        "stated": stated, "computed": computed,
                        "issue": f"합계 불일치 (표기: {stated}, 계산: {computed})"
                    })
            pending_numbers = []  # 합계 줄을 지났으니(숫자 유무와 무관하게) 다음 구간을 위해 초기화
        elif numbers_in_line:
            pending_numbers.extend(numbers_in_line)

    return issues


def run_all_checks(text: str):
    """모든 규칙 기반 체크를 실행해 결과를 정리해 반환한다."""
    numbering_issues = check_item_numbering(text)
    arithmetic_issues = check_arithmetic(text)
    return {
        "numbering_issues": numbering_issues,
        "arithmetic_issues": arithmetic_issues,
        "total_issue_count": len(numbering_issues) + len(arithmetic_issues),
    }


if __name__ == "__main__":
    sample = """1. 추진 배경
가. 주민 민원 증가
나. 소음 문제 청취 필요
2. 세부 계획
1) 일정: 다음 주 화요일
3) 장소: 회의실
3. 예산
- 인건비: 500,000
- 재료비: 300,000
- 합계: 900,000
"""
    result = run_all_checks(sample)
    print("항목 번호 오류:", result["numbering_issues"])
    print("계산 오류:", result["arithmetic_issues"])
    print("총 오류 수:", result["total_issue_count"])
