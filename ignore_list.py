"""ignore_list.py — 숫자검증에서 "이건 괜찮아, 무시해"로 지정한 값을 기억해
다음 검증부터 오류(빨강)로 표시하지 않게 하는 오탐 학습 도구.
speed_tracker.py와 동일한 순수 함수 + 로컬 JSON 파일 조합 관례를 따른다."""
import json
import os

DEFAULT_IGNORE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ignore_list.json")


def _load_raw_list(path: str) -> list[str]:
    if not os.path.exists(path):
        return []
    with open(path, "r", encoding="utf-8") as f:
        try:
            return json.load(f)
        except json.JSONDecodeError:
            return []


def _save_raw_list(values: list[str], path: str) -> None:
    with open(path, "w", encoding="utf-8") as f:
        json.dump(values, f, ensure_ascii=False)


def record_ignored_value(raw: str, path: str = DEFAULT_IGNORE_PATH) -> None:
    """raw 값을 무시 목록에 추가한다(중복 저장 방지). speed_tracker.py의
    record_call_speed()와 달리 최근 N건 제한은 두지 않는다 - 무시 목록은
    "한 번 안전하다고 확인된 값"의 누적 화이트리스트라서, 오래된 항목이라도
    계속 유효해야 한다(속도 기록처럼 최신 것만 의미있는 데이터가 아님)."""
    values = _load_raw_list(path)
    if raw not in values:
        values.append(raw)
        _save_raw_list(values, path)


def load_ignored_values(path: str = DEFAULT_IGNORE_PATH) -> list[str]:
    """무시 목록에 등록된 값들을 리스트로 반환한다. 파일이 없으면 빈 리스트."""
    return _load_raw_list(path)


def _selftest_record_and_load_ignored_values():
    import tempfile
    path = os.path.join(tempfile.gettempdir(), "_test_ignore_list.json")
    if os.path.exists(path):
        os.remove(path)
    try:
        assert load_ignored_values(path) == []
        record_ignored_value("9999999", path=path)
        record_ignored_value("1850000", path=path)
        assert load_ignored_values(path) == ["9999999", "1850000"]
        print("record_ignored_value + load_ignored_values 통과")
    finally:
        if os.path.exists(path):
            os.remove(path)


def _selftest_record_ignored_value_deduplicates():
    """같은 값을 두 번 기록해도 목록에는 한 번만 남아야 한다."""
    import tempfile
    path = os.path.join(tempfile.gettempdir(), "_test_ignore_list_중복.json")
    if os.path.exists(path):
        os.remove(path)
    try:
        record_ignored_value("9999999", path=path)
        record_ignored_value("9999999", path=path)
        assert load_ignored_values(path) == ["9999999"], load_ignored_values(path)
        print("record_ignored_value(중복 방지) 통과")
    finally:
        if os.path.exists(path):
            os.remove(path)


def _selftest_load_ignored_values_missing_file_returns_empty():
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_존재하지_않는_ignore_list.json")
    assert load_ignored_values(path) == []
    print("load_ignored_values(파일없음) 통과")


if __name__ == "__main__":
    _selftest_record_and_load_ignored_values()
    _selftest_record_ignored_value_deduplicates()
    _selftest_load_ignored_values_missing_file_returns_empty()
