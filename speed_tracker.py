"""speed_tracker.py — Ollama 로컬 모델 호출 속도를 기록해서 다음 호출의
예상 소요시간을 계산한다. 순수 함수 + 로컬 JSON 파일(speed_log.json,
최근 20건만 유지) 조합."""
import json
import os

_DEFAULT_LOG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "speed_log.json")
_MAX_RECORDS = 20


def _load_speeds(path: str) -> list[float]:
    if not os.path.exists(path):
        return []
    with open(path, "r", encoding="utf-8") as f:
        try:
            return json.load(f)
        except json.JSONDecodeError:
            return []


def _save_speeds(speeds: list[float], path: str) -> None:
    with open(path, "w", encoding="utf-8") as f:
        json.dump(speeds, f)


def record_call_speed(eval_count: int, eval_duration_ns: int, path: str = _DEFAULT_LOG_PATH) -> None:
    """Ollama 응답의 eval_count(생성된 토큰 수)와 eval_duration(나노초)을
    받아 초당 토큰수를 계산해 path의 JSON 파일에 누적 기록한다(최근
    _MAX_RECORDS건만 유지). eval_count나 eval_duration_ns가 0 이하이면
    (0으로 나누기 방지 + 의미 없는 기록 배제) 아무 것도 기록하지 않는다."""
    if eval_count <= 0 or eval_duration_ns <= 0:
        return
    speed = eval_count / (eval_duration_ns / 1e9)
    speeds = _load_speeds(path)
    speeds.append(speed)
    speeds = speeds[-_MAX_RECORDS:]
    _save_speeds(speeds, path)


def estimate_seconds(expected_tokens: int, path: str = _DEFAULT_LOG_PATH) -> tuple[float, float] | None:
    """기록된 속도들의 평균으로 expected_tokens를 생성하는 데 걸릴 예상
    시간(초)을 계산해, 평균 대비 ±20% 오차범위를 (최소, 최대) 튜플로
    반환한다. 기록이 하나도 없으면 아직 예측 불가하므로 None."""
    speeds = _load_speeds(path)
    if not speeds:
        return None
    avg_speed = sum(speeds) / len(speeds)
    base_seconds = expected_tokens / avg_speed
    return (base_seconds * 0.8, base_seconds * 1.2)


def _selftest_estimate_seconds_none_without_records():
    import tempfile
    path = os.path.join(tempfile.gettempdir(), "_test_speed_log_없음.json")
    if os.path.exists(path):
        os.remove(path)
    try:
        assert estimate_seconds(100, path=path) is None
        print("estimate_seconds 통과(기록 없음 -> None)")
    finally:
        if os.path.exists(path):
            os.remove(path)


def _selftest_record_and_estimate_seconds():
    import tempfile
    path = os.path.join(tempfile.gettempdir(), "_test_speed_log.json")
    if os.path.exists(path):
        os.remove(path)
    try:
        record_call_speed(100, 10_000_000_000, path=path)  # 10초에 100토큰 = 초당 10토큰
        result = estimate_seconds(200, path=path)
        assert result is not None, result
        lo, hi = result
        assert 15 <= lo <= 20 <= hi <= 25, result  # 200토큰 / 초당10토큰 = 20초 근방(±20%)
        print("record_call_speed + estimate_seconds 통과:", result)
    finally:
        if os.path.exists(path):
            os.remove(path)


if __name__ == "__main__":
    _selftest_estimate_seconds_none_without_records()
    _selftest_record_and_estimate_seconds()
