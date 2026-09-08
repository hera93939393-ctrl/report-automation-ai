"""code_fix_tool.py — D05 최소 코딩 시나리오용 도구 3개.
`coding_fixture/` 폴더의 결함 함수 1개 + pytest 테스트 1개로 범위를 고정한다
(INTERFACES.md 실행 도구 계약). 폴더 밖 경로나 지정 테스트 외 임의 명령은
다루지 않는다.

read_fixture_code()/propose_code_fix()는 모델이 직접 호출하는 도구다.
apply_fix_and_test()는 의도적으로 모델 도구 목록에 넣지 않는다 - 모델이
스스로 approved=True를 만들어 호출하면 R05(승인 없는 변경 금지)가 그대로
무력화되기 때문이다. 이 함수는 chat_assistant.py가 사용자의 실제 "승인"/
"거절" 메시지를 결정론적으로 감지했을 때만 직접 호출한다."""
import os
import subprocess
import sys
import uuid

_FIXTURE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "coding_fixture")
_TARGET_FILE = os.path.join(_FIXTURE_DIR, "split_cost.py")
_TEST_FILE = os.path.join(_FIXTURE_DIR, "test_split_cost.py")
_TEST_TIMEOUT_SECONDS = 10

# approval_id -> 아직 결정 안 된 제안 코드. apply_fix_and_test가 승인/거절
# 처리 후 여기서 빼서 _resolved_approvals로 옮긴다.
_pending_proposals: dict[str, str] = {}
# approval_id -> 이미 처리된 결과. 같은 id로 다시 apply_fix_and_test가 오면
# (중복 승인 클릭 등) 재실행하지 않고 이 캐시를 그대로 돌려준다(멱등성).
_resolved_approvals: dict[str, dict] = {}


def _run_fixture_test() -> dict:
    """coding_fixture/test_split_cost.py만 별도 프로세스로 실행한다(임의
    명령 실행 불가). 10초를 넘기면 강제 종료하고 FAIL로 보고한다 - 이
    프로젝트에서 유일하게 subprocess 자체 timeout으로 하드 타임아웃을 거는
    도구(agent_loop.py의 trace 시간 기록은 사후 관찰일 뿐 강제 중단이
    아님, agent_loop.py 참고)."""
    try:
        completed = subprocess.run(
            [sys.executable, "-m", "pytest", os.path.basename(_TEST_FILE), "-q"],
            cwd=_FIXTURE_DIR, capture_output=True, text=True, timeout=_TEST_TIMEOUT_SECONDS,
        )
        output = (completed.stdout + completed.stderr)[-2000:]
        return {"test_result": "PASS" if completed.returncode == 0 else "FAIL", "test_output": output}
    except subprocess.TimeoutExpired:
        return {"test_result": "FAIL", "test_output": f"테스트 실행이 {_TEST_TIMEOUT_SECONDS}초를 넘어 중단됨(TEST_TIMEOUT)"}


def read_fixture_code() -> dict:
    """coding_fixture/split_cost.py의 현재 코드와, 지정 테스트를 지금 돌려본
    결과를 함께 돌려준다 - 모델이 수정 전 상태(무엇이 왜 실패하는지)를
    정확히 보고 판단하게 한다."""
    if not os.path.isfile(_TARGET_FILE):
        return {"ok": False, "error": {"code": "FIXTURE_NOT_FOUND", "message": "coding_fixture/split_cost.py가 없습니다."}}
    with open(_TARGET_FILE, "r", encoding="utf-8") as handle:
        code = handle.read()
    return {"ok": True, "code": code, **_run_fixture_test()}


def propose_code_fix(new_code: str) -> dict:
    """new_code를 파일에 바로 쓰지 않고 제안만 만든다 - 승인 전에는 파일이
    바뀌지 않는다(R05). approval_id는 이 제안 1건에만 쓰이는 1회용 이름표다.
    구문 오류가 있는 코드는 애초에 제안으로 받지 않는다(적용 단계까지
    미루지 않고 여기서 바로 거부)."""
    if not new_code or not new_code.strip():
        return {"ok": False, "error": {"code": "EMPTY_CHANGE", "message": "빈 변경은 제안할 수 없습니다."}}
    try:
        compile(new_code, _TARGET_FILE, "exec")
    except SyntaxError as error:
        return {"ok": False, "error": {"code": "SYNTAX_ERROR", "message": str(error)}}

    with open(_TARGET_FILE, "r", encoding="utf-8") as handle:
        current_code = handle.read()
    approval_id = uuid.uuid4().hex[:12]
    _pending_proposals[approval_id] = new_code
    return {"ok": True, "approval_id": approval_id, "current_code": current_code, "proposed_code": new_code}


def apply_fix_and_test(approval_id: str, approved: bool) -> dict:
    """approved=True면 제안된 코드를 실제로 적용하고 지정 테스트를 실행한다.
    approved=False면 파일을 건드리지 않고 applied=False만 돌려준다(A04
    "거절하면 파일 불변"의 실행 지점). 같은 approval_id로 다시 호출되면
    (중복 승인) 파일을 또 쓰거나 테스트를 다시 돌리지 않고 직전 결과를
    그대로 돌려준다."""
    if approval_id in _resolved_approvals:
        return _resolved_approvals[approval_id]
    if approval_id not in _pending_proposals:
        return {"ok": False, "error": {
            "code": "UNKNOWN_APPROVAL_ID", "message": f"알 수 없거나 이미 처리된 제안입니다: {approval_id}",
        }}

    new_code = _pending_proposals.pop(approval_id)
    if approved:
        with open(_TARGET_FILE, "w", encoding="utf-8") as handle:
            handle.write(new_code)
        result = {"ok": True, "applied": True, **_run_fixture_test()}
    else:
        result = {"ok": True, "applied": False}
    _resolved_approvals[approval_id] = result
    return result


def _selftest_read_fixture_code_reports_current_failure():
    result = read_fixture_code()
    assert result["ok"] is True, result
    assert "def total_cost" in result["code"], result
    assert result["test_result"] == "FAIL", "결함이 이미 고쳐져 있음 - fixture 상태 확인 필요"
    print("read_fixture_code 통과(현재 결함/테스트 상태 보고)")


def _selftest_propose_code_fix_does_not_write_file():
    original = open(_TARGET_FILE, "r", encoding="utf-8").read()
    try:
        result = propose_code_fix(
            "def total_cost(participant_count, per_person_cost, venue_fee):\n"
            "    return participant_count * per_person_cost + venue_fee\n"
        )
        assert result["ok"] is True, result
        assert "approval_id" in result, result
        after = open(_TARGET_FILE, "r", encoding="utf-8").read()
        assert after == original, "제안 단계에서 파일이 이미 바뀜(승인 없이 적용됨)"
    finally:
        _pending_proposals.clear()
        _resolved_approvals.clear()
    print("propose_code_fix 통과(제안만 하고 파일은 그대로)")


def _selftest_propose_code_fix_rejects_empty_change():
    result = propose_code_fix("   ")
    assert result == {"ok": False, "error": {"code": "EMPTY_CHANGE", "message": "빈 변경은 제안할 수 없습니다."}}, result
    print("propose_code_fix 통과(빈 변경 거부)")


def _selftest_propose_code_fix_rejects_syntax_error():
    result = propose_code_fix("def total_cost(:\n  broken")
    assert result["ok"] is False, result
    assert result["error"]["code"] == "SYNTAX_ERROR", result
    print("propose_code_fix 통과(구문 오류 거부)")


def _selftest_apply_fix_and_test_rejection_leaves_file_unchanged():
    """A04: 거절하면 파일 내용이 변경되지 않아야 한다."""
    original = open(_TARGET_FILE, "r", encoding="utf-8").read()
    try:
        proposal = propose_code_fix(
            "def total_cost(participant_count, per_person_cost, venue_fee):\n"
            "    return participant_count * per_person_cost + venue_fee\n"
        )
        result = apply_fix_and_test(proposal["approval_id"], approved=False)
        assert result == {"ok": True, "applied": False}, result
        after = open(_TARGET_FILE, "r", encoding="utf-8").read()
        assert after == original, "거절했는데 파일이 바뀜"
    finally:
        with open(_TARGET_FILE, "w", encoding="utf-8") as handle:
            handle.write(original)
        _pending_proposals.clear()
        _resolved_approvals.clear()
    print("apply_fix_and_test 통과(거절 시 파일 불변, A04)")


def _selftest_apply_fix_and_test_approval_applies_and_runs_test():
    """A03: 승인하면 실제로 적용되고, 지정 테스트가 통과해야 한다."""
    original = open(_TARGET_FILE, "r", encoding="utf-8").read()
    try:
        proposal = propose_code_fix(
            "def total_cost(participant_count: int, per_person_cost: int, venue_fee: int) -> int:\n"
            "    return participant_count * per_person_cost + venue_fee\n"
        )
        result = apply_fix_and_test(proposal["approval_id"], approved=True)
        assert result["ok"] is True and result["applied"] is True, result
        assert result["test_result"] == "PASS", result
        after = open(_TARGET_FILE, "r", encoding="utf-8").read()
        assert "participant_count * venue_fee" not in after, "결함이 그대로 남아있음"
    finally:
        with open(_TARGET_FILE, "w", encoding="utf-8") as handle:
            handle.write(original)
        _pending_proposals.clear()
        _resolved_approvals.clear()
    print("apply_fix_and_test 통과(승인 시 적용+테스트 통과, A03)")


def _selftest_apply_fix_and_test_unknown_id_rejected():
    result = apply_fix_and_test("존재하지-않는-id", approved=True)
    assert result["ok"] is False, result
    assert result["error"]["code"] == "UNKNOWN_APPROVAL_ID", result
    print("apply_fix_and_test 통과(모르는 approval_id 거부)")


def _selftest_apply_fix_and_test_duplicate_approval_is_idempotent():
    """같은 approval_id로 두 번 승인 요청이 와도(중복 클릭 등) 파일을 두 번
    쓰거나 테스트를 두 번 돌리지 않고 같은 결과를 돌려줘야 한다."""
    original = open(_TARGET_FILE, "r", encoding="utf-8").read()
    try:
        proposal = propose_code_fix(
            "def total_cost(participant_count: int, per_person_cost: int, venue_fee: int) -> int:\n"
            "    return participant_count * per_person_cost + venue_fee\n"
        )
        first = apply_fix_and_test(proposal["approval_id"], approved=True)
        second = apply_fix_and_test(proposal["approval_id"], approved=True)
        assert first == second, (first, second)
    finally:
        with open(_TARGET_FILE, "w", encoding="utf-8") as handle:
            handle.write(original)
        _pending_proposals.clear()
        _resolved_approvals.clear()
    print("apply_fix_and_test 통과(중복 승인 요청은 멱등적)")


if __name__ == "__main__":
    _selftest_read_fixture_code_reports_current_failure()
    _selftest_propose_code_fix_does_not_write_file()
    _selftest_propose_code_fix_rejects_empty_change()
    _selftest_propose_code_fix_rejects_syntax_error()
    _selftest_apply_fix_and_test_rejection_leaves_file_unchanged()
    _selftest_apply_fix_and_test_approval_applies_and_runs_test()
    _selftest_apply_fix_and_test_unknown_id_rejected()
    _selftest_apply_fix_and_test_duplicate_approval_is_idempotent()
