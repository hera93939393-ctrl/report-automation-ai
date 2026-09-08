"""agent_loop.py — 모델(Ollama/vLLM)의 도구 요청을 실제 실행과 연결하고 결과를
다시 모델에 돌려주는 반복 루프. chat_assistant.py의 UI/COM 코드와 완전히 분리된
순수 로직 — 도구 dispatch 테이블(이름→함수)만 넘겨주면 어떤 화면에서도 재사용된다.

(2026-09-08, harness-design-kit 재설계) 기존 route_intent()는 모델이 "6개 중
뭘 실행할지 이름 하나"만 고르고, 그 결과를 다시 모델에 보여주지 않는 1회성
호출이었다 - 과제(INTERFACES.md)가 요구하는 "요청→도구 해석→검사·실행→결과
반환→다음 판단" 반복 구조가 아니었다. 이 모듈이 그 반복 루프를 실제로 구현한다.
"""
import json
import time

import ollama
import requests

from file_answer import DEFAULT_MODEL, VLLM_BASE_URL, VLLM_MODEL

MAX_MODEL_CALLS = 8
TOOL_TIMEOUT_SECONDS = 30
TOTAL_TIMEOUT_SECONDS = 120


def _normalize_tool_calls(raw_tool_calls, arguments_are_json_string: bool) -> list[dict]:
    """provider가 돌려준 tool_calls를 {"id", "function": {"name", "arguments"(dict)}}
    형태로 통일한다. Ollama는 arguments를 이미 dict로 주지만, OpenAI 호환(vLLM)은
    JSON 문자열로 준다 - 이 차이를 호출부가 몰라도 되게 여기서 흡수한다."""
    normalized = []
    for tc in raw_tool_calls:
        arguments = tc["function"]["arguments"]
        if arguments_are_json_string:
            arguments = json.loads(arguments) if arguments else {}
        else:
            arguments = dict(arguments) if arguments else {}
        call_id = tc.get("id")
        normalized.append({
            "id": call_id if call_id else None,
            "function": {"name": tc["function"]["name"], "arguments": arguments},
        })
    return normalized


def call_model(messages: list[dict], tools: list[dict], model: str = DEFAULT_MODEL) -> dict | None:
    """Ollama를 먼저 호출하고, 예외가 나거나 도구호출도 답변도 없는 빈 응답이면
    OpenAI 호환 vLLM으로 한 번 폴백한다(file_answer.answer_file_question과 같은
    정책). 반환은 {"content": str, "tool_calls": [...]} 로 통일된 plain dict -
    두 provider 모두 실패하면 None(호출자가 failed로 처리).

    한계(2026-09-08 기준 미검증): vLLM 쪽 도구 호출 응답이 OpenAI 표준 형식
    (tool_calls[].function.arguments가 JSON 문자열)을 따른다고 가정했다 - 실제
    Colab vLLM 엔드포인트로 도구 호출을 검증한 적은 없다(D08은 vLLM을 문맥 비교
    용도로만 실측했음). 이 가정이 틀리면 vLLM 경로의 도구 호출은 실패로
    처리된다(아래 except가 잡아 None 또는 다음 폴백으로 넘어감)."""
    try:
        response = ollama.chat(model=model, messages=messages, tools=tools)
        msg = response.get("message", {})
        content = (msg.get("content") or "").strip()
        tool_calls = _normalize_tool_calls(msg.get("tool_calls") or [], arguments_are_json_string=False)
        if content or tool_calls:
            return {"content": content, "tool_calls": tool_calls}
    except Exception:
        pass

    try:
        response = requests.post(
            f"{VLLM_BASE_URL}/chat/completions",
            json={
                "model": VLLM_MODEL, "messages": messages, "tools": tools, "stream": False,
                "chat_template_kwargs": {"enable_thinking": False},
            },
            headers={"ngrok-skip-browser-warning": "1"},
            timeout=120,
        )
        response.raise_for_status()
        choice = response.json()["choices"][0]["message"]
        content = (choice.get("content") or "").strip()
        tool_calls = _normalize_tool_calls(choice.get("tool_calls") or [], arguments_are_json_string=True)
        if content or tool_calls:
            return {"content": content, "tool_calls": tool_calls}
        return None
    except Exception:
        return None


def _missing_required_fields(arguments: dict, schema: dict) -> list[str]:
    required = schema.get("parameters", {}).get("required", [])
    return [field for field in required if field not in arguments]


def run_agent_loop(
    user_message: str,
    tools_schema: list[dict],
    dispatch: dict,
    model: str = DEFAULT_MODEL,
    max_model_calls: int = MAX_MODEL_CALLS,
    total_timeout_seconds: float = TOTAL_TIMEOUT_SECONDS,
    on_step=None,
    call_model_fn=call_model,
) -> dict:
    """사용자 메시지 하나를 받아 "모델 호출 → 도구 요청 해석 → 검사·실행 → 결과를
    모델에 반환 → 다음 판단"을 모델이 도구 호출 없는 답을 내거나 한도에 도달할
    때까지 반복한다.

    dispatch[도구이름](**모델이_준_인자) 은 항상 {"ok": True, ...} 또는
    {"ok": False, "error": {"code":.., "message":..}} 형태의 dict를 반환해야
    한다(INTERFACES.md의 도구 계약) - 이 계약을 지키지 않고 예외를 던지면
    TOOL_EXECUTION_ERROR로 잡아 모델에게 실패로 알린다(성공으로 위장하지 않음).

    반환: {"final_text": str|None, "status": "completed"|"failed",
           "reason": str|None, "trace": [{"call_id","tool","arguments","result","elapsed_seconds"}...]}
    trace는 실행 순서 그대로 쌓이므로, 채팅창 로그나 TROUBLESHOOTING.md가 요구하는
    "단계 이름 → 결과 요약" 실행 기록으로 그대로 쓸 수 있다.

    (2026-09-08) 순수 함수 도구(HWP COM 호출 등)는 실행 도중 강제로 중단시킬
    안전한 방법이 없어(스레드/시그널로 끊으면 COM 핸들이 깨진 상태로 남을 위험),
    tool 실행 자체에 타임아웃을 강제하지는 않는다 - 대신 실제 걸린 시간을
    trace에 기록해 한도 초과가 있었는지 사후에 알 수 있게 한다. 진짜 강제
    타임아웃이 필요한 도구(예: 코딩 시나리오의 pytest 실행)는 그 도구 함수
    자신이 subprocess timeout으로 구현한다(code_fix_tool.py 참고).

    call_model_fn은 기본값이 실제 call_model()이다 - self-test가 이 모듈을
    `python agent_loop.py`로 직접 실행할 때(__name__ == "__main__") 자기
    자신을 `import agent_loop`로 다시 불러 monkeypatch하면 별도의 모듈
    사본이 생겨 패치가 실제로 실행되는 코드에 반영되지 않는다(직접 재현:
    패치가 씹혀 진짜 네트워크 호출로 새면서 vLLM 요청이 120초 타임아웃까지
    멈춰있었음) - 그래서 몽키패치 대신 이 매개변수로 가짜 함수를 주입한다.
    """
    schema_by_name = {t["function"]["name"]: t["function"] for t in tools_schema}
    messages = [{"role": "user", "content": user_message}]
    trace = []
    start = time.monotonic()

    for _ in range(max_model_calls):
        if time.monotonic() - start > total_timeout_seconds:
            return {"final_text": None, "status": "failed", "reason": "TOTAL_TIMEOUT", "trace": trace}

        if on_step:
            on_step("모델 응답 기다리는 중...")
        reply = call_model_fn(messages, tools_schema, model)
        if reply is None:
            return {"final_text": None, "status": "failed", "reason": "PROVIDER_UNAVAILABLE", "trace": trace}

        tool_calls = reply["tool_calls"]
        if not tool_calls:
            return {"final_text": reply["content"], "status": "completed", "reason": None, "trace": trace}

        messages.append({
            "role": "assistant", "content": reply["content"],
            "tool_calls": [{"function": tc["function"]} for tc in tool_calls],
        })
        for index, tc in enumerate(tool_calls):
            name = tc["function"]["name"]
            arguments = tc["function"]["arguments"]
            call_id = tc["id"] or f"call_{len(trace)}_{index}"

            elapsed_seconds = 0.0
            schema = schema_by_name.get(name)
            missing = _missing_required_fields(arguments, schema) if schema else []
            if schema is None:
                result = {"ok": False, "error": {"code": "UNKNOWN_TOOL", "message": f"등록되지 않은 도구: {name}"}}
            elif missing:
                result = {"ok": False, "error": {
                    "code": "INVALID_ARGUMENTS",
                    "message": f"필수 인자 누락: {', '.join(missing)}",
                }}
            else:
                if on_step:
                    on_step(f"{name} 실행 중...")
                tool_start = time.monotonic()
                try:
                    result = dispatch[name](**arguments)
                except Exception as error:
                    result = {"ok": False, "error": {
                        "code": "TOOL_EXECUTION_ERROR", "message": str(error),
                    }}
                elapsed_seconds = round(time.monotonic() - tool_start, 3)
            trace.append({
                "call_id": call_id, "tool": name, "arguments": arguments,
                "result": result, "elapsed_seconds": elapsed_seconds,
            })
            messages.append({"role": "tool", "content": json.dumps(result, ensure_ascii=False)})

    return {"final_text": None, "status": "failed", "reason": "MAX_MODEL_CALLS_REACHED", "trace": trace}


def _selftest_run_agent_loop_completes_without_tool_call():
    """도구 호출 없이 바로 텍스트로 답하면, 도구를 하나도 안 부르고 completed로
    끝나야 한다."""
    def fake_call_model(messages, tools, model):
        return {"content": "안녕하세요", "tool_calls": []}

    result = run_agent_loop("안녕", tools_schema=[], dispatch={}, call_model_fn=fake_call_model)
    assert result["status"] == "completed", result
    assert result["final_text"] == "안녕하세요", result
    assert result["trace"] == [], result
    print("run_agent_loop 통과(도구 호출 없이 바로 완료)")


def _selftest_run_agent_loop_executes_tool_and_feeds_result_back():
    """도구 호출 1번 → 결과를 모델에 돌려줌 → 모델이 최종 텍스트로 답하는
    2턴짜리 정상 흐름."""
    calls = []

    def fake_call_model(messages, tools, model):
        calls.append(len(messages))
        if len(calls) == 1:
            return {"content": "", "tool_calls": [
                {"id": "call_1", "function": {"name": "read_file", "arguments": {"path": "a.txt"}}},
            ]}
        # 두 번째 호출 - 도구 결과가 실제로 대화에 들어갔는지 확인
        assert messages[-1]["role"] == "tool"
        assert "오후 3시" in messages[-1]["content"], messages[-1]["content"]
        return {"content": "회의는 오후 3시입니다.", "tool_calls": []}

    def fake_read_file(path):
        assert path == "a.txt", path
        return {"ok": True, "text": "회의 시작: 오후 3시"}

    result = run_agent_loop(
        "a.txt 읽고 회의시간 알려줘",
        tools_schema=[{"type": "function", "function": {
            "name": "read_file", "parameters": {"type": "object", "properties": {"path": {}}, "required": ["path"]},
        }}],
        dispatch={"read_file": fake_read_file},
        call_model_fn=fake_call_model,
    )
    assert result["status"] == "completed", result
    assert result["final_text"] == "회의는 오후 3시입니다.", result
    assert len(result["trace"]) == 1, result["trace"]
    assert result["trace"][0]["tool"] == "read_file", result["trace"]
    assert result["trace"][0]["result"] == {"ok": True, "text": "회의 시작: 오후 3시"}, result["trace"]
    print("run_agent_loop 통과(도구 실행 결과가 모델에 전달됨)")


def _selftest_run_agent_loop_executes_multiple_tool_calls_in_order():
    """한 응답에 도구 호출이 여러 개 오면 순서대로 전부 실행해야 한다."""
    executed = []

    def fake_call_model(messages, tools, model):
        if len(messages) == 1:
            return {"content": "", "tool_calls": [
                {"id": "c1", "function": {"name": "step", "arguments": {"n": 1}}},
                {"id": "c2", "function": {"name": "step", "arguments": {"n": 2}}},
            ]}
        return {"content": "done", "tool_calls": []}

    def fake_step(n):
        executed.append(n)
        return {"ok": True, "n": n}

    result = run_agent_loop(
        "두 단계 실행해줘",
        tools_schema=[{"type": "function", "function": {
            "name": "step", "parameters": {"type": "object", "properties": {"n": {}}, "required": ["n"]},
        }}],
        dispatch={"step": fake_step},
        call_model_fn=fake_call_model,
    )
    assert executed == [1, 2], executed
    assert len(result["trace"]) == 2, result["trace"]
    print("run_agent_loop 통과(여러 도구 호출을 순서대로 실행)")


def _selftest_run_agent_loop_missing_required_argument_skips_execution():
    """필수 인자가 빠지면 도구 함수를 아예 부르지 않고 INVALID_ARGUMENTS를
    모델에 돌려줘야 한다 - 모델의 잘못된 요청을 그대로 실행하지 않는다."""
    called = []

    def fake_call_model(messages, tools, model):
        if len(messages) == 1:
            return {"content": "", "tool_calls": [
                {"id": "c1", "function": {"name": "read_file", "arguments": {}}},
            ]}
        assert "INVALID_ARGUMENTS" in messages[-1]["content"], messages[-1]["content"]
        return {"content": "경로를 다시 알려주세요.", "tool_calls": []}

    def fake_read_file(path):
        called.append(path)
        return {"ok": True}

    result = run_agent_loop(
        "파일 읽어줘",
        tools_schema=[{"type": "function", "function": {
            "name": "read_file", "parameters": {"type": "object", "properties": {"path": {}}, "required": ["path"]},
        }}],
        dispatch={"read_file": fake_read_file},
        call_model_fn=fake_call_model,
    )
    assert called == [], "필수 인자 없이도 도구 함수가 실행됨"
    assert result["trace"][0]["result"]["error"]["code"] == "INVALID_ARGUMENTS", result["trace"]
    print("run_agent_loop 통과(필수 인자 누락 시 도구를 실행하지 않음)")


def _selftest_run_agent_loop_tool_exception_becomes_error_result_not_crash():
    """도구 함수가 예외를 던져도 전체 루프가 죽지 않고, 실패를 모델에게
    정직하게 알려야 한다(성공처럼 넘어가지 않음)."""
    def fake_call_model(messages, tools, model):
        if len(messages) == 1:
            return {"content": "", "tool_calls": [
                {"id": "c1", "function": {"name": "boom", "arguments": {}}},
            ]}
        assert "TOOL_EXECUTION_ERROR" in messages[-1]["content"], messages[-1]["content"]
        return {"content": "도구 실행에 실패했습니다.", "tool_calls": []}

    def fake_boom():
        raise RuntimeError("디스크 오류")

    result = run_agent_loop(
        "실행해줘",
        tools_schema=[{"type": "function", "function": {
            "name": "boom", "parameters": {"type": "object", "properties": {}, "required": []},
        }}],
        dispatch={"boom": fake_boom},
        call_model_fn=fake_call_model,
    )
    assert result["status"] == "completed", result
    assert result["trace"][0]["result"]["error"]["code"] == "TOOL_EXECUTION_ERROR", result["trace"]
    print("run_agent_loop 통과(도구 예외를 오류 결과로 잡아 모델에 전달)")


def _selftest_run_agent_loop_unknown_tool_name_reported_as_error():
    """스키마에 등록 안 된 도구 이름이 오면(모델의 환각 등) UNKNOWN_TOOL로
    거부하고 dispatch를 뒤지지 않는다."""
    def fake_call_model(messages, tools, model):
        if len(messages) == 1:
            return {"content": "", "tool_calls": [
                {"id": "c1", "function": {"name": "never_registered", "arguments": {}}},
            ]}
        return {"content": "실패", "tool_calls": []}

    result = run_agent_loop("아무거나 해줘", tools_schema=[], dispatch={}, call_model_fn=fake_call_model)
    assert result["trace"][0]["result"]["error"]["code"] == "UNKNOWN_TOOL", result["trace"]
    print("run_agent_loop 통과(등록 안 된 도구 이름 거부)")


def _selftest_run_agent_loop_max_model_calls_reached_reports_failed_not_completed():
    """모델이 끝없이 도구만 요청하면(반복 한도), completed가 아니라 failed +
    이유를 반환해야 한다 - 한도 도달을 성공처럼 숨기지 않는다."""
    def fake_call_model(messages, tools, model):
        return {"content": "", "tool_calls": [
            {"id": "c", "function": {"name": "loop_tool", "arguments": {}}},
        ]}

    result = run_agent_loop(
        "계속 반복해줘",
        tools_schema=[{"type": "function", "function": {
            "name": "loop_tool", "parameters": {"type": "object", "properties": {}, "required": []},
        }}],
        dispatch={"loop_tool": lambda: {"ok": True}},
        max_model_calls=3,
        call_model_fn=fake_call_model,
    )
    assert result["status"] == "failed", result
    assert result["reason"] == "MAX_MODEL_CALLS_REACHED", result
    assert len(result["trace"]) == 3, result["trace"]
    print("run_agent_loop 통과(반복 한도 도달 시 failed로 정직하게 보고)")


def _selftest_run_agent_loop_provider_unavailable_reports_failed():
    """call_model이 두 provider 모두 실패해 None을 주면, 예외를 내지 않고
    failed 상태로 정리해야 한다."""
    def fake_call_model(messages, tools, model):
        return None

    result = run_agent_loop("아무거나", tools_schema=[], dispatch={}, call_model_fn=fake_call_model)
    assert result["status"] == "failed", result
    assert result["reason"] == "PROVIDER_UNAVAILABLE", result
    print("run_agent_loop 통과(두 provider 모두 실패 시 failed)")


def _selftest_normalize_tool_calls_parses_json_string_arguments():
    """vLLM(OpenAI 호환) 응답처럼 arguments가 JSON 문자열이면 dict로 파싱해야
    한다."""
    raw = [{"id": "x1", "function": {"name": "f", "arguments": '{"path": "a.txt"}'}}]
    normalized = _normalize_tool_calls(raw, arguments_are_json_string=True)
    assert normalized == [{"id": "x1", "function": {"name": "f", "arguments": {"path": "a.txt"}}}], normalized
    print("_normalize_tool_calls 통과(JSON 문자열 인자 파싱)")


def _selftest_normalize_tool_calls_keeps_dict_arguments_as_is():
    """Ollama 응답처럼 arguments가 이미 dict면 그대로 쓴다(재파싱 안 함)."""
    raw = [{"id": None, "function": {"name": "f", "arguments": {"path": "a.txt"}}}]
    normalized = _normalize_tool_calls(raw, arguments_are_json_string=False)
    assert normalized == [{"id": None, "function": {"name": "f", "arguments": {"path": "a.txt"}}}], normalized
    print("_normalize_tool_calls 통과(dict 인자는 그대로 유지)")


if __name__ == "__main__":
    _selftest_normalize_tool_calls_parses_json_string_arguments()
    _selftest_normalize_tool_calls_keeps_dict_arguments_as_is()
    _selftest_run_agent_loop_completes_without_tool_call()
    _selftest_run_agent_loop_executes_tool_and_feeds_result_back()
    _selftest_run_agent_loop_executes_multiple_tool_calls_in_order()
    _selftest_run_agent_loop_missing_required_argument_skips_execution()
    _selftest_run_agent_loop_tool_exception_becomes_error_result_not_crash()
    _selftest_run_agent_loop_unknown_tool_name_reported_as_error()
    _selftest_run_agent_loop_max_model_calls_reached_reports_failed_not_completed()
    _selftest_run_agent_loop_provider_unavailable_reports_failed()
