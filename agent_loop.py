"""agent_loop.py — PRD 16-3 루프 그래프의 뼈대(⑤단계 준비, 2026-10-01 야간).

[문맥 수집] → [계획] → [행동] → [도구 실행] → [점검] → (쓰기면) [승인] → [다음 판단]
                 ↑                                                          │
                 └──────────── 재계획 ◄─────────────────────────────────────┤
                                                                     완료 → [답변]

아직 채팅 앱에 연결하지 않았다. 이 파일의 목적은 서버(9B) 없이도 확인할 수
있는 것, 즉 **랭그래프 interrupt로 멈췄다가 사람 승인 뒤 같은 자리에서 재개되는
메커니즘**(PRD 16-5 ⑤의 사전 확인 항목)과 도구 호출 상한·재계획 규칙을 가짜
LLM/가짜 도구로 검증해 두는 것이다. 실제 LLM 계획(OllamaPlanner)은 서버를 켠
뒤 9B의 계획 JSON 일관성을 따로 재 봐야 한다(모듈 끝의 확인 목록).

설계 결정(16-6 실측 반영): 쓰기 도구는 루프 안에서 문서를 직접 고치지 않고
"제안"만 만든다 → approve 노드가 interrupt로 사람에게 전/후를 보여주고,
승인되면 그때 apply 콜백을 부른다. 변경추적 대신 propose/apply/revert 방식
(section_edit_tool과 동일)이라 승인 전엔 문서가 바뀌지 않는다.

LLM 인터페이스(주입식, 테스트에선 FakePlanner):
    planner.plan(user_message, context) -> list[str]      # 단계 목록(최대 MAX_STEPS)
    planner.choose(step, context, tools, history) -> dict # {"tool": name, "args": {...}}
                                                          # 또는 {"tool": "answer", "text": "..."}
도구 레지스트리(주입식): {name: {"fn": callable(**args)->dict, "desc": str, "write": bool}}
쓰기 도구의 fn은 문서를 바꾸지 않고 {"proposal": {...}} 를 돌려주고, 레지스트리의
"apply": callable(proposal)->dict 가 승인 뒤 실제 반영을 맡는다."""
import json
import re
from typing import Any, Optional, TypedDict

from langgraph.checkpoint.memory import MemorySaver
from langgraph.graph import END, START, StateGraph
from langgraph.types import Command, interrupt

MAX_STEPS = 6
MAX_TOOL_CALLS = 12
MAX_SAME_TOOL_FAILURES = 2


class LoopState(TypedDict, total=False):
    user_message: str
    context: dict
    plan: list
    step_idx: int
    tool_calls_made: int
    history: list          # [{"step", "tool", "args", "result"|"error"}]
    pending_write: Optional[dict]
    citations: list
    failures: dict         # tool name -> 연속 실패 수
    replans: int
    answer: str
    notes: list


def _numbers_in(text: str) -> set:
    return set(re.findall(r"\d[\d,]*(?:\.\d+)?", text or ""))


def check_numbers(after: str, allowed_sources: list) -> list:
    """제안문 속 숫자 중 원문·문맥·도구 결과 어디에도 없는 것을 돌려준다(점검
    노드의 "근거 없는 숫자" 표시용). 완전한 검증(verify_numbers)이 아니라
    '어디서 나온 숫자인지 모르겠다'를 걸러내는 최소 장치다."""
    allowed = set()
    for src in allowed_sources:
        allowed |= _numbers_in(src if isinstance(src, str) else json.dumps(src, ensure_ascii=False))
    return sorted(n for n in _numbers_in(after) if n not in allowed)


def build_graph(planner, tools: dict, context_provider, apply_write=None):
    """planner/tools/context_provider를 주입해 컴파일된 그래프를 돌려준다.
    apply_write(proposal)->dict 는 승인된 쓰기를 실제로 반영하는 콜백."""

    def gather_context(state: LoopState) -> dict:
        return {"context": context_provider() or {}, "step_idx": 0, "tool_calls_made": 0,
                "history": [], "pending_write": None, "citations": [], "failures": {},
                "replans": state.get("replans", 0), "answer": "", "notes": []}

    def plan(state: LoopState) -> dict:
        try:
            steps = planner.plan(state["user_message"], state["context"])
        except Exception as e:
            steps = []
            state.setdefault("notes", []).append(f"계획 실패: {e}")
        steps = [s for s in (steps or []) if isinstance(s, str) and s.strip()][:MAX_STEPS]
        if not steps:
            steps = [state["user_message"]]  # 계획을 못 세우면 요청 자체를 한 단계로
        return {"plan": steps, "step_idx": 0}

    def act(state: LoopState) -> dict:
        step = state["plan"][state["step_idx"]]
        try:
            choice = planner.choose(step, state["context"], tools, state["history"])
        except Exception as e:
            choice = {"tool": "answer", "text": f"도구를 고르지 못했어요: {e}"}
        return {"history": state["history"] + [{"step": step, "choice": choice}]}

    def execute_tool(state: LoopState) -> dict:
        entry = dict(state["history"][-1])
        choice = entry["choice"]
        name, args = choice.get("tool"), choice.get("args") or {}
        failures = dict(state["failures"])
        if name == "answer":
            entry["result"] = {"answer": choice.get("text", "")}
            return {"history": state["history"][:-1] + [entry]}
        spec = tools.get(name)
        if spec is None:
            entry["error"] = f"없는 도구: {name}"
            failures[name] = failures.get(name, 0) + 1
            return {"history": state["history"][:-1] + [entry], "failures": failures,
                    "tool_calls_made": state["tool_calls_made"] + 1}
        try:
            result = spec["fn"](**args)
            entry["result"] = result
            failures[name] = 0
        except Exception as e:
            entry["error"] = str(e)
            failures[name] = failures.get(name, 0) + 1
        return {"history": state["history"][:-1] + [entry], "failures": failures,
                "tool_calls_made": state["tool_calls_made"] + 1}

    def check(state: LoopState) -> dict:
        entry = state["history"][-1]
        notes = list(state["notes"])
        pending = None
        result = entry.get("result") or {}
        if entry["choice"].get("tool") in tools and tools[entry["choice"]["tool"]].get("write"):
            proposal = result.get("proposal")
            if proposal:
                sources = [proposal.get("before", ""), json.dumps(state["context"], ensure_ascii=False)]
                sources += [json.dumps(h.get("result", {}), ensure_ascii=False) for h in state["history"][:-1]]
                unknown = check_numbers(proposal.get("after", ""), sources)
                proposal = dict(proposal)
                proposal["unverified_numbers"] = unknown
                if unknown:
                    notes.append(f"근거를 못 찾은 숫자: {', '.join(unknown)} (확인 필요)")
                pending = proposal
        citations = list(state["citations"])
        for c in result.get("citations", []) if isinstance(result, dict) else []:
            if c not in citations:
                citations.append(c)
        return {"pending_write": pending, "notes": notes, "citations": citations}

    def approve(state: LoopState) -> dict:
        proposal = state["pending_write"]
        decision = interrupt({"type": "approve_write", "title": proposal.get("title", ""),
                              "before": proposal.get("before", ""), "after": proposal.get("after", ""),
                              "unverified_numbers": proposal.get("unverified_numbers", [])})
        history = list(state["history"])
        entry = dict(history[-1])
        if decision and decision.get("approved"):
            applied = apply_write(proposal) if apply_write else {"applied": False, "reason": "apply 콜백 없음"}
            entry["applied"] = applied
        else:
            entry["applied"] = {"applied": False, "reason": "사용자가 적용하지 않음"}
        history[-1] = entry
        return {"history": history, "pending_write": None}

    def next_step(state: LoopState) -> dict:
        return {"step_idx": state["step_idx"] + 1}

    def _stop_reason(state: LoopState) -> Optional[str]:
        """상한/연속 실패로 멈춘 경우의 안내문. 분기 함수(after_next)는 상태를
        바꾸면 안 되므로(랭그래프 관례 — 분기 함수의 변경은 저장되지 않을 수
        있음) 여기 answer 노드에서 같은 조건을 다시 판정해 노트로 남긴다."""
        if state["tool_calls_made"] >= MAX_TOOL_CALLS:
            return f"도구 호출 상한({MAX_TOOL_CALLS}회)에 도달해 여기서 멈췄어요."
        last = state["history"][-1] if state["history"] else {}
        failed_tool = last.get("choice", {}).get("tool")
        if (last.get("error") and state["failures"].get(failed_tool, 0) >= MAX_SAME_TOOL_FAILURES
                and state.get("replans", 0) >= 1):
            return f"'{failed_tool}' 도구가 계속 실패해서 더 진행하지 못했어요. 다른 방법으로 다시 말씀해 주세요."
        return None

    def answer(state: LoopState) -> dict:
        if state.get("answer"):
            return {}
        last = state["history"][-1] if state["history"] else {}
        text = (last.get("result") or {}).get("answer") if isinstance(last.get("result"), dict) else None
        if not text:
            done = [h for h in state["history"] if h.get("applied", {}).get("applied")]
            text = f"{len(done)}건을 문서에 반영했어요." if done else "요청을 처리했어요."
        notes = list(state["notes"])
        stop = _stop_reason(state)
        if stop:
            notes.append(stop)
        if state["citations"]:
            text += "\n\n출처: " + "; ".join(state["citations"])
        if notes:
            text += "\n\n" + "\n".join(notes)
        return {"answer": text, "notes": notes}

    # ---- 분기 규칙
    def after_execute(state: LoopState) -> str:
        entry = state["history"][-1]
        if entry["choice"].get("tool") == "answer":
            return "answer"
        return "check"

    def after_check(state: LoopState) -> str:
        return "approve" if state.get("pending_write") else "next_step"

    def after_next(state: LoopState) -> str:
        # 순수 판정만 한다(상태 변경 금지) — 안내문은 answer 노드의 _stop_reason이 붙인다.
        if state["tool_calls_made"] >= MAX_TOOL_CALLS:
            return "answer"
        last = state["history"][-1] if state["history"] else {}
        failed_tool = last.get("choice", {}).get("tool")
        if last.get("error") and state["failures"].get(failed_tool, 0) >= MAX_SAME_TOOL_FAILURES:
            if state.get("replans", 0) >= 1:
                return "answer"
            return "replan"
        if state["step_idx"] >= len(state["plan"]):
            return "answer"
        return "act"

    def replan(state: LoopState) -> dict:
        note = f"'{state['history'][-1]['choice'].get('tool')}' 도구가 두 번 실패해 계획을 다시 세웁니다."
        try:
            steps = planner.plan(state["user_message"] + "\n(참고: " + note + ")", state["context"])
        except Exception:
            steps = []
        steps = [s for s in (steps or []) if isinstance(s, str) and s.strip()][:MAX_STEPS] or [state["user_message"]]
        return {"plan": steps, "step_idx": 0, "replans": state.get("replans", 0) + 1,
                "failures": {}, "notes": state["notes"] + [note]}

    g = StateGraph(LoopState)
    for name, fn in [("gather_context", gather_context), ("plan", plan), ("act", act),
                     ("execute_tool", execute_tool), ("check", check), ("approve", approve),
                     ("next_step", next_step), ("replan", replan), ("answer", answer)]:
        g.add_node(name, fn)
    g.add_edge(START, "gather_context")
    g.add_edge("gather_context", "plan")
    g.add_edge("plan", "act")
    g.add_edge("act", "execute_tool")
    g.add_conditional_edges("execute_tool", after_execute, {"answer": "answer", "check": "check"})
    g.add_conditional_edges("check", after_check, {"approve": "approve", "next_step": "next_step"})
    g.add_edge("approve", "next_step")
    g.add_conditional_edges("next_step", after_next, {"act": "act", "answer": "answer", "replan": "replan"})
    g.add_edge("replan", "act")
    g.add_edge("answer", END)
    return g.compile(checkpointer=MemorySaver())


def run(graph, user_message: str, thread_id: str) -> dict:
    """한 요청을 시작한다. 승인이 필요하면 {"status": "needs_approval", "payload": {...}}
    로 멈추고, 끝나면 {"status": "done", "answer": str, "state": ...}."""
    config = {"configurable": {"thread_id": thread_id}}
    result = graph.invoke({"user_message": user_message}, config)
    return _status(result)


def resume(graph, thread_id: str, approved: bool) -> dict:
    """승인/거부로 멈춘 자리에서 이어간다. 또 승인이 필요하면 다시 멈춘다."""
    config = {"configurable": {"thread_id": thread_id}}
    result = graph.invoke(Command(resume={"approved": approved}), config)
    return _status(result)


def _status(result: dict) -> dict:
    interrupts = result.get("__interrupt__") or []
    if interrupts:
        return {"status": "needs_approval", "payload": interrupts[0].value, "state": result}
    return {"status": "done", "answer": result.get("answer", ""), "state": result}


# ---------------------------------------------------------------- 실제 LLM 계획자(서버 필요, 아침 확인 대상)

class OllamaPlanner:
    """서버 qwen3.5:9b로 계획·도구 선택을 한다. 계획은 JSON 배열 문자열로 받고
    (format='json'), 도구 선택은 정식 tool-calling으로 받는다. think=False."""

    def __init__(self, client=None, model=None):
        from ollama_client import GENERATION_MODEL, get_client
        self.client = client or get_client()
        self.model = model or GENERATION_MODEL

    def plan(self, user_message: str, context: dict) -> list:
        prompt = (
            "사용자의 요청을 처리하기 위한 단계 목록을 JSON으로 만드세요. 형식: "
            '{"steps": ["...", "..."]} — 단계는 최대 6개, 각 단계는 한 문장, 도구 이름을 '
            "직접 쓰지 말고 무엇을 할지만 적습니다.\n\n[문맥]\n"
            + json.dumps(context, ensure_ascii=False)[:3000] + f"\n\n[요청]\n{user_message}"
        )
        response = self.client.chat(model=self.model, messages=[{"role": "user", "content": prompt}],
                                    think=False, format="json")
        data = json.loads(response["message"]["content"] or "{}")
        return list(data.get("steps", []))

    def choose(self, step: str, context: dict, tools: dict, history: list) -> dict:
        tool_defs = [{"type": "function", "function": {"name": n, "description": t["desc"],
                      "parameters": t.get("params", {"type": "object", "properties": {}})}}
                     for n, t in tools.items()]
        tool_defs.append({"type": "function", "function": {"name": "answer",
                          "description": "더 할 일이 없을 때 사용자에게 최종 답변을 한다",
                          "parameters": {"type": "object", "properties": {"text": {"type": "string"}}, "required": ["text"]}}})
        recent = json.dumps(history[-3:], ensure_ascii=False)[:2000]
        messages = [{"role": "user", "content": f"[현재 단계]\n{step}\n\n[최근 도구 결과]\n{recent}\n\n이 단계를 수행할 도구 하나를 호출하세요."}]
        response = self.client.chat(model=self.model, messages=messages, tools=tool_defs, think=False)
        calls = response.get("message", {}).get("tool_calls") or []
        if not calls:
            return {"tool": "answer", "text": response.get("message", {}).get("content", "") or "처리할 도구를 찾지 못했어요."}
        fn = calls[0]["function"]
        args = fn.get("arguments") or {}
        if fn["name"] == "answer":
            return {"tool": "answer", "text": args.get("text", "")}
        return {"tool": fn["name"], "args": args}


# ---------------------------------------------------------------- self-tests (서버·한글 불필요)

class FakePlanner:
    def __init__(self, steps, choices):
        self.steps, self.choices, self.plan_calls = list(steps), list(choices), 0

    def plan(self, user_message, context):
        self.plan_calls += 1
        return self.steps

    def choose(self, step, context, tools, history):
        return self.choices.pop(0)


def _fake_tools(doc: dict):
    """가짜 문서(dict)를 다루는 도구 3개 + apply 콜백."""
    def read_section(section_id):
        return {"text": doc[section_id], "citations": [f"문서 · {section_id}"]}

    def propose_write(section_id, text):
        return {"proposal": {"title": section_id, "section_id": section_id, "before": doc[section_id], "after": text}}

    def read_attachment(name):
        return {"text": "회의결과: 건수 120건", "citations": [f"{name} · 본문"]}

    def flaky():
        raise RuntimeError("COM 오류")

    tools = {
        "read_section": {"fn": read_section, "desc": "구간 읽기", "write": False},
        "write_section": {"fn": propose_write, "desc": "구간 쓰기(제안)", "write": True},
        "read_attachment": {"fn": read_attachment, "desc": "첨부 읽기", "write": False},
        "flaky": {"fn": flaky, "desc": "항상 실패", "write": False},
    }

    def apply(proposal):
        doc[proposal["section_id"]] = proposal["after"]
        return {"applied": True}
    return tools, apply


def _selftest_interrupt_and_resume_apply():
    doc = {"s2": "심사 절차를 표준화한다.\n심사 기간을 단축한다."}
    tools, apply = _fake_tools(doc)
    planner = FakePlanner(
        ["구간을 읽는다", "첨부를 읽는다", "구간을 고쳐 쓴다"],
        [{"tool": "read_section", "args": {"section_id": "s2"}},
         {"tool": "read_attachment", "args": {"name": "회의결과.txt"}},
         {"tool": "write_section", "args": {"section_id": "s2", "text": "심사 절차를 표준화하고 기간을 단축한다(120건 기준)."}}],
    )
    graph = build_graph(planner, tools, lambda: {"cursor_section": "s2"}, apply_write=apply)
    r = run(graph, "회의결과 반영해서 가. 목적 구간 줄여줘", "t1")
    assert r["status"] == "needs_approval", r
    assert r["payload"]["before"].startswith("심사 절차를") and "120건" in r["payload"]["after"], r["payload"]
    assert r["payload"]["unverified_numbers"] == [], r["payload"]  # 120은 첨부 결과에 있음
    assert doc["s2"].startswith("심사 절차를 표준화한다.")  # 승인 전엔 안 바뀜
    r = resume(graph, "t1", approved=True)
    assert r["status"] == "done", r
    assert doc["s2"] == "심사 절차를 표준화하고 기간을 단축한다(120건 기준)."
    assert "1건을 문서에 반영했어요" in r["answer"] and "회의결과.txt · 본문" in r["answer"] and "문서 · s2" in r["answer"], r["answer"]
    assert r["state"]["tool_calls_made"] == 3
    print("루프: interrupt로 멈춤 → 승인 재개 → 반영·출처 답변 통과")


def _selftest_reject_keeps_document_and_flags_unknown_numbers():
    doc = {"s2": "원문"}
    tools, apply = _fake_tools(doc)
    planner = FakePlanner(["고친다"], [{"tool": "write_section", "args": {"section_id": "s2", "text": "예산 1,850,000원으로 수정"}}])
    graph = build_graph(planner, tools, lambda: {}, apply_write=apply)
    r = run(graph, "예산 넣어줘", "t2")
    assert r["status"] == "needs_approval" and r["payload"]["unverified_numbers"] == ["1,850,000"], r["payload"]
    r = resume(graph, "t2", approved=False)
    assert r["status"] == "done" and doc["s2"] == "원문", (r, doc)
    assert "근거를 못 찾은 숫자: 1,850,000" in r["answer"] and "요청을 처리했어요" in r["answer"], r["answer"]
    print("루프: 거부 시 문서 유지 + 근거 없는 숫자 표시 통과")


def _selftest_answer_tool_ends_loop_and_call_cap():
    global MAX_TOOL_CALLS  # 상한 안내문 확인용으로 잠시 낮춘다(노드 함수들이 전역을 호출 시점에 읽음)
    tools, apply = _fake_tools({"s1": "x"})
    planner = FakePlanner(["답한다"], [{"tool": "answer", "text": "2024년 건수는 120건입니다."}])
    graph = build_graph(planner, tools, lambda: {}, apply_write=apply)
    r = run(graph, "건수?", "t3")
    assert r["status"] == "done" and r["answer"].startswith("2024년 건수는 120건") and r["state"]["tool_calls_made"] == 0, r

    # 상한: 같은 읽기 도구를 13번 계획해도 12번에서 멈춘다
    choices = [{"tool": "read_section", "args": {"section_id": "s1"}}] * 13
    planner = FakePlanner([f"읽기 {i}" for i in range(13)], choices)  # MAX_STEPS로 6개만 남지만 상한 검증엔 충분
    graph = build_graph(planner, tools, lambda: {}, apply_write=apply)
    r = run(graph, "많이 읽어", "t4")
    assert r["status"] == "done" and r["state"]["tool_calls_made"] <= MAX_TOOL_CALLS, r["state"]["tool_calls_made"]
    # 계획이 MAX_STEPS(6)로 잘려 상한(12)엔 못 미친다 — 상한 안내문은 상한을 잠시 낮춰 확인
    planner = FakePlanner([f"읽기 {i}" for i in range(6)], list(choices))
    graph = build_graph(planner, tools, lambda: {}, apply_write=apply)
    saved = MAX_TOOL_CALLS
    try:
        MAX_TOOL_CALLS = 3
        r = run(graph, "많이 읽어", "t4b")
        assert r["state"]["tool_calls_made"] == 3 and "도구 호출 상한(3회)" in r["answer"], (r["state"]["tool_calls_made"], r["answer"])
    finally:
        MAX_TOOL_CALLS = saved
    print("루프: answer 도구로 종료 + 호출 상한 통과")


def _selftest_failures_trigger_replan_then_stop():
    tools, apply = _fake_tools({"s1": "x"})
    planner = FakePlanner(["실패 도구", "실패 도구"],
                          [{"tool": "flaky", "args": {}}, {"tool": "flaky", "args": {}},   # 1차 계획: 2번 실패 → 재계획
                           {"tool": "flaky", "args": {}}, {"tool": "flaky", "args": {}}])  # 2차 계획: 또 2번 실패 → 중단
    graph = build_graph(planner, tools, lambda: {}, apply_write=apply)
    r = run(graph, "뭔가 해줘", "t5")
    assert r["status"] == "done" and planner.plan_calls == 2, (r["status"], planner.plan_calls)
    assert "계획을 다시 세웁니다" in r["answer"] and "다른 방법으로 다시 말씀" in r["answer"], r["answer"]
    print("루프: 같은 도구 2회 실패 → 재계획 1회 → 되묻기 종료 통과")


if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8")
    _selftest_interrupt_and_resume_apply()
    _selftest_reject_keeps_document_and_flags_unknown_numbers()
    _selftest_answer_tool_ends_loop_and_call_cap()
    _selftest_failures_trigger_replan_then_stop()
    print("확인 필요(서버 켠 뒤): OllamaPlanner.plan의 JSON 일관성(요청문 10개), choose의 도구호출 성공률, "
          "그리고 채팅 앱 연결 시 interrupt 대기 중 Tk 이벤트 루프와의 스레드 분리.")
