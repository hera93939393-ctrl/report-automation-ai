"""loop_runner.py — agent_loop(랭그래프 루프)를 채팅 앱에 잇는 접착층(PRD 16-5 ⑤).

세 가지 문제를 여기서 푼다.

1. **스레드 분리.** 루프(LLM 호출 포함)는 작업 스레드에서 돌려 채팅창이 멈추지
   않게 한다. 그런데 pyhwpx/COM 객체는 그걸 만든 메인 스레드에서만 안전하게
   쓸 수 있으므로(COM 아파트), 한글을 만지는 도구 호출은 `MainThreadBridge`로
   메인 스레드에 넘기고 결과를 기다린다. Tk는 스레드 안전하지 않아 UI 갱신도
   같은 다리로 넘긴다(`bridge.post` = 결과를 기다리지 않는 버전).
2. **승인 대기(interrupt).** 그래프가 쓰기 승인에서 멈추면 작업 스레드는 끝나고,
   UI에 승인 카드를 요청한다. 사용자가 누르면 `resume()`이 새 작업 스레드에서
   `Command(resume=…)`로 같은 자리부터 이어간다(agent_loop의 체크포인터).
3. **진행 표시.** `graph.stream(stream_mode="updates")`가 노드마다 내는 갱신을
   받아 "계획 수립 중 → 2/3단계: … → 도구 실행: read_section" 식으로 알린다.
   토큰 스트리밍은 아니지만 "멈춘 건지 처리 중인지"는 알 수 있다.

도구 레지스트리(`build_tools`)는 부품 1~4(doc_sections, attachments)와 polish를
루프용으로 감싼다. 쓰기 도구는 문서를 바꾸지 않고 제안(proposal)만 돌려주고,
승인 뒤 `apply_write`가 doc_sections.write_section으로 반영한다(16-6)."""
import os
import queue
import threading
import traceback

from langgraph.types import Command

import agent_loop
from doc_sections import format_outline, read_outline, read_section, section_at_cursor, write_section


class MainThreadBridge:
    """작업 스레드 → 메인 스레드 호출 다리. 메인 스레드는 Tk `after()`로 `pump()`를
    주기적으로 불러 큐를 비운다(테스트에서는 직접 pump()). 메인 스레드에서
    run()을 부르면 큐를 거치지 않고 바로 실행한다(교착 방지)."""

    def __init__(self):
        self._jobs = queue.Queue()
        self._main_thread = threading.current_thread()

    def run(self, fn, *args, **kwargs):
        if threading.current_thread() is self._main_thread:
            return fn(*args, **kwargs)
        done = threading.Event()
        box = {}

        def job():
            try:
                box["result"] = fn(*args, **kwargs)
            except BaseException as e:  # noqa: BLE001 — 작업 스레드로 그대로 전달
                box["error"] = e
            finally:
                done.set()

        self._jobs.put(job)
        done.wait()
        if "error" in box:
            raise box["error"]
        return box.get("result")

    def post(self, fn, *args, **kwargs):
        """결과를 기다리지 않는 메인 스레드 호출(UI 갱신용)."""
        if threading.current_thread() is self._main_thread:
            fn(*args, **kwargs)
            return
        self._jobs.put(lambda: fn(*args, **kwargs))

    def pump(self, max_jobs: int = 50) -> int:
        """메인 스레드에서 대기 중인 작업을 실행한다. 실행한 개수를 돌려준다."""
        n = 0
        while n < max_jobs:
            try:
                job = self._jobs.get_nowait()
            except queue.Empty:
                break
            job()
            n += 1
        return n


def build_tools(report, source_paths: list, bridge: MainThreadBridge, llm_client_factory=None) -> dict:
    """루프용 도구 레지스트리. 한글을 만지는 도구는 bridge.run으로 메인 스레드에서."""

    def _outline():
        return bridge.run(read_outline, report)

    def tool_read_outline():
        outline = _outline()
        return {"outline": format_outline(outline), "sections": [
            {"id": s["id"], "title": s["title"], "kind": s["kind"]} for s in outline["sections"]]}

    def tool_read_section(section_id: str):
        sec = bridge.run(read_section, report, section_id)
        if "error" in sec:
            raise ValueError(sec["error"])
        return {"title": sec["title"], "text": sec["body"], "tables": sec["tables"],
                "citations": [f"현재 문서 · {sec['title']}"]}

    def tool_write_section(section_id: str, text: str, mode: str = "replace_body"):
        sec = bridge.run(read_section, report, section_id)
        if "error" in sec:
            raise ValueError(sec["error"])
        return {"proposal": {"title": sec["title"], "section_id": section_id, "mode": mode,
                             "before": sec["body"] if mode != "insert_after" else "", "after": text}}

    def tool_read_attachment(name: str = ""):
        from attachments import read_attachments
        docs = read_attachments(source_paths)
        if name:
            docs = [d for d in docs if name in os.path.basename(d["path"])] or docs
        texts, citations = [], []
        for d in docs:
            if d.get("error"):
                texts.append(f"[{os.path.basename(d['path'])}] 읽기 실패: {d['error']}")
                continue
            texts.append(d["text"])
            citations += [f"{os.path.basename(d['path'])} · {p['label']}" for p in d["parts"]]
        return {"text": "\n\n".join(texts)[:20000], "citations": citations}

    def tool_polish(text: str):
        from polish_tool import _generate_formal_style
        client = llm_client_factory() if llm_client_factory else None
        polished = _generate_formal_style(text, client=client)
        if polished is None:
            raise ConnectionError("서버에 연결할 수 없습니다")
        return {"text": polished}

    def tool_search_archive(query: str, year: int = None):
        from archive_index import load_folders, reindex, search_archive
        folders = load_folders()
        if not folders:
            raise ValueError("과거 문서 폴더가 등록되지 않았습니다(python archive_index.py --add-folder)")
        reindex(folders)
        results = search_archive(query, k=6, year=year)
        return {"results": [{"file": r["file"], "label": r["label"], "text": r["text"], "year": r["year"]} for r in results],
                "citations": [f"{r['file']} · {r['label']}" for r in results]}

    return {
        "search_archive": {"fn": tool_search_archive, "write": False,
                           "desc": "과거 보고서 아카이브에서 질문과 관련된 부분을 찾는다(파일명·위치·원문). year로 연도를 좁힐 수 있다",
                           "params": {"type": "object", "properties": {"query": {"type": "string"}, "year": {"type": "integer"}}, "required": ["query"]}},
        "read_outline": {"fn": tool_read_outline, "write": False,
                         "desc": "지금 열려 있는 문서의 구간 목록(id, 제목, 단계)을 읽는다",
                         "params": {"type": "object", "properties": {}}},
        "read_section": {"fn": tool_read_section, "write": False,
                         "desc": "구간 id의 제목·본문·표를 읽는다",
                         "params": {"type": "object", "properties": {"section_id": {"type": "string"}}, "required": ["section_id"]}},
        "write_section": {"fn": tool_write_section, "write": True,
                          "desc": "구간 id의 본문을 text로 바꾸는 제안을 만든다(사용자 승인 뒤 반영). mode: replace_body|append|insert_after",
                          "params": {"type": "object", "properties": {"section_id": {"type": "string"}, "text": {"type": "string"}, "mode": {"type": "string"}}, "required": ["section_id", "text"]}},
        "read_attachment": {"fn": tool_read_attachment, "write": False,
                            "desc": "첨부 파일(회의결과·엑셀·PDF) 내용을 읽는다. name으로 파일명 일부를 줄 수 있다",
                            "params": {"type": "object", "properties": {"name": {"type": "string"}}}},
        "polish": {"fn": tool_polish, "write": False,
                   "desc": "문장을 공문서 개조식 문체로 다듬어 돌려준다(문서는 바꾸지 않음)",
                   "params": {"type": "object", "properties": {"text": {"type": "string"}}, "required": ["text"]}},
    }


def build_context_provider(report, source_paths: list, bridge: MainThreadBridge):
    """문맥 수집 노드용: 선택 텍스트, 커서 구간, 목차, 첨부 목록(메인 스레드에서 읽음)."""
    def provider():
        def gather():
            hwp = report.hwp
            selection = hwp.get_selected_text(keep_select=True) if hwp.SelectionMode != 0 else ""
            outline = read_outline(report)
            cur = section_at_cursor(report, outline)
            return {
                "selection": selection or "",
                "cursor_section": {"id": cur["id"], "title": cur["title"]} if cur else None,
                "outline": format_outline(outline),
                "attachments": [os.path.basename(p) for p in source_paths],
            }
        return bridge.run(gather)
    return provider


def build_apply_write(report, bridge: MainThreadBridge):
    def apply(proposal: dict) -> dict:
        def do():
            return write_section(report, proposal["section_id"], proposal["after"], proposal.get("mode", "replace_body"))
        result = bridge.run(do)
        return {"applied": bool(result.get("applied")), "reason": result.get("reason")}
    return apply


_NODE_LABELS = {
    "gather_context": "문서와 첨부를 살펴보는 중…",
    "plan": "계획을 세우는 중…",
    "act": "다음 할 일을 고르는 중…",
    "execute_tool": "도구 실행 중…",
    "check": "결과를 점검하는 중…",
    "replan": "계획을 다시 세우는 중…",
    "answer": "답변 정리 중…",
}


class LoopSession:
    """채팅 한 요청 = 세션 하나. start()/resume()는 작업 스레드를 띄우고 바로
    돌아오며, 결과는 콜백(메인 스레드에서 실행됨)으로 온다:
        on_progress(text)                 진행 상황 한 줄
        on_needs_approval(payload)        쓰기 승인 필요 (payload: title/before/after/unverified_numbers)
        on_done(answer_text)              완료
        on_error(message)                 예외
        on_cancelled()                    사용자 중단으로 끝남 (없으면 on_error로 대신 알림)
    """
    _counter = 0

    def __init__(self, graph, bridge: MainThreadBridge, on_progress, on_needs_approval, on_done, on_error,
                 on_cancelled=None):
        LoopSession._counter += 1
        self.graph = graph
        self.bridge = bridge
        self.config = {"configurable": {"thread_id": f"chat-{LoopSession._counter}"}}
        self.on_progress, self.on_needs_approval, self.on_done, self.on_error = (
            on_progress, on_needs_approval, on_done, on_error)
        self.on_cancelled = on_cancelled
        self.thread = None
        self.cancelled = False
        self._awaiting_approval = False

    def cancel(self):
        """(U2) 중단 요청. LLM·도구 호출을 중간에 끊을 수는 없으므로, 돌고 있는
        단계는 끝까지 간 뒤 다음 노드 경계에서 멈춘다. 승인 카드를 기다리는
        중이면(작업 스레드가 이미 끝난 상태라 경계 검사가 없음) 여기서 바로
        끝났다고 알린다."""
        self.cancelled = True
        if self._awaiting_approval:
            self._awaiting_approval = False
            self._notify_cancelled()

    def _notify_cancelled(self):
        if self.on_cancelled is not None:
            self.bridge.post(self.on_cancelled)
        else:
            self.bridge.post(self.on_error, "사용자가 중단했습니다")

    def start(self, user_message: str):
        self._launch({"user_message": user_message})

    def resume(self, approved: bool):
        if self.cancelled:
            return  # 중단 뒤 늦게 눌린 승인 버튼 방어(버튼 비활성화가 1차 방어)
        self._awaiting_approval = False
        self._launch(Command(resume={"approved": approved}))

    def _launch(self, graph_input):
        self.thread = threading.Thread(target=self._run, args=(graph_input,), daemon=True)
        self.thread.start()

    def _run(self, graph_input):
        try:
            if self.cancelled:  # 시작/재개 전에 이미 중단된 경우
                self._notify_cancelled()
                return
            state_snapshot = None
            for chunk in self.graph.stream(graph_input, self.config, stream_mode="updates"):
                if self.cancelled:
                    # 노드 경계 중단 — 체크포인터에 상태가 남지만 세션을 다시
                    # 쓰지 않으므로(요청당 새 thread_id) 정리는 불필요.
                    self._notify_cancelled()
                    return
                if "__interrupt__" in chunk:
                    payload = chunk["__interrupt__"][0].value
                    self._awaiting_approval = True
                    self.bridge.post(self.on_needs_approval, payload)
                    return
                for node, update in chunk.items():
                    label = _NODE_LABELS.get(node, node)
                    if node == "act" and isinstance(update, dict):
                        state_snapshot = update
                        hist = update.get("history") or []
                        if hist:
                            label = f"다음 할 일: {hist[-1].get('step', '')}"
                    elif node == "plan" and isinstance(update, dict):
                        steps = update.get("plan") or []
                        label = "계획: " + " → ".join(steps) if steps else label
                    elif node == "execute_tool" and isinstance(update, dict):
                        hist = update.get("history") or []
                        if hist:
                            tool = hist[-1].get("choice", {}).get("tool", "")
                            label = f"도구 실행: {tool}" + (" (실패)" if hist[-1].get("error") else "")
                    self.bridge.post(self.on_progress, label)
            if self.cancelled:  # 마지막 노드 처리 중 눌린 중단
                self._notify_cancelled()
                return
            final = self.graph.get_state(self.config)
            if final.next:
                # 멈췄는데 interrupt 청크를 못 받은 경우(방어) — 상태에서 직접 꺼낸다
                for task in final.tasks:
                    if task.interrupts:
                        self._awaiting_approval = True
                        self.bridge.post(self.on_needs_approval, task.interrupts[0].value)
                        return
            self.bridge.post(self.on_done, final.values.get("answer", ""))
        except Exception as e:  # noqa: BLE001
            self.bridge.post(self.on_error, f"{e}\n{traceback.format_exc(limit=2)}")


# ---------------------------------------------------------------- self-tests (서버·한글 불필요)

def _selftest_bridge_runs_on_main_thread_and_propagates_errors():
    bridge = MainThreadBridge()
    seen = {}

    def work():
        seen["thread"] = threading.current_thread()
        return 42

    results = {}

    def worker():
        results["value"] = bridge.run(work)
        try:
            bridge.run(lambda: 1 / 0)
        except ZeroDivisionError:
            results["error_propagated"] = True
        bridge.post(lambda: results.setdefault("posted", True))

    t = threading.Thread(target=worker)
    t.start()
    while t.is_alive() or bridge.pump():
        bridge.pump()
        t.join(timeout=0.01)
    bridge.pump()
    assert results == {"value": 42, "error_propagated": True, "posted": True}, results
    assert seen["thread"] is threading.main_thread()
    assert bridge.run(lambda: "direct") == "direct"  # 메인 스레드에서는 즉시 실행
    print("MainThreadBridge(메인 스레드 실행·예외 전달·post) 통과")


def _selftest_session_streams_progress_interrupts_and_resumes():
    doc = {"s2": "원문"}
    tools, apply = agent_loop._fake_tools(doc)
    planner = agent_loop.FakePlanner(
        ["구간을 읽는다", "구간을 고친다"],
        [{"tool": "read_section", "args": {"section_id": "s2"}},
         {"tool": "write_section", "args": {"section_id": "s2", "text": "수정문"}}])
    bridge = MainThreadBridge()
    graph = agent_loop.build_graph(planner, tools, lambda: {}, apply_write=apply)
    events = []
    session = LoopSession(graph, bridge,
                          on_progress=lambda t: events.append(("progress", t)),
                          on_needs_approval=lambda p: events.append(("approve", p)),
                          on_done=lambda a: events.append(("done", a)),
                          on_error=lambda m: events.append(("error", m)))

    def wait():
        while session.thread.is_alive() or bridge.pump():
            bridge.pump()
            session.thread.join(timeout=0.01)
        bridge.pump()

    session.start("고쳐줘")
    wait()
    kinds = [k for k, _ in events]
    assert kinds[-1] == "approve" and "progress" in kinds, kinds
    assert any("계획: 구간을 읽는다 → 구간을 고친다" == t for k, t in events if k == "progress"), events
    assert any(t == "도구 실행: read_section" for k, t in events if k == "progress"), events
    assert events[-1][1]["after"] == "수정문" and doc["s2"] == "원문"
    session.resume(approved=True)
    wait()
    assert events[-1][0] == "done" and "1건을 문서에 반영했어요" in events[-1][1], events[-1]
    assert doc["s2"] == "수정문"
    print("LoopSession(진행 알림 → 승인 대기 → 재개 → 완료) 통과")


def _selftest_session_cancel():
    """(U2) 중단 경로 3가지: ① 시작 전 중단 → 아무것도 실행 안 하고 cancelled
    알림 ② 승인 카드 대기 중 중단 → 즉시 cancelled 알림 ③ 중단 뒤 늦게 눌린
    resume은 무시(문서가 바뀌지 않음)."""
    def make_session(events):
        doc = {"s2": "원문"}
        tools, apply = agent_loop._fake_tools(doc)
        planner = agent_loop.FakePlanner(
            ["구간을 읽는다", "구간을 고친다"],
            [{"tool": "read_section", "args": {"section_id": "s2"}},
             {"tool": "write_section", "args": {"section_id": "s2", "text": "수정문"}}])
        bridge = MainThreadBridge()
        graph = agent_loop.build_graph(planner, tools, lambda: {}, apply_write=apply)
        session = LoopSession(graph, bridge,
                              on_progress=lambda t: events.append(("progress", t)),
                              on_needs_approval=lambda p: events.append(("approve", p)),
                              on_done=lambda a: events.append(("done", a)),
                              on_error=lambda m: events.append(("error", m)),
                              on_cancelled=lambda: events.append(("cancelled", None)))
        return session, bridge, doc

    def wait(session, bridge):
        while session.thread.is_alive() or bridge.pump():
            bridge.pump()
            session.thread.join(timeout=0.01)
        bridge.pump()

    # ① 시작 전 중단 — 노드가 하나도 돌지 않고 cancelled만 온다
    events = []
    session, bridge, doc = make_session(events)
    session.cancel()
    session.start("고쳐줘")
    wait(session, bridge)
    assert [k for k, _ in events] == ["cancelled"] and doc["s2"] == "원문", events

    # ② 승인 대기 중 중단 → ③ 그 뒤의 resume은 무시
    events = []
    session, bridge, doc = make_session(events)
    session.start("고쳐줘")
    wait(session, bridge)
    assert events[-1][0] == "approve", events
    session.cancel()  # 메인 스레드라 post가 즉시 실행됨
    assert events[-1][0] == "cancelled", events
    session.resume(approved=True)  # 중단 뒤 승인 — 무시돼야 함
    if session.thread is not None and session.thread.is_alive():
        wait(session, bridge)
    bridge.pump()
    assert doc["s2"] == "원문" and events[-1][0] == "cancelled", (doc, events)
    print("LoopSession.cancel(시작 전·승인 대기 중·중단 뒤 resume 무시) 통과")


if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8")
    _selftest_bridge_runs_on_main_thread_and_propagates_errors()
    _selftest_session_streams_progress_interrupts_and_resumes()
    _selftest_session_cancel()
