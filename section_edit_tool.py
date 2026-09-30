"""section_edit_tool.py — "구간 편집" 도구: 부품 1·2·3(doc_sections)과 서버 LLM을
조합해, 커서가 있는 구간(또는 선택 텍스트)을 자연어 지시대로 고친 제안을 만들고,
사용자가 승인하면 문서에 반영한다.

PRD 16-5의 ③단계 산출물 — 루프 그래프(⑤)가 오기 전에 "선택하고 고쳐줘"를
지금 채팅창에서 써 볼 수 있게 부품을 하나의 도구로 노출한 것. 제안(propose)과
적용(apply)을 일부러 두 함수로 나눴다: 사이에 사람 승인 카드가 들어가야
하고(PRD 16-3 "쓰기면 승인"), 나중에 루프 그래프의 interrupt 노드가 같은
두 함수를 그대로 쓰면 되기 때문이다.

LLM 호출 규약(ollama_client.py): 서버의 GENERATION_MODEL, think=False(사고형
모델이 빈 답으로 끝나는 문제 회피), 서버 미접속은 error="server_unreachable"로
구분 보고(polish_tool과 동일).

변경추적을 기본으로 켜지 않는 이유(2026-10-01 실측): 추적이 켜진 채 구간을
교체하면 삭제된 원문이 "삭제 표시"로 문단에 그대로 남아, 그 뒤 read_outline/
read_section이 원문+수정문이 합쳐진 문단("…단축한다.심사 절차를 표준화하고…")을
읽고, get_text()는 승인/거부 전까지 빈 문자열을 돌려준다(hwp_report.py의
추적 self-test가 기록한 현상과 같음). 즉 다음 편집 요청이 잘못된 본문을 보게
된다. 그래서 되돌리기는 한글의 검토 기능 대신 이 도구가 직접 제공한다:
apply가 돌려준 결과로 revert_section_edit()을 부르면 원문을 다시 써 넣는다
(채팅 카드의 "되돌리기" 버튼). 변경추적이 꼭 필요하면 track_changes=True로
켤 수 있지만, 그 뒤 편집 도구를 이어서 쓰기 전에 한글에서 승인/거부를 끝내야 한다."""
import re

from doc_sections import read_outline, read_section, section_at_cursor, write_section
from ollama_client import GENERATION_MODEL, get_client

_SYSTEM_PROMPT = (
    "당신은 대한민국 공공기관 보고서 편집 보조입니다. 사용자가 지시한 대로 주어진 "
    "구간의 본문을 고쳐 씁니다. 규칙: ① 고친 본문만 출력하고 설명·머리말·코드블록을 "
    "붙이지 않는다 ② 문단 구분은 줄바꿈으로 한다 ③ 원문에 있는 항목기호(1. 가. 1) "
    "가) (1) ① □ ○ - 등)와 개조식 문체를 유지한다 ④ 지시나 원문에 없는 사실·수치를 "
    "새로 지어내지 않는다 ⑤ 제목 줄은 다시 쓰지 않는다(본문만)."
)


def _clean_llm_output(text: str) -> str:
    """코드블록 울타리, "결과:" 같은 머리말을 걷어낸다."""
    t = text.strip()
    t = re.sub(r"^```[a-zA-Z]*\s*", "", t)
    t = re.sub(r"\s*```$", "", t)
    t = re.sub(r"^(결과|수정안|수정 결과|고친 본문)\s*[:：]\s*", "", t)
    return t.strip()


def resolve_target(report, outline: dict = None) -> dict:
    """무엇을 고칠지 정한다. 선택 텍스트가 있으면 그것, 없으면 커서가 있는 구간.
    반환: {"kind": "selection"|"section", ...} 또는 {"kind": None, "reason": ...}."""
    hwp = report.hwp
    if hwp.SelectionMode != 0:
        selected = hwp.get_selected_text(keep_select=True)
        if selected and selected.strip():
            return {"kind": "selection", "before": selected, "title": "선택한 텍스트"}
    if outline is None:
        outline = read_outline(report)
    section = section_at_cursor(report, outline)
    if section is None:
        return {"kind": None, "reason": "커서가 있는 구간을 찾지 못했어요. 고칠 문단을 선택하거나 그 안에 커서를 두세요."}
    if section["kind"] == "table":
        return {"kind": None, "reason": "표 안은 이 도구로 고칠 수 없어요(표 도구를 쓰세요)."}
    content = read_section(report, section["id"], outline)
    if not content["body"].strip():
        return {"kind": None, "reason": f"'{content['title']}' 구간에 고칠 본문이 없어요."}
    if content["tables"]:
        return {"kind": None, "reason": f"'{content['title']}' 구간 안에 표가 있어 통째로 고칠 수 없어요. 문단을 선택해서 요청해 주세요."}
    return {"kind": "section", "section_id": section["id"], "title": content["title"],
            "before": content["body"], "outline": outline}


def _ask_llm(instruction: str, title: str, before: str, client=None, max_attempts: int = 3):
    """서버 LLM에 고쳐 쓰기를 요청. 빈 답은 최대 3회 재시도, 연결 실패는 None."""
    if client is None:
        client = get_client()
    user_prompt = (
        f"[지시]\n{instruction}\n\n[구간 제목]\n{title}\n\n[구간 본문]\n{before}\n\n"
        "위 본문을 지시대로 고쳐 쓴 결과만 출력하세요."
    )
    for _ in range(max_attempts):
        try:
            response = client.chat(
                model=GENERATION_MODEL,
                messages=[{"role": "system", "content": _SYSTEM_PROMPT},
                          {"role": "user", "content": user_prompt}],
                think=False,
            )
        except Exception:
            return None
        after = _clean_llm_output(response["message"]["content"])
        if after:
            return after
    return ""


def propose_section_edit(report, instruction: str, client=None, outline: dict = None) -> dict:
    """제안만 만들고 문서는 건드리지 않는다.
    반환: {"ok": True, "kind", "section_id"?, "title", "before", "after", "outline"?}
        또는 {"ok": False, "error": "server_unreachable"|"empty_response"|"no_target"|"no_change", "reason": str}."""
    target = resolve_target(report, outline)
    if target["kind"] is None:
        return {"ok": False, "error": "no_target", "reason": target["reason"]}
    after = _ask_llm(instruction, target["title"], target["before"], client=client)
    if after is None:
        return {"ok": False, "error": "server_unreachable",
                "reason": "서버에 연결할 수 없어서 제안을 만들지 못했어요 — 서버가 켜져 있는지 확인해보세요."}
    if not after:
        return {"ok": False, "error": "empty_response", "reason": "모델 응답이 비어 있었어요. 다시 시도해주세요."}
    if after.strip() == target["before"].strip():
        return {"ok": False, "error": "no_change", "reason": "모델이 바꿀 게 없다고 판단했어요(원문과 동일)."}
    proposal = {"ok": True, "kind": target["kind"], "title": target["title"],
                "before": target["before"], "after": after}
    if target["kind"] == "section":
        proposal["section_id"] = target["section_id"]
        proposal["outline"] = target["outline"]
    return proposal


def _insert_lines(hwp, text: str) -> None:
    lines = text.replace("\r\n", "\n").split("\n")
    for k, line in enumerate(lines):
        if k > 0:
            hwp.BreakPara()
        hwp.insert_text(line)


def apply_section_edit(report, proposal: dict, track_changes: bool = False) -> dict:
    """승인된 제안을 문서에 반영한다. 저장하지 않는다. 되돌리기는
    revert_section_edit()(모듈 docstring의 변경추적 실측 참고).
    반환: {"applied": True, "kind", "before", "after", "section_id"?, "outline"?}
        또는 {"applied": False, "reason"}."""
    if not proposal.get("ok"):
        return {"applied": False, "reason": "승인할 제안이 없어요"}
    hwp = report.hwp
    if track_changes:
        report.enable_track_changes()
    if proposal["kind"] == "selection":
        # polish_tool과 같은 이유로, 시간이 흐른 뒤에도 "정말 그 원문이 선택된
        # 상태"를 보장하려고 find()로 다시 선택한 뒤 덮어쓴다.
        hwp.Cancel()
        if not hwp.find(proposal["before"], direction="AllDoc"):
            return {"applied": False, "reason": "원래 선택했던 텍스트를 다시 찾지 못했어요(문서가 바뀌었나요?)"}
        _insert_lines(hwp, proposal["after"])
        return {"applied": True, "kind": "selection", "before": proposal["before"], "after": proposal["after"]}
    result = write_section(report, proposal["section_id"], proposal["after"], "replace_body",
                           proposal.get("outline"))
    if not result["applied"]:
        return {"applied": False, "reason": result.get("reason", "알 수 없는 이유")}
    return {"applied": True, "kind": "section", "section_id": proposal["section_id"],
            "before": proposal["before"], "after": proposal["after"], "outline": result["outline"]}


def revert_section_edit(report, applied: dict) -> dict:
    """apply_section_edit()이 돌려준 결과로 원문을 되살린다(채팅 카드 "되돌리기").
    구간 교체는 제목 문단이 움직이지 않으므로 같은 section_id를 apply가 돌려준
    새 outline과 함께 쓰면 된다. 선택 교체는 수정문을 다시 찾아 원문으로 덮는다."""
    if not applied.get("applied"):
        return {"reverted": False, "reason": "되돌릴 적용 결과가 없어요"}
    hwp = report.hwp
    if applied["kind"] == "selection":
        hwp.Cancel()
        if not hwp.find(applied["after"], direction="AllDoc"):
            return {"reverted": False, "reason": "적용했던 텍스트를 다시 찾지 못했어요(그 사이 문서가 바뀌었나요?)"}
        _insert_lines(hwp, applied["before"])
        return {"reverted": True}
    result = write_section(report, applied["section_id"], applied["before"], "replace_body", applied["outline"])
    if not result["applied"]:
        return {"reverted": False, "reason": result.get("reason", "알 수 없는 이유")}
    return {"reverted": True, "outline": result["outline"]}


# ---------------------------------------------------------------- self-tests

class _FakeClient:
    def __init__(self, replies):
        self.replies = list(replies)
        self.calls = []

    def chat(self, **kwargs):
        self.calls.append(kwargs)
        reply = self.replies.pop(0)
        if isinstance(reply, Exception):
            raise reply
        return {"message": {"content": reply}}


def _selftest_clean_llm_output():
    assert _clean_llm_output("```\n가. 목적\n본문\n```") == "가. 목적\n본문"
    assert _clean_llm_output("결과: 심사 절차 표준화") == "심사 절차 표준화"
    assert _clean_llm_output("  본문  ") == "본문"
    print("_clean_llm_output 통과")


def _selftest_ask_llm_paths():
    fake = _FakeClient(["", "```\n고친 본문\n```"])
    assert _ask_llm("줄여줘", "가. 목적", "원문", client=fake) == "고친 본문"
    assert fake.calls[0]["model"] == GENERATION_MODEL and fake.calls[0]["think"] is False
    assert fake.calls[0]["messages"][0]["role"] == "system"
    assert _ask_llm("줄여줘", "가. 목적", "원문", client=_FakeClient([ConnectionError()])) is None
    assert _ask_llm("줄여줘", "가. 목적", "원문", client=_FakeClient(["", "", ""])) == ""
    print("_ask_llm(재시도/미접속/빈응답) 통과")


def _selftest_live_propose_apply_revert():
    """실제 한글 + 가짜 LLM: 커서 구간 제안 → 적용 → 되돌리기 → 재적용,
    선택 텍스트 경로(적용·되돌리기), 서버 미접속 보고, 표 안 커서 거부."""
    import os
    from doc_sections import _make_probe_report
    report, path, _keep = _make_probe_report()
    try:
        report.hwp.set_pos(0, 3, 0)  # "가. 목적" 본문 첫 문단
        fake = _FakeClient(["심사 절차를 표준화하고 기간을 단축한다."])
        proposal = propose_section_edit(report, "두 문장을 한 문장으로 줄여줘", client=fake)
        assert proposal["ok"] and proposal["kind"] == "section", proposal
        assert proposal["title"] == "가. 목적" and "표준화한다." in proposal["before"], proposal
        assert "[지시]\n두 문장을 한 문장으로 줄여줘" in fake.calls[0]["messages"][1]["content"]
        # 제안만으로는 문서가 안 바뀜
        assert "심사 절차를 표준화한다." in report.get_text()

        result = apply_section_edit(report, proposal)
        assert result["applied"], result
        assert report._track_changes_enabled is False  # 기본은 추적 안 켬(모듈 docstring)
        assert "심사 절차를 표준화하고 기간을 단축한다." in report.get_text()
        sec = read_section(report, "s2", result["outline"])
        assert sec["paragraphs"] == ["심사 절차를 표준화하고 기간을 단축한다."], sec
        # 되돌리기 → 원문 두 문단 복원, 다시 적용도 가능
        rv = revert_section_edit(report, result)
        assert rv["reverted"], rv
        sec = read_section(report, "s2", rv["outline"])
        assert sec["paragraphs"] == ["심사 절차를 표준화한다.", "심사 기간을 단축한다."], sec
        proposal["outline"] = rv["outline"]
        result = apply_section_edit(report, proposal)
        assert result["applied"] and "표준화하고" in report.get_text(), result

        # 선택 텍스트 경로
        assert report.hwp.find("관련 규정 제3조", direction="AllDoc")
        proposal = propose_section_edit(report, "조 번호를 4조로", client=_FakeClient(["관련 규정 제4조"]))
        assert proposal["ok"] and proposal["kind"] == "selection", proposal
        result = apply_section_edit(report, proposal)
        assert result["applied"], result
        assert "관련 규정 제4조" in report.get_text()
        rv = revert_section_edit(report, result)
        text = report.get_text()
        assert rv["reverted"] and "관련 규정 제3조" in text and "제4조" not in text, (rv, text)

        # 서버 미접속
        report.hwp.Cancel()
        report.hwp.set_pos(0, 3, 0)
        proposal = propose_section_edit(report, "줄여줘", client=_FakeClient([ConnectionError()]))
        assert proposal["ok"] is False and proposal["error"] == "server_unreachable" and "서버" in proposal["reason"], proposal
        # 표 안에 커서 → no_target
        tbl_para = [s for s in read_outline(report)["sections"] if s["kind"] == "table"][0]["para"]
        report.hwp.set_pos(0, tbl_para, 0)
        report.hwp.MoveRight()  # 표 안으로
        proposal = propose_section_edit(report, "줄여줘", client=_FakeClient(["x"]))
        assert proposal["ok"] is False and proposal["error"] == "no_target", proposal
        print("propose/apply/revert_section_edit(실제 한글 + 가짜 LLM) 통과")
    finally:
        report.close(save=False)
        os.remove(path)


if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8")
    _selftest_clean_llm_output()
    _selftest_ask_llm_paths()
    _selftest_live_propose_apply_revert()
