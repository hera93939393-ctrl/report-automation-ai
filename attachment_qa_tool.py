"""attachment_qa_tool.py — 첨부문서 즉석 질의응답(PRD 16-5 ④, 웍스AI 분석 아이디어 8).

RAG(색인·임베딩) 없이, 지금 "+"로 첨부한 파일들을 부품 4(attachments)로 읽어
그 텍스트만 보고 답한다. 답에는 반드시 "[파일명 · 위치]" 출처를 붙이게 하고,
문서에 없는 내용은 "문서에서 찾지 못했습니다"라고 말하게 한다(지어내기 방지 —
PRD 16-4의 출처 원칙과 같음). 나중에 ⑥단계(search_archive)가 오면 "찾은
문서를 이 도구에 넘긴다"는 형태로 그대로 이어진다.

컨텍스트 한도: 9B 모델의 서버 기본 num_ctx는 작을 수 있어(Ollama 기본값은
모델별 4096 등) 문서가 조용히 잘리면 엉뚱한 답이 나온다. 그래서 ① 요청 시
num_ctx를 명시하고(_NUM_CTX), ② 문서가 예산(_MAX_PROMPT_CHARS)을 넘으면
질문과 단어가 겹치는 부분(파일·위치 단위)만 골라 넣는다(단순 키워드 점수 —
임베딩 없이도 "관련 부분 우선"은 됨). 그래도 넘치면 잘렸다고 답에 표시한다."""
import os
import re

from attachments import read_attachments
from ollama_client import GENERATION_MODEL, get_client

_NUM_CTX = 12288          # 토큰. 한국어 ~1.5자/토큰 가정 시 아래 예산과 대략 맞춤
_MAX_PROMPT_CHARS = 14000
_PART_MAX_CHARS = 6000    # 한 부분(쪽/시트/본문)이 너무 길면 앞부분만

_SYSTEM_PROMPT = (
    "당신은 공공기관 실무자의 문서 검색 보조입니다. 아래 [첨부 문서] 안의 내용만 근거로 "
    "질문에 답합니다. 규칙: ① 답의 근거가 된 부분마다 그 머리말([파일명 · 위치])을 그대로 "
    "인용해 출처로 붙인다 ② 문서에 없는 내용은 추측하지 말고 '문서에서 찾지 못했습니다'라고 "
    "답한다 ③ 숫자는 문서에 적힌 그대로 옮긴다 ④ 간결한 개조식으로 답한다."
)


def _tokens(text: str) -> set:
    return set(t for t in re.findall(r"[가-힣A-Za-z0-9]{2,}", text.lower()))


def _score(question_tokens: set, text: str) -> int:
    return len(question_tokens & _tokens(text))


def select_context(docs: list, question: str, max_chars: int = _MAX_PROMPT_CHARS) -> tuple:
    """첨부 읽기 결과 목록(read_attachments 반환)에서 프롬프트에 넣을 부분을 고른다.
    전부 들어가면 그대로, 넘치면 질문과 겹치는 단어가 많은 부분부터.
    반환: (context_text, truncated: bool, used_labels: [str])"""
    blocks = []
    for d in docs:
        if d.get("error"):
            continue
        name = os.path.basename(d["path"])
        for part in d["parts"]:
            text = part["text"]
            if len(text) > _PART_MAX_CHARS:
                text = text[:_PART_MAX_CHARS] + "\n…(이 부분 뒷부분 생략)"
            blocks.append({"label": f"{name} · {part['label']}", "text": text})
    if not blocks:
        return "", False, []
    total = sum(len(b["text"]) + len(b["label"]) + 6 for b in blocks)
    truncated = False
    if total > max_chars:
        q = _tokens(question)
        ranked = sorted(blocks, key=lambda b: _score(q, b["text"]), reverse=True)
        chosen, used = [], 0
        for b in ranked:
            size = len(b["text"]) + len(b["label"]) + 6
            if used + size > max_chars:
                truncated = True
                continue
            chosen.append(b)
            used += size
        # 문서 순서를 되살린다(모델이 앞뒤 맥락을 덜 헷갈리게)
        order = {id(b): i for i, b in enumerate(blocks)}
        blocks = sorted(chosen, key=lambda b: order[id(b)])
    context = "\n\n".join(f"[{b['label']}]\n{b['text']}" for b in blocks)
    return context, truncated, [b["label"] for b in blocks]


def ask_attachments(source_paths: list, question: str, client=None) -> dict:
    """첨부 파일들을 읽고 질문에 답한다.
    반환: {"ok": True, "answer": str, "used": [label], "truncated": bool,
           "unreadable": [(path, error)]}
        또는 {"ok": False, "error": "no_sources"|"unreadable"|"server_unreachable"|"empty_response", "reason": str}."""
    if not source_paths:
        return {"ok": False, "error": "no_sources", "reason": "먼저 '+'로 첨부 파일을 올려주세요."}
    docs = read_attachments(source_paths)
    unreadable = [(d["path"], d["error"]) for d in docs if d.get("error")]
    context, truncated, used = select_context(docs, question)
    if not context:
        detail = "; ".join(f"{os.path.basename(p)}: {e}" for p, e in unreadable) or "읽을 수 있는 내용이 없습니다"
        return {"ok": False, "error": "unreadable", "reason": f"첨부에서 읽은 내용이 없어요 — {detail}"}

    r = answer_from_context(context, question, client=client)
    if not r["ok"]:
        return r
    return {"ok": True, "answer": r["answer"], "used": used, "truncated": truncated, "unreadable": unreadable}


def answer_from_context(context: str, question: str, client=None, heading: str = "첨부 문서") -> dict:
    """[머리말이 붙은 문서 텍스트] + 질문 → 출처 인용 답변. 첨부 QA와 아카이브
    QA(archive_qa_tool)가 같은 규칙(출처 강제·없으면 없다고·num_ctx 명시)을 쓴다."""
    if client is None:
        client = get_client()
    user_prompt = f"[{heading}]\n{context}\n\n[질문]\n{question}"
    answer = ""
    for _ in range(3):
        try:
            response = client.chat(
                model=GENERATION_MODEL,
                messages=[{"role": "system", "content": _SYSTEM_PROMPT},
                          {"role": "user", "content": user_prompt}],
                think=False,
                options={"num_ctx": _NUM_CTX},
            )
        except Exception:
            return {"ok": False, "error": "server_unreachable",
                    "reason": "서버에 연결할 수 없어서 답하지 못했어요 — 서버가 켜져 있는지 확인해보세요."}
        answer = (response["message"]["content"] or "").strip()
        if answer:
            break
    if not answer:
        return {"ok": False, "error": "empty_response", "reason": "모델 응답이 비어 있었어요. 다시 시도해주세요."}
    return {"ok": True, "answer": answer}


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


def _selftest_select_context_prefers_relevant_parts_when_over_budget():
    docs = [{"path": r"C:\x\실적.xlsx", "error": None, "parts": [
                {"label": "시트 '서류심사'", "text": "연도 | 건수\n2024 | 120\n2025 | 135"},
                {"label": "시트 '현장점검'", "text": "구분 | 회수\n상반기 | 7"}]},
            {"path": r"C:\x\회의.txt", "error": None, "parts": [{"label": "본문", "text": "잡담 " * 50}]},
            {"path": r"C:\x\깨짐.hwp", "error": "파일이 없습니다", "parts": []}]
    ctx, truncated, used = select_context(docs, "2024년 서류심사 건수", max_chars=10_000)
    assert not truncated and len(used) == 3 and ctx.startswith("[실적.xlsx · 시트 '서류심사']"), (truncated, used)
    ctx, truncated, used = select_context(docs, "2024년 서류심사 건수", max_chars=70)
    assert truncated and used == ["실적.xlsx · 시트 '서류심사'"], (truncated, used)
    # 예산이 조금 더 크면 관련도 높은 것부터 채우되, 들어가는 작은 부분은 마저 넣는다
    ctx, truncated, used = select_context(docs, "2024년 서류심사 건수", max_chars=120)
    assert truncated and used == ["실적.xlsx · 시트 '서류심사'", "실적.xlsx · 시트 '현장점검'"], (truncated, used)
    assert select_context([docs[2]], "?") == ("", False, [])
    print("select_context(예산 초과 시 관련 부분 우선) 통과")


def _selftest_ask_attachments_paths():
    import tempfile, shutil
    d = tempfile.mkdtemp(prefix="_qa_")
    try:
        p = os.path.join(d, "회의결과.txt")
        with open(p, "w", encoding="utf-8") as f:
            f.write("2024년 서류심사 건수는 120건이었다.")
        fake = _FakeClient(["- 2024년 서류심사 건수: 120건 [회의결과.txt · 본문]"])
        r = ask_attachments([p], "재작년 서류심사 몇 건?", client=fake)
        assert r["ok"] and "120건" in r["answer"] and r["used"] == ["회의결과.txt · 본문"] and not r["truncated"], r
        sent = fake.calls[0]["messages"][1]["content"]
        assert "[회의결과.txt · 본문]\n2024년 서류심사 건수는 120건이었다." in sent and "[질문]\n재작년" in sent, sent
        assert fake.calls[0]["options"]["num_ctx"] == _NUM_CTX and fake.calls[0]["think"] is False

        assert ask_attachments([], "?", client=fake)["error"] == "no_sources"
        r = ask_attachments([os.path.join(d, "없음.pdf")], "?", client=fake)
        assert r["error"] == "unreadable" and "없음.pdf: 파일이 없습니다" in r["reason"], r
        r = ask_attachments([p], "?", client=_FakeClient([ConnectionError()]))
        assert r["error"] == "server_unreachable", r
        r = ask_attachments([p], "?", client=_FakeClient(["", "", ""]))
        assert r["error"] == "empty_response", r
        print("ask_attachments(정상/첨부없음/못읽음/미접속/빈응답) 통과")
    finally:
        shutil.rmtree(d, ignore_errors=True)


if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8")
    _selftest_select_context_prefers_relevant_parts_when_over_budget()
    _selftest_ask_attachments_paths()
