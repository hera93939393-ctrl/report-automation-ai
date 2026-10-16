"""law_mcp_client.py — 법령·판례 MCP 서버(airealty-lawmcp) 클라이언트 + 외부 전송 관문.

airealty-lawmcp(MIT, Node)는 법제처 국가법령정보 Open API를 감싼 MCP 서버다.
조문 원문·별표·연혁·판례·해석례 조회와 "글 속 법령 인용이 실제로 있는지/현행인지"
판정(citation_audit)을 제공한다. 기존 F5(law_search.py)는 법령 목록 검색만 했다.

**연결 방식(2026-10-16 사용자 결정)**: 작성자 호스팅 서버(https://law.airealty.co.kr/mcp,
Bearer 토큰)를 쓴다. 이 경우 요청이 법제처가 아니라 개인 개발자의 서버를 거치므로,
이 모듈은 나가기 전에 `check_outbound()`로 다음을 **코드가 강제**한다(호스팅일 때):
  - 허용된 조회 도구만(_HOSTED_ALLOWED). 문서 전문을 받는 contract_risk_scan,
    research_run, toolbox(내부 도구가 무엇을 받을지 통제 불가)는 차단.
  - `text`/`claim` 인자 금지 — 문서 본문·문장은 절대 보내지 않는다. 인용 검증은
    노트북에서 `extract_citations()`로 「법령명」+조문만 뽑아 `citations`로 보낸다.
  - 모든 문자열 값 100자 이하(법령명·조문번호·사건번호·키워드 수준만).
  - 개인정보 패턴(privacy_guard)이 하나라도 있으면 차단.
로컬로 띄운 서버(LAW_MCP_URL이 localhost/127.0.0.1)면 관문을 완화한다(모든 도구 허용) —
그 경우 데이터는 이 PC와 법제처 사이만 오간다.

설정: 토큰은 환경변수 LAW_MCP_TOKEN 또는 `law_mcp_token.txt`(gitignore), 주소는
LAW_MCP_URL(기본 호스팅). 호스팅 제한: 초당 2회, 하루 2,000회(tools/call 기준).
실패는 예외(LawMcpError, kind = unreachable|auth|blocked|tool_error|bad_response)로
알린다 — 호출자가 "서버 꺼짐"과 "막힘"을 구분해 안내할 수 있게."""
import json
import os
import re
from urllib.parse import urlparse

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_URL = "https://law.airealty.co.kr/mcp"
TOKEN_FILE = os.path.join(HERE, "law_mcp_token.txt")
_TIMEOUT_SEC = 30
_MAX_STR = 100

# 키워드·식별자만 받는 조회 도구들(입력 스키마 2026-10-16 tools/list로 확인)
_HOSTED_ALLOWED = {
    "statute_find", "article_get", "statute_outline", "annex_get", "statute_history",
    "source_search", "source_get", "precedent_trace", "citation_audit",
}
_FORBIDDEN_KEYS = {"text", "claim"}
_CITATION_KEYS = {"law", "article", "title", "caseNo"}


class LawMcpError(Exception):
    def __init__(self, kind: str, message: str):
        super().__init__(message)
        self.kind = kind


def _load_token() -> str:
    token = os.environ.get("LAW_MCP_TOKEN", "").strip()
    if token:
        return token
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, encoding="utf-8") as f:
            return f.read().strip()
    return ""


def is_hosted(url: str) -> bool:
    host = (urlparse(url).hostname or "").lower()
    return host not in ("localhost", "127.0.0.1", "::1")


# ---------------------------------------------------------------- 외부 전송 관문

def _iter_strings(value, path="args"):
    if isinstance(value, str):
        yield path, value
    elif isinstance(value, dict):
        for k, v in value.items():
            yield from _iter_strings(v, f"{path}.{k}")
    elif isinstance(value, (list, tuple)):
        for i, v in enumerate(value):
            yield from _iter_strings(v, f"{path}[{i}]")


def check_outbound(tool: str, args: dict, hosted: bool) -> None:
    """나가는 요청을 검사하고, 규칙 위반이면 LawMcpError(kind="blocked")를 던진다."""
    if not hosted:
        return
    if tool not in _HOSTED_ALLOWED:
        raise LawMcpError("blocked", f"'{tool}' 도구는 호스팅 서버로 보낼 수 없습니다(문서 내용을 받는 도구)")
    bad_keys = _FORBIDDEN_KEYS & set(args)
    if bad_keys:
        raise LawMcpError("blocked", f"문서 본문 인자({', '.join(sorted(bad_keys))})는 호스팅 서버로 보내지 않습니다")
    for item in args.get("citations", []) or []:
        extra = set(item) - _CITATION_KEYS
        if extra:
            raise LawMcpError("blocked", f"인용 항목에 허용되지 않은 필드: {', '.join(sorted(extra))}")
    from privacy_guard import detect_pii_patterns
    for path, s in _iter_strings(args):
        if len(s) > _MAX_STR:
            raise LawMcpError("blocked", f"{path} 값이 {len(s)}자 — 호스팅 서버로는 {_MAX_STR}자 이하(키워드·법령명 수준)만 보냅니다")
        pii = detect_pii_patterns(s)
        if pii:
            raise LawMcpError("blocked", f"{path} 에 개인정보로 보이는 패턴({', '.join(pii)})이 있어 보내지 않았습니다")


# ---------------------------------------------------------------- 인용 추출(로컬)

_ARTICLE = r"제\s*\d+\s*조(?:\s*의\s*\d+)?"
_BRACKETED = re.compile(r"「(?P<law>[^」]{2,60})」\s*(?P<art>" + _ARTICLE + r")?(?:\s*\((?P<title>[^()]{1,30})\))?")
_PLAIN = re.compile(r"(?<![가-힣])(?P<law>[가-힣]{1,20}(?:법률|법|시행령|시행규칙|조례))\s*(?P<art>" + _ARTICLE + r")(?:\s*\((?P<title>[^()]{1,30})\))?")
_SAME_LAW = re.compile(r"(?:같은\s*법|동법|동\s*법)\s*(?P<art>" + _ARTICLE + r")")
# 낱말 법령명 패턴(_PLAIN)에 걸리지만 실제 법령명이 아닌 지시어들
_REFERENCE_WORDS = {"동법", "본법", "같은법", "이법", "해당법", "상기법", "위법", "관계법", "개별법", "특별법", "일반법"}


def _norm_article(a: str) -> str:
    return re.sub(r"\s+", "", a) if a else ""


def extract_citations(text: str) -> list:
    """문서 텍스트에서 법령 인용을 뽑는다(노트북에서만 실행, 서버로 텍스트를 보내지 않기 위해).
    「법령명」 제N조(제목) / 단일 낱말 법령명(주택법 제5조) / 같은 법·동법 제N조(직전 법령으로 해석).
    반환: [{"law", "article", "title"?}] — 같은 (법령, 조문)은 한 번만."""
    found = []
    bracket_spans = []
    for m in _BRACKETED.finditer(text):
        bracket_spans.append(m.span())
        found.append((m.start(), m.group("law").strip(), _norm_article(m.group("art")), (m.group("title") or "").strip()))
    for m in _PLAIN.finditer(text):
        # 「」 인용 범위와 실제로 겹칠 때만 건너뜀(「학교급식법」 안의 '학교급식법' 등)
        if any(m.start() < e and m.end() > s for s, e in bracket_spans):
            continue
        if m.group("law") in _REFERENCE_WORDS:  # "동법 제3조"는 법령명이 아니라 지시어(_SAME_LAW가 처리)
            continue
        found.append((m.start(), m.group("law"), _norm_article(m.group("art")), (m.group("title") or "").strip()))
    found.sort()
    out, seen, last_law = [], set(), None
    same_positions = [(m.start(), _norm_article(m.group("art"))) for m in _SAME_LAW.finditer(text)]
    events = [(p, "cite", law, art, title) for p, law, art, title in found] + [(p, "same", None, art, "") for p, art in same_positions]
    for pos, kind, law, art, title in sorted(events, key=lambda e: e[0]):
        if kind == "same":
            if not last_law:
                continue
            law = last_law
        else:
            last_law = law
        key = (law, art)
        if key in seen:
            continue
        seen.add(key)
        item = {"law": law[:_MAX_STR]}
        if art:
            item["article"] = art
        if title:
            item["title"] = title
        out.append(item)
    return out


# ---------------------------------------------------------------- 전송(JSON-RPC over Streamable HTTP)

class _HttpTransport:
    def __init__(self, url: str, token: str):
        self.url, self.token, self.session = url, token, None

    def post(self, body: dict):
        import requests
        headers = {"Content-Type": "application/json", "Accept": "application/json, text/event-stream"}
        if self.token:
            headers["Authorization"] = f"Bearer {self.token}"
        if self.session:
            headers["Mcp-Session-Id"] = self.session
        try:
            r = requests.post(self.url, headers=headers, json=body, timeout=_TIMEOUT_SEC)
        except Exception as e:
            raise LawMcpError("unreachable", f"법령 서버에 연결할 수 없습니다: {e}")
        if r.status_code in (401, 403):
            raise LawMcpError("auth", "법령 서버 인증 실패(토큰 확인 필요)")
        if r.status_code == 429:
            raise LawMcpError("tool_error", "법령 서버 호출 한도 초과(잠시 후 다시)")
        if r.status_code >= 300:
            raise LawMcpError("unreachable", f"법령 서버 응답 오류 HTTP {r.status_code}")
        self.session = r.headers.get("Mcp-Session-Id") or self.session
        if not r.text.strip():
            return None
        if "event-stream" in r.headers.get("content-type", ""):
            for line in r.text.splitlines():
                if line.startswith("data:"):
                    return json.loads(line[5:].strip())
            return None
        return r.json()


class LawMcpClient:
    def __init__(self, url: str = None, token: str = None, transport=None):
        self.url = url or os.environ.get("LAW_MCP_URL") or DEFAULT_URL
        self.hosted = is_hosted(self.url)
        token = _load_token() if token is None else token
        if self.hosted and not token and transport is None:
            raise LawMcpError("auth", "법령 서버 토큰이 없습니다(law_mcp_token.txt 또는 LAW_MCP_TOKEN)")
        self.transport = transport or _HttpTransport(self.url, token)
        self._ready = False
        self._next_id = 1

    def _rpc(self, method: str, params: dict, notify: bool = False):
        body = {"jsonrpc": "2.0", "method": method, "params": params}
        if not notify:
            body["id"] = self._next_id
            self._next_id += 1
        resp = self.transport.post(body)
        if notify:
            return None
        if not isinstance(resp, dict):
            raise LawMcpError("bad_response", f"{method} 응답이 비어 있습니다")
        if resp.get("error"):
            raise LawMcpError("tool_error", f"{method} 오류: {resp['error'].get('message', resp['error'])}")
        return resp.get("result", {})

    def _ensure_ready(self):
        if self._ready:
            return
        self._rpc("initialize", {"protocolVersion": "2025-06-18", "capabilities": {},
                                 "clientInfo": {"name": "hangul-jipsa", "version": "0.1"}})
        self._rpc("notifications/initialized", {}, notify=True)
        self._ready = True

    def call(self, tool: str, args: dict):
        """도구를 부르고 결과를 돌려준다(JSON 텍스트면 dict로 풀어 줌)."""
        check_outbound(tool, args, self.hosted)
        self._ensure_ready()
        result = self._rpc("tools/call", {"name": tool, "arguments": args})
        texts = [c.get("text", "") for c in result.get("content", []) if c.get("type") == "text"]
        joined = "\n".join(texts)
        if result.get("isError"):
            raise LawMcpError("tool_error", joined[:300] or f"{tool} 실패")
        try:
            return json.loads(joined)
        except Exception:
            return joined

    # 자주 쓰는 호출(루프·채팅 도구가 그대로 씀)
    def find_statute(self, query: str, size: int = 5):
        return self.call("statute_find", {"query": query, "size": size})

    def get_articles(self, name: str, articles: list, as_of: str = None):
        args = {"name": name, "articles": articles}
        if as_of:
            args["asOf"] = as_of
        return self.call("article_get", args)

    def audit_citations(self, citations: list, as_of: str = None):
        args = {"citations": citations}
        if as_of:
            args["asOf"] = as_of
        return self.call("citation_audit", args)


_STATUS_KO = {
    "VERIFIED": ("확인됨", "ok"),
    "CASE_VERIFIED": ("판례 확인됨", "ok"),
    "ARTICLE_NOT_FOUND": ("없는 조문", "bad"),
    "PARAGRAPH_NOT_FOUND": ("없는 항", "bad"),
    "LAW_NOT_FOUND": ("법령명을 찾지 못함", "bad"),
    "CASE_NOT_FOUND": ("판례를 찾지 못함", "bad"),
    "TITLE_MISMATCH": ("조문 제목 불일치", "bad"),
    "CONTENT_MISMATCH": ("내용 불일치", "bad"),
    "NOT_IN_FORCE": ("기준일에 시행 중 아님", "bad"),
    "UNCHECKED": ("확인 못 함(서버 사정)", "unknown"),
}


def format_audit(audit: dict) -> dict:
    """citation_audit 응답을 채팅용 한국어 요약으로. 반환: {"summary", "lines", "bad_count"}"""
    lines, bad = [], 0
    for it in (audit or {}).get("items", []):
        label, level = _STATUS_KO.get(it.get("status"), (it.get("status", "?"), "unknown"))
        cite = f"「{it.get('law', '')}」 {it.get('article', '')}".strip()
        extra = ""
        if it.get("status") == "ARTICLE_NOT_FOUND" and it.get("validRange"):
            extra = f" — 이 법령의 조문 범위: {it['validRange']}"
        elif it.get("status") == "TITLE_MISMATCH" and (it.get("actual") or {}).get("title"):
            extra = f" — 실제 제목: {it['actual']['title']}"
        elif it.get("status") == "VERIFIED" and (it.get("actual") or {}).get("effectiveDate"):
            extra = f" (현행 시행일 {it['actual']['effectiveDate']})"
        mark = {"ok": "✓", "bad": "✗", "unknown": "?"}[level]
        if level == "bad":
            bad += 1
        lines.append(f"{mark} {cite}: {label}{extra}")
    total = len(lines)
    summary = (f"법령 인용 {total}건 중 {bad}건이 확인 필요해요." if bad
               else f"법령 인용 {total}건 모두 현행 법령에서 확인됐어요.")
    return {"summary": summary, "lines": lines, "bad_count": bad}


def audit_document_citations(text: str, client: "LawMcpClient" = None, as_of: str = None) -> dict:
    """문서 텍스트에서 인용만 로컬로 뽑아 검증한다(본문은 보내지 않음).
    반환: {"citations": [...보낸 것...], "result": 서버 응답 | None}"""
    citations = extract_citations(text)
    if not citations:
        return {"citations": [], "result": None}
    client = client or LawMcpClient()
    return {"citations": citations, "result": client.audit_citations(citations[:100], as_of)}


# ---------------------------------------------------------------- self-tests

class _FakeTransport:
    def __init__(self, tool_result=None):
        self.bodies = []
        self.tool_result = tool_result if tool_result is not None else {"ok": True}

    def post(self, body):
        self.bodies.append(body)
        if body["method"] == "initialize":
            return {"jsonrpc": "2.0", "id": body["id"], "result": {"serverInfo": {"name": "fake"}}}
        if body["method"] == "tools/call":
            return {"jsonrpc": "2.0", "id": body["id"],
                    "result": {"content": [{"type": "text", "text": json.dumps(self.tool_result, ensure_ascii=False)}]}}
        return None


def _selftest_extract_citations():
    text = ("「공공기관의 운영에 관한 법률」 제4조(공공기관) 및 같은 법 제 5 조의 2에 따라 지정한다. "
            "「학교급식법」 제10조, 주택법 제5조도 참고. 또 「공공기관의 운영에 관한 법률」 제4조는 중복. "
            "「국가계약법」만 언급.")
    got = extract_citations(text)
    assert got == [
        {"law": "공공기관의 운영에 관한 법률", "article": "제4조", "title": "공공기관"},
        {"law": "공공기관의 운영에 관한 법률", "article": "제5조의2"},
        {"law": "학교급식법", "article": "제10조"},
        {"law": "주택법", "article": "제5조"},
        {"law": "국가계약법"},
    ], got
    assert extract_citations("법령 인용 없는 문장입니다.") == []
    assert extract_citations("동법 제3조") == []  # 앞선 법령이 없으면 해석하지 않음
    assert extract_citations("민법 제618조와 본법 제2조") == [{"law": "민법", "article": "제618조"}]
    print("extract_citations(「」·낱말 법령·같은 법·중복 제거) 통과")


def _selftest_check_outbound_blocks():
    def blocked(tool, args):
        try:
            check_outbound(tool, args, hosted=True)
        except LawMcpError as e:
            assert e.kind == "blocked", e.kind
            return str(e)
        return None
    assert blocked("statute_find", {"query": "학교급식법"}) is None
    assert blocked("citation_audit", {"citations": [{"law": "학교급식법", "article": "제10조"}]}) is None
    assert "도구" in blocked("contract_risk_scan", {"text": "계약서"})
    assert "도구" in blocked("toolbox", {"action": "x"})
    assert "본문" in blocked("citation_audit", {"text": "문서 전문"})
    assert "필드" in blocked("citation_audit", {"citations": [{"law": "a", "claim": "주장 문장"}]})
    assert "자" in blocked("statute_find", {"query": "가" * 101})
    assert "개인정보" in blocked("statute_find", {"query": "010-1234-5678"})
    check_outbound("contract_risk_scan", {"text": "가" * 5000}, hosted=False)  # 로컬 서버는 통과
    assert is_hosted(DEFAULT_URL) and not is_hosted("http://127.0.0.1:3457/mcp")
    print("check_outbound(허용 도구·본문 금지·길이·개인정보, 로컬 완화) 통과")


def _selftest_client_with_fake_transport():
    fake = _FakeTransport({"ok": True, "items": [{"name": "학교급식법"}]})
    c = LawMcpClient(url=DEFAULT_URL, token="t", transport=fake)
    r = c.find_statute("학교급식법")
    assert r["items"][0]["name"] == "학교급식법", r
    methods = [b["method"] for b in fake.bodies]
    assert methods == ["initialize", "notifications/initialized", "tools/call"], methods
    assert fake.bodies[2]["params"] == {"name": "statute_find", "arguments": {"query": "학교급식법", "size": 5}}
    c.find_statute("주택법")
    assert [b["method"] for b in fake.bodies].count("initialize") == 1  # 초기화는 한 번만
    # 관문에서 막히면 아무것도 나가지 않는다
    n = len(fake.bodies)
    try:
        c.call("contract_risk_scan", {"text": "x"})
        raise AssertionError("막혀야 함")
    except LawMcpError as e:
        assert e.kind == "blocked"
    assert len(fake.bodies) == n
    # 문서 인용 검증: 본문 대신 인용 목록만 나간다
    out = audit_document_citations("「학교급식법」 제10조에 따라 010-1234-5678로 문의", client=c)
    sent = fake.bodies[-1]["params"]["arguments"]
    assert sent == {"citations": [{"law": "학교급식법", "article": "제10조"}]}, sent
    assert "010" not in json.dumps(sent, ensure_ascii=False) and out["citations"] == sent["citations"]
    print("LawMcpClient(가짜 전송: 초기화 1회·관문 차단 시 미전송·인용만 전송) 통과")


def _selftest_format_audit():
    audit = {"items": [
        {"status": "VERIFIED", "law": "공공기관의 운영에 관한 법률", "article": "제4조",
         "actual": {"title": "공공기관", "effectiveDate": "2026-10-02"}},
        {"status": "ARTICLE_NOT_FOUND", "law": "공공기관의 운영에 관한 법률", "article": "제999조", "validRange": "제1조~제56조"},
        {"status": "LAW_NOT_FOUND", "law": "없는법률이름", "article": "제1조"},
        {"status": "UNCHECKED", "law": "학교급식법", "article": "제10조"},
    ]}
    f = format_audit(audit)
    assert f["bad_count"] == 2 and f["summary"] == "법령 인용 4건 중 2건이 확인 필요해요.", f
    assert f["lines"][0] == "✓ 「공공기관의 운영에 관한 법률」 제4조: 확인됨 (현행 시행일 2026-10-02)", f["lines"][0]
    assert f["lines"][1].endswith("없는 조문 — 이 법령의 조문 범위: 제1조~제56조") and f["lines"][3].startswith("? "), f["lines"]
    assert format_audit({"items": [audit["items"][0]]})["summary"] == "법령 인용 1건 모두 현행 법령에서 확인됐어요."
    print("format_audit(판정별 한국어 요약) 통과")


def _selftest_live_hosted(optional: bool = True):
    """실제 호스팅 서버: 토큰이 있고 연결되면 키워드 검색·인용 검증 1회씩(문서 내용 없음)."""
    if not _load_token():
        print("live 건너뜀 — 토큰 없음")
        return
    try:
        c = LawMcpClient()
        r = c.find_statute("공공기관의 운영에 관한 법률", size=2)
        names = [it.get("name") for it in r.get("items", [])]
        assert "공공기관의 운영에 관한 법률" in names, names
        a = c.audit_citations([{"law": "공공기관의 운영에 관한 법률", "article": "제4조"},
                               {"law": "공공기관의 운영에 관한 법률", "article": "제999조"}])
        print("live 통과 — statute_find:", names, "| citation_audit 응답 키:", list(a.keys())[:8] if isinstance(a, dict) else type(a))
    except LawMcpError as e:
        if optional and e.kind in ("unreachable", "auth"):
            print(f"live 건너뜀 — {e.kind}: {e}")
            return
        raise


if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8")
    _selftest_extract_citations()
    _selftest_check_outbound_blocks()
    _selftest_client_with_fake_transport()
    _selftest_format_audit()
    if "--live" in sys.argv:
        _selftest_live_hosted()
