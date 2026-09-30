"""doc_sections.py — 범용 부품 1·2·3: 열려 있는 한글 문서를 "구간" 단위로 읽고 쓴다.

PRD 16-2의 부품 1(read_outline), 2(read_section), 3(write_section)의 구현.
업무별 도구(표 삽입, 번호 매기기 …)가 아니라, 어떤 요청이든 모델이 조립해
쓸 수 있는 작은 부품이라는 점이 설계의 핵심이다(PRD 16-1).

(2026-10-01 실측, 야간 자율 작업) 설계 근거가 된 pyhwpx 동작:
- 본문은 list 0이고 문단은 0부터 번호가 붙는다. `set_pos(0, i, 0)`이 실패
  (False)하면 문단 범위 밖이다 → 이걸로 문단 수를 센다.
- `select_text(i, 0, i, -1, 0)` + `get_selected_text()`로 문단 i의 텍스트를
  읽는다. 표가 앵커된 문단을 선택하면 셀 내용이 "\\r\\n"으로 이어진 문자열이
  나온다(표 자체는 list 0의 문단 하나로 셈).
- 표 컨트롤의 `GetAnchorPos(0)`이 (0, para, 0)을 돌려줘 어느 문단에 표가
  있는지 정확히 알 수 있다.
- `select_text(s, 0, e, -1, 0)` + `insert_text(...)`로 여러 문단을 한 번에
  교체할 수 있고, 교체 뒤 뒤쪽 문단 번호와 표 앵커가 함께 당겨진다(밀린다).
  → **쓰기 뒤에는 구간 id(문단 번호 기반)가 무효**가 되므로 write_section은
  새 outline을 함께 돌려주고, 호출자는 그걸 다시 써야 한다.
- `set_pos(0, i, -1)` + `BreakPara()` + `insert_text()`로 문단 i 뒤에 새
  문단을 끼울 수 있다.

구간(section)의 정의: 항목기호(1. 가. 1) 가) (1) (가) ① ㉮, 그리고 □ ○ - ·
같은 보고서식 글머리)로 시작하는 문단을 "제목 문단"으로 보고, 그 제목부터
같은 단계 이상의 다음 제목 직전까지를 한 구간으로 본다. 항목기호가 하나도
없는 문서는 문단 하나하나를 구간으로 본다(폴백). 표는 그 자체로 한 구간
(kind="table")이며 write_section이 표 문단을 덮어쓰지 않도록 막는다 —
표 조작은 별도 부품/도구(table_tool 등)의 몫이다.

이 모듈은 HwpReport(hwp_report.py)를 받아서 동작하고, 문서를 저장하지
않는다(이 프로젝트의 일관된 원칙 — 자동저장 안 함, 사용자가 확인 후 저장)."""
import re

# 행정업무운영편람의 8단계 항목기호(.tmp/gov-hwp-formatting-rules-2026-09-07.md
# 규칙 1) + 보고서에서 흔한 글머리 기호. 앞쪽이 상위 단계. format_checker.py의
# LEVEL_PATTERNS(4단계)는 번호 "순서 검사"용이라 그대로 두고, 여기서는 "단계
# 판정"만 하므로 더 넓게 잡는다. 글머리 기호 계열(□ ○ - ·)은 편람 항목기호와
# 섞어 쓰는 문서가 많아, 편람 8단계 뒤에 낮은 단계로 이어 붙였다 — 완벽한
# 규칙은 아니고 "같은 기호끼리는 같은 단계"라는 최소한의 보장만 한다.
_LEVEL_PATTERNS = [
    re.compile(r"^\d+\.\s"),          # 1.
    re.compile(r"^[가-힣]\.\s"),       # 가.
    re.compile(r"^\d+\)\s"),          # 1)
    re.compile(r"^[가-힣]\)\s"),       # 가)
    re.compile(r"^\(\d+\)\s"),        # (1)
    re.compile(r"^\([가-힣]\)\s"),     # (가)
    re.compile(r"^[①-⑳]\s?"),         # ①
    re.compile(r"^[㉮-㉻]\s?"),        # ㉮
    re.compile(r"^[□■]\s?"),          # □
    re.compile(r"^[○◦]\s?"),          # ○
    re.compile(r"^[-–—]\s"),          # -
    re.compile(r"^[·•ㆍ]\s?"),        # ·
]


def level_of(text: str):
    """문단 텍스트의 항목기호 단계(0이 최상위). 항목기호가 없으면 None."""
    stripped = text.lstrip()
    for level, pattern in enumerate(_LEVEL_PATTERNS):
        if pattern.match(stripped):
            return level
    return None


def _read_paragraphs(hwp) -> list:
    """본문(list 0)의 모든 문단 텍스트를 순서대로 읽는다. 표가 앵커된 문단은
    셀 내용이 이어진 문자열로 나온다(호출자가 표 앵커 목록으로 구분)."""
    texts = []
    i = 0
    while hwp.set_pos(0, i, 0):
        hwp.select_text(i, 0, i, -1, 0)
        texts.append(hwp.get_selected_text() or "")
        hwp.Cancel()
        i += 1
    return texts


def _table_anchor_paras(hwp) -> list:
    """본문(list 0)에 앵커된 표들의 문단 번호를 문서 순서대로 돌려준다
    (hwp_report._read_table_grids와 같은 HeadCtrl 순회 → table_to_df(n)의
    n과 순서가 같다)."""
    paras = []
    ctrl = hwp.HeadCtrl
    while ctrl:
        if ctrl.CtrlID == "tbl":
            try:
                anchor = ctrl.GetAnchorPos(0)
                if anchor.Item("List") == 0:
                    paras.append(anchor.Item("Para"))
                else:
                    paras.append(None)  # 표 안의 표 등 — 본문 순회 대상 아님
            except Exception:
                paras.append(None)
        ctrl = ctrl.Next
    return paras


def build_outline(paragraphs: list, table_paras: list) -> list:
    """문단 텍스트 목록과 표 앵커 문단 번호로 구간 목록을 만든다(순수 함수 —
    한글 없이 테스트 가능). 각 구간은 dict:
      id, kind("heading"|"paragraph"|"table"), level(int|None), para(제목 문단),
      title(제목 문단 텍스트, 앞 60자), body_start, body_end(본문 문단 범위,
      제목 제외, 포함 범위; 본문이 없으면 None), table_index(kind=table일 때
      table_to_df용 n), preview(표면 첫 행 요약)"""
    table_index_by_para = {p: n for n, p in enumerate(table_paras) if p is not None}
    n = len(paragraphs)
    is_heading = [False] * n
    levels = [None] * n
    for i, text in enumerate(paragraphs):
        if i in table_index_by_para:
            continue
        lv = level_of(text)
        levels[i] = lv
        is_heading[i] = lv is not None

    any_heading = any(is_heading)
    sections = []
    for i, text in enumerate(paragraphs):
        if i in table_index_by_para:
            first_row = text.replace("\r\n", "\n").split("\n")
            first_row = [c for c in first_row if c][:4]
            sections.append({
                "id": f"s{i}", "kind": "table", "level": None, "para": i,
                "title": "[표] " + " | ".join(first_row), "body_start": None, "body_end": None,
                "table_index": table_index_by_para[i], "preview": first_row,
            })
            continue
        if any_heading and not is_heading[i]:
            continue  # 제목 문단이 있는 문서에선 본문 문단은 구간에 포함될 뿐 따로 나열하지 않음
        if not any_heading and not text.strip():
            continue  # 폴백 모드에서 빈 문단은 건너뜀
        lv = levels[i]
        # 본문 범위: 다음 "같은 단계 이상의 제목" 또는 문서 끝 직전까지.
        # 표는 본문 범위 안에 포함될 수 있다(read_section이 격자로 같이 보여줌)
        end = n - 1
        for j in range(i + 1, n):
            if is_heading[j] and levels[j] is not None and lv is not None and levels[j] <= lv:
                end = j - 1
                break
            if not any_heading:
                end = i
                break
        body_start, body_end = (i + 1, end) if end >= i + 1 else (None, None)
        sections.append({
            "id": f"s{i}", "kind": "heading" if any_heading else "paragraph", "level": lv,
            "para": i, "title": text.strip()[:60], "body_start": body_start, "body_end": body_end,
            "table_index": None, "preview": None,
        })
    return sections


def read_outline(report) -> dict:
    """부품 1. 열려 있는 문서의 구간 목록을 돌려준다.
    반환: {"para_count": n, "sections": [...]} (build_outline 참고)."""
    hwp = report.hwp
    paragraphs = _read_paragraphs(hwp)
    table_paras = _table_anchor_paras(hwp)
    hwp.set_pos(0, 0, 0)
    return {"para_count": len(paragraphs), "sections": build_outline(paragraphs, table_paras)}


def format_outline(outline: dict) -> str:
    """모델/사람에게 보여줄 목차 문자열. 단계만큼 들여쓴다."""
    lines = []
    for s in outline["sections"]:
        indent = "  " * (s["level"] or 0)
        span = ""
        if s["body_start"] is not None:
            span = f" (본문 문단 {s['body_start']}~{s['body_end']})"
        lines.append(f"{indent}[{s['id']}] {s['title']}{span}")
    return "\n".join(lines)


def _find_section(outline: dict, section_id: str):
    for s in outline["sections"]:
        if s["id"] == section_id:
            return s
    return None


def section_at_para(outline: dict, para: int):
    """문단 번호가 속한 구간(제목 문단이거나 본문 범위 안)을 돌려준다. 여러
    단계가 겹치면 가장 깊은(마지막에 나열된) 구간을 고른다."""
    chosen = None
    for s in outline["sections"]:
        if s["para"] == para:
            return s
        if s["body_start"] is not None and s["body_start"] <= para <= s["body_end"]:
            chosen = s
    return chosen


def section_at_cursor(report, outline: dict = None):
    """커서가 있는 구간. 커서가 표 안(list != 0)이면 그 표 구간을 찾아준다."""
    hwp = report.hwp
    if outline is None:
        pos = hwp.get_pos()
        outline = read_outline(report)
        hwp.set_pos(*pos)
    lst, para, _ = hwp.get_pos()
    if lst != 0:
        # 표 안: 표 컨트롤의 본문 앵커로 환산
        try:
            anchor = hwp.get_ctrl_pos(hwp.ParentCtrl, option=1)
            para = anchor[1]
        except Exception:
            return None
    return section_at_para(outline, para)


def read_section(report, section_id: str, outline: dict = None) -> dict:
    """부품 2. 구간의 제목·본문 텍스트와, 본문 범위 안의 표 격자를 돌려준다.
    반환: {"id", "title", "body": str, "paragraphs": [str], "tables": [grid]}.
    없는 id면 {"error": ...}."""
    hwp = report.hwp
    if outline is None:
        outline = read_outline(report)
    s = _find_section(outline, section_id)
    if s is None:
        return {"error": f"구간 {section_id}이(가) 없습니다"}
    paragraphs = _read_paragraphs(hwp)
    table_paras = _table_anchor_paras(hwp)
    hwp.set_pos(0, 0, 0)

    if s["kind"] == "table":
        grid = _table_grid(hwp, s["table_index"])
        return {"id": s["id"], "title": s["title"], "body": _grid_to_text(grid),
                "paragraphs": [], "tables": [grid]}

    body_paras, tables = [], []
    if s["body_start"] is not None:
        for i in range(s["body_start"], s["body_end"] + 1):
            if i in table_paras:
                grid = _table_grid(hwp, table_paras.index(i))
                tables.append(grid)
                body_paras.append(_grid_to_text(grid))
            else:
                body_paras.append(paragraphs[i])
    return {"id": s["id"], "title": paragraphs[s["para"]], "body": "\n".join(body_paras),
            "paragraphs": body_paras, "tables": tables}


def _table_grid(hwp, table_index: int) -> list:
    try:
        df = hwp.table_to_df(table_index, cols=0)
    except Exception:
        return []
    return [[str(c) for c in df.columns]] + [[str(v) for v in row] for row in df.to_numpy().tolist()]


def _grid_to_text(grid: list) -> str:
    return "\n".join(" | ".join(row) for row in grid)


def _insert_paragraphs(hwp, lines: list) -> None:
    """현재 커서/선택 위치에 줄 목록을 문단으로 넣는다(첫 줄은 선택을 덮어씀)."""
    for k, line in enumerate(lines):
        if k > 0:
            hwp.BreakPara()
        hwp.insert_text(line)


def write_section(report, section_id: str, text: str, mode: str = "replace_body",
                  outline: dict = None) -> dict:
    """부품 3. 구간에 텍스트를 쓴다. 문서를 저장하지 않는다.

    mode:
      - "replace_body": 제목 문단은 그대로 두고 본문 문단들만 text로 교체.
        본문이 없던 구간이면 제목 뒤에 새로 끼운다.
      - "replace": 제목 문단까지 포함해 구간 전체를 text로 교체.
      - "append": 구간 본문 끝에 text를 새 문단(들)으로 덧붙임.
      - "insert_after": 구간(본문 포함) 바로 뒤에 새 문단(들)로 끼움.
    text의 줄바꿈("\\n")마다 문단을 나눈다.

    안전장치: 교체 범위 안에 표가 있으면 덮어쓰지 않고 거부한다(표는 이 부품의
    대상이 아님 — 표를 잃는 실수를 막기 위해). 승인/변경추적은 호출자(채팅
    도구)가 맡는다 — 이 함수는 조립용 부품이라 정책을 갖지 않는다.

    반환: {"applied": bool, "reason"?: str, "before": str, "after": str,
           "outline": 새 outline}. 쓰기 뒤에는 문단 번호가 바뀌므로 기존
    outline/section id는 더 이상 유효하지 않다 — 반드시 반환된 outline을 쓸 것."""
    hwp = report.hwp
    if mode not in ("replace_body", "replace", "append", "insert_after"):
        return {"applied": False, "reason": f"알 수 없는 mode: {mode}"}
    if outline is None:
        outline = read_outline(report)
    s = _find_section(outline, section_id)
    if s is None:
        return {"applied": False, "reason": f"구간 {section_id}이(가) 없습니다"}
    if s["kind"] == "table":
        return {"applied": False, "reason": "표 구간은 이 부품으로 쓸 수 없습니다(표 도구를 쓰세요)"}

    lines = [ln.rstrip("\r") for ln in text.replace("\r\n", "\n").split("\n")]
    if not any(ln.strip() for ln in lines):
        return {"applied": False, "reason": "쓸 내용이 비어 있습니다"}

    paragraphs = _read_paragraphs(hwp)
    table_paras = set(p for p in _table_anchor_paras(hwp) if p is not None)
    body_start, body_end = s["body_start"], s["body_end"]
    section_end = body_end if body_end is not None else s["para"]

    if mode in ("replace_body", "replace"):
        start = s["para"] if mode == "replace" else body_start
        end = section_end
        if start is not None and any(p in table_paras for p in range(start, end + 1)):
            return {"applied": False, "reason": "교체 범위 안에 표가 있어 덮어쓰지 않았습니다"}
        if start is None:
            # 본문이 없는 구간: 제목 문단 뒤에 새로 끼움
            before = ""
            hwp.set_pos(0, s["para"], -1)
            hwp.BreakPara()
            _insert_paragraphs(hwp, lines)
        else:
            before = "\n".join(paragraphs[start:end + 1])
            hwp.select_text(start, 0, end, -1, 0)
            _insert_paragraphs(hwp, lines)
    elif mode == "append":
        before = "\n".join(paragraphs[body_start:body_end + 1]) if body_start is not None else ""
        hwp.set_pos(0, section_end, -1)
        hwp.BreakPara()
        _insert_paragraphs(hwp, lines)
    else:  # insert_after
        before = ""
        hwp.set_pos(0, section_end, -1)
        hwp.BreakPara()
        _insert_paragraphs(hwp, lines)

    hwp.Cancel()
    new_outline = read_outline(report)
    hwp.set_pos(0, s["para"], 0)
    return {"applied": True, "before": before, "after": "\n".join(lines), "outline": new_outline}


# ---------------------------------------------------------------- self-tests

def _selftest_build_outline_pure():
    """한글 없이: 항목기호 단계·본문 범위·표 구간·폴백 규칙."""
    paras = ["서류심사 기본계획(안)", "1. 추진 배경", "가. 목적", "절차 표준화", "기간 단축",
             "나. 근거", "규정 제3조", "2. 추진 계획", "구분\r\n건수\r\n2024\r\n120\r\n", "가. 일정", "3월 착수"]
    sections = build_outline(paras, [8])
    ids = [(s["id"], s["kind"], s["level"], s["body_start"], s["body_end"]) for s in sections]
    assert ids == [
        ("s1", "heading", 0, 2, 6),      # 1. 추진 배경: 가.~규정 제3조
        ("s2", "heading", 1, 3, 4),      # 가. 목적: 두 문단
        ("s5", "heading", 1, 6, 6),      # 나. 근거
        ("s7", "heading", 0, 8, 10),     # 2. 추진 계획: 표 + 가. 일정 + 3월 착수
        ("s8", "table", None, None, None),
        ("s9", "heading", 1, 10, 10),
    ], ids
    assert sections[4]["title"].startswith("[표] 구분 | 건수"), sections[4]["title"]
    # 제목 문단 자체가 문서 제목(항목기호 없음)이면 구간으로 나열되지 않는다
    assert all(s["para"] != 0 for s in sections)
    # 폴백: 항목기호가 전혀 없으면 문단 하나하나가 구간
    plain = build_outline(["첫 문단", "", "둘째 문단"], [])
    assert [(s["id"], s["kind"], s["body_start"]) for s in plain] == [("s0", "paragraph", None), ("s2", "paragraph", None)], plain
    # 글머리 기호 계열
    assert level_of("□ 추진 방향") == 8 and level_of("○ 세부 과제") == 9 and level_of("- 일정") == 10
    assert level_of("(1) 세부") == 4 and level_of("① 항목") == 6 and level_of("일반 문장.") is None
    # section_at_para: 겹치면 가장 깊은 구간
    outline = {"sections": sections}
    assert section_at_para(outline, 4)["id"] == "s2"
    assert section_at_para(outline, 6)["id"] == "s5"
    assert section_at_para(outline, 10)["id"] == "s9"
    assert section_at_para(outline, 1)["id"] == "s1"
    print("build_outline/level_of/section_at_para(순수) 통과")


def _make_probe_report():
    """실제 한글로 프로브와 같은 가짜 문서를 만들어 HwpReport로 연다."""
    import os, tempfile
    from hwp_report import HwpReport
    from hwp_session import new_hwp
    path = os.path.join(tempfile.gettempdir(), "_test_doc_sections.hwp")
    setup = new_hwp(visible=False)
    for ln in ["서류심사 기본계획(안)", "1. 추진 배경", "가. 목적", "심사 절차를 표준화한다.",
               "심사 기간을 단축한다.", "나. 근거", "관련 규정 제3조", "2. 추진 계획"]:
        setup.insert_text(ln); setup.BreakPara()
    setup.create_table(2, 2)
    setup.insert_text("구분"); setup.TableRightCell(); setup.insert_text("건수")
    setup.TableRightCell(); setup.insert_text("2024"); setup.TableRightCell(); setup.insert_text("120")
    setup.MoveDocEnd()
    setup.insert_text("가. 일정"); setup.BreakPara(); setup.insert_text("3월 착수")
    setup.save_as(path)
    setup.quit()
    # setup 객체를 호출자에게 같이 넘겨 테스트가 끝날 때까지 살려 둔다 —
    # pyhwpx Hwp의 __del__이 CoUninitialize()를 호출해 같은 프로세스의 다른
    # 살아있는 인스턴스(여기선 HwpReport)의 COM 연결을 끊는다(hwp-report-tool
    # 메모 2026-09-04 항목, 이 함수가 반환되며 setup이 수거되자 실제로
    # "개체가 열려 있지 않거나 등록되지 않았습니다"로 재현됨).
    return HwpReport(path), path, setup


def _selftest_live_read_outline_and_section():
    """실제 한글: outline이 프로브와 같은 구조로 나오고 read_section이 본문·표를 읽는다."""
    import os
    report, path, _keep_alive = _make_probe_report()
    try:
        outline = read_outline(report)
        assert outline["para_count"] == 11, outline["para_count"]
        ids = [(s["id"], s["kind"]) for s in outline["sections"]]
        assert ids == [("s1", "heading"), ("s2", "heading"), ("s5", "heading"), ("s7", "heading"),
                       ("s8", "table"), ("s9", "heading")], ids
        sec = read_section(report, "s2", outline)
        assert sec["title"] == "가. 목적" and sec["paragraphs"] == ["심사 절차를 표준화한다.", "심사 기간을 단축한다."], sec
        sec7 = read_section(report, "s7", outline)
        assert sec7["tables"] and sec7["tables"][0][0] == ["구분", "건수"] and sec7["tables"][0][1] == ["2024", "120"], sec7["tables"]
        assert "가. 일정" in sec7["body"] and "3월 착수" in sec7["body"], sec7["body"]
        tbl = read_section(report, "s8", outline)
        assert tbl["tables"][0][1] == ["2024", "120"], tbl
        # 커서를 문단 4에 두면 s2가 잡힌다
        report.hwp.set_pos(0, 4, 0)
        assert section_at_cursor(report, outline)["id"] == "s2"
        print("read_outline/read_section/section_at_cursor(실제 한글) 통과")
        print(format_outline(outline))
    finally:
        report.close(save=False)
        os.remove(path)


def _selftest_live_write_section_modes():
    """실제 한글: replace_body / append / insert_after / replace, 표 보호, id 갱신."""
    import os
    report, path, _keep_alive = _make_probe_report()
    try:
        outline = read_outline(report)
        r = write_section(report, "s2", "심사 절차를 표준화하고 기간을 단축한다.", "replace_body", outline)
        assert r["applied"], r
        assert r["before"] == "심사 절차를 표준화한다.\n심사 기간을 단축한다.", r["before"]
        outline = r["outline"]
        assert outline["para_count"] == 10, outline["para_count"]
        sec = read_section(report, "s2", outline)
        assert sec["paragraphs"] == ["심사 절차를 표준화하고 기간을 단축한다."], sec
        # 표 앵커가 함께 당겨졌는지(문단 8 → 7)
        assert [s["para"] for s in outline["sections"] if s["kind"] == "table"] == [7]

        r = write_section(report, "s2", "추가 문장 1\n추가 문장 2", "append", outline)
        assert r["applied"], r
        outline = r["outline"]
        sec = read_section(report, "s2", outline)
        assert sec["paragraphs"] == ["심사 절차를 표준화하고 기간을 단축한다.", "추가 문장 1", "추가 문장 2"], sec

        # 본문이 없는 구간(나. 근거의 본문을 비운 뒤)에 replace_body → 제목 뒤에 끼움
        sec_na = [s for s in outline["sections"] if s["title"] == "나. 근거"][0]
        r = write_section(report, sec_na["id"], "다. 임시", "insert_after", outline)
        assert r["applied"], r
        outline = r["outline"]
        titles = [s["title"] for s in outline["sections"]]
        assert "다. 임시" in titles, titles

        # 표가 포함된 구간(2. 추진 계획)의 본문 교체는 거부
        sec2 = [s for s in outline["sections"] if s["title"] == "2. 추진 계획"][0]
        r = write_section(report, sec2["id"], "덮어쓰기 시도", "replace_body", outline)
        assert r["applied"] is False and "표" in r["reason"], r
        # 표 구간 자체도 거부
        tbl = [s for s in outline["sections"] if s["kind"] == "table"][0]
        r = write_section(report, tbl["id"], "x", "replace_body", outline)
        assert r["applied"] is False, r
        # 표는 그대로 살아 있음
        assert read_section(report, tbl["id"], outline)["tables"][0][1] == ["2024", "120"]

        # replace: 제목까지 교체
        sec_il = [s for s in outline["sections"] if s["title"] == "가. 일정"][0]
        r = write_section(report, sec_il["id"], "나. 추진 일정\n4월 착수", "replace", outline)
        assert r["applied"], r
        outline = r["outline"]
        titles = [s["title"] for s in outline["sections"]]
        assert "나. 추진 일정" in titles and "가. 일정" not in titles, titles
        # 빈 텍스트 거부
        r = write_section(report, "s2", "  \n ", "replace_body", outline)
        assert r["applied"] is False, r
        print("write_section(4 mode + 표 보호 + id 갱신, 실제 한글) 통과")
        print(format_outline(outline))
    finally:
        report.close(save=False)
        os.remove(path)


if __name__ == "__main__":
    import sys, time
    sys.stdout.reconfigure(encoding="utf-8")
    _selftest_build_outline_pure()
    _selftest_live_read_outline_and_section()
    time.sleep(2)
    _selftest_live_write_section_modes()
