"""eval_mock.py — 사용자 제공 목업 문서로 도는 "실제 모양" 회귀(평가셋 2단계의 첫 조각).

가짜 문서(self-test가 만드는 것)는 모양이 단순해서, 실제 aT 보고서의 구조
(로마숫자 장 제목 상자, ❍ 글머리, 공백 들여쓰기, 표지·목차 표)에서 나는
문제를 못 잡는다. `_목업/`에 사용자가 정제해 준 목업(업체명·담당자·연락처 없음,
수치 변경)이 있으면 그 문서로 아래를 확인한다. 파일이 없으면(저장소엔 안
올라감) 건너뜀으로 끝낸다 — 실패가 아니다.

확인 항목(2026-10-01 실측값 기준, 바뀌면 의도한 변경인지 검토할 것):
  1. kordoc 읽기: 3쪽, "공급업체 관리 개요" 포함, 1초 안팎
  2. outline: "Ⅰ. 공급업체 관리 개요"가 0단계 제목(제목 상자 표), ❍ 구간 2개,
     표지 날짜 줄("2026. 1.")은 제목이 아님
  3. 구간 편집: 선택 경로(❍ 문장 교체→되돌리기), 구간 경로(공백 들여쓰기 보존→되돌리기)
  4. 아카이브: _목업 폴더 색인 후 "서류심사 항목" 검색이 이 문서를 찾음
문서는 절대 저장하지 않는다(save=False, 색인은 읽기만)."""
import os
import sys
import tempfile
import time

HERE = os.path.dirname(os.path.abspath(__file__))
MOCK_DIR = os.path.join(HERE, "_목업")
MOCK_HWP = os.path.join(MOCK_DIR, "목업_eaT_공급업체관리결과_3장.hwp")


class _Shim:
    """HwpReport 대신 숨김 한글 인스턴스를 감싸는 최소 객체(편집 도구가 쓰는 속성만)."""
    _track_changes_enabled = False

    def __init__(self, hwp):
        self.hwp = hwp

    def get_text(self):
        return self.hwp.GetTextFile("TEXT", "")


def _check_kordoc():
    from attachments import kordoc_available, read_attachment
    if not kordoc_available():
        print("  kordoc 미설치 — 건너뜀(cd node_tools && npm install)")
        return
    t0 = time.time()
    r = read_attachment(MOCK_HWP, reader="kordoc")
    dt = time.time() - t0
    assert r["error"] is None and r["reader"] == "kordoc", r
    labels = [p["label"] for p in r["parts"]]
    assert labels == ["1페이지", "2페이지", "3페이지"], labels
    assert "공급업체 관리 개요" in r["text"] and "<table>" in r["text"], r["text"][:200]
    assert dt < 10, f"kordoc 읽기 {dt:.1f}s — 평소 1초 안팎"
    print(f"  kordoc 읽기 통과: {labels}, {len(r['text'])}자, {dt:.1f}s")


def _check_outline_and_edit():
    from doc_sections import _read_paragraphs, read_outline
    from hwp_session import new_hwp, open_document
    from section_edit_tool import _FakeClient, apply_section_edit, propose_section_edit, revert_section_edit
    hwp = new_hwp(visible=False)
    try:
        assert open_document(hwp, MOCK_HWP, unattended=True), "목업을 열지 못함"
        r = _Shim(hwp)
        outline = read_outline(r)
        titles = [(s["kind"], s["level"], s["title"]) for s in outline["sections"]]
        chapters = [t for t in titles if t[0] == "heading" and t[1] == 0]
        assert chapters == [("heading", 0, "Ⅰ. 공급업체 관리 개요")], chapters
        bullets = [t for t in titles if t[0] == "heading" and t[2].startswith("❍")]
        assert len(bullets) == 2, bullets
        assert not any("2026. 1." in t[2] for t in titles if t[0] == "heading"), titles
        print("  outline 통과:", [t[2][:18] for t in titles])

        # 선택 경로
        target = "상시 모니터링 통한 투명하고 안전한 급식 식재료 공급체계 지원"
        assert hwp.find(target, direction="AllDoc")
        p = propose_section_edit(r, "짧게", client=_FakeClient(["상시 모니터링으로 안전한 급식 식재료 공급 지원"]))
        assert p["ok"] and p["kind"] == "selection", p
        a = apply_section_edit(r, p)
        assert a["applied"] and "공급 지원" in r.get_text(), a
        rv = revert_section_edit(r, a)
        assert rv["reverted"] and target in r.get_text() and "공급 지원" not in r.get_text(), rv

        # 구간 경로: ❍ 구간의 본문(공백 들여쓰기 줄) 교체 → 들여쓰기 보존 → 되돌리기
        hwp.Cancel()
        paras = _read_paragraphs(hwp)
        outline = read_outline(r)
        sec = [s for s in outline["sections"] if s["title"].startswith("❍")][0]
        hwp.set_pos(0, sec["body_start"], 0)
        before_para = paras[sec["body_start"]]
        indent = before_para[:len(before_para) - len(before_para.lstrip())]
        assert indent, "목업의 본문 줄이 공백 들여쓰기를 쓴다는 전제가 깨짐"
        p = propose_section_edit(r, "다른 말로", client=_FakeClient(["관리 체계 개요"]))
        assert p["ok"] and p["kind"] == "section", p
        a = apply_section_edit(r, p)
        assert a["applied"], a
        assert _read_paragraphs(hwp)[sec["body_start"]] == indent + "관리 체계 개요"
        rv = revert_section_edit(r, a)
        assert rv["reverted"] and _read_paragraphs(hwp)[sec["body_start"]] == before_para, rv
        print("  구간 편집(선택·구간·들여쓰기 보존·되돌리기) 통과")
    finally:
        hwp.quit()


def _check_archive():
    from archive_index import reindex, search_archive
    db = os.path.join(tempfile.mkdtemp(prefix="_eval_mock_"), "idx.sqlite")
    st = reindex([MOCK_DIR], db_path=db)
    assert st["indexed"] == 1 and st["failed"] == [], st
    hits = search_archive("서류심사 항목 심사", db_path=db, k=3)
    assert hits and hits[0]["file"].startswith("목업_eaT") and "서류심사" in hits[0]["text"], hits[:1]
    print("  아카이브 색인·검색 통과:", hits[0]["label"], hits[0]["score"])


def main():
    if not os.path.exists(MOCK_HWP):
        print("목업 문서 없음 — 건너뜀 (사용자 정제 목업을 _목업/ 에 두면 실행됨)")
        return 0
    _check_kordoc()
    _check_outline_and_edit()
    _check_archive()
    print("eval_mock 전체 통과")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.exit(main())
