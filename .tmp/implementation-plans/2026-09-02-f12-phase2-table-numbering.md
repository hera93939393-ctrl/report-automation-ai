# F12 2단계 (표/번호서식 미리보기·선택) Implementation Plan

> For the agent worker: required sub skill. Implement task by task with `subagent-driven-development`. Track steps with checkbox (`- [ ]`) syntax.

Goal: 채팅창에서 "표 만들어줘"/"번호 매겨줘"라고 입력하면, 스타일/서식 미리보기 팝업에서 직접 보고 고른 뒤 그 결과가 열려있는 한글 문서의 커서 위치에 삽입된다.

Architecture: (1) 순수 로직 계층 — `table_tool.py`(엑셀→표 변환+스타일 적용), `numbering_tool.py`(프리픽스 삽입), 둘 다 hwp_report.py의 기존 HwpReport 핸들을 받아 동작(F12 1단계와 동일한 "문서 1회 열기" 전제) → (2) 미리보기 팝업 계층 — `chat_assistant.py`에 `_show_table_style_picker()`/`_show_numbering_style_picker()` 추가, CustomTkinter 위젯으로 그린 축소 모형 → (3) 라우팅 계층 — `route_intent()`에 두 도구 등록, `_on_submit()`이 팝업을 띄우고 팝업 콜백이 실제 삽입 함수를 호출하도록 연결.

Tech stack: 기존과 동일(Python 3.13, pyhwpx, customtkinter, openpyxl, ollama) — 신규 의존성 없음(matplotlib 등 차트 관련 라이브러리는 PRD 14-2에 따라 이번 라운드에서 의도적으로 배제).

PRD 참고: `PRD.md` "## 14. F12 — 2단계" (14-1~14-6). 특히 14-2(비목표)의 축소 지점(번호서식은 텍스트 프리픽스, 표 스타일은 배경색+굵게 2~3종, 차트 제외, 미리보기는 위젯 모형)을 반드시 그대로 따를 것 — 임의로 범위를 넓히지 말 것.

---

## 파일 구조

| 파일 | 책임 |
|---|---|
| `table_tool.py` (신규) | 원본자료(엑셀)를 읽어 표로 변환, 스타일(배경색+굵게) 적용, 커서 위치에 삽입하는 순수 도구 함수 |
| `numbering_tool.py` (신규) | 선택된 번호서식 프리픽스 문자열을 커서 위치에 삽입하는 순수 도구 함수 |
| `chat_assistant.py` (수정) | 두 도구를 route_intent에 등록, 두 개의 미리보기 팝업 메서드 추가, `_on_submit()`에서 팝업을 띄우도록 분기 |

---

### Task 1: table_tool.py — 원본자료를 표로 변환하는 순수 로직 ✅ 완료 (커밋 0cecfa8, 코드품질 리뷰 반영 45033d1)

**구현 시 변경사항**: Step 3의 "헤더 행 수동 블록선택(TableCellBlockRow 등)" 코드는 실제로 필요 없었다 — pyhwpx의 `table_from_data()`가 이미 `cell_fill` 파라미터를 내장하고 있어(직접 소스 5087번 줄 확인) `report.hwp.table_from_data(excel_path, header=True, index=False, header_bold=True, cell_fill=header_fill or False)` 한 줄로 대체함. `TableCellBlockRow`라는 메서드는 pyhwpx에 존재하지 않음을 확인함(계획서가 예상한 불확실 지점이 실재했음).

**코드품질 리뷰 반영**: "엑셀 없음" 경로에 대한 회귀 테스트가 없어서 `_selftest_insert_table_from_source_no_excel_source()` 추가.

**코드품질 재검토에서 발견된 크래시 버그 수정**: 문서에 텍스트가 선택된 상태로 `insert_table_from_source()`를 호출하면 프로그램이 크래시하는 결정적(재현성 100%) 버그가 발견됨. 원인: 선택된 상태로 `report.hwp.table_from_data(...)`를 호출하면, 그 내부의 `create_table()`이 실행하는 "TableCreate" HAction이 `pywintypes.com_error: (-2147417851, ...)`를 던지고, 이어서 pyhwpx `core.py`의 `create_table()` finally 블록(5916번 줄 근처)에서 `ctrl = self.hwp.CurSelectedCtrl or self.hwp.ParentCtrl`가 표 생성 실패로 인해 `None`인 채로 `ctrl.Properties = pset`을 실행하다 `AttributeError: 'NoneType' object has no attribute 'Properties'`로 이어짐. TDD로 먼저 회귀 테스트 `_selftest_insert_table_from_source_with_selection_does_not_crash()`를 추가해 수정 전 실제로 이 크래시(동일 스택트레이스)가 재현되는 것을 직접 실행으로 확인한 뒤, `insert_table_from_source()`에서 `report.hwp.table_from_data(...)` 호출 직전에 `report.hwp.Cancel()`을 추가해 고쳤다 — numbering_tool.py의 `insert_numbering_prefix()`(커밋 7e228e9)가 이미 같은 부류의 "선택된 콘텐츠" 위험에 대해 적용한 것과 동일한 방어 철학. numbering_tool.py의 텍스트 프리픽스 케이스와 달리 표는 새로 삽입되는 콘텐츠라 선택했던 텍스트가 지워지거나 대체될 위험 자체는 원래 없었다(Cancel()은 오직 table_from_data/create_table이 COM 에러 없이 안정적으로 동작하게 만들기 위한 방어). 수정 후 회귀 테스트로 직접 확인한 동작: "가나다라마바사"에서 "다라마"를 선택([find](AllDoc))한 채 표를 삽입하면, 결과 문서는 `가나다라마` + [표: 항목/금액/인건비/1000000] + `바사`가 된다 — 선택했던 "다라마"는 그대로 보존되고, 표는 Cancel() 이후 커서가 남는 위치(선택 영역의 끝점, "마"와 "바" 사이)에 삽입된다. 수정 후 기존 self-test 3종 + 신규 회귀 테스트 1종 전부(`python table_tool.py`) 통과 확인(exit code 0).

Files

- Create: `table_tool.py`

- [ ] Step 1: 실패하는 테스트 작성

`table_tool.py` 맨 아래에 작성 (아직 `insert_table_from_source` 함수는 없으므로 실행하면 ImportError):

```python
def _selftest_insert_table_from_source_default_style():
    """엑셀 원본자료 하나를 기본 스타일(배경색 없음, 헤더만 굵게)로 표 삽입하면,
    문서에 표가 실제로 생기고 헤더/데이터 값이 포함되는지 확인한다."""
    import os, tempfile, openpyxl
    from pyhwpx import Hwp
    from hwp_report import HwpReport

    test_dir = os.path.join(tempfile.gettempdir(), "_test_표원본")
    os.makedirs(test_dir, exist_ok=True)
    source_path = os.path.join(test_dir, "원본.xlsx")
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Sheet1"
    ws.append(["항목", "금액"])
    ws.append(["인건비", 1000000])
    ws.append(["운영비", 500000])
    wb.save(source_path)

    report_path = os.path.join(tempfile.gettempdir(), "_test_보고서_표.hwp")
    setup = Hwp(visible=False, new=True)
    setup.save_as(report_path)  # 빈 문서
    setup.quit()

    report = None
    try:
        report = HwpReport(report_path)
        result = insert_table_from_source(report, source_paths=[source_path], style="default")
        assert result["inserted"] is True, result

        final_text = report.get_text()
        assert "항목" in final_text, final_text
        assert "인건비" in final_text, final_text
        assert "1000000" in final_text, final_text
        print("insert_table_from_source(기본 스타일) 통과")
    finally:
        if report is not None:
            report.close(save=False)
        import shutil
        shutil.rmtree(test_dir)
        os.remove(report_path)


if __name__ == "__main__":
    _selftest_insert_table_from_source_default_style()
```

- [ ] Step 2: 실패 확인

Run: `python table_tool.py`
Expected: `NameError: name 'insert_table_from_source' is not defined` (또는 ImportError, 함수가 아직 없으므로)

- [ ] Step 3: 최소 구현

`table_tool.py` 맨 위에 함수 정의를 추가한다 (테스트 함수보다 위):

```python
"""table_tool.py — 원본자료(엑셀)를 읽어 한글 표로 변환해 커서 위치에 삽입하는 도구.
verify_tool.py/polish_tool.py와 나란한 세 번째 도구. PRD 14-2에 따라 이번 라운드는
표 스타일을 배경색(cell_fill)+헤더 굵게 조합으로만 한정한다(테두리 스타일 등은 다루지 않음)."""
import os

from hwp_report import HwpReport

# PRD 14-2: 표 스타일은 이 3종으로 한정 — 미리보기 팝업(chat_assistant.py)에 쓰이는
# 값과 반드시 일치해야 한다. "style" 파라미터는 이 dict의 key만 허용한다.
TABLE_STYLES = {
    "default": {"label": "기본형 (테두리만)", "header_fill": None},
    "gray_header": {"label": "회색 헤더", "header_fill": (217, 217, 217)},
    "blue_header": {"label": "파란 헤더", "header_fill": (198, 224, 241)},
}


def _first_excel_source(source_paths: list[str]) -> str | None:
    """PRD 14-2: 원본자료가 여러 개면 첫 번째 엑셀 파일을 기본으로 선택한다
    (여러 개 중 고르는 UI는 이번 라운드에 만들지 않음). source_paths에는 파일
    경로와 폴더 경로가 섞여 있을 수 있다(F12 1단계의 "+" 첨부 방식) — 폴더는
    이 함수에서는 건드리지 않고 개별 파일 경로만 검사한다(폴더까지 훑는 건
    source_reader.read_source_files의 책임이라 여기서 중복하지 않음).
    """
    for path in source_paths:
        if os.path.isfile(path) and path.lower().endswith((".xlsx", ".xls")):
            return path
    return None


def insert_table_from_source(report: HwpReport, source_paths: list[str], style: str) -> dict:
    """원본자료 중 첫 엑셀 파일을 읽어 커서 위치에 표로 삽입한다.

    style은 TABLE_STYLES의 key 중 하나여야 한다. header_fill이 있으면 표 삽입
    직후 커서가 표 안에 있는 상태를 이용해 헤더 행에 cell_fill()을 적용한다
    (table_from_data가 표 삽입 후 커서를 표 첫 열 시작 위치로 되돌려 놓는 것을
    이용 — pyhwpx 소스의 table_from_data 마지막 줄 TableColBegin()/Cancel() 참고).

    원본자료에 엑셀 파일이 하나도 없으면 문서를 건드리지 않고
    inserted=False를 반환한다.
    """
    if style not in TABLE_STYLES:
        raise ValueError(f"알 수 없는 표 스타일: {style}")

    excel_path = _first_excel_source(source_paths)
    if excel_path is None:
        return {"inserted": False, "reason": "원본자료 중 엑셀 파일이 없습니다"}

    report.hwp.table_from_data(excel_path, header=True, index=False, header_bold=True)

    header_fill = TABLE_STYLES[style]["header_fill"]
    if header_fill is not None:
        # table_from_data 종료 직후 커서는 표의 첫 번째 열(헤더 행 포함) 맨 위에
        # 있다 — 그 상태에서 첫 행(헤더)만 블록 선택 후 cell_fill을 적용한다.
        report.hwp.TableCellBlockExtend()
        report.hwp.Right()  # 다음 열로 이동하며 블록을 헤더 행 전체로 확장하려면
        # 실제로는 첫 행 전체를 선택해야 하므로, 아래처럼 행 단위 선택으로 교체한다.
        report.hwp.Cancel()
        report.hwp.TableColBegin()
        report.hwp.TableCellBlockRow()
        report.hwp.cell_fill(header_fill)
        report.hwp.Cancel()

    return {"inserted": True, "style": style, "source_file": excel_path}
```

- [ ] Step 4: 통과 확인

Run: `python table_tool.py`
Expected: `insert_table_from_source(기본 스타일) 통과` 출력, exit 0

만약 `TableCellBlockRow` 같은 메서드가 pyhwpx에 없다는 AttributeError가 나면(사전 조사에서 실제 존재를 확인하지 못한 메서드명이므로 가능성 있음), pyhwpx의 `site-packages/pyhwpx/core.py`에서 "표 전체 행 선택" 또는 "TableCellBlock"으로 검색해 실제 존재하는 메서드로 교체할 것. 대안으로 `report.hwp.hwp.HAction.Run("TableCellBlockRow")` 형태의 원시 HAction 호출도 pyhwpx 소스에 선례가 있는지 확인 후 사용 가능. **이 부분(헤더 행 선택 방법)이 이 태스크에서 가장 불확실한 지점이니, 실제 실행 결과를 반드시 확인할 것.**

- [ ] Step 5: 회색/파란 헤더 스타일도 실제로 다른 색을 적용하는지 확인하는 테스트 추가

`table_tool.py`에 추가:

```python
def _selftest_insert_table_from_source_colored_header():
    """gray_header 스타일을 쓰면 헤더 행 셀의 배경색이 실제로 바뀌는지 확인한다."""
    import os, tempfile, openpyxl
    from pyhwpx import Hwp
    from hwp_report import HwpReport

    test_dir = os.path.join(tempfile.gettempdir(), "_test_표스타일")
    os.makedirs(test_dir, exist_ok=True)
    source_path = os.path.join(test_dir, "원본.xlsx")
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Sheet1"
    ws.append(["항목", "금액"])
    ws.append(["인건비", 1000000])
    wb.save(source_path)

    report_path = os.path.join(tempfile.gettempdir(), "_test_보고서_표스타일.hwp")
    setup = Hwp(visible=False, new=True)
    setup.save_as(report_path)
    setup.quit()

    report = None
    try:
        report = HwpReport(report_path)
        result = insert_table_from_source(report, source_paths=[source_path], style="gray_header")
        assert result["inserted"] is True, result
        # 셀 배경색 자체를 읽는 API가 이 프로젝트엔 아직 없으므로, 여기서는
        # "예외 없이 완료됨 + 표 내용은 정상"만 확인한다 — 실제 색상이 눈으로
        # 봤을 때 맞는지는 아침에 스크린샷으로 사람이 확인해야 한다(PRD 14-6).
        final_text = report.get_text()
        assert "인건비" in final_text, final_text
        print("insert_table_from_source(gray_header) 통과 (배경색 육안 확인은 아침에 필요)")
    finally:
        if report is not None:
            report.close(save=False)
        import shutil
        shutil.rmtree(test_dir)
        os.remove(report_path)
```

`if __name__ == "__main__":` 블록에 이 함수 호출도 추가.

- [ ] Step 6: 통과 확인

Run: `python table_tool.py`
Expected: 두 self-test 모두 통과 메시지, exit 0

- [ ] Step 7: 커밋

```bash
git add table_tool.py
git commit -m "F12 2단계: table_tool.py 추가 - 원본자료를 표로 변환해 삽입 (기본/회색헤더/파란헤더 3종 스타일)"
```

---

### Task 2: numbering_tool.py — 번호서식 프리픽스 삽입 ✅ 완료 (커밋 02faa94, 스펙 검토 버그 수정 7e228e9)

**스펙 검토에서 발견된 버그 수정**: `insert_numbering_prefix()`가 `report.hwp.insert_text(prefix)`를 호출하기 전에 "선택 없음" 상태를 전제만 하고 실제로 보장하지 않아, 문서에 선택된 텍스트가 있는 채로 호출하면 `insert_text()`가 그 선택 영역을 통째로 지우고 대체해버리는 사고가 재현됨(직접 재현: "가나다라마바사"에서 "다라마"를 선택한 채 호출 → "다라마"가 사라지고 "가나1. 바사"가 됨). 회귀 테스트 `_selftest_insert_numbering_prefix_preserves_selection()`을 먼저 추가해 이 버그가 실패로 재현되는 것을 확인한 뒤(TDD), 삽입 직전에 `report.hwp.Cancel()`을 호출해 선택을 항상 해제하도록 고침. 직접 테스트로 확인한 바, `Cancel()` 이후 커서는 선택 영역의 시작점이 아니라 "끝점"에 남는다(예: "다라마" 선택 후 Cancel() → 커서는 "마"와 "바" 사이) — 수정 후 위 테스트는 `'가나다라마1. 바사'`를 만들어내며 통과함(선택했던 "다라마"는 그대로 보존, 프리픽스는 선택 끝점 바로 뒤에 삽입).

Files

- Create: `numbering_tool.py`

- [ ] Step 1: 실패하는 테스트 작성

```python
def _selftest_insert_numbering_prefix():
    """선택된 번호서식 프리픽스가 커서 위치에 삽입되는지 확인한다."""
    import os, tempfile
    from pyhwpx import Hwp
    from hwp_report import HwpReport

    report_path = os.path.join(tempfile.gettempdir(), "_test_보고서_번호.hwp")
    setup = Hwp(visible=False, new=True)
    setup.save_as(report_path)
    setup.quit()

    report = None
    try:
        report = HwpReport(report_path)
        result = insert_numbering_prefix(report, style="arabic_dot")
        assert result["inserted"] is True, result

        final_text = report.get_text()
        assert final_text.startswith("1. "), repr(final_text)
        print("insert_numbering_prefix(arabic_dot) 통과:", repr(final_text))
    finally:
        if report is not None:
            report.close(save=False)
        os.remove(report_path)


if __name__ == "__main__":
    _selftest_insert_numbering_prefix()
```

- [ ] Step 2: 실패 확인

Run: `python numbering_tool.py`
Expected: `NameError`

- [ ] Step 3: 최소 구현

파일 맨 위에 추가:

```python
"""numbering_tool.py — 번호서식(문단번호 스타일) 프리픽스를 커서 위치에 삽입하는 도구.

PRD 14-2: 한/글의 "진짜" 개요번호(자동 채번) 객체가 아니라, 텍스트 프리픽스를
직접 삽입하는 방식이다 — 순서가 바뀌어도 자동으로 재채번되지 않는다는 한계가
있음을 명시적으로 문서화해둔다(다음 라운드 재검토 대상)."""

from hwp_report import HwpReport

# PRD 14-2: 미리보기 팝업(chat_assistant.py)에 쓰이는 값과 반드시 일치해야 한다.
NUMBERING_STYLES = {
    "arabic_dot": {"label": "1. 아라비아숫자", "prefix": "1. "},
    "arabic_paren": {"label": "1) 괄호숫자", "prefix": "1) "},
    "double_paren": {"label": "(1) 이중괄호", "prefix": "(1) "},
    "circled": {"label": "① 원문자", "prefix": "① "},
}


def insert_numbering_prefix(report: HwpReport, style: str) -> dict:
    """style에 해당하는 프리픽스 문자열을 커서 위치에 삽입한다.
    선택된 텍스트가 있으면(F12 1단계 polish_to_formal_style과 같은 관례로)
    그 앞에 프리픽스만 덧붙이는 것이 아니라, 프리픽스를 그대로 삽입 지점에
    끼워넣는다 — insert_text()의 기본 동작(선택 영역 앞에 삽입되지 않고
    선택 영역을 대체함)과 혼동하지 않도록, 이 함수는 항상 "선택 없음" 상태를
    전제로 한다(미리보기 팝업 플로우에서 선택 여부를 별도로 다루지 않음,
    PRD 14-2 범위 밖)."""
    if style not in NUMBERING_STYLES:
        raise ValueError(f"알 수 없는 번호서식: {style}")

    prefix = NUMBERING_STYLES[style]["prefix"]
    applied = report.hwp.insert_text(prefix)
    return {"inserted": bool(applied), "style": style}
```

- [ ] Step 4: 통과 확인

Run: `python numbering_tool.py`
Expected: 통과 메시지, exit 0

- [ ] Step 5: 4종 스타일 전부 프리픽스가 올바른지 확인하는 테스트 추가

```python
def _selftest_all_styles_have_correct_prefix():
    """NUMBERING_STYLES의 4종 모두 실제로 올바른 프리픽스로 삽입되는지 확인."""
    import os, tempfile
    from pyhwpx import Hwp
    from hwp_report import HwpReport

    expected = {
        "arabic_dot": "1. ",
        "arabic_paren": "1) ",
        "double_paren": "(1) ",
        "circled": "① ",
    }
    for style, prefix in expected.items():
        report_path = os.path.join(tempfile.gettempdir(), f"_test_번호_{style}.hwp")
        setup = Hwp(visible=False, new=True)
        setup.save_as(report_path)
        setup.quit()

        report = None
        try:
            report = HwpReport(report_path)
            result = insert_numbering_prefix(report, style=style)
            assert result["inserted"] is True, result
            final_text = report.get_text()
            assert final_text.startswith(prefix), (style, repr(final_text))
        finally:
            if report is not None:
                report.close(save=False)
            os.remove(report_path)
    print("insert_numbering_prefix 4종 스타일 전부 통과")
```

`if __name__ == "__main__":` 블록에 호출 추가.

- [ ] Step 6: 통과 확인

Run: `python numbering_tool.py`
Expected: 두 테스트 모두 통과, exit 0

- [ ] Step 7: 커밋

```bash
git add numbering_tool.py
git commit -m "F12 2단계: numbering_tool.py 추가 - 번호서식 프리픽스 삽입 (4종 스타일)"
```

---

### Task 3: chat_assistant.py — 표 스타일 미리보기 팝업 + route_intent 등록 ✅ 완료 (커밋 65697f8, 스펙검토 반영 2b006ca, 코드품질 반영 aa8cb61)

**발견·수정된 버그 2건**:
1. `CTkToplevel`이 내부적으로 생성자에서 `self.withdraw()` 후 5ms 뒤 `deiconify()`하는 타이밍 때문에, 팝업이 topmost 메인 창 뒤에 가려지는 현상이 실제로 재현됨(win32gui로 실측) → `picker.after(50, lambda: (picker.lift(), picker.focus_force()))`로 수정. **Task 4의 번호서식 팝업도 이 패턴을 그대로 적용할 것.**
2. `choose()` 콜백(버튼 클릭 시 실행)이 `_on_submit()`의 try/except 범위 밖이라, `insert_table_from_source`가 예외를 던지면 채팅창에 아무 표시 없이 조용히 실패하던 문제 → 콜백 안에 자체 try/except 추가. **Task 4에도 동일하게 적용할 것.**

첫 스크린샷 시도에서 `ImageGrab.grab()`을 bbox 없이 호출해 사용자의 개인 브라우저 화면(강의 페이지 등)이 그대로 찍힌 사고가 있었음(오케스트레이터가 발견해 파일 삭제) — **이후 스크린샷은 반드시 bbox로 앱 창 좌표만 좁혀서 찍고, 저장 전 직접 열어 개인정보 없는지 확인할 것.**

Files

- Modify: `chat_assistant.py`

- [ ] Step 1: route_intent에 insert_table 도구 등록 — 실패하는 테스트 작성

`chat_assistant.py`의 `_selftest_route_intent()` 함수 안, 기존 테스트들 뒤에 추가:

```python
    # 6) F12 2단계: 새로 추가된 insert_table 도구가 키워드로 잡히는지 확인
    table_choice = route_intent("이 데이터로 표 만들어줘")
    assert table_choice == "insert_table", table_choice
    print("route_intent 통과 (표 삽입):", table_choice)
```

- [ ] Step 2: 실패 확인

Run: `python chat_assistant.py --selftest`
Expected: `AssertionError` (insert_table을 아직 아무도 반환 안 함 — route_intent가 None 또는 다른 값을 반환)

- [ ] Step 3: 최소 구현 — _TOOLS/_VERIFY_KEYWORDS 등 아래에 새 상수·라우팅 분기 추가

`chat_assistant.py`의 `_TOOLS` 리스트에 세 번째 항목 추가 (`polish_to_formal_style` 항목 뒤):

```python
    {
        "type": "function",
        "function": {
            "name": "insert_table",
            "description": "첨부된 원본자료(엑셀)를 표로 변환해 지금 열려있는 한글 문서의 커서 위치에 삽입한다",
            "parameters": {"type": "object", "properties": {}, "required": []},
        },
    },
```

`_POLISH_KEYWORDS` 아래에 새 키워드 목록 추가:

```python
_TABLE_KEYWORDS = ["표", "테이블", "표로", "표 만들어"]
```

`route_intent()` 함수 안, `if any(keyword in cleaned_message for keyword in _POLISH_KEYWORDS):` 블록 바로 다음(즉 `return None` 앞)에 추가:

```python
    if any(keyword in cleaned_message for keyword in _TABLE_KEYWORDS):
        return "insert_table"
```

- [ ] Step 4: 통과 확인

Run: `python chat_assistant.py --selftest`
Expected: 새 테스트 포함 전부 통과, exit 0

- [ ] Step 5: 표 스타일 미리보기 팝업 메서드 추가 (self-test는 팝업 콜백 함수를 직접 호출하는 방식으로 작성 — 실제 클릭 시뮬레이션은 CustomTkinter GUI 이벤트 루프가 필요해 무인 테스트로는 다루지 않음)

`ChatAssistant` 클래스에 `_attach_source` 메서드 뒤에 추가:

```python
    def _show_table_style_picker(self):
        """"표 만들어줘" 요청 시 뜨는 스타일 미리보기 팝업. 실제 한글 문서를
        미리 그리지 않고, CustomTkinter 위젯으로 각 스타일의 축소 모형을 직접
        그려서 보여준다(PRD 14-2 — 위젯 모형으로 "직접 보고 선택"이라는
        목표를 달성). 스타일을 고르면 즉시 insert_table_from_source를 실행하고
        팝업을 닫는다 — 팝업이 뜨기 전까지는 문서를 건드리지 않는다."""
        from table_tool import TABLE_STYLES, insert_table_from_source

        picker = ctk.CTkToplevel(self)
        picker.title("표 스타일 선택")
        picker.geometry("360x260")
        picker.attributes("-topmost", True)

        def choose(style_key):
            picker.destroy()
            result = insert_table_from_source(self.report, self.source_paths, style_key)
            if result["inserted"]:
                self._log(f"도우미: 표를 삽입했어요 ({TABLE_STYLES[style_key]['label']})")
            else:
                self._log(f"도우미: 표를 삽입하지 못했어요 - {result.get('reason', '알 수 없는 이유')}")

        for style_key, style_info in TABLE_STYLES.items():
            row = ctk.CTkFrame(picker)
            row.pack(pady=6, padx=10, fill="x")

            # 축소 모형: 2열짜리 미니 표를 Label 격자로 직접 그린다. 헤더 행에만
            # header_fill 색을 적용해 실제 표 삽입 결과와 시각적으로 대응시킨다.
            preview = ctk.CTkFrame(row)
            preview.pack(side="left", padx=(0, 10))
            header_color = style_info["header_fill"]
            header_hex = "#{:02x}{:02x}{:02x}".format(*header_color) if header_color else "#3a3a3a"
            for col, text in enumerate(["항목", "금액"]):
                ctk.CTkLabel(preview, text=text, fg_color=header_hex, width=50, height=20).grid(row=0, column=col, padx=1, pady=1)
            for col, text in enumerate(["인건비", "1,000,000"]):
                ctk.CTkLabel(preview, text=text, width=50, height=20).grid(row=1, column=col, padx=1, pady=1)

            ctk.CTkButton(row, text=style_info["label"], command=lambda k=style_key: choose(k)).pack(side="left", fill="x", expand=True)
```

- [ ] Step 6: _on_submit에서 insert_table 라우팅 시 팝업을 띄우도록 연결

`_on_submit()`의 `elif tool_name == "polish_to_formal_style":` 블록 뒤에 추가:

```python
            elif tool_name == "insert_table":
                self._show_table_style_picker()
```

- [ ] Step 7: 직접 실행으로 팝업이 뜨는지 확인 (사람이 볼 수 없는 야간이므로, 코드 레벨로 예외 없이 인스턴스화되는지만 확인)

아래 스크립트를 임시로 실행해 `_show_table_style_picker`가 예외 없이 위젯을 생성하는지 확인한다 (실제 파일 경로/보고서가 필요 없도록 mock report 사용):

```python
import customtkinter as ctk
from chat_assistant import ChatAssistant

app = ChatAssistant()
app.source_paths = []  # 원본자료 없이 팝업 UI 자체만 확인
app._show_table_style_picker()
app.after(500, app.destroy)  # 0.5초 후 자동 종료 (무인 실행이므로 창이 계속 떠있지 않게)
app.mainloop()
```

Run: 위 스크립트를 `python -c "..."` 또는 임시 파일로 실행
Expected: 예외 없이 종료(exit 0). 창이 실제로 잘 그려졌는지는 사람이 봐야 하므로, 스크린샷을 찍어 아침 확인용으로 남길 것(가능하면 `PIL.ImageGrab` 등으로 저장).

- [ ] Step 8: 커밋

```bash
git add chat_assistant.py
git commit -m "F12 2단계: insert_table 도구 등록 + 표 스타일 미리보기 팝업 추가"
```

---

### Task 4: chat_assistant.py — 번호서식 미리보기 팝업 + route_intent 등록 ✅ 완료 (커밋 4049781, 코드품질 리뷰 반영 e861f22)

Task 3에서 발견된 두 패턴(팝업 topmost 가려짐 방지, 콜백 예외처리)을 처음부터 반영해 구현함. 스펙검토에서 "팝업이 foreground 전환 안 됨" 의심을 5가지 시나리오로 재현 조사 — 실사용 시나리오(마우스 조작 전제) 8/8 성공, 문제는 비현실적 테스트 타이밍에서만 나타나는 아티팩트로 결론(코드 수정 불필요).

**코드품질 리뷰 반영**: `_show_table_style_picker`/`_show_numbering_style_picker`의 중복 보일러플레이트를 `_make_topmost_popup()`/`_run_tool_safely()` 공통 헬퍼로 추출. `_selftest_route_intent()`에 표/번호서식 동시 등장 tie-break 회귀 테스트 추가("표에 번호 매겨줘" → insert_table 우선, F12 1단계 Task8과 같은 맥락).

**다음 라운드 후보로 남긴 것**: `_attach_source()`의 `choice_window`(원본자료 첨부 선택창)도 topmost 가려짐 버그가 잠재할 수 있으나 이번 Task 4 범위 밖이라 손대지 않음 — `_make_topmost_popup()`으로 통일 검토 필요.

Files

- Modify: `chat_assistant.py`

- [ ] Step 1: route_intent에 insert_numbering 등록 — 실패하는 테스트 작성

`_selftest_route_intent()`에 추가:

```python
    # 7) F12 2단계: insert_numbering 도구가 키워드로 잡히는지 확인
    numbering_choice = route_intent("이 목록에 번호 매겨줘")
    assert numbering_choice == "insert_numbering", numbering_choice
    print("route_intent 통과 (번호서식):", numbering_choice)
```

- [ ] Step 2: 실패 확인

Run: `python chat_assistant.py --selftest`
Expected: `AssertionError`

- [ ] Step 3: 최소 구현

`_TOOLS`에 네 번째 항목 추가:

```python
    {
        "type": "function",
        "function": {
            "name": "insert_numbering",
            "description": "번호서식(1. / 1) / (1) / ① 등)을 미리보기에서 고른 뒤 커서 위치에 삽입한다",
            "parameters": {"type": "object", "properties": {}, "required": []},
        },
    },
```

`_TABLE_KEYWORDS` 아래에:

```python
_NUMBERING_KEYWORDS = ["번호", "번호매겨", "번호 매겨", "번호서식", "순번"]
```

`route_intent()`에서 `_TABLE_KEYWORDS` 체크 블록 뒤, `return None` 앞에:

```python
    if any(keyword in cleaned_message for keyword in _NUMBERING_KEYWORDS):
        return "insert_numbering"
```

- [ ] Step 4: 통과 확인

Run: `python chat_assistant.py --selftest`
Expected: 전부 통과, exit 0

- [ ] Step 5: 번호서식 미리보기 팝업 메서드 추가

`_show_table_style_picker` 메서드 뒤에 추가:

```python
    def _show_numbering_style_picker(self):
        """"번호 매겨줘" 요청 시 뜨는 서식 미리보기 팝업. 각 서식이 실제로
        어떻게 보이는지 예시 문장으로 직접 보여준 뒤, 고른 프리픽스를
        커서 위치에 삽입한다."""
        from numbering_tool import NUMBERING_STYLES, insert_numbering_prefix

        picker = ctk.CTkToplevel(self)
        picker.title("번호서식 선택")
        picker.geometry("280x200")
        picker.attributes("-topmost", True)

        def choose(style_key):
            picker.destroy()
            result = insert_numbering_prefix(self.report, style_key)
            if result["inserted"]:
                self._log(f"도우미: 번호를 삽입했어요 ({NUMBERING_STYLES[style_key]['label']})")
            else:
                self._log("도우미: 번호 삽입에 실패했어요.")

        for style_key, style_info in NUMBERING_STYLES.items():
            example = f"{style_info['prefix']}예시 항목입니다"
            ctk.CTkButton(
                picker, text=example, command=lambda k=style_key: choose(k)
            ).pack(pady=4, padx=10, fill="x")
```

- [ ] Step 6: _on_submit 연결

`elif tool_name == "insert_table":` 블록 뒤에 추가:

```python
            elif tool_name == "insert_numbering":
                self._show_numbering_style_picker()
```

- [ ] Step 7: 직접 실행 확인 (Task 3의 Step 7과 동일한 방식)

```python
import customtkinter as ctk
from chat_assistant import ChatAssistant

app = ChatAssistant()
app._show_numbering_style_picker()
app.after(500, app.destroy)
app.mainloop()
```

Run 후 exit 0 확인.

- [ ] Step 8: 커밋

```bash
git add chat_assistant.py
git commit -m "F12 2단계: insert_numbering 도구 등록 + 번호서식 미리보기 팝업 추가"
```

---

### Task 5: 통합 검증 + 계획 마무리

Files

- 없음(검증 전용 태스크)

- [ ] Step 1: 전체 self-test 재실행

```bash
python table_tool.py
python numbering_tool.py
python chat_assistant.py --selftest
python verify_tool.py
python source_reader.py
python polish_tool.py
python hwp_report.py
python window_layout.py
```

Expected: 전부 exit 0, 회귀 없음(F12 1단계 self-test들도 여전히 통과해야 함).

- [ ] Step 2: 실사용 시나리오 스크립트로 직접 검증 (사람이 볼 수 없는 야간이므로 스크린샷을 남길 것)

아래 흐름을 실제 pyhwpx로 재현하는 검증 스크립트를 작성해 실행한다:
1. 임시 엑셀 원본자료(항목/금액 2열, 데이터 2행) 생성
2. 임시 빈 hwp 문서 생성 → HwpReport로 열기
3. `insert_table_from_source(report, [엑셀경로], "blue_header")` 호출 → 문서에 표가 생겼는지 `get_text()`로 확인
4. 커서를 문서 맨 끝으로 이동(`report.hwp.MoveDocEnd()`) 후 `insert_numbering_prefix(report, "circled")` 호출 → 번호 프리픽스가 삽입됐는지 확인
5. 가능하면 `PIL.ImageGrab.grab()` 또는 `win32gui`로 문서 창을 캡처해 스크린샷 파일로 저장 (아침에 사용자가 육안 확인할 수 있도록)
6. `report.close(save=False)`

Expected: 예외 없이 완료, 스크린샷 파일 생성됨.

- [ ] Step 3: 계획 파일에 최종 결과 기록

이 계획 파일(`.tmp/implementation-plans/2026-09-02-f12-phase2-table-numbering.md`) 맨 아래에 "## 최종 검증 결과" 섹션을 추가해 다음을 기록:
- 각 태스크의 스펙 준수/코드 품질 리뷰 결과 요약
- Step 2 통합 검증의 실제 출력(표/번호 삽입 후 `get_text()` 결과 원문)
- 스크린샷 파일 경로
- PRD 14-2에서 이미 밝힌 축소 지점 재확인(사용자가 아침에 가장 먼저 확인해야 할 부분으로 강조)
- 발견된 버그·미결 이슈가 있으면 전부 기록(숨기지 말 것 — F12 1단계 Task 10의 관례를 그대로 따름)

- [ ] Step 4: 커밋

```bash
git add .tmp/implementation-plans/2026-09-02-f12-phase2-table-numbering.md
git commit -m "F12 2단계: 최종 통합 검증 결과 기록"
```

---

## 최종 검증 결과 (Task 5, 2026-09-02 야간)

### 1. Task 1~4 스펙 준수 / 코드 품질 리뷰 종합

| Task | 구현 | 스펙 검토 | 코드품질 검토 | 발견된 버그 |
|---|---|---|---|---|
| Task1 table_tool.py | 0cecfa8 | (구현에 포함) | 45033d1 (엑셀없음 회귀테스트 추가) | **크래시 버그**: 텍스트 선택 상태로 `insert_table_from_source()` 호출 시 `table_from_data()` 내부 `create_table()`이 COM 에러(`pywintypes.com_error`) 후 `ctrl.Properties` AttributeError로 이어져 100% 재현 크래시. `report.hwp.Cancel()`을 `table_from_data()` 호출 직전에 추가해 수정(44e4c83). TDD로 회귀 테스트 선행 확인. |
| Task2 numbering_tool.py | 02faa94 | 7e228e9에서 발견·수정 | (스펙검토에 통합) | **데이터 손실 버그**: 선택된 텍스트가 있는 상태로 `insert_numbering_prefix()` 호출 시 `insert_text()`가 선택 영역을 통째로 삭제·대체("다라마"가 사라지고 "가나1. 바사"가 됨). `Cancel()`을 삽입 직전에 추가해 수정. Cancel() 이후 커서는 선택 시작점이 아니라 끝점에 남는다는 것도 직접 테스트로 확인. |
| Task3 chat_assistant.py (표 팝업) | 65697f8 | 2b006ca | aa8cb61 | **버그 2건**: (1) `CTkToplevel`이 생성자 내부에서 `withdraw()` 후 5ms 뒤 `deiconify()`하는 타이밍 때문에 topmost 메인 창 뒤로 팝업이 가려짐 → `picker.after(50, lambda: (picker.lift(), picker.focus_force()))`로 수정. (2) 팝업 `choose()` 콜백이 `_on_submit()`의 try/except 범위 밖에 있어 `insert_table_from_source` 예외가 조용히 삼켜짐 → 콜백 내부에 자체 try/except 추가. |
| Task4 chat_assistant.py (번호 팝업) | 4049781 | (Task3 패턴 재사용, 5개 시나리오로 재현조사: 8/8 성공, 코드수정 불필요 결론) | e861f22 (팝업 보일러플레이트를 `_make_topmost_popup()`/`_run_tool_safely()`로 공통화, tie-break 회귀 테스트 추가) | 신규 버그 없음 — Task3의 두 패턴을 처음부터 반영해 구현. "다음 라운드 후보"로 `_attach_source()`의 `choice_window`도 동일한 topmost 취약점이 잠재할 수 있다고 명시(이번 범위 밖, 손 안 댐). |

### 2. Step 1: 전체 self-test 재실행 결과 — **일부 미완료(중요)**

| 파일 | 결과 |
|---|---|
| `table_tool.py` | ✅ exit 0 (3개 self-test 전부 통과) |
| `numbering_tool.py` | ✅ exit 0 (3개 self-test 전부 통과) |
| `verify_tool.py` | ✅ exit 0 |
| `source_reader.py` | ✅ exit 0 |
| `hwp_report.py` | ✅ exit 0 |
| `window_layout.py` | ✅ exit 0 |
| `chat_assistant.py --selftest` | ⚠️ **완주 실패** — 아래 상세 참고 |
| `polish_tool.py` | ⚠️ **완주 실패** — 아래 상세 참고 |

**정직하게 기록**: `chat_assistant.py --selftest`와 `polish_tool.py`는 오늘 밤 여러 번(각 4회 이상) 재시도했으나 끝까지 완주하는 exit 0을 한 번도 얻지 못했다. 원인을 추적한 결과:

- **1차 시도**(`chat_assistant.py --selftest`, 백그라운드 전환 없이 완주)에서는 실제로 끝까지 실행됐지만, 2번째 테스트케이스(`route_intent("오늘 날씨 어때")`는 `None`이어야 함)에서 `AssertionError: verify_numbers`로 실패했다. `route_intent()`는 로컬 LLM(qwen3.5:2b)의 도구호출을 먼저 시도하고(코드 주석에 "F11에서 실측된 도구호출 성공률 약 33%"라고 이미 명시됨), 실패했을 때만 키워드 안전망으로 넘어가는 구조인데, 이번엔 LLM이 "오늘 날씨 어때"에 대해 엉뚱하게 `verify_numbers` 도구를 호출해버렸다. 같은 문장을 격리된 새 프로세스에서 다시 호출하니 정상적으로 `None`이 나왔다 — **코드 버그가 아니라 로컬 LLM 도구호출의 알려진 비결정성**으로 판단.
- **2차~4차 시도**(전체 self-test를 처음부터 끝까지 다시 실행)에서는 매번 정확히 **7번째 ollama.chat() 호출**(`route_intent("이 데이터로 표 만들어줘")` 차례) 지점에서 프로세스가 완전히 멈췄다 — CPU 사용량이 몇 초 뒤로 전혀 늘지 않고(`TotalProcessorTime` 고정), 5~13분을 기다려도 응답이 없어 강제 종료함. 동일한 현상이 `chat_assistant.py`의 전체 실행 2회 + 별도로 작성한 진단 스크립트(9개 메시지를 순서대로 호출) 1회, 총 **3회 독립 재현**됐다. 반면 그 7번째 메시지만 새 프로세스에서 단독으로 호출하면 즉시 정상 응답(`insert_table`)이 왔다. 즉 **같은 프로세스에서 ollama.chat()을 여러 번 연속 호출하면 특정 시점(오늘 밤은 7번째)에서 멈추는 문제**로 보이며, 새로 추가된 라우팅 로직 자체의 결함이 아니라 이 환경(로컬 ollama 서버/클라이언트)의 세션 안정성 문제로 판단된다.
- `polish_tool.py`도 같은 양상이었다 — 4번 재시도 모두 **첫 번째** `ollama.chat()` 호출(`_generate_formal_style()` 내부, `polish_to_formal_style` 자체는 F12 1단계에 이미 있던 기존 코드로 이번 라운드에 손대지 않음)에서 멈췄다. 재시도 사이에 `ollama ps`로 서버 상태를 확인했고, 한 번은 모델이 `Stopping...` 상태였다가 곧 정상(`4 minutes from now`)으로 돌아왔음을 확인한 뒤 바로 재시도했는데도 다시 멈췄다. `ollama run qwen3.5:2b "1+1"` 같은 단순 CLI 호출은 그 사이사이 항상 정상적으로 빠르게 응답했다 — 즉 ollama 서버 자체는 살아있지만, 파이썬 `ollama` 패키지의 `chat()` 호출이 이 세션에서 간헐적으로 멈추는 것으로 보인다.
- 강제 종료한 프로세스는 모두 `Stop-Process -Force`로 정리했고, 남은 임시 `.hwp` 파일이나 좀비 `Hwp.exe`가 없는지 확인했다(각 self-test의 `finally` 블록이 `report.close()` + `os.remove()`를 수행하므로, 강제종료 시 정리가 안 될 수 있어 별도 확인함 — 확인 결과 파일 잠금 문제는 없었음).

**격리된 개별 호출로 대신 확인한 결과** (실제 프로덕션 코드는 그대로 두고, 각각 새 프로세스에서 `route_intent()`/`polish_to_formal_style()`을 직접 호출):

```
route_intent("숫자 검증해줘") → verify_numbers  ✅
route_intent("오늘 날씨 어때") → None  ✅ (격리 실행 시 정상 — 위 1차 시도의 실패는 LLM 비결정성)
route_intent("이거 검토 점검하고 오류 있는지 검사해줘") → verify_numbers  ✅
route_intent("파일 선택 확인했어") → None  ✅
route_intent("체크카드로 결제했어요") → None  ✅
route_intent("이 문장 공문서체로 다듬어줘") → polish_to_formal_style  ✅
route_intent("이 데이터로 표 만들어줘") → insert_table  ✅
route_intent("이 목록에 번호 매겨줘") → insert_numbering  ✅
route_intent("표에 번호 매겨줘") → insert_numbering  ⚠️ (아래 참고)
```

**새로 발견한 이슈(코드 결함은 아니지만 자기테스트의 근본적 한계)**: 계획서 Task4의 tie-break 회귀 테스트는 `route_intent("표에 번호 매겨줘")`가 항상 `insert_table`을 반환한다고 가정한다(`_TABLE_KEYWORDS`가 `_NUMBERING_KEYWORDS`보다 먼저 체크되므로). 그런데 이 우선순위는 **LLM 도구호출이 실패했을 때만** 적용되는 안전망이다. 오늘 밤 격리 실행에서는 LLM이 이 문장에 대해 스스로 도구호출에 성공해서 `insert_numbering`을 선택해버렸다 — 키워드 우선순위 로직 자체는 코드 그대로 정확하게 구현돼 있지만(직접 코드 읽기로 확인, `_TABLE_KEYWORDS` 체크가 `_NUMBERING_KEYWORDS` 체크보다 앞에 있음), **LLM이 도구호출에 성공하느냐 실패하느냐에 따라 이 self-test의 결과가 실행마다 달라질 수 있다**는 뜻이다. 새로 만든 버그는 아니고 F11부터 있던 "LLM 우선, 키워드는 안전망" 설계의 연장선이지만, 도구가 2개(F12 1단계)에서 4개(F12 2단계)로 늘면서 이런 식으로 tie-break 테스트가 흔들릴 여지도 함께 커졌다 — 다음 라운드에서 self-test를 "키워드 로직만 직접 호출해서 검증"하는 방식으로 분리하는 걸 검토할 만하다.

`polish_to_formal_style`은 `find()`/`SelectionMode` 감지까지는 격리 실행으로 정상 확인했으나, 정작 핵심인 LLM 생성 호출 자체가 격리 실행에서도 멈춰서 `applied=True` 끝까지의 결과는 오늘 밤 확인하지 못했다(이 함수는 F12 1단계 기존 코드로 이번 라운드에서 변경하지 않았다).

**결론**: table_tool.py/numbering_tool.py(이번 라운드 신규 코드)와 chat_assistant.py의 키워드 라우팅 로직은 실제 동작을 개별 확인했고 문제없다. 다만 `chat_assistant.py --selftest`와 `polish_tool.py`를 명령 그대로 실행해 마지막 줄까지 exit 0을 받는 것은 오늘 밤 끝내 성공하지 못했다 — 이는 로컬 ollama 세션의 환경적 불안정성으로 보이며, 아침에 사용자가 직접 `python chat_assistant.py --selftest`와 `python polish_tool.py`를 한 번 더 돌려서 정상 완주하는지 확인해주시는 게 가장 확실하다.

### 3. Step 2: 실사용 시나리오 통합 검증 — ✅ 성공 (ollama 미사용 경로라 위 불안정성의 영향 없음)

`table_tool.py`/`numbering_tool.py`는 ollama를 전혀 쓰지 않으므로(순수 pyhwpx 조작), 계획서가 요구한 통합 시나리오는 깨끗하게 한 번에 성공했다.

실행 흐름: 임시 엑셀(항목/금액 2열, 인건비/운영비 2행) 생성 → 빈 hwp 문서 생성 후 `HwpReport`로 열기 → `insert_table_from_source(report, [엑셀경로], "blue_header")` → `report.hwp.MoveDocEnd()` → `insert_numbering_prefix(report, "circled")` → `win32gui.GetWindowRect()`로 한글 창 bbox만 정확히 잘라 스크린샷 → `report.close(save=False)`.

실제 출력(원문 그대로):

```
insert_table_from_source 결과: {'inserted': True, 'style': 'blue_header', 'source_file': 'C:\\Users\\Public\\Documents\\ESTsoft\\CreatorTemp\\_test_task5_통합검증\\원본.xlsx'}
표 삽입 후 get_text():
'\r\n항목\r\n금액\r\n인건비\r\n1000000\r\n운영비\r\n500000\r\n\r\n'
insert_numbering_prefix 결과: {'inserted': True, 'style': 'circled'}
번호서식 삽입 후 get_text() (최종):
'\r\n항목\r\n금액\r\n인건비\r\n1000000\r\n운영비\r\n500000\r\n\r\n① '
한글 창 rect: (0, 0, 1440, 1032)
스크린샷 후보 저장: D:\보고서자동화\.worktrees\f12-phase2-table-numbering\.tmp\screenshots\_task5_raw_candidate.png (1440, 1032)
최종 문서 텍스트 (참고용 전체):
'\r\n항목\r\n금액\r\n인건비\r\n1000000\r\n운영비\r\n500000\r\n\r\n① '
report.close(save=False) 완료
```

표(파란 헤더)와 원문자 번호(①)가 실제 문서에 삽입됐고, `report.close(save=False)`도 정상 완료됨을 확인했다.

### 4. 스크린샷

- 경로: `.tmp/screenshots/task5_integration_table_and_numbering.png`
- 촬영 방법: `report.get_window_handle()` → `win32gui.GetWindowRect(hwnd)`로 **한글 창의 정확한 좌표(bbox)만** 얻어 `PIL.ImageGrab.grab(bbox=rect)`로 캡처 — 화면 전체나 임의 좌표를 찍지 않았다.
- **저장 전 직접 이미지를 열어 확인**: 한글 문서 창 하나만 보이고, 개인정보나 무관한 화면(브라우저 등)은 전혀 없음을 확인한 뒤 최종 파일명(`task5_integration_table_and_numbering.png`)으로 저장했다. 내용은 파란 헤더 표(항목/금액/인건비/1000000/운영비/500000) + 그 아래 원문자 "①" — 위 get_text() 출력과 일치.
- **(오케스트레이터 사후 확인, Task 5 완료 후)**: `.tmp/screenshots/task3_table_style_picker_zoom.png`(위에서 "왼쪽 끝에 다른 창 텍스트 잘림"으로 지적된 그 파일)를 직접 열어 재확인한 결과, 왼쪽 끝의 잘린 글자가 실제로는 사용자의 개인 브라우저 배경 창(다른 탭 제목의 일부)이었음을 확인했다 — 미커밋 상태였으므로 git 이력에는 남지 않고, 파일 자체를 삭제했다. 같은 시점에 Task 3 구현 중 처음 촬영된(bbox 없이 전체화면을 찍은) 스크린샷도 개인 브라우저 화면(강의 페이지 등)이 그대로 찍혀 별도로 이미 삭제된 바 있다 — 두 사고 모두 최종적으로 파일이 남지 않도록 정리됐다. 현재 저장소에 남아있는 스크린샷은 `task4_numbering_style_picker.png`, `task5_integration_table_and_numbering.png` 둘 뿐이며, 둘 다 오케스트레이터가 직접 열어 개인정보 없음을 확인했다.

### 5. PRD 14-2 축소 지점 재확인 (아침에 가장 먼저 확인해야 할 부분)

이번 라운드는 사용자의 원래 기대와 다를 수 있는 지점들이 있다 — PRD.md "## 14. F12 — 2단계"의 14-2에 이미 적혀있지만 다시 강조한다:

1. **차트 자동삽입은 이번에 없다.** matplotlib 의존성/무인 디버깅 리스크 때문에 제외, 표 삽입만 안정화했다.
2. **번호서식은 진짜 개요번호(자동 채번) 객체가 아니라 텍스트 프리픽스 삽입이다.** "1. ", "1) ", "(1) ", "① " 를 텍스트로 그냥 끼워넣는 방식이라, 항목 순서가 바뀌어도 번호가 자동으로 재정렬되지 않는다.
3. **표 스타일은 기본형/회색 헤더/파란 헤더 3종으로 한정**돼 있다(배경색+헤더 굵게 조합). 테두리 두께나 선 스타일 다양화는 없다.
4. **미리보기는 실제 한글 렌더링이 아니라 CustomTkinter 위젯으로 그린 축소 모형**이다(표는 격자 Label, 번호서식은 예시 문장 버튼).
5. 원본자료에 엑셀이 여러 개 첨부돼 있어도 **항상 첫 번째 엑셀 파일만** 표로 변환된다 — 고르는 UI는 없다.

### 6. 다음 라운드 후보로 남긴 것

- `_attach_source()`의 `choice_window`(원본자료 첨부 선택창)에 Task3에서 발견한 것과 같은 topmost 가려짐 버그가 잠재할 수 있음 — `_make_topmost_popup()` 공통 헬퍼로 통일 검토 필요(Task4 완료 기록에서 이미 명시, 이번 Task5에서도 손대지 않음).
- `_selftest_route_intent()`의 tie-break 테스트(및 유사한 LLM 우선 테스트들)가 LLM 도구호출 성공 여부에 따라 결과가 흔들릴 수 있음이 이번에 실측됨 — 다음 라운드에서 "키워드 안전망 로직만 별도 함수로 분리해 LLM 호출 없이 직접 테스트"하는 리팩터링을 검토할 만하다.
- 오늘 밤 반복 재현된 "동일 프로세스에서 ollama.chat() 연속 호출 시 특정 시점에서 멈춤" 현상의 근본 원인은 밝혀내지 못했다 — ollama 파이썬 클라이언트의 연결 재사용/keep-alive 문제인지, 로컬 서버 리소스(CPU 전용 추론, 2.4GB 모델) 한계인지 추가 조사가 필요하다.

### 7. 미해결 이슈 요약 (숨기지 않고 전부 기록)

- `chat_assistant.py --selftest`와 `polish_tool.py`를 명령 그대로 실행해 exit 0까지 완주시키는 데 오늘 밤 끝내 실패했다(각각 4회 이상 재시도, 매번 ollama.chat() 호출 중 무응답으로 멈춤 → 강제종료). 아침에 사용자가 직접 재실행해 확인이 필요하다.
- `.tmp/screenshots/task3_table_style_picker_zoom.png`가 git에 커밋되지 않은 채 남아있다(Task 3 범위라 이번엔 손대지 않음).
- ollama.chat() 연속 호출 시 멈추는 근본 원인 미상.

---

## 사용자 아침 확인 체크리스트 (구현 완료 후 이 섹션은 그대로 남겨둘 것)

아침에 확인하실 때 특히 아래 순서로 봐주시면 빠르게 판단하실 수 있습니다:

1. PRD.md "## 14. F12 — 2단계"의 14-2(비목표)를 먼저 읽어주세요 — 제가 임의로 축소한 지점이 전부 여기 있습니다.
2. 이 계획 파일 맨 아래 "## 최종 검증 결과"의 스크린샷을 확인해주세요.
3. 실제로 `python chat_assistant.py`를 실행해 "표 만들어줘"/"번호 매겨줘"를 직접 입력해보시면 가장 정확합니다.
4. 마음에 안 드는 부분이 있으면 F11→F12 전환 때처럼 편하게 말씀해주세요 — 이 브랜치(`feature/f12-phase2-table-numbering`)는 아직 master에 병합 안 했습니다.
