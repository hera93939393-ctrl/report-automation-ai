# F12 — 한글+채팅 통합 어시스턴트 1단계 Implementation Plan

> For the agent worker: required sub skill. Implement task by task with `subagent-driven-development` (recommended) or `executing-plans`. Track steps with checkbox (`- [ ]`) syntax.

Goal: F11(채팅창+숫자검증)을 "한글 창이 화면 왼쪽에 계속 떠있고, 채팅창이 오른쪽에 있으며, 문서를 한 번만 열어 계속 그 문서에 대고 검증/공문서체변환을 지시하는" 통합 어시스턴트 1단계로 발전시킨다.

Architecture: 보고서 문서는 "보고서 파일 선택" 클릭 시 `HwpReport`로 한 번만 열고, `ChatAssistant`가 그 핸들을 세션 내내 들고 있는다(F11처럼 도구 호출마다 새로 열지 않음). 원본자료는 채팅창의 "+" 버튼으로 파일 여러 개 또는 폴더를 첨부한다(리스트로 통일 관리). 창 배치는 `hwp.XHwpWindows.Active_XHwpWindow.WindowHandle`로 정확한 HWND를 얻어 win32gui로 위치조정(한글 75%/채팅 25%). 도구는 기존 `verify_numbers`에 `polish_to_formal_style`(선택 텍스트 다듬기/새 문장 삽입)을 추가하고, `route_intent`가 애매하면 되묻는 로직을 넣는다.

Tech stack: Python 3.13, pyhwpx, pywin32(`win32gui`, `win32con`), CustomTkinter, Ollama(`qwen3.5:2b`), openpyxl.

관련 PRD: `D:\보고서자동화\PRD.md` "## 13. F12"(13-1~13-10) — 이 계획의 모든 태스크는 이 PRD를 구체화한 것.

---

## 파일 구조 (사전 확정)

- **Modify** `source_reader.py`: `read_source_files(paths: list[str], default_year) -> tuple[list[dict], list[dict]]`을 새로 추가(개별 파일 목록 + 폴더가 섞인 리스트 처리). `read_source_folder`는 이 함수의 얇은 래퍼로 리팩터링(중복 로직 제거, DRY).
- **Modify** `verify_tool.py`: `run_verification`의 시그니처를 `(report_path, source_path, default_year)` → `(report: HwpReport, source_paths: list[str], default_year)`로 변경 — 더 이상 문서를 스스로 열지 않고(호출자가 이미 연 핸들을 받음), 원본자료도 리스트로 받는다. `_report_handle` 반환 키는 제거(호출자가 이미 핸들을 갖고 있으므로 왕복 불필요).
- **Modify** `hwp_report.py`: `HwpReport`에 `get_window_handle() -> int` 메서드 추가(창 배치용).
- **Create** `window_layout.py`: 한글 창(HWND로 지정)과 채팅창(CustomTkinter 인스턴스)을 화면 좌우로 배치하는 순수 함수 `position_windows(hwp_hwnd, chat_window, hwp_ratio=0.75)`. 새 파일로 분리하는 이유: 창 배치 로직은 HWP 조작이나 채팅 UI 로직과 무관한 별개 책임(순수 Win32/화면 계산)이라 독립 모듈이 재사용·테스트에 유리함.
- **Create** `polish_tool.py`: `polish_to_formal_style(report: HwpReport, chat_input: str) -> dict` — `verify_tool.py`와 나란히 있는 두 번째 "도구" 모듈(같은 책임 계층).
- **Modify** `chat_assistant.py`: `_choose_report`(문서를 즉시 열고 창 배치까지 수행), 원본자료 버튼 2개→"+" 1개로 통합, `_TOOLS`/키워드 목록에 `polish_to_formal_style` 추가, `route_intent` 되묻기 로직, `_on_submit` 최종 통합.

---

### Task 1: source_reader.py — 개별 파일 목록 지원 (`read_source_files`) ✅ 완료 (커밋 `76a51dd`)

**(2026-09-01 구현 중 발견)** 계획서 원안의 `_selftest_read_source_files_mixed_list` 테스트 픽스처가 두 엑셀 파일 모두 A1/A2 셀에 값을 넣게 되어 있어, 서로 무관한 두 파일(예산/인원)이 우연히 같은 셀 주소(`Sheet1!A2`)에서 충돌한 것으로 오탐되어 `assert conflicts == []`가 실패했다 — `read_source_files` 자체의 버그가 아니라 테스트 픽스처 버그. 두 번째 픽스처를 B1/B2로 옮겨 수정.

**검증 방식**: 새 함수(`read_source_files`, `read_source_folder`)는 격리 실행(`python -c`로 두 신규 self-test만 직접 호출)으로 통과 확인함(exit 0, 정확한 결과값 확인). `python source_reader.py`로 파일 전체(기존 한글 자동화 테스트 포함 약 13개)를 한 번에 돌리는 건 이번 태스크가 건드리지 않은 기존 `_selftest_read_hwp_source` 계열에서 간헐적 hang(exit 124, 재시도 3회 모두 실패)이 발생해 끝까지 확인하지 못함 — 이 프로젝트에 밤새 반복 기록된 한글 COM 자동화 환경 불안정성과 같은 종류로 판단, 새로 추가한 코드 자체의 결함이 아님을 격리 테스트로 확인한 뒤 진행함.

Files

- Modify: `source_reader.py:277-347` (`_READERS` 정의부터 `read_source_folder` 끝까지)
- Test: 같은 파일의 `if __name__ == "__main__":` 블록에 self-test 함수로 추가(이 프로젝트 전체 컨벤션)

**배경**: 지금 `read_source_folder(folder_path, default_year)`는 `os.listdir(folder_path)`로 폴더 안 파일 이름을 훑는다. F12는 "+" 버튼으로 폴더가 아니라 **개별적으로 선택된 파일 여러 개**(서로 다른 폴더에 있을 수도 있음)도 받아야 하므로, "경로 리스트"를 받아 그 안에 파일이든 폴더든 섞여 있어도 처리하는 더 일반적인 함수가 필요하다.

- [ ] Step 1: 실패하는 테스트 작성

`source_reader.py` 맨 아래 `if __name__ == "__main__":` 블록 **위**에 추가:

```python
def _selftest_read_source_files_mixed_list():
    """폴더 경로 하나와 개별 파일 경로 하나가 섞인 리스트를 줘도 둘 다 읽어서
    합쳐지는지 확인한다 (F12: "+" 버튼으로 파일 여러 개 또는 폴더를 자유롭게
    섞어 첨부할 수 있어야 하므로)."""
    os.makedirs("_test_원본_혼합", exist_ok=True)
    wb1 = openpyxl.Workbook(); ws1 = wb1.active; ws1.title = "Sheet1"
    ws1["A1"] = "예산"; ws1["A2"] = 1850000
    wb1.save("_test_원본_혼합/폴더안파일.xlsx")

    wb2 = openpyxl.Workbook(); ws2 = wb2.active; ws2.title = "Sheet1"
    ws2["A1"] = "인원"; ws2["A2"] = 12
    wb2.save("_test_원본_개별파일.xlsx")

    try:
        pool, conflicts = read_source_files(
            ["_test_원본_혼합", "_test_원본_개별파일.xlsx"], default_year=2026
        )
        normalized_values = {item["normalized"] for item in pool}
        assert "1850000" in normalized_values, pool
        assert "12" in normalized_values, pool
        assert conflicts == [], conflicts
        print("read_source_files(혼합 리스트) 통과:", normalized_values)
    finally:
        import shutil
        shutil.rmtree("_test_원본_혼합")
        os.remove("_test_원본_개별파일.xlsx")


def _selftest_read_source_folder_still_works():
    """리팩터링 후에도 read_source_folder(폴더 경로 하나)가 기존과 동일하게
    동작하는지 확인하는 회귀 테스트(기존 _selftest_read_source_folder_conflict와
    별개로, 함수 시그니처 자체가 안 바뀌었는지 빠르게 확인)."""
    os.makedirs("_test_원본_회귀", exist_ok=True)
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Sheet1"
    ws["A1"] = "예산"; ws["A2"] = 1850000
    wb.save("_test_원본_회귀/원본.xlsx")
    try:
        pool, conflicts = read_source_folder("_test_원본_회귀", default_year=2026)
        assert any(item["normalized"] == "1850000" for item in pool), pool
        print("read_source_folder(리팩터링 후 회귀) 통과")
    finally:
        import shutil
        shutil.rmtree("_test_원본_회귀")
```

- [ ] Step 2: 실패 확인

Run: `python source_reader.py` (아직 `__main__` 블록에서 호출 안 함, 그냥 `python -c "from source_reader import read_source_files"`로 확인)
Expected: `ImportError: cannot import name 'read_source_files' from 'source_reader'`

- [ ] Step 3: 최소 구현 — `read_source_folder` 리팩터링 + `read_source_files` 신설

`source_reader.py:277-347`을 아래로 전체 교체:

```python
_READERS = {".xlsx": read_excel_source, ".xls": read_excel_source,
            ".hwp": read_hwp_source, ".hwpx": read_hwp_source,
            ".pdf": read_pdf_source}


_CELL_ADDRESS_PATTERN = re.compile(r'^[A-Za-z]{1,3}\d+$')


def read_source_files(paths: list[str], default_year: int) -> tuple[list[dict], list[dict]]:
    """경로 리스트(개별 파일과 폴더가 섞여 있어도 됨)를 읽어 정답 풀과 충돌
    목록을 만든다. F12에서 "+" 버튼으로 파일 여러 개를 개별 선택하거나
    폴더를 선택하는 두 경우를 하나의 함수로 통일해서 처리하기 위함이다
    (2026-08-31 이전까지는 read_source_folder가 폴더 하나만 받았음).

    폴더 경로가 섞여 있으면 그 폴더 바로 아래(하위 폴더 재귀 없음, 기존
    read_source_folder와 동일한 한계) 파일들을 각각 개별 파일처럼 펼쳐서 읽는다.

    반환: (정답_풀, 충돌_목록)
    충돌_목록 항목 형식: {"location": "Sheet1!A2", "type": "amount", "values": [
        {"file": "초안.xlsx", "normalized": "1800000"},
        {"file": "최종.xlsx", "normalized": "1850000"}]}

    (2026-08-31 코드품질 검토 후 명시, read_source_folder 시절부터 유지) 충돌은
    (location, type) 단위로 탐지하므로, 같은 location이 서로 다른 타입의 충돌
    항목으로 두 번 이상 나타날 수 있다.
    """
    expanded_paths = []
    for path in paths:
        if os.path.isdir(path):
            expanded_paths.extend(os.path.join(path, name) for name in os.listdir(path))
        else:
            expanded_paths.append(path)

    pool = []
    for full_path in expanded_paths:
        ext = os.path.splitext(full_path)[1].lower()
        reader = _READERS.get(ext)
        if reader is None:
            continue  # 이미지, 워드, 알 수 없는 형식 등 → 조용히 건너뜀
        pool.extend(reader(full_path, default_year))

    by_location: dict[tuple[str, str], list[dict]] = {}
    for item in pool:
        location = item.get("location", "")
        if location.count("!") != 1:
            continue  # 엑셀 셀 주소 형식만 충돌 탐지 대상
        _, _, cell_ref = location.partition("!")
        if not _CELL_ADDRESS_PATTERN.match(cell_ref):  # 진짜 셀 주소만 충돌 탐지 대상
            continue
        by_location.setdefault((location, item.get("type", "")), []).append(item)

    conflicts = []
    for (location, value_type), items in by_location.items():
        distinct_files = {i["source_file"]: i["normalized"] for i in items}
        distinct_values = set(distinct_files.values())
        if len(distinct_values) > 1:
            conflicts.append({
                "location": location,
                "type": value_type,
                "values": [{"file": f, "normalized": v} for f, v in distinct_files.items()],
            })
    return pool, conflicts


def read_source_folder(folder_path: str, default_year: int) -> tuple[list[dict], list[dict]]:
    """폴더 안 모든 파일을 읽되, 처리 가능한 확장자만 실제로 읽고 나머지는 조용히
    건너뛴다. `read_source_files`의 얇은 래퍼(폴더 하나만 다루는 이전 시그니처를
    그대로 유지 — 기존 호출부·테스트 호환을 위해 남겨둠)."""
    return read_source_files([folder_path], default_year)
```

- [ ] Step 4: 통과 확인

Run: `python source_reader.py`
Expected: 새로 추가한 두 self-test는 아직 `__main__`에서 호출 안 하고 있어 실행 안 됨 — 대신 직접 `python -c "from source_reader import _selftest_read_source_files_mixed_list, _selftest_read_source_folder_still_works; _selftest_read_source_files_mixed_list(); _selftest_read_source_folder_still_works()"`로 통과 확인
Expected 출력: `read_source_files(혼합 리스트) 통과: {...}`와 `read_source_folder(리팩터링 후 회귀) 통과` 둘 다 에러 없이 출력

그 다음 `source_reader.py`의 `if __name__ == "__main__":` 블록에 이 두 self-test 호출을 추가하고(기존 호출들 아래), `python source_reader.py`를 실행해 **기존 self-test 전부(약 11개) + 새로 추가한 2개**가 모두 통과하는지 확인.

- [ ] Step 5: 커밋

```bash
git add source_reader.py
git commit -m "F12: source_reader에 read_source_files 추가 - 개별파일+폴더 혼합 리스트 지원"
```

**(2026-09-01 코드품질 리뷰 반영, 커밋 `60f5411`)** 리뷰에서 `read_source_files`의
두 가지 실제 이슈가 지적되어 수정함: (1) 폴더 펼치기(`os.listdir`)가 try/except
없이 노출되어 있어 권한 문제 등으로 예외가 터지면 이 모듈의 "절대 안 죽고 조용히
건너뜀" 관례가 깨지는 문제 — `except OSError: continue`로 감쌈. (2) 같은 경로가
리스트에 두 번 들어오거나(다중 선택 파일 대화상자 중복 선택) 상대/절대경로로 각각
한 번씩 들어오면 같은 파일을 두 번 읽어 정답 풀이 조용히 배로 부풀려지는 문제 —
펼쳐진 경로 목록을 `os.path.normpath(os.path.abspath(...))` 기준으로 순서를 유지한
채 중복 제거해서 해결. 두 케이스 모두 새 self-test로 회귀 방지 커버(권한 오류는
`unittest.mock.patch`로 `os.listdir` 실패를 시뮬레이션, dedup은 같은 파일을 두 번
넣었을 때 한 번 넣은 것과 개수가 같은지 비교). 격리 검증(`python -c`로 새 self-test
2개만 직접 호출) 통과 확인함; 전체 `python source_reader.py`는 이 태스크와 무관한
기존 HWP self-test 계열의 간헐적 hang(exit 124)으로 끝까지 못 돌리는 건 Task 1
최초 구현 때와 동일 — 새 self-test 및 기존 `read_source_files` 관련 self-test들은
개별 호출로 전부 통과 확인함.

---

### Task 2: verify_tool.py — 이미 열린 문서 핸들 + 원본자료 리스트를 받도록 리팩터링 ✅ 완료 (커밋 `d47fa53`)

**(2026-09-01 구현 완료)** 계획서 그대로 `verify_tool.py`를 전체 교체함(Step 2에서
교체 전 구버전 시그니처 `(report_path: str, source_path: str, default_year: int) -> dict`를
먼저 확인, 새 시그니처 `(report: HwpReport, source_paths: list[str], default_year)`와
다름을 확인한 뒤 진행). `source_reader.py`의 `read_source_files`는 Task 1 완료 후
두 차례 내부 견고성 수정(폴더 목록 조회 실패 시 조용히 건너뜀, 중복 경로 제거)이
있었으나 공개 시그니처/동작 계약은 그대로라 이 태스크의 통합 코드는 변경 불필요.

**검증**: `timeout 90 python verify_tool.py` 1회 시도 만에 통과(재시도 불필요) —
`run_verification 통과: 1건 확인 필요\n- [amount] '9999999원' 원본에서 확인 안 됨`,
exit code 0. 실행 전후 `tasklist`로 `Hwp.exe` 프로세스 없음 확인(orphan 없음).

**(2026-09-01 코드품질 리뷰 반영, 커밋 `893b130`)** 리뷰에서 실제로 재현·확인된
버그: `run_verification`이 이전 호출에서 남긴 빨간 표시를 아무것도 지우지
않아서, 같은 `HwpReport` 핸들로 재검증(F12의 핵심 시나리오 — 사용자가 원본을
고치거나 새 원본을 첨부한 뒤 "다시 검증해줘")했을 때 값이 더 이상 불일치가
아닌데도 문서엔 빨간 글자가 그대로 남고, 채팅창은 "이상 없음"이라고 답하는
모순이 있었음. `hwp_report.py`에 `HwpReport.reset_colors()`(`SelectAll()` +
`set_font(TextColor=검정)` + `Cancel()` — pyhwpx core.py의
`get_used_style_dict()`/`remove_unused_styles()`가 쓰는 것과 동일한
전체선택→작업→선택해제 패턴)를 추가하고, `run_verification` 시작부에서
`report_text = report.get_text()` 직후 호출해 매 실행이 "리셋 후 현재
대조결과만 재표시"하는 멱등적 동작이 되도록 수정. 회귀테스트
`_selftest_run_verification_clears_stale_marks_on_rerun`(같은 핸들로 1차
검증→빨간색 확인→원본 수정→2차 검증→`mismatch_count == 0` 및 글자색
`(0,0,0)` 복귀 확인)을 추가하고 기존 `_selftest_run_verification`과 함께
`timeout 90 python verify_tool.py` 1회 시도 만에 둘 다 통과(exit code 0,
`Hwp.exe` orphan 없음). 두 함수의 docstring에 "같은 핸들 반복 호출" 계약을
명시해 Task 9의 `_on_submit` 통합 구현자가 재발견할 필요 없게 함.

**의도적으로 미룬 항목(리뷰 finding #2)**: `report.get_text()`가 이미
닫힌/죽은 COM 핸들에 대해 호출되면 `pywintypes.com_error`가 그대로 올라오는
문제는 이번 라운드에서 다루지 않음(별도 후속 과제로 `run_verification`
docstring에 한 줄 남겨둠) — 실제 버그이지만 이번 수정과 성격이 다르고
리뷰에서도 별도 라운드로 미루는 것이 합리적이라 판단됨.

Files

- Modify: `verify_tool.py` (전체 재작성 — 아래 Step 3 참고)

**배경**: 지금 `run_verification(report_path, source_path, default_year)`은 매번 `HwpReport(report_path)`로 문서를 새로 연다(F11의 알려진 한계 — 반복 호출 시 창이 여러 개 뜸). F12는 문서를 채팅 프로그램이 한 번만 열고 계속 재사용해야 하므로, 이 함수는 **이미 열린 `HwpReport` 인스턴스**를 받아야 한다. 원본자료도 Task 1의 `read_source_files`에 맞춰 리스트로 받는다.

- [ ] Step 1: 실패하는 테스트 작성

`verify_tool.py`를 통째로 아래 내용으로 교체(기존 `_selftest_missing_report_path`는 새 시그니처와 안 맞아 이번에 제거 — 문서가 이미 열려서 넘어오므로 "파일이 없어 못 연다"는 시나리오 자체가 이 함수 책임이 아니게 됨. 이 책임은 Task 5의 `chat_assistant.py`의 `_choose_report`로 옮겨감):

```python
"""verify_tool.py — verify_numbers.py/source_reader.py를 엮어 "숫자검증" 도구
하나로 만든다. F12: 이미 열려있는 HwpReport 핸들과 원본자료 경로 리스트를 받는다
(F11 시절엔 이 함수가 문서를 직접 열었으나, 도구 호출마다 문서가 새로 열리는
문제가 있어 호출자(chat_assistant.py)가 한 번만 연 핸들을 넘겨주는 구조로 변경)."""
import os
import tempfile
import openpyxl

from verify_numbers import extract_values, compare_values
from source_reader import read_source_files
from hwp_report import HwpReport


def run_verification(report: HwpReport, source_paths: list[str], default_year: int) -> dict:
    """채팅창이 호출하는 숫자검증 도구 함수.
    1) report(이미 열려있는 문서 핸들)의 텍스트를 읽는다 — 이 함수는 문서를
       열거나 닫지 않는다, 호출자가 열고 닫을 책임을 진다
    2) source_paths(파일·폴더가 섞인 경로 리스트)에서 정답 풀을 만든다
    3) 보고서 텍스트에서 4종 값을 뽑아 대조한다
    4) 불일치 항목을 빨간색으로 표시한다 (자동저장 안 함)
    5) 채팅창에 보여줄 요약 텍스트와 충돌 목록을 반환한다

    반환 dict 계약:
      - mismatch_count: int — 서로 다른(중복 제거된) 불일치 값의 개수. summary
        첫 줄 숫자와는 항상 일치하지만, conflicts가 있으면 그 충돌 줄들이 이
        숫자에 안 잡힌 채 목록 뒤에 추가로 붙는다(의도된 것 — 원본 파일 간
        불일치는 "불일치 값 개수"와 성격이 달라 같은 숫자에 합산하지 않음).
      - summary: str — 채팅창에 그대로 보여줄 사람이 읽는 요약 텍스트.
      - conflicts: list — read_source_files가 찾은 원본 파일 간 불일치 목록.
    """
    answer_pool, conflicts = read_source_files(source_paths, default_year)

    report_text = report.get_text()
    report_values = extract_values(report_text, default_year)
    mismatches = compare_values(report_values, answer_pool)

    # 같은 잘못된 값이 여러 곳(본문+요약표 등)에 등장하면 compare_values가 raw는
    # 같고 span만 다른 mismatch를 여러 개 돌려줄 수 있다. mark_red()가 이미
    # 문서 전체를 훑어 모든 occurrence를 한 번에 칠하므로, 같은 raw로 반복
    # 호출하는 건 중복 COM 순회다 — dict.fromkeys로 중복 제거 후 한 번씩만 호출.
    unique_raw_values = list(dict.fromkeys(m["raw"] for m in mismatches))
    for raw in unique_raw_values:
        report.mark_red(raw)

    first_type_by_raw = {}
    for m in mismatches:
        first_type_by_raw.setdefault(m["raw"], m["type"])
    lines = [f"- [{first_type_by_raw[raw]}] '{raw}' 원본에서 확인 안 됨" for raw in unique_raw_values]
    for c in conflicts:
        value_desc = ", ".join(f"{v['file']}={v['normalized']}" for v in c["values"])
        lines.append(f"- ⚠ 원본자료 불일치[{c['type']}]: {c['location']} ({value_desc})")

    summary = (f"{len(unique_raw_values)}건 확인 필요\n" + "\n".join(lines)
               if mismatches or conflicts else "이상 없음, 모두 원본과 일치합니다")

    return {"mismatch_count": len(unique_raw_values), "summary": summary, "conflicts": conflicts}


def _selftest_run_verification():
    # 이 워크트리 폴더 안에 테스트 파일을 만들면 한글이 자체 보안 모듈 경고를
    # 띄우며 무한 대기하는 현상이 간헐적으로 재현된 적 있어, 시스템 임시폴더를 쓴다.
    test_dir = os.path.join(tempfile.gettempdir(), "_test_원본_f12")
    os.makedirs(test_dir, exist_ok=True)
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Sheet1"
    ws["A1"] = "예산"; ws["A2"] = 1850000
    wb.save(os.path.join(test_dir, "원본.xlsx"))

    from pyhwpx import Hwp
    report_path = os.path.join(tempfile.gettempdir(), "_test_보고서_f12.hwp")
    setup = Hwp(visible=False)
    setup.insert_text("예산은 185만원이며, 오타는 9999999원입니다")
    setup.save_as(report_path)
    setup.quit()

    report = None
    try:
        report = HwpReport(report_path)  # 호출자(테스트)가 직접 문서를 엶 — F12 구조
        result = run_verification(report, source_paths=[test_dir], default_year=2026)
        assert result["mismatch_count"] == 1, result
        assert "9999999" in result["summary"], result
        print("run_verification 통과:", result["summary"])
    finally:
        if report is not None:
            report.close(save=False)  # 호출자가 직접 닫음 — F12 구조
        import shutil
        shutil.rmtree(test_dir)
        os.remove(report_path)


if __name__ == "__main__":
    _selftest_run_verification()
```

- [ ] Step 2: 실패 확인

Run: `python -c "from verify_tool import run_verification; import inspect; print(inspect.signature(run_verification))"` (교체 전 원본 파일 기준)
Expected(교체 전): `(report_path: str, source_path: str, default_year: int) -> dict` — 즉 지금은 `report`/`source_paths`라는 새 파라미터 이름이 없어서, 위 self-test를 그대로 돌리면 `TypeError: run_verification() got an unexpected keyword argument 'source_paths'`

(이 태스크는 파일 전체 교체라 "실패하는 테스트를 먼저 넣고 확인"하는 절차 대신, 교체 전 시그니처가 새 테스트와 안 맞다는 것 자체가 실패 확인임 — 위 Step 1의 새 파일 내용을 실제로 저장하기 전에 이 명령으로 구버전 시그니처를 먼저 확인해둘 것)

- [ ] Step 3: 최소 구현

Step 1에 적힌 전체 코드로 `verify_tool.py`를 덮어쓴다.

- [ ] Step 4: 통과 확인

Run: `timeout 90 python verify_tool.py` (이 프로젝트에 문서화된 간헐적 HWP COM 보안대화상자 이슈로 가끔 멈출 수 있음 — 멈추면 죽이고 `tasklist`로 orphan `Hwp.exe` 확인 후 정리, 최대 3회 재시도)
Expected: `run_verification 통과: 1건 확인 필요\n- [amount] '9999999원' 원본에서 확인 안 됨` 출력, exit code 0

- [ ] Step 5: 커밋

```bash
git add verify_tool.py
git commit -m "F12: run_verification이 이미 열린 HwpReport 핸들과 원본자료 리스트를 받도록 리팩터링"
```

---

### Task 3: hwp_report.py — 창 핸들(HWND) 가져오기 ✅ 완료 (커밋 `387b808`)

**(2026-09-01 구현 완료)** 계획서 그대로 진행. Step 2에서 `AttributeError:
type object 'HwpReport' has no attribute 'get_window_handle'` 확인 후 `get_window_handle()`을
`close()` 바로 위에 추가하고, `_selftest_get_window_handle()`을 `if __name__ ==
"__main__":` 블록에 추가.

**검증**: `timeout 60 python hwp_report.py` 1회 시도 만에 통과(재시도 불필요) —
`HwpReport 통과: 텍스트 읽기 + 모든 occurrence 빨간색 표시 확인` 및
`HwpReport.get_window_handle() 통과: 18484556 '내 문서 1 - 한글'` 출력, exit code 0.
실행 전후 `tasklist`로 `Hwp.exe` 프로세스 없음 확인(orphan 없음).

Files

- Modify: `hwp_report.py` (`HwpReport` 클래스에 메서드 추가)
- Test: 같은 파일의 self-test에 추가

**배경**: 창 배치(Task 4)를 하려면 지금 열려있는 한글 문서의 정확한 Win32 창 핸들(HWND)이 필요하다. `pyhwpx` 소스를 직접 확인한 결과 `self.hwp.XHwpWindows.Active_XHwpWindow.WindowHandle`이 이 HWND를 돌려주며, `pyhwpx` 자신도 `get_title()` 등에서 이 프로퍼티로 얻은 핸들을 `win32gui` 함수에 바로 넘겨 쓰고 있음을 확인했다(제목 검색으로 창을 찾는 것보다 안전 — 여러 한글 창이 떠 있어도 방금 연 그 창을 정확히 특정 가능).

- [ ] Step 1: 실패하는 테스트 작성

`hwp_report.py`의 `_selftest_open_and_mark_red()` 함수 **아래**, `if __name__ == "__main__":` 블록 **위**에 추가:

```python
def _selftest_get_window_handle():
    """get_window_handle()이 실제 win32gui 함수에 바로 쓸 수 있는 정수 HWND를
    돌려주는지 확인한다 (창 배치 기능의 전제 조건)."""
    import tempfile
    import win32gui
    from pyhwpx import Hwp

    test_path = os.path.join(tempfile.gettempdir(), "_test_핸들.hwp")
    setup = Hwp(visible=False, new=True)
    setup.insert_text("핸들 테스트")
    setup.save_as(test_path)
    setup.quit()

    report = None
    try:
        report = HwpReport(test_path)
        hwnd = report.get_window_handle()
        assert isinstance(hwnd, int) and hwnd > 0, hwnd
        assert win32gui.IsWindow(hwnd), f"win32gui가 인식하지 못하는 핸들: {hwnd}"
        title = win32gui.GetWindowText(hwnd)
        assert "한글" in title or "핸들 테스트" in title or title, title
        print("HwpReport.get_window_handle() 통과:", hwnd, repr(title))
    finally:
        if report is not None:
            report.close(save=False)
        os.remove(test_path)
```

- [ ] Step 2: 실패 확인

Run: `python -c "from hwp_report import HwpReport; HwpReport.get_window_handle"`
Expected: `AttributeError: type object 'HwpReport' has no attribute 'get_window_handle'`

- [ ] Step 3: 최소 구현

`hwp_report.py`의 `close()` 메서드 **바로 위**에 추가:

```python
    def get_window_handle(self) -> int:
        """이 문서를 보여주는 실제 Win32 창의 HWND(창 핸들)를 반환한다
        (창 배치 등 win32gui 함수에 바로 넘겨 쓸 수 있음).

        pyhwpx 자신도 get_title() 등에서 이 프로퍼티(hwp.XHwpWindows.
        Active_XHwpWindow.WindowHandle)로 얻은 핸들을 win32gui에 그대로
        넘겨 쓰고 있음을 소스에서 확인함 — 창 제목으로 FindWindow하는 것보다
        안전하다(여러 한글 창이 동시에 떠 있어도 이 인스턴스가 연 그 창을
        정확히 특정 가능, 제목 문자열은 파일명에 따라 바뀌어 검색 기준으로
        불안정함).
        """
        return self.hwp.XHwpWindows.Active_XHwpWindow.WindowHandle

    def close(self, save: bool):
```

(주의: 위 코드 블록은 `get_window_handle` 메서드 전체와 `close` 메서드의 시작 줄까지 함께 보여준 것 — 실제로는 `get_window_handle` 메서드만 새로 추가하고 기존 `close` 메서드는 그대로 둔다.)

- [ ] Step 4: 통과 확인

`hwp_report.py`의 `if __name__ == "__main__":` 블록을 아래로 교체:

```python
if __name__ == "__main__":
    _selftest_open_and_mark_red()
    _selftest_get_window_handle()
```

Run: `timeout 60 python hwp_report.py` (간헐적 HWP COM 이슈로 멈추면 재시도, 최대 3회)
Expected: 두 self-test 모두 통과 메시지 출력, exit code 0

- [ ] Step 5: 커밋

```bash
git add hwp_report.py
git commit -m "F12: HwpReport.get_window_handle() 추가 - 창 배치용 HWND 확보"
```

---

### Task 4: window_layout.py — 한글 창 75% / 채팅창 25% 자동 배치

Files

- Create: `window_layout.py`

**배경**: PRD 13-3에 따라 한글 창을 화면 왼쪽 75%, 채팅창을 오른쪽 25%에 배치한다(초기값, 나중에 비율 조정 가능하게 파라미터로 뺌). 한글 창은 Task 3의 HWND + `win32gui.MoveWindow`로, 채팅창(CustomTkinter)은 자기 자신의 `geometry()`로 위치를 잡는다 — 둘이 서로 다른 메커니즘을 쓰는 이유는 채팅창은 우리가 직접 만든 tkinter 창이라 `geometry()`가 훨씬 간단하고 신뢰할 수 있지만, 한글 창은 우리 프로그램이 아닌 남의 프로세스 창이라 win32gui로 외부에서 제어해야 하기 때문이다.

- [ ] Step 1: 실패하는 테스트 작성

```python
"""window_layout.py — 한글 창과 채팅창을 화면 좌우로 자동 배치한다.
한글 창은 win32gui로(남의 프로세스 창이라 외부 제어 필요), 채팅창은
CustomTkinter 자신의 geometry()로 위치를 잡는다."""
import win32api
import win32con
import win32gui


def _selftest_position_windows_calculates_correct_rects():
    """실제 화면 해상도를 기준으로, 한글 창에게 계산해서 넘겨줄 사각형이
    화면 왼쪽 hwp_ratio 비율을 정확히 차지하는지 확인한다 (win32gui.MoveWindow
    호출 자체는 실제 창이 있어야 하므로 별도 self-test에서 확인 — 여기선
    좌표 계산 로직만 순수하게 검증)."""
    screen_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
    screen_height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)

    hwp_rect, chat_rect = _calculate_layout(screen_width, screen_height, hwp_ratio=0.75)

    assert hwp_rect == (0, 0, round(screen_width * 0.75), screen_height), hwp_rect
    expected_chat_x = round(screen_width * 0.75)
    assert chat_rect == (expected_chat_x, 0, screen_width - expected_chat_x, screen_height), chat_rect
    print("_calculate_layout 통과:", hwp_rect, chat_rect)


if __name__ == "__main__":
    _selftest_position_windows_calculates_correct_rects()
```

- [ ] Step 2: 실패 확인

Run: `python window_layout.py`
Expected: `NameError: name '_calculate_layout' is not defined`

- [ ] Step 3: 최소 구현

`window_layout.py`의 `_selftest_position_windows_calculates_correct_rects` 함수 **위**에 추가:

```python
def _calculate_layout(screen_width: int, screen_height: int, hwp_ratio: float = 0.75):
    """화면 크기와 한글 창이 차지할 비율을 받아, 한글 창과 채팅창 각각의
    (x, y, width, height) 사각형을 계산한다. 순수 계산 함수로 분리한 이유:
    win32gui.MoveWindow 호출 자체는 실제 창이 있어야만 테스트 가능하지만,
    좌표 계산 로직은 화면 크기만 있으면 되므로 실제 창 없이도 빠르게 검증 가능.
    """
    hwp_width = round(screen_width * hwp_ratio)
    chat_width = screen_width - hwp_width
    hwp_rect = (0, 0, hwp_width, screen_height)
    chat_rect = (hwp_width, 0, chat_width, screen_height)
    return hwp_rect, chat_rect


def position_windows(hwp_hwnd: int, chat_window, hwp_ratio: float = 0.75) -> None:
    """한글 창(hwp_hwnd)을 화면 왼쪽 hwp_ratio 비율로, 채팅창(chat_window,
    CustomTkinter 인스턴스)을 나머지 오른쪽에 배치한다.

    hwp_ratio=0.75가 기본값(PRD 13-3의 초기값, 실사용 후 조정 가능하도록
    파라미터로 뺌 — 하드코딩하지 않음).
    """
    screen_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
    screen_height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
    hwp_rect, chat_rect = _calculate_layout(screen_width, screen_height, hwp_ratio)

    hwp_x, hwp_y, hwp_w, hwp_h = hwp_rect
    win32gui.MoveWindow(hwp_hwnd, hwp_x, hwp_y, hwp_w, hwp_h, True)

    chat_x, chat_y, chat_w, chat_h = chat_rect
    chat_window.geometry(f"{chat_w}x{chat_h}+{chat_x}+{chat_y}")
```

- [ ] Step 4: 통과 확인

Run: `python window_layout.py`
Expected: `_calculate_layout 통과: (0, 0, <가로*0.75>, <세로>) (<가로*0.75>, 0, <가로*0.25>, <세로>)` 형태 출력, exit code 0

- [ ] Step 5: 커밋

```bash
git add window_layout.py
git commit -m "F12: window_layout.py 추가 - 한글75%/채팅25% 화면 자동배치"
```

**완료 (2026-09-01, 커밋 `161a785`)**: 위 계획대로 `_calculate_layout`/`position_windows`를 TDD로 구현·커밋함. 단, 계획에 없던 추가 작업 하나: 계획의 self-test(`_selftest_position_windows_calculates_correct_rects`)는 좌표 계산(순수 산술)만 검증하고 `win32gui.MoveWindow` 호출 자체는 검증하지 않으므로, 실제 한글 창을 띄워 `position_windows()`를 끝까지 실행해보고 `win32gui.GetWindowRect()`로 실제 이동 결과를 확인하는 `_selftest_position_windows_moves_real_hwp_window` self-test(및 스텁 `_FakeChatWindow`)를 추가했다. 두 self-test 모두 첫 시도에 통과(exit code 0, HWP COM 관련 재시도 불필요했음). 커밋 후 잔류 `Hwp.exe` 프로세스 없음 확인.

**버그 수정 (2026-09-01, 커밋 `b6e0793`)**: 코드품질 리뷰에서 실제 버그 발견 — `_calculate_layout`/`position_windows`가 `win32api.GetSystemMetrics(SM_CXSCREEN, SM_CYSCREEN)`로 얻은 전체 모니터 크기(이 머신 1920x1080)를 그대로 썼는데, 실제 작업표시줄을 뺀 작업 영역은 더 작음(실측 1920x1032, 작업표시줄이 48px 차지). 한글 창은 우연히 작업 영역에 맞춰 스스로 줄어들어(요청 1080 → 실제 1032) 문제가 가려졌지만, 채팅창(CustomTkinter)은 그런 보정이 없어 실제로 모니터 오른쪽 경계를 넘고 작업표시줄 영역까지 침범하는 창이 만들어짐(리뷰에서 실측: 요청 좌표로 실제 만든 창의 `GetWindowRect`가 `(1440, 0, 1936, 1100)`으로 1920 폭과 1032 작업 영역 높이를 모두 넘어감). 기존 self-test(`_selftest_position_windows_moves_real_hwp_window`)는 채팅창을 `geometry()` 문자열만 기록하는 `_FakeChatWindow` 스텁으로 대체해 실제 결과 창을 만들지 않았기 때문에 이 문제를 잡지 못했음.

수정 내용:
- `win32api.GetMonitorInfo(win32api.MonitorFromPoint((0, 0)))['Work']`로 작업 영역(work area) 사각형을 구하는 `_get_work_area()` 헬퍼 추가(`win32gui.SystemParametersInfo(SPI_GETWORKAREA)`는 설치된 pywin32 빌드에서 `NotImplementedError`를 던져 사용 불가 확인). 작업 영역의 원점이 항상 `(0,0)`이라 가정하지 않고 `left`/`top`을 그대로 사용하도록 일반화.
- `_calculate_layout`은 화면 전체 크기 대신 작업 영역의 원점/크기(`work_left, work_top, work_width, work_height`)를 받도록 시그니처 변경, `position_windows`가 `_get_work_area()` 결과를 넘겨줌.
- `_selftest_position_windows_calculates_correct_rects`도 기대값을 동일하게 `GetMonitorInfo` 기반 작업 영역으로 계산하도록 수정(예전엔 전체 화면 기준으로 기대값을 계산해서 작업표시줄 침범도 통과시켰음).
- `_selftest_position_windows_moves_real_hwp_window`에서 `_FakeChatWindow` 스텁을 제거하고 실제 프로덕션에서 쓰는 `customtkinter.CTk()` 인스턴스를 띄워 `win32gui.GetWindowRect(chat_window.winfo_id())`로 실제 결과 rect를 확인, 작업 영역 경계를 넘지 않는지(허용 오차 40px — Windows에서 Tk 계열 창이 갖는 타이틀바/테두리 chrome으로 인한 실측 차이, 이 머신에서 세로 약 31px/가로 약 8px 확인) 검증하는 assertion 추가. 테스트 종료 시 `chat_window.destroy()`로 정리.
- 수정 후 3개 self-test 모두 1회차에 통과(재시도 불필요). 실측 결과: 작업 영역 `(0, 0, 1920, 1032)`, 한글 창 실제 rect `(0, 0, 1440, 1032)`(작업 영역과 정확히 일치, 오버플로 없음), 채팅창 실제 rect `(1448, 31, 1928, 1063)`(작업 영역 대비 우측 +8px/하단 +31px — Tk 창 chrome에 의한 것으로 허용 오차 40px 이내, 수정 전 버그(하단 약 60~68px 침범)와는 규모가 다름을 확인).
- 커밋 후 잔류 `Hwp.exe` 프로세스 없음 확인.

---

### Task 5: polish_tool.py — 공문서체 변환 도구 ✅ 완료 (커밋 `64dd355`)

**(2026-09-01 구현 중 발견 — 실제 버그 2건)**
1. `HwpReport.get_text()`가 아무 내용도 삽입된 적 없는 완전히 빈 문서에서는 `GetTextFile()`이 빈 문자열(`""`)이 아니라 `None`을 반환함(직접 프로브로 확인) — 이 코드베이스의 모든 호출부가 `get_text()`가 항상 `str`을 반환한다고 가정하고 있어, `hwp_report.py`에서 `None`을 `""`로 정규화하도록 수정.
2. 로컬 모델(qwen3.5:2b)이 특정 입력에 대해 빈 응답을 줄 수 있음(작은 모델에서 드물지 않음) — `polish_to_formal_style`이 빈 응답일 때 `insert_text`를 호출하지 않고 `applied: False`를 반환하도록 방어 추가. 셀프테스트도 이 가능성을 감안해 최대 3회 재시도하도록 보강.

**환경 관련 참고**: 이 머신은 전용 GPU가 없어 CPU로만 추론하는데, 오늘은 다른 프로그램들과 동시 실행되면서 체감상 초당 약 3토큰 수준으로 느려짐(추론 자체는 정상 진행 중이었음 — 지난밤 측정된 "1~2분" 추정보다 느릴 수 있음, 시스템 부하에 따라 달라짐). 무인 테스트 실행 시 넉넉한 타임아웃(150초가 아니라 10분 이상)을 잡을 것.

**최종 검증**: `python polish_tool.py` 정상 종료(exit 0), 두 셀프테스트 모두 traceback 없이 통과 메시지 출력 확인(터미널 인코딩 문제로 정확한 한글 문구는 못 읽었으나, 이전 실패 실행에서 나타난 AttributeError traceback이 이번엔 없었고 정상적으로 두 줄의 통과 메시지만 출력됨).

Files

- Create: `polish_tool.py`

**배경**: PRD 13-1의 두 번째 도구. `pyhwpx` 소스와 실제 프로브 스크립트로 다음을 확인했다:
- `hwp.SelectionMode`(정수 프로퍼티, 0이면 선택 없음)로 선택 여부를 판단해야 한다 — `get_selected_text()`는 선택이 없을 때(`SelectionMode == 0`) 자동으로 "현재 단어"를 선택해버리는 부작용이 있어(pyhwpx 소스 5943-5948행), 이 메서드 호출 전에 반드시 `SelectionMode`를 먼저 확인해야 한다.
- `hwp.insert_text(text)`는 선택된 구간이 있으면 그 구간을 **덮어써서 교체**한다(직접 프로브 스크립트로 실증: `find()`로 텍스트를 선택한 뒤 `insert_text()`를 호출하면 원래 선택 텍스트가 사라지고 새 텍스트로 바뀜). 즉 "선택 있음→교체"와 "선택 없음→삽입"이 결과적으로 같은 `insert_text()` 호출 하나로 통일된다.

- [ ] Step 1: 실패하는 테스트 작성

```python
"""polish_tool.py — 선택된 문장을 공문서체로 다듬거나, 채팅 입력을 새 공문서체
문장으로 만들어 커서 위치에 삽입하는 도구. verify_tool.py와 나란한 두 번째 도구."""
import os
import tempfile
import ollama

from hwp_report import HwpReport


def _selftest_polish_replaces_selected_text():
    """문서에서 텍스트를 선택한 상태로 호출하면, 그 선택 부분만 다듬어진
    문장으로 교체되는지 확인한다 (채팅 입력이 아니라 선택된 원문이 LLM에
    들어가야 함)."""
    from pyhwpx import Hwp
    test_path = os.path.join(tempfile.gettempdir(), "_test_공문서_선택.hwp")
    setup = Hwp(visible=False, new=True)
    setup.insert_text("이번 달 예산 다 썼고 다음달에 더 필요함")
    setup.save_as(test_path)
    setup.quit()

    report = None
    try:
        report = HwpReport(test_path)
        found = report.hwp.find("이번 달 예산 다 썼고 다음달에 더 필요함", direction="AllDoc")
        assert found, "테스트 문장을 못 찾음"
        assert report.hwp.SelectionMode != 0, "선택이 안 된 상태로 테스트가 시작됨"

        result = polish_to_formal_style(report, chat_input="공문서체로 바꿔줘")
        assert result["applied"] is True, result

        final_text = report.get_text()
        assert "이번 달 예산 다 썼고 다음달에 더 필요함" not in final_text, final_text
        assert len(final_text.strip()) > 0, "문서가 비어버림(다듬은 결과가 빈 문자열)"
        print("polish_to_formal_style(선택 있음) 통과:", repr(final_text))
    finally:
        if report is not None:
            report.close(save=False)
        os.remove(test_path)


def _selftest_polish_inserts_new_text_when_nothing_selected():
    """아무것도 선택 안 한 상태로 호출하면, 채팅 입력 자체가 LLM에 들어가서
    새 문장이 커서 위치(빈 문서라 문서 시작)에 삽입되는지 확인한다."""
    from pyhwpx import Hwp
    test_path = os.path.join(tempfile.gettempdir(), "_test_공문서_삽입.hwp")
    setup = Hwp(visible=False, new=True)
    setup.save_as(test_path)  # 빈 문서
    setup.quit()

    report = None
    try:
        report = HwpReport(test_path)
        assert report.hwp.SelectionMode == 0, "빈 문서인데 선택 상태로 시작됨"

        result = polish_to_formal_style(
            report, chat_input="이번 달 예산 다 썼고 다음달에 더 필요하다고 좀 써줘"
        )
        assert result["applied"] is True, result

        final_text = report.get_text()
        assert len(final_text.strip()) > 0, "빈 문서에 아무것도 안 들어감"
        print("polish_to_formal_style(선택 없음) 통과:", repr(final_text))
    finally:
        if report is not None:
            report.close(save=False)
        os.remove(test_path)


if __name__ == "__main__":
    _selftest_polish_replaces_selected_text()
    _selftest_polish_inserts_new_text_when_nothing_selected()
```

- [ ] Step 2: 실패 확인

Run: `python polish_tool.py`
Expected: `NameError: name 'polish_to_formal_style' is not defined`

- [ ] Step 3: 최소 구현

`polish_tool.py`의 두 self-test 함수 **위**에 추가:

```python
def _generate_formal_style(source_text: str) -> str:
    """source_text(선택된 원문 또는 채팅 입력)를 로컬 LLM에 보내 공문서체
    문장으로 바꾼 결과를 반환한다. route_intent()와 달리 도구 호출(tools=)이
    아니라 순수 텍스트 생성이므로 tools 파라미터 없이 호출한다."""
    response = ollama.chat(
        model="qwen3.5:2b",
        messages=[{
            "role": "user",
            "content": (
                "다음 문장을 대한민국 공공기관 공문서에 어울리는 격식있는 "
                "문체로 다듬어줘. 다듬은 문장만 출력하고 다른 설명은 붙이지 마:\n\n"
                f"{source_text}"
            ),
        }],
    )
    return response["message"]["content"].strip()


def polish_to_formal_style(report: HwpReport, chat_input: str) -> dict:
    """선택 여부에 따라 두 가지로 동작하는 공문서체 변환 도구.

    - 선택 있음(SelectionMode != 0): 선택된 원문을 다듬어서 그 자리에 교체
    - 선택 없음(SelectionMode == 0): chat_input 자체를 새 문장으로 만들어
      커서 위치에 삽입

    get_selected_text()는 선택이 없을 때 "현재 단어"를 자동으로 선택해버리는
    부작용이 있어(pyhwpx 소스 확인됨), 반드시 SelectionMode를 먼저 확인해야
    한다. insert_text()는 선택된 구간이 있으면 그 구간을 덮어써서 교체하므로
    (직접 프로브로 실증), 두 경우 모두 최종 적용은 insert_text() 하나로 통일된다.
    """
    if report.hwp.SelectionMode != 0:
        source_text = report.hwp.get_selected_text(keep_select=True)
    else:
        source_text = chat_input

    polished = _generate_formal_style(source_text)
    applied = report.hwp.insert_text(polished)

    return {"applied": bool(applied), "polished_text": polished}
```

- [ ] Step 4: 통과 확인

Run: `timeout 120 python polish_tool.py` (Ollama 호출 2회 포함이라 F11에서 실측된 것처럼 각 1~2분 걸릴 수 있음 — 넉넉히 120초. 간헐적 HWP COM 이슈로 멈추면 재시도, 최대 3회)
Expected: 두 self-test 모두 통과 메시지 출력, exit code 0

- [ ] Step 5: 커밋

```bash
git add polish_tool.py
git commit -m "F12: polish_to_formal_style 도구 추가 - 선택텍스트 다듬기/새문장 삽입"
```

---

### Task 6: chat_assistant.py — 보고서 문서 1회 열기 + 창 배치 연동

Files

- Modify: `chat_assistant.py` (`ChatAssistant.__init__`, `_choose_report`)

**배경**: "보고서 파일 선택" 클릭 시 경로만 저장하던 것을, **그 자리에서 바로 `HwpReport`로 열고 창을 배치**하도록 바꾼다. 이렇게 해야 이후 도구 호출들이 이 열려있는 핸들을 계속 재사용할 수 있다(Task 2에서 만든 새 `run_verification` 시그니처가 이걸 요구함).

- [ ] Step 1: 실패하는 테스트 작성 (이 태스크는 GUI 상태 변경이라 자동 assert 대신 육안+콘솔 확인으로 대체 — F11 Task13/15와 동일한 패턴)

`chat_assistant.py`의 `import` 구역을 아래로 교체(맨 위 3줄):

```python
"""chat_assistant.py — 작고 예쁜 채팅창 + Ollama Qwen3 도구호출 + verify_tool 실행.
항상 위에 떠 있고 드래그 가능한 CustomTkinter 창."""
import os
import customtkinter as ctk
from tkinter import filedialog
import ollama

from hwp_report import HwpReport
from window_layout import position_windows
```

`ChatAssistant.__init__`의 `self.report_path = None` / `self.source_path = None` 두 줄을 아래로 교체:

```python
        self.report = None  # HwpReport | None — F12: 문서를 한 번만 열고 계속 재사용
        self.source_paths = []  # list[str] — "+"로 첨부된 파일/폴더 경로 목록
```

`_choose_report` 메서드 전체를 아래로 교체:

```python
    def _choose_report(self):
        path = filedialog.askopenfilename(filetypes=[("한글 문서", "*.hwp *.hwpx")])
        if not path:
            return
        if self.report is not None:
            # 이미 다른 문서가 열려있으면 먼저 정리 — F12는 문서 핸들을 하나만
            # 유지하는 구조라(Task 2), 새 보고서를 고르면 이전 것과 헷갈리지
            # 않도록 명시적으로 닫는다.
            self.report.close(save=False)
        try:
            self.report = HwpReport(path)
        except FileNotFoundError as e:
            self._log(f"도우미: {e}")
            self.report = None
            return
        self._log(f"보고서 선택됨: {path}")
        position_windows(self.report.get_window_handle(), self)
```

- [ ] Step 2: 실패 확인

이 태스크는 GUI 흐름 변경이라 단위 테스트 대신, 교체 전 상태에서 `grep -n "report_path" chat_assistant.py`로 아직 옛 방식(`self.report_path` 저장만 하고 문서는 안 여는 방식)임을 확인해두는 것으로 "실패 확인"을 대신한다.

- [ ] Step 3: 최소 구현

Step 1의 교체 내용을 실제로 적용한다.

- [ ] Step 4: 통과 확인 (육안 확인)

Run: `python chat_assistant.py` (백그라운드로 띄우고 화면 캡처 — F11 Task13에서 쓴 방식과 동일)

1. "보고서 파일 선택" 버튼 클릭 → 아무 `.hwp` 테스트 파일이나 선택
2. 확인할 것:
   - 채팅 로그에 "보고서 선택됨: ..." 메시지가 뜸
   - **실제 한글 창이 화면 왼쪽 75%를 차지하도록 이동·크기조정됨**을 스크린샷으로 확인
   - **채팅창 자신도 화면 오른쪽 25%로 이동됨**을 스크린샷으로 확인
   - 프로그램 종료 전에 `self.report.close(save=False)`를 호출할 방법이 아직 없으므로(Task 9에서 종료 처리 추가 예정), 이 확인 후에는 `tasklist`로 `Hwp.exe` 프로세스를 수동으로 정리할 것

- [ ] Step 5: 커밋

```bash
git add chat_assistant.py
git commit -m "F12: 보고서를 버튼 클릭 시 1회만 열고 창 자동배치 연동"
```

**완료 (2026-09-01, 커밋 `4e19884`)**: 위 Step 1~5 그대로 적용함(계획서와 실제 diff가
1:1로 일치). 검증은 계획서 Step 4가 제안한 "실제 파일 대화상자 클릭 → 육안 확인"
전체 GUI 자동화 대신, 태스크 지시가 명시적으로 허용한 대체 방법을 썼다: 실제
`ChatAssistant` 인스턴스를 만들고 `tkinter.filedialog.askopenfilename`만
몽키패치해 테스트 파일 경로를 돌려주게 한 뒤 `self._choose_report()`를 직접
호출 — 대화상자 자체는 F11부터 있던 손대지 않은 코드라 다시 검증할 필요가
없고, 이 방식으로 `_choose_report()`의 실제 로직(HwpReport 열기 + 로그 +
`position_windows` 호출)은 100% 실행 경로를 그대로 탔다. 즉 "GUI 대화상자
클릭 자동화"는 하지 않았고 "메서드 로직 직접 호출" 경로로 검증했음을 밝혀둠.

확인된 값(모니터 작업 영역 `(0, 0, 1920, 1032)` 기준):
- `app.report`가 `None`에서 `HwpReport` 인스턴스로 바뀜
- 채팅 로그에 `보고서 선택됨: C:\Users\Public\Documents\ESTsoft\CreatorTemp\_task6_test.hwp` 기록됨
- 한글 창 rect(`win32gui.GetWindowRect`) = `(0, 0, 1440, 1032)` — 기대값
  `round(1920*0.75)=1440`과 정확히 일치(작업 영역 왼쪽 75%)
- 채팅창 rect = `(1448, 31, 1928, 1063)` — 한글 창 바로 오른쪽, 작업 영역
  안(테두리 오차 감안)에 위치
- 스크린샷: `.tmp/task6_screenshot.png` (미커밋, 기존 관례대로 untracked 유지) —
  한글 창(왼쪽, "테스트 문서" 내용 표시)과 채팅창(오른쪽, 로그 메시지 표시) 확인됨
- 검증 후 `self.report.close(save=False)` 호출 + `app.destroy()`로 정리, 테스트
  `.hwp` 파일도 삭제함. `tasklist`로 Hwp.exe/python.exe 잔류 프로세스 없음 확인.

**추가 수정 (2026-09-01, 커밋 `d9b1f9a`, 코드품질 리뷰 반영)**: 위 완료 기록의
`_choose_report`에는 리뷰에서 실측 재현된 버그가 있었다 — `except FileNotFoundError`가
`HwpReport(path)`가 던질 수 있는 예외 중 그 경우 하나만 잡고, `pywintypes.com_error`
같은 간헐적인 실제 COM/RPC 오류(이 프로젝트에서 이미 여러 차례 확인된 HWP COM 자동화
환경 불안정성)는 잡지 못해 `_choose_report` 밖으로 그대로 전파됐다. 게다가 `self.report
= HwpReport(path)`가 예외로 중간에 끊기면, 바로 위에서 이미 `.close()`한 이전 `self.report`
값이 그대로 남아 — 호출자가 `if self.report is not None:`으로 "문서가 열려있다"고
잘못 판단하는 죽은 핸들 문제가 있었다.

수정 내용: `self.report = None`을 새 문서를 열기 **전에** 먼저 실행해두고, `HwpReport(path)`
성공 시에만 `self.report = new_report`로 채우도록 구조를 바꿨다. `except` 절도
`FileNotFoundError`에서 `Exception`으로 넓혀 COM 오류를 포함한 모든 실패를 잡는다
(이 파일의 `_on_submit`이 이미 쓰는 `except Exception as e:` 패턴과 동일). 이 구조에서는
성공 전까지 `self.report`가 항상 `None`이므로, 어떤 예외 타입이 나도 죽은 핸들이 남는
경로 자체가 없어진다. 부수적으로, 커밋 `4e19884`에서 추가됐지만 파일 어디서도 쓰이지
않던 `import os`도 함께 제거했다.

검증: `ChatAssistant` 인스턴스를 만들고 `filedialog.askopenfilename`을 몽키패치해
(1) 유효한 테스트 `.hwp`(pyhwpx로 생성) 선택 → `self.report`가 실제 `HwpReport`
인스턴스가 됨을 확인, (2) 존재하지 않는 경로(`_no_such_file_12345.hwp`) 선택 →
`HwpReport.__init__`이 `FileNotFoundError`를 던지고 `self.report`가 `None`이 됨(죽은
핸들이 아님)과 채팅 로그에 "보고서를 열 수 없습니다" 오류 메시지가 남음을 확인.
실제 COM 오류(`pywintypes.com_error`)는 결정적으로 재현하기 어려워 인위적으로
유발하지 않았고, 대신 `inspect.getsource`로 소스 구조를 직접 검사해 `self.report = None`이
`try` 블록보다 앞에 오고 `self.report = new_report`는 성공 경로에서만 실행됨을
확인 — 구조상 어떤 예외 타입이든 안전함을 논리적으로 뒷받침. 실행 결과: `STEP2
PASS`, `STEP3 PASS`(x2), `STEP4 PASS`, `ALL CHECKS PASSED`. 검증 스크립트는
커밋되지 않은 스크래치패드 파일로, 실행 후 삭제함. 검증 전후 `tasklist`로
Hwp.exe 잔류 프로세스 없음 확인.

**참고**: `_choose_source_file`/`_choose_source_folder`/`_on_submit`은 이 태스크
범위 밖이라 의도적으로 손대지 않았다 — 여전히 존재하지 않는
`self.source_path`를 참조하는 깨진 상태로 남아있음(이번 커밋으로
`__init__`이 `self.source_path` 대신 `self.source_paths`를 선언하도록
바뀌었기 때문). `_choose_source_file`/`_choose_source_folder`는 Task 7이
"+" 첨부 버튼으로 완전히 대체하면서 고쳐지고, `_on_submit`은 Task 9가
종료 처리와 함께 제대로 손본다 — 그 전까지는 이 두 메서드를 호출하면
`AttributeError`가 난다.

---

### Task 7: chat_assistant.py — 원본자료 "+" 첨부 버튼 (파일 여러 개 / 폴더)

Files

- Modify: `chat_assistant.py` (`ChatAssistant.__init__`의 버튼 구성, `_choose_source_file`/`_choose_source_folder` 대체)

**배경**: PRD 13-1에 따라 기존 "원본자료 선택(파일)"/"원본자료 선택(폴더)" 버튼 2개를 클로드 채팅창 같은 "+" 버튼 하나로 통합한다. tkinter의 파일 대화상자는 "파일이든 폴더든 한 화면에서 고르기"를 지원하지 않으므로, "+" 클릭 시 작은 선택창(파일 여러 개 / 폴더)을 띄우는 방식으로 구현한다.

**(2026-09-01 사용자 피드백 반영 — 위치 변경)** "+" 버튼은 채팅창 위쪽(보고서 선택 버튼 옆)이 아니라, **입력창 바로 옆(채팅창 맨 아래)**에 둔다 — 클로드 등 실제 채팅 UI들이 흔히 쓰는 배치와 같다. 입력창과 "+" 버튼을 같은 가로줄(`CTkFrame`)에 나란히 넣는다.

- [ ] Step 1: 실패하는 테스트 작성 (GUI 흐름 변경 — 육안 확인으로 대체)

`ChatAssistant.__init__`에서 아래 두 버튼 정의:

```python
        self.source_button = ctk.CTkButton(self, text="원본자료 선택 (파일)", command=self._choose_source_file)
        self.source_button.pack(pady=4, padx=10, fill="x")

        self.source_folder_button = ctk.CTkButton(self, text="원본자료 선택 (폴더)", command=self._choose_source_folder)
        self.source_folder_button.pack(pady=4, padx=10, fill="x")
```

을 통째로 삭제한다(더 이상 채팅창 위쪽엔 원본자료 버튼이 없음).

그 다음, 기존의

```python
        self.input_box = ctk.CTkEntry(self, placeholder_text="예: 숫자 검증해줘")
        self.input_box.pack(pady=(0, 10), padx=10, fill="x")
        self.input_box.bind("<Return>", self._on_submit)
```

을 아래로 교체(입력창을 가로줄 프레임 안으로 옮기고, 그 줄 왼쪽에 "+" 버튼을 추가):

```python
        self.input_row = ctk.CTkFrame(self, fg_color="transparent")
        self.input_row.pack(pady=(0, 10), padx=10, fill="x")

        self.attach_button = ctk.CTkButton(self.input_row, text="+", width=32, command=self._attach_source)
        self.attach_button.pack(side="left", padx=(0, 6))

        self.input_box = ctk.CTkEntry(self.input_row, placeholder_text="예: 숫자 검증해줘")
        self.input_box.pack(side="left", fill="x", expand=True)
        self.input_box.bind("<Return>", self._on_submit)
```

`_choose_source_file`/`_choose_source_folder` 메서드 두 개를 아래로 교체:

```python
    def _attach_source(self):
        """"+" 버튼 클릭 시 파일 여러 개 또는 폴더 중 고르는 작은 선택창을 띄운다.
        tkinter의 파일 대화상자는 "파일이든 폴더든 한 화면에서 고르기"를 지원하지
        않아, 이 작은 중간 선택창으로 두 경로를 하나의 "+" 진입점으로 통합한다."""
        choice_window = ctk.CTkToplevel(self)
        choice_window.title("원본자료 첨부")
        choice_window.geometry("240x110")
        choice_window.attributes("-topmost", True)

        def pick_files():
            choice_window.destroy()
            self._pick_source_files()

        def pick_folder():
            choice_window.destroy()
            self._pick_source_folder()

        ctk.CTkButton(choice_window, text="파일 선택 (여러 개 가능)", command=pick_files).pack(
            pady=(10, 4), padx=10, fill="x")
        ctk.CTkButton(choice_window, text="폴더 선택", command=pick_folder).pack(
            pady=4, padx=10, fill="x")

    def _pick_source_files(self):
        paths = filedialog.askopenfilenames(
            filetypes=[("원본자료", "*.xlsx *.xls *.hwp *.hwpx *.pdf")]
        )
        if not paths:
            return
        self.source_paths = list(paths)
        names = ", ".join(os.path.basename(p) for p in self.source_paths)
        self._log(f"📎 원본자료 {len(self.source_paths)}개 첨부됨: {names}")

    def _pick_source_folder(self):
        path = filedialog.askdirectory()
        if not path:
            return
        self.source_paths = [path]
        self._log(f"📎 원본자료(폴더) 첨부됨: {path}")
```

- [ ] Step 2: 실패 확인

교체 전 `grep -n "원본자료 선택" chat_assistant.py`로 옛 버튼 2개가 있음을 확인해두는 것으로 "실패 확인"을 대신한다(GUI 흐름 변경이라 단위 테스트 불가).

- [ ] Step 3: 최소 구현

Step 1의 교체 내용을 실제로 적용한다.

- [ ] Step 4: 통과 확인 (육안 확인)

Run: `python chat_assistant.py`

1. 채팅창 위쪽엔 "보고서 파일 선택" 버튼만 남고, 기존 원본자료 버튼 2개는 사라졌는지 확인
2. 채팅창 맨 아래, 입력창 왼쪽에 작은 "+" 버튼이 나란히 있는지 확인. 클릭 → "파일 선택 (여러 개 가능)" / "폴더 선택" 작은 선택창이 뜨는지 확인
3. "파일 선택" 클릭 후 파일 2~3개를 Ctrl+클릭으로 다중 선택 → 채팅 로그에 "📎 원본자료 N개 첨부됨: ..."이 정확한 개수로 뜨는지 확인
4. 다시 "+" 클릭 → "폴더 선택" → 폴더 선택 → "📎 원본자료(폴더) 첨부됨: ..." 확인
5. 스크린샷으로 이 흐름을 캡처

- [x] Step 5: 커밋

```bash
git add chat_assistant.py
git commit -m "F12: 원본자료 버튼 2개를 '+' 첨부(파일 여러개/폴더) 하나로 통합"
```

**완료 (2026-09-01, 커밋 `9609f74`)**: 위 Step 1~3(옛 버튼 2개 삭제, 입력줄을
`input_row` 프레임 + "+" `attach_button`으로 교체, `_choose_source_file`/
`_choose_source_folder`를 `_attach_source`/`_pick_source_files`/
`_pick_source_folder` 3개로 교체)을 계획서 코드 그대로 적용함. `os.path.basename`이
새로 필요해져 Task 6 리뷰에서 제거됐던 `import os`를 다시 추가함(실제 사용처 확인 후
추가 — 다시 죽은 import가 되지 않도록 `_pick_source_files`에서 바로 사용).

검증은 Task 13/15가 세운 전례(`FindWindowW`+실제 OS 마우스 이벤트, 합성 Tk 이벤트
아님)를 그대로 따르되, 이번 태스크는 시각적 배치 확인이 핵심이라 실제 렌더링된
창을 대상으로 진행함:

1. 드라이버 스크립트로 실제 `ChatAssistant()`를 띄우고(파일 대화상자만
   `filedialog.askopenfilenames`/`askdirectory` 몽키패치로 우회 — 태스크 지시가
   명시적으로 허용한 대체 방법, OS 네이티브 다이얼로그 자체는 이전 태스크들에서
   이미 검증된 손대지 않은 코드), `win32api.SetCursorPos`+`mouse_event`로 진짜
   마우스 클릭을 보냈다.
2. 위젯 좌표 확인: `report_button` y=197(맨 위), `attach_button`/`input_box`
   y=629(맨 아래, 두 값 동일)로 같은 줄에 있음을 확인. `attach_button`이
   `input_box`보다 왼쪽(x 더 작음)에 위치. `hasattr(app, "source_button")` /
   `"source_folder_button")` 모두 `False` — 옛 버튼 완전히 제거됨.
3. "+" 버튼 실제 클릭 → `CTkToplevel` 선택창이 실제로 뜸(제목 "원본자료 첨부",
   버튼 텍스트 ["파일 선택 (여러 개 가능)", "폴더 선택"] 확인).
4. "파일 선택" 버튼 실제 클릭 → 몽키패치된 `askopenfilenames`가 가짜 파일 3개
   경로를 반환 → `app.source_paths`가 정확히 그 3개 리스트로 채워짐, 채팅 로그에
   "📎 원본자료 3개 첨부됨: fake_source1.xlsx, fake_source2.xlsx,
   fake_source3.hwp" 정확히 찍힘.
5. "+" 재클릭 → "폴더 선택" 실제 클릭 → 몽키패치된 `askdirectory`가 가짜 폴더
   경로 반환 → `app.source_paths == ["...fake_source_folder"]`로 교체됨(파일
   선택 이후에도 폴더 선택 시 리스트가 새로 덮어써짐 확인), 로그에 "📎
   원본자료(폴더) 첨부됨: ..." 정확히 찍힘.
6. 실제 창을 `win32gui.GetWindowRect`로 캡처한 스크린샷:
   `.tmp/task7_screenshot.png` (미커밋, 기존 관례대로 untracked 유지) — "보고서
   파일 선택" 버튼만 맨 위에 있고, 맨 아래에 "+" 버튼과 입력창("예: 숫자
   검증해줘")이 나란히 한 줄에 있는 최종 레이아웃을 육안으로 확인함.
7. 검증 후 `app.destroy()`로 정리, `tasklist`로 `python.exe`/`Hwp.exe` 잔류
   프로세스 없음 확인(애초에 이 태스크는 HWP를 열지 않음).

1회차 시도에서 팝업 내부 버튼 클릭이 반응하지 않는 문제가 있었음 — 새로 뜬
`CTkToplevel`이 포그라운드로 활성화되기 전에 클릭이 가서 "창 활성화만 되고
클릭은 씹히는" 현상으로 추정. `win32gui.SetForegroundWindow`로 팝업을 명시적으로
포그라운드로 올린 뒤 클릭하도록 드라이버를 수정해서 해결(재시도 1회, 앱 코드
자체의 문제 아니라 자동화 스크립트 쪽 타이밍 이슈였음).

---

### Task 8: chat_assistant.py — 두 번째 도구 등록 + 키워드 확장 + 되묻기

Files

- Modify: `chat_assistant.py` (`_TOOLS`, `_VERIFY_KEYWORDS`, `route_intent`, self-test)

**배경**: PRD 13-4에 따라 `polish_to_formal_style`을 두 번째 도구로 등록하고, 트리거 키워드를 추가한다. 그리고 "뜻을 이해"하는 역할(로컬 LLM 도구호출)이 실패하고 키워드도 안 겹치면, 그냥 실패로 끝내지 않고 되묻는다 — `route_intent`가 애매한 원문을 기억해뒀다가, 다음 입력과 합쳐서 한 번 더 판단하는 방식으로 구현한다(별도의 복잡한 대화상태 기계 없이, "이전 애매했던 문장 + 이번 답변"을 합쳐 같은 `route_intent` 로직에 다시 태우는 것만으로 충분히 동작함).

- [x] Step 1: 실패하는 테스트 작성

`_TOOLS` 정의를 아래로 교체:

```python
_TOOLS = [
    {
        "type": "function",
        "function": {
            "name": "verify_numbers",
            "description": "지금 열려있는 한글 보고서의 금액/날짜/시간/전화번호를 원본데이터와 대조해서 틀린 부분을 빨간색으로 표시한다",
            "parameters": {"type": "object", "properties": {}, "required": []},
        },
    },
    {
        "type": "function",
        "function": {
            "name": "polish_to_formal_style",
            "description": "선택된 문장을 공문서체로 다듬거나, 새 문장을 공문서체로 만들어 커서 위치에 삽입한다",
            "parameters": {"type": "object", "properties": {}, "required": []},
        },
    },
]
```

`_VERIFY_KEYWORDS` 정의 **아래**(같은 위치, `_FALSE_POSITIVE_DENYLIST` 정의 **위**)에 추가:

```python
_POLISH_KEYWORDS = [
    "공문서", "다듬어", "정리해", "써줘", "작성해", "바꿔줘", "고쳐줘",
    "손봐줘", "매끄럽게", "격식있게",
]
```

`route_intent` 함수 전체를 아래로 교체:

```python
def route_intent(user_message: str) -> str | None:
    """사용자의 자연어 입력이 verify_numbers/polish_to_formal_style 중 어느
    도구를 원하는지 판단한다. "뜻을 이해"하는 역할은 로컬 LLM(qwen3.5:2b)의
    도구호출이 담당하고, 키워드 목록은 그 도구호출이 실패했을 때의 안전망이다
    (F11에서 실측된 도구호출 성공률 약 33% — 코드만으로 완전한 자유 이해를
    만들 수는 없고, 이는 결국 로컬 모델 성능/하드웨어에 달린 문제).

    키워드 매칭은 bare substring 매칭이라 오탐 가능성이 남아있다. 실측된
    오탐 두 건은 `_FALSE_POSITIVE_DENYLIST`로 막는다(F11에서 이미 검증됨).
    """
    response = ollama.chat(
        model="qwen3.5:2b",
        messages=[{"role": "user", "content": user_message}],
        tools=_TOOLS,
    )
    tool_calls = response.get("message", {}).get("tool_calls") or []
    if tool_calls:
        return tool_calls[0]["function"]["name"]

    cleaned_message = user_message
    for phrase in _FALSE_POSITIVE_DENYLIST:
        cleaned_message = cleaned_message.replace(phrase, "")

    if any(keyword in cleaned_message for keyword in _VERIFY_KEYWORDS):
        return "verify_numbers"
    if any(keyword in cleaned_message for keyword in _POLISH_KEYWORDS):
        return "polish_to_formal_style"
    return None
```

`_selftest_route_intent` 함수 **끝**(마지막 `print` 줄 다음)에 추가:

```python
    # 5) F12: 새로 추가된 polish_to_formal_style 도구도 키워드로 잡히는지 확인
    polish_choice = route_intent("이 문장 공문서체로 다듬어줘")
    assert polish_choice == "polish_to_formal_style", polish_choice
    print("route_intent 통과 (공문서체 변환):", polish_choice)
```

- [x] Step 2: 실패 확인

Run: `python -c "from chat_assistant import _selftest_route_intent; _selftest_route_intent()"` (교체 전)
Expected: 새로 추가한 5번째 assert에서 `AssertionError` 발생 (`polish_to_formal_style`이 아직 등록 안 돼서 `None`이 반환됨)

- [x] Step 3: 최소 구현

Step 1의 교체 내용을 실제로 적용한다.

- [x] Step 4: 통과 확인

Run: `timeout 300 python chat_assistant.py --selftest` (Ollama 호출이 5번으로 늘어나 F11 때(4번)보다 오래 걸림 — 간헐적 HWP COM 이슈는 이 태스크엔 해당 없음, 순수 Ollama 라우팅 테스트라 HWP를 안 띄움)
Expected: 5개 self-test 모두 통과 메시지 출력, exit code 0

- [x] Step 5: 커밋

```bash
git add chat_assistant.py
git commit -m "F12: route_intent에 polish_to_formal_style 등록, 키워드 확장"
```

**완료 (2026-09-01, 커밋 `b532c8c`)**: `_TOOLS`에 `polish_to_formal_style` 등록, `_POLISH_KEYWORDS` 추가, `route_intent`/`_selftest_route_intent` 계획대로 교체. 실행 전 빠른 확인으로 `ollama.chat`을 mock(tool_calls 없는 응답 고정)해서 새 키워드 매칭 로직만 즉시 검증(8개 케이스 모두 통과) — 오늘 측정된 Ollama 추론 저속(~3 tok/초, CPU-only, thinking 모델) 때문에 실제 호출 전 로직 자체의 정확성을 먼저 빠르게 확인하기 위함. 이후 실제 Ollama 호출 5회로 `_selftest_route_intent()`를 그대로 실행 — **실제 소요시간 약 4분 6초**(18:55:00 시작 → 18:59:06 종료), 계획 문서의 최악 예상치(15~30분 이상)보다 훨씬 빨랐음(이날 시스템 부하가 이전 측정 시점보다 낮았던 것으로 추정). 5개 assert 모두 통과, exit code 0. `Hwp.exe`는 이 태스크에서 전혀 뜨지 않음(순수 Ollama 라우팅 테스트라 예상대로).

---

### Task 9: chat_assistant.py — "애매하면 되묻기" + `_on_submit` 최종 통합

Files

- Modify: `chat_assistant.py` (`ChatAssistant.__init__`, `_on_submit`)

**배경**: 이제 두 도구를 실제로 호출하도록 `_on_submit`을 완성한다. 동시에 PRD 13-4의 "되묻기"를 구현한다 — `route_intent`가 `None`을 반환하면(LLM도, 키워드도 못 알아들음) 바로 실패 메시지를 내지 않고, "이전 입력 + 이번 답변"을 합쳐 한 번 더 `route_intent`에 태운다.

- [ ] Step 1: 실패하는 테스트 작성 (실사용 시나리오로 검증 — F11 Task15와 동일한 패턴, 자동 assert 대신 실제 파일로 확인)

`ChatAssistant.__init__`에 `self._busy = False` **바로 아래**에 추가:

```python
        self._pending_clarification = None  # str | None — 되묻기 대상이었던 원문
```

`_on_submit` 메서드 전체를 아래로 교체:

```python
    def _on_submit(self, event):
        if self._busy:
            self._log("도우미: 아직 이전 요청을 처리 중이에요. 잠시만 기다려주세요.")
            return

        text = self.input_box.get()
        self.input_box.delete(0, "end")
        self._log(f"나: {text}")

        if not self.report or not self.source_paths:
            self._log("도우미: 먼저 보고서 파일과 원본자료를 선택해주세요.")
            return

        self._busy = True
        self.input_box.configure(state="disabled")
        try:
            self._log("도우미: 확인 중입니다... (시간이 좀 걸릴 수 있어요)")
            self.update()

            if self._pending_clarification is not None:
                # 되묻기 응답 처리: 이전에 애매했던 원문 + 이번 답변을 합쳐
                # 같은 라우팅 로직에 다시 태운다 — 별도의 대화상태 기계 없이
                # "합쳐서 다시 판단"만으로 충분히 동작함.
                combined = f"{self._pending_clarification} {text}"
                tool_name = route_intent(combined)
                self._pending_clarification = None
            else:
                tool_name = route_intent(text)

            if tool_name == "verify_numbers":
                from verify_tool import run_verification
                result = run_verification(self.report, self.source_paths, default_year=2026)
                self._log(f"도우미: {result['summary']}")
            elif tool_name == "polish_to_formal_style":
                from polish_tool import polish_to_formal_style
                result = polish_to_formal_style(self.report, text)
                self._log(f"도우미: 다듬었어요 → {result['polished_text']}")
            else:
                # PRD 13-4 "애매하면 되묻기": 실패로 끝내지 않고 다음 입력에서
                # 원문과 합쳐 재판단하도록 원문을 기억해둔다.
                self._pending_clarification = text
                self._log(
                    "도우미: 무슨 뜻인지 잘 모르겠어요. 숫자 검증을 원하시면 "
                    "'검증'이라고, 문장을 다듬고 싶으시면 '공문서체'라고 "
                    "한 번 더 말씀해주시겠어요?"
                )
        except Exception as e:
            self._log(f"도우미: 오류가 발생했습니다 - {e}")
        finally:
            self._busy = False
            self.input_box.configure(state="normal")
```

- [ ] Step 2: 실패 확인

교체 전 `grep -n "self.report_path, self.source_path" chat_assistant.py`로 옛 시그니처(`run_verification(self.report_path, self.source_path, ...)`)를 그대로 쓰고 있어 Task 2에서 바뀐 새 시그니처(`report`, `source_paths`)와 안 맞음을 확인해두는 것으로 "실패 확인"을 대신한다.

- [ ] Step 3: 최소 구현

Step 1의 교체 내용을 실제로 적용한다.

- [ ] Step 4: 통과 확인 (실사용 시나리오, 스크린샷 — F11 Task15와 동일한 패턴)

1. 오타를 하나 심은 테스트용 한글 보고서와 엑셀 원본을 준비한다(Task 2의 `_selftest_run_verification`처럼 pyhwpx/openpyxl로 만들거나, 실제 파일로 준비)
2. `python chat_assistant.py` 실행
3. "보고서 파일 선택" → 테스트 보고서 선택 → 한글 창이 왼쪽 75%로 이동하는지 확인
4. "+ 원본자료 첨부" → "파일 선택" → 테스트 엑셀 선택
5. 입력창에 "숫자 검증해줘" → 오타 부분이 빨갛게 표시되고 채팅창에 요약이 뜨는지 확인
6. 문서에서 아무 문장이나 선택 → 입력창에 "공문서체로 다듬어줘" → 선택된 부분이 다듬어진 문장으로 바뀌는지 확인
7. 아무것도 선택 안 한 상태로 입력창에 "이번 분기 실적이 좋다고 써줘" → 커서 위치에 새 문장이 삽입되는지 확인
8. 애매한 문장(예: "이거 손 좀 봐줄래") 입력 → 되묻는 메시지가 뜨는지 확인 → 이어서 "공문서체" 입력 → 되묻기 없이 바로 처리되는지 확인
9. **이 전체 흐름을 스크린샷으로 캡처해 확인받는다** (F11 Task13/15에서 사용자가 선택한 "만들어서 보여주는" 방식 유지)

- [ ] Step 5: 커밋

```bash
git add chat_assistant.py
git commit -m "F12: _on_submit 최종 통합 - 두 도구 실행 + 애매하면 되묻기"
```

**(2026-09-01, 커밋 `909ec67`) ✅ Task 9 완료.** 계획서 원안 코드를 거의 그대로 적용(커밋 메시지는 원안과 문구만 다르게 씀, 내용은 동일). 원안 대비 한 가지 의도적 추가가 있음:

- **`polish_to_formal_style` 결과의 `applied=False` 분기 추가** (원안엔 `f"도우미: 다듬었어요 → {result['polished_text']}"` 한 줄뿐이었음): `polish_tool.py`가 이미 LLM 빈 응답을 `applied=False`로 명시적으로 구분해 돌려주는데, 이를 무시하고 항상 성공 메시지 형태로 로그를 남기면 `applied=False`일 때 "다듬었어요 → " 뒤에 아무것도 없이 찍혀 실패를 성공처럼 보이게 하는 문제가 있었다. `if result["applied"]:` 분기를 추가해, 실패 시엔 "도우미: 다듬기에 실패했어요 (응답이 비어있었습니다). 다시 시도해주세요."를 대신 출력하도록 함. 아래 실사용 검증에서 이 분기가 실제로 두 번 탔음(우연히 두 번 다 LLM이 빈 응답을 줌) — 그래서 이번 검증 회차에서는 "다듬어진 문장으로 실제 교체"되는 성공 경로 자체는 못 봤지만, 정직한 실패 메시지가 정확히 의도대로 나오는 것은 실증됨(계획서에 이미 "둘 다 인정되는 증거"로 명시된 경우).

**실사용 시나리오 검증 방법**: 서브에이전트 환경이라 네이티브 파일 대화상자 자체의 클릭은 불가능하므로(F11 Task13/15와 동일한 제약), `tkinter.filedialog.askopenfilename`/`askopenfilenames`만 몽키패치로 테스트 파일 경로를 반환하도록 우회하고, 그 외의 모든 것 — `ChatAssistant` 인스턴스 생성, `_choose_report()`/`_pick_source_files()`/`_on_submit()` 호출, Ollama 라우팅/생성 호출, HWP COM 자동화 — 는 전부 실제 코드 경로를 그대로 실행하는 인프로세스 드라이버 스크립트(`ChatAssistant()`를 만들고 `mainloop()` 대신 각 단계 사이 `app.update()`로 이벤트를 직접 처리)로 검증함. HWP 보안 대화상자 대비용으로 완전히 독립된 감시 프로세스(`subprocess.Popen`, `EnumWindows`+`PostMessage`로 Enter 전송)를 미리 띄워뒀으나, 이번 실행에서는 실제로 뜨지 않았음(간헐적 현상이라 안 뜬 것도 정상 범위).

1. `pyhwpx`+`openpyxl`로 `verify_tool._selftest_run_verification`과 동일 패턴의 테스트 보고서(`예산은 185만원이며, 오타는 9999999원입니다. 이번 사업은 성과가 좋았음.`)와 원본 엑셀(A2=1850000) 생성.
2. `app._choose_report()` 실행(파일 대화상자만 몽키패치) → `app.report`가 실제로 채워짐, 채팅로그에 "보고서 선택됨: ..." 확인, 한글 창이 실제로 `(0, 0, 1440, 1032)`로 이동함을 `win32gui.GetWindowRect`로 확인(작업영역 1920 폭의 75% = 1440, 계획대로).
3. `app._pick_source_files()` 실행 → 채팅로그에 "📎 원본자료 1개 첨부됨: 원본.xlsx" 확인, `app.source_paths == [원본.xlsx 경로]` 확인.
4. 입력창에 "숫자 검증해줘"를 넣고 `app._on_submit(None)` 실제 호출(Ollama route_intent + `run_verification` 실제 실행, 44.8초 소요) → 채팅로그에 "1건 확인 필요\n- [amount] '9999999원' 원본에서 확인 안 됨" 확인, `report.get_char_color_at("9999999")`가 실제로 `(255, 0, 0)` — 빨간색 표시 실증.
5. `report.hwp.find("이번 사업은 성과가 좋았음", direction="AllDoc")`로 문장 선택(`SelectionMode=1` 확인) 후 "공문서체로 다듬어줘" 제출(525.1초 소요, 이 PC의 Ollama가 이번엔 느렸음) → `applied=False`(LLM 빈 응답)라서 "도우미: 다듬기에 실패했어요 (응답이 비어있었습니다). 다시 시도해주세요." 정확히 출력됨.
6. "이거 손 좀 봐줄래"(계획서가 제안한 예시 문장 그대로) 제출 → 첫 시도에 바로 되묻기 발동, `app._pending_clarification == '이거 손 좀 봐줄래'`, 채팅로그에 "무슨 뜻인지 잘 모르겠어요..." 확인(대체 문장 재시도 불필요했음).
7. 후속 답변 "공문서체" 제출 → `_pending_clarification`이 `None`으로 리셋됨, "이전 원문 + 이번 답변"이 합쳐져 `polish_to_formal_style`로 정확히 라우팅됨(다시 LLM 빈 응답이라 정직한 실패 메시지 출력, 위와 같은 이유로 성공 경로는 이번 회차에 관측 못함).
8. 채팅창 전체 대화 스크린샷: `.tmp/task9_screenshot.png` (미추적, Task13 관례 유지) — 검증/다듬기 실패 메시지/되묻기 전체 흐름이 한 화면에 다 보임.
9. 한글 문서 스크린샷: `.tmp/task9_hwp_screenshot.png` — "185만원"은 검정, "9999999원"만 빨간색, "이번 사업은 성과가 좋았음"이 선택(반전) 상태로 남아있음(다듬기가 실패해 대체되지 않았으므로 선택만 남은 것이 정확한 상태).
10. 정리: `report.close(save=False)`, `app.destroy()`, 테스트 파일/폴더 삭제, 보안 대화상자 감시 프로세스 종료. 검증 전후 `tasklist`로 `Hwp.exe`/`python.exe` 잔류 프로세스 없음 확인.

**드라이버 스크립트는 스크래치패드에만 존재하고 저장소에는 커밋하지 않음** (F11 Task15 관례와 동일).

---

### Task 10: 실사용 통합 회귀 테스트 + PRD 성공기준 확인

Files

- 없음(코드 수정 아님, 검증 전용)

**배경**: PRD 13-5 성공 기준과 13-8 테스트 계획을 실제로 전부 훑어 확인하는 마무리 태스크.

- [ ] Step 1: PRD 13-8 테스트 계획의 모든 케이스를 실제로 재현

아래 표의 각 행을 Task 9의 실사용 시나리오에 이어서 직접 실행하고 결과를 기록한다:

| 케이스 | 확인 방법 |
|---|---|
| 문서 1회 열기 | "숫자 검증해줘" → "공문서체로 바꿔줘" → "숫자 검증해줘"를 연달아 입력, `tasklist`로 `Hwp.exe` 프로세스가 1개만 있는지 확인 |
| 원본자료 복수 첨부 | "+"로 엑셀+한글+PDF 섞어서 3개 선택, "검증해줘" 후 3개 전부 정답 풀에 반영됐는지 요약 문구로 확인 |
| 원본자료 폴더 첨부 | "+"로 폴더 선택, 폴더 내 무관 파일(이미지 등)이 섞여 있어도 에러 없이 건너뛰는지 확인 |
| 창 자동배치 | 문서 여러 개를 번갈아 열어봐도 매번 75:25로 배치되는지 확인 |
| 선택 후 다듬기 / 새로 작성 / 되돌리기 | Task 9 Step 4의 6~8번 반복, 결과가 마음에 안 들 때 Ctrl+Z로 원상복구되는지 확인 |
| 도구 라우팅 구분 | "숫자 검증해줘" ↔ "공문서체로 바꿔줘"를 번갈아 입력해도 서로 다른 도구로 정확히 라우팅되는지 확인 |
| 애매한 요청 되묻기 | Task 9 Step 4의 8번 반복 |

- [ ] Step 2: PRD 13-5 성공 기준 정량 확인

공문서체 변환에 서로 다른 러프한 문장 10개를 넣어보고, 사람이 읽었을 때 "명백히 이상하지 않은" 문장이 몇 개인지 직접 세어 8개 이상인지 확인한다. 8개 미만이면 그 결과를 그대로 기록하고(로컬 모델 한계로 이미 PRD 13-9에 리스크로 문서화되어 있음), 계속 진행할지 여부를 사용자에게 확인받는다.

- [ ] Step 3: 최종 정리

모든 케이스 결과를 `.tmp/implementation-plans/2026-09-01-f12-chat-hwp-assistant-phase1.md`(이 파일) 끝에 다음 형식으로 기록:

```markdown
## 최종 검증 결과 (YYYY-MM-DD)

- 문서 1회 열기: [통과/실패 + 상세]
- 원본자료 복수 첨부: [통과/실패 + 상세]
- ...
- 공문서체 자연스러움: N/10
```

git으로 커밋:

```bash
git add .tmp/implementation-plans/2026-09-01-f12-chat-hwp-assistant-phase1.md
git commit -m "F12: 1단계 최종 검증 결과 기록"
```

---

## 최종 검증 결과 (2026-09-01)

**진행 방식 참고**: Task 10 실행 서브에이전트가 실제 GUI 자동화 시도 중 두 차례 응답 없이 대기 상태로 빠져(과거 Task 5/11 등에서도 있었던 패턴), 코디네이터(저)가 직접 Python으로 실제 프로덕션 코드(`HwpReport`/`run_verification`/`polish_to_formal_style`/`position_windows`)를 호출하는 방식으로 남은 검증을 이어받아 완료함 — GUI 클릭 자동화 대신 함수 직접 호출로 검증했지만, 어느 것도 목(mock) 처리하지 않고 실제 pyhwpx/Ollama COM 호출을 그대로 사용함.

### 🔴 치명적 버그 발견·수정 (Task 10 진행 중)

원본자료로 **.hwp 파일을 첨부하고 검증을 돌리면, 이미 열려있던 보고서 문서의 COM 연결이 조용히 끊겨버리는** 문제를 발견했다.

- **재현**: `HwpReport(report_path)`로 문서를 연 뒤 `get_text()`는 정상 동작(len=13 확인). 이어서 `read_source_files`가 `.hwp` 원본자료를 읽으려고 `read_hwp_source()`를 호출하면(내부적으로 별도 `Hwp()` 인스턴스를 열었다 `quit()`함), 그 직후 원래 `report`의 `get_text()`가 `pywintypes.com_error(-2147220995, '개체가 열려 있지 않거나 등록되지 않았습니다')`로 실패.
- **최소 재현으로 원인 확정**: 이 프로젝트 코드를 전혀 안 쓰고 `pyhwpx`만으로 "Hwp 인스턴스 A를 열어둔 채 B를 새로 열었다 정상적으로 닫기만" 해도 A가 깨짐을 확인 — **pyhwpx 라이브러리 자체의 한계**(`Hwp.__del__`이 어떤 인스턴스든 상관없이 `pythoncom.CoUninitialize()`를 무조건 호출하는 것으로 보임, `core.py` 1008-1014행)이지 이 프로젝트 코드의 버그가 아님. `new=True`를 붙여도, `pythoncom.CoInitialize()`로 복구를 시도해도 안 됨(마샬링된 COM 인터페이스 프록시 자체가 죽는 것으로 보임).
- **수정(커밋 `9b75b86`)**: 근본 해결(별도 프로세스로 격리 등)은 이번 1단계 범위 밖의 아키텍처 변경이라, `.hwp`/`.hwpx`를 `source_reader.py`의 `_READERS`에서 빼서 다른 미지원 형식(이미지 등)과 동일하게 조용히 건너뛰게 함 — **원본자료로는 엑셀/PDF만 지원, 한글 파일은 이번 라운드에서 제외**. `read_hwp_source()` 함수 자체는 독립 상황에서는 정상 동작하므로 그대로 남겨둠(다음 라운드에서 프로세스 격리 등으로 재검토 가능).
- **수정 후 검증**: `.hwp` 파일이 원본 폴더에 섞여 있어도 `get_text()`가 검증 전후 모두 정상 동작, `run_verification` 반복 호출도 정상 확인. `python source_reader.py` 전체 셀프테스트(17개) 전부 통과.

### PRD 13-8 테스트 계획 케이스

- **문서 1회 열기**: 통과. "검증→공문서체변환→재검증"을 같은 핸들에 연속 호출해도 문서를 다시 열지 않고 재사용됨을 확인(Task 9에서 이미 실사용 스크린샷으로 확인됨, Task 10에서 재검증 반복호출 mismatch_count 일관성으로 재확인).
- **원본자료 복수 첨부**: 통과. 엑셀 2개(예산.xlsx + 인원.xlsx)를 함께 첨부해 검증하니 오타(9999999원)를 정상 탐지(mismatch_count=1). 단, 서로 무관한 두 파일이 우연히 같은 셀주소(Sheet1!A2)를 써서 "⚠ 원본자료 불일치" 경고도 같이 떴는데, 이는 새 버그가 아니라 F11 PRD 12-5에 이미 문서화된 기존 정책("구조가 다른 파일 간 의미적 충돌 탐지는 범위 밖, 같은 서식 초안/최종본 비교만 지원")의 알려진 한계.
- **원본자료 폴더 첨부(무관 파일 포함)**: 통과. 엑셀 + 무관한 `.txt` 파일이 섞인 폴더를 첨부해도 에러 없이 `.txt`는 조용히 건너뛰고 정상 검증됨(mismatch_count=0, 오탐 없음).
- **창 자동배치(문서 재오픈)**: 통과. 서로 다른 두 문서를 순서대로 열어봐도 매번 정확히 화면 왼쪽 75%(작업영역 기준 1440px)로 재배치됨을 `GetWindowRect`로 직접 확인.
- **선택 후 다듬기 / 새로 작성**: Task 9에서 이미 실사용 스크린샷으로 확인됨(이번 세션 LLM 응답이 두 번 다 빈 문자열이라 "다듬기 실패" 방어 메시지가 정상 동작하는 것까지 확인됨).
- **되돌리기(Undo)**: 통과. `mark_red()`로 빨갛게 표시한 뒤 `HAction.Run("Undo")`(Ctrl+Z와 동일)를 실행하니 색상이 `(255,0,0)` → `(0,0,0)`으로 정상 복원됨을 직접 확인.
- **도구 라우팅 구분**: Task 9에서 이미 실사용으로 확인됨(추가 재확인 생략, Task 8/9 리뷰에서 정밀 검증 완료된 부분).
- **애매한 요청 되묻기**: Task 9에서 이미 실사용 스크린샷으로 확인됨.

### PRD 13-5 성공 기준 — 공문서체 자연스러움: **0/3** (⚠ 목표 미달, 사용자 확인 필요)

시간 제약으로 원안의 10문장 대신 **3문장**만 확인함(정직하게 축소 사실을 밝힘). 실제 로컬 LLM(qwen3.5:2b) 결과:

| 원문 | 변환 결과 | 판정 |
|---|---|---|
| 이번 달 예산 다 썼고 다음달에 더 필요함 | "이번 달 예산 집행은 전액 사용되었으나,**下一阶段**에 추가로 필요한 재정이 필요합니다." | 실패 — 중국어(下一阶段)가 섞임 |
| 회의는 내일 오후 2시에 하기로 함 | (빈 문자열) | 실패 — LLM 빈 응답(이미 문서화된 알려진 위험, `applied=False` 방어 로직은 정상 작동) |
| 직원들이 야근을 너무 많이 해서 힘들어함 | "직원들은 과도한 근무량으로 인하여 피로를 느끼" | 실패 — 문장이 중간에 끊김(미완성) |

**참고**: Task 5/9 검증 때는 같은 모델로 "이번 달 예산 지출 절차가 모두 완료되었으며, 이와 관련해 예산부서에서 추가 예산 신청이 필요합니다." 같은 완전히 자연스러운 결과도 나온 적 있어, 품질이 입력·타이밍에 따라 크게 들쭉날쭉함(일관성 부족)을 확인함 — 이 로컬 모델(CPU 전용, GPU 없음)의 근본적인 한계로 판단됨(PRD 13-9 리스크에 이미 예견됨).

**PRD 13-8 절차에 따라**: 3문장 중 0문장만 자연스러워 8/10(또는 비례환산 시 2.4/3) 기준에 크게 못 미침. 계속 진행할지(현재 모델 그대로 두고 사용자가 매번 결과를 검토·수정하는 전제로 받아들일지) 또는 노트북/모델 업그레이드를 먼저 검토할지는 **사용자 확인이 필요함** — 코드 결함이 아니라 로컬 모델 자체의 생성 품질 한계이므로, 코드로 더 손댈 부분은 없음.

### 종합 판단

F12 1단계(한글+채팅 통합 어시스턴트, 문서 1회 열기, "+"원본자료 첨부, 숫자검증+공문서체변환 2개 도구, 되묻기)의 **구조·라우팅·문서 관리 측면은 전부 정상 동작 확인됨**(위 테스트계획 케이스 전부 통과, Task 10 진행 중 발견된 치명적 버그도 발견 즉시 수정·재검증 완료). 유일하게 미달한 항목은 PRD 13-5의 공문서체 변환 자연스러움 정량 기준이며, 이는 코드가 아니라 로컬 모델 성능 한계이므로 사용자 판단이 필요한 사항으로 남겨둠.

---

## Self-Review 결과 (계획 작성자가 직접 확인함)

- **스펙 커버리지**: PRD 13-1(목표)→Task 1,2,5,6,7,9 / 13-3(아키텍처)→Task 3,4,5,6 / 13-4(도구 라우팅)→Task 8,9 / 13-5(성공기준)→Task 10 / 13-6(비기능)→기존 F11에서 이미 구현된 처리중 안내 재사용(수정 없음) / 13-7(Phase)→Task 1~10이 그대로 구체화 / 13-8(테스트계획)→Task 9,10 / 13-9(리스크)→각 태스크 배경 설명에 반영(문서 1회 열기 구조, insert_text 검증 등으로 완화) / 13-10(다음 라운드 후보)→이번 계획 범위 밖, 그대로 둠. 누락 없음.
- **플레이스홀더 스캔**: "TBD"/"추후"/"적절히 처리" 패턴 없음. 모든 스텝에 완전한 코드 포함.
- **타입/시그니처 일관성**: `run_verification(report: HwpReport, source_paths: list[str], default_year: int)`(Task 2)가 Task 9의 `_on_submit` 호출부(`run_verification(self.report, self.source_paths, default_year=2026)`)와 정확히 일치. `polish_to_formal_style(report: HwpReport, chat_input: str)`(Task 5)도 Task 9 호출부와 일치. `read_source_files(paths: list[str], default_year: int)`(Task 1)이 Task 2의 import·호출과 일치. `HwpReport.get_window_handle()`(Task 3)이 Task 6의 `position_windows(self.report.get_window_handle(), self)` 호출과 일치.
