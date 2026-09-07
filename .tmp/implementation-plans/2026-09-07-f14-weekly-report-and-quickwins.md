# F14 (주간업무보고 취합 + 4종 퀵윈) Implementation Plan

> For the agent worker: required sub skill. Implement task by task with `subagent-driven-development`. Track steps with checkbox (`- [ ]`) syntax.

Goal: F11~F13에서 완성된 한글+채팅 통합 어시스턴트에 다음 5개 기능을 추가한다.
1. **주간업무보고(주보) 취합 자동화** — 여러 사람이 보낸 고정서식 문서(표 하나,
   "이번주(...)"/"다음주(...)" 2열)에서 파란색으로 써넣은 내용만 뽑아, 이미
   열려있는 대상 보고서의 같은 칸으로 옮긴다. 자리가 겹치면 덮어쓰지 않고
   그 아래에 이어붙인다.
2. **"한 페이지에 맞춤" 자동 서식조정** — 행간→자간→글자크기 순으로 조금씩
   줄여가며 1페이지에 맞춘다.
3. **컬럼 간 파생값(증감률·평균) 계산** — `compute_column_sums()`(합계만)을
   확장해 두 컬럼 간 증감률과 단일 컬럼 평균을 정답 풀에 추가한다.
4. **오탐 학습("이건 무시해")** — 숫자검증 결과 중 특정 항목을 무시 목록에
   등록하면 다음부터 오류로 표시하지 않는다.
5. **첨부문서 미리보기** — "+"로 첨부한 문서를 첨부 즉시 작은 미리보기(HWP는
   이미지, 엑셀/PDF는 텍스트)로 보여준다.

Architecture: 5개 기능 전부 기존 관례(도구 하나당 파일 하나, `HwpReport` 핸들을
호출자가 열고 닫음, pyhwpx self-test는 `Hwp(visible=False, new=True)`, 같은
프로세스 안 두 번째 Hwp 인스턴스가 필요하면 반드시 `source_reader.py`의
`read_hwp_source_isolated` 패턴처럼 subprocess로 격리)을 그대로 따른다. 새 도구
모듈 4개(`weekly_report_tool.py`, `fit_to_page_tool.py`, `ignore_list.py`,
`attachment_preview.py`)를 만들고, `verify_numbers.py`/`verify_tool.py`/
`chat_assistant.py`를 확장한다.

Tech stack: 기존과 동일(Python 3.13, pyhwpx 1.7.2, customtkinter, openpyxl,
pdfplumber, ollama) — 신규 의존성 없음.

## 사전조사 실측 확인 요약 (계획서 작성 전 `Hwp(visible=False, new=True)`로 직접 실행해 확인함)

이 섹션은 PRD/요구사항이 "확인 필요"라고 표시한 지점들을 실제로 실행해본
결과다 — 아래 Task들의 코드는 전부 이 결과를 근거로 작성됐다.

1. **표 안 특정 칸의 텍스트 읽기**: 커서를 그 칸으로 옮긴 뒤
   `hwp.get_selected_text()` 하나로 읽힌다. pyhwpx의 `get_selected_text()`는
   `SelectionMode==0`이고 `hwp.is_cell()`이 참이면 내부적으로
   `TableCellBlock()`을 실행해 "현재 칸 전체"를 선택하므로, 칸 안에서 그냥
   호출하면 된다. 여러 문단이 있는 칸도 `"\r\n"`으로 이어진 하나의 문자열로
   정확히 반환됨을 확인(직접 재현: 2문단짜리 칸이 `'줄1\r\n줄2'`로 읽힘).

2. **칸 안 글자색 판별**: `hwp.find(문단텍스트, direction="Forward")`로 찾은 뒤
   `hwp.CharShape.Item("TextColor")`를 읽으면 된다 — `hwp_report.py`의
   `get_char_color_at()`와 완전히 같은 방식이 표 안에서도 그대로 통한다(직접
   재현: 같은 칸 안에서 문단1=파랑, 문단2=검정으로 서로 다르게 칠해도 각각
   정확히 구분해서 읽음).

3. **"이번주"/"다음주" 헤더로 표 찾기**: pyhwpx에 표 전용 열거 API
   (`get_into_nth_table()`)는 있지만, 이 프로젝트의 전제(문서 안에 이런
   구조의 표가 하나뿐)에서는 `hwp.find(키워드, direction="Forward")`로 찾은
   위치가 `hwp.is_cell()`이고 `hwp.get_cell_addr()`의 행 번호가 1(헤더 행)인
   첫 occurrence를 헤더로 채택하는 방식이 더 단순하고 헤더 텍스트 자체로
   검증까지 겸한다. 확인됨: `get_cell_addr()`은 "A1"/"A2"/"B2" 같은 문자열을
   반환하고, `get_row_num()`/`get_col_num()`은 칸 안에서 그 표의 전체
   행/열 개수를 반환한다(2행 2열 표에서 각각 2, 2 확인).

4. **자리가 겹칠 때 "덮어쓰지 않고 이어붙이기" — 이번 사전조사에서 가장
   중요한 발견**: 두 가지 커서 위치 규칙이 확연히 다르다.
   - `hwp.TableCellBlock()`(칸 전체 선택) 후 `hwp.Cancel()`(또는 `Right`
     방향키)을 하면, 커서는 선택 영역의 **시작점**(칸 맨 앞)에 남는다 — 이
     상태에서 그냥 삽입하면 기존 내용 뒤가 아니라 **앞**에 새 내용이
     끼어든다(직접 재현: 기존 `'기존내용1\r\n기존내용2'`가 있는 칸에서 이
     방식으로 삽입하면 `'새내용기존내용1\r\n기존내용2'`가 됨 — 버그).
   - 반면 `hwp.find(마지막_줄_텍스트, direction="Forward")`로 칸 안의 마지막
     문단을 찾아 선택한 뒤 `hwp.Cancel()`을 하면, 커서는 그 선택 영역의
     **끝점**(그 줄의 맨 끝)에 남는다 — 이 뒤에 `BreakPara`+`insert_text`를
     하면 정확히 기존 내용 뒤에 이어붙는다(직접 재현: 결과가
     `'기존내용1\r\n기존내용2\r\n새내용'`이 됨 — 의도한 동작).
   이 프로젝트의 다른 코드(`table_tool.py`, `hwp_report.py`)가 "Cancel() 이후
   커서는 선택 영역의 끝점에 남는다"고 문서화해둔 것은 전부 `find()`로 만든
   선택에 대한 얘기였다 — `TableCellBlock()`으로 만든 "칸 전체" 선택에는 그
   규칙이 적용되지 않는다는 게 이번에 새로 확인된 사실이다. 그래서
   `weekly_report_tool.py`는 반드시 후자(마지막 줄을 find로 다시 찾아
   Cancel) 방식만 쓴다.

5. **행간/자간/글자크기 API**: `hwp.PageCount`, `hwp.SelectAll()` +
   `hwp.set_linespacing(value, method="Percent")`(160%→130%까지 5%p씩 줄여도
   실제로 페이지 수가 줄어듦을 확인)까지는 사전 확인된 그대로였다. 다만
   **자간 API는 사전 확인 내용과 실제로 달랐다** — `hwp.set_para(Spacing=...)`
   라는 API 자체가 pyhwpx에 없다(`set_para()`의 실제 파라미터는 `Condense`
   등이고 `Spacing`이 없음, pyhwpx 소스로 직접 확인). 자간은
   `hwp.set_font(Spacing=값)`(-50~50, `set_font()`의 파라미터)가 맞다 — 직접
   테스트로 자간 -10 적용 시 페이지 수가 줄어듦을 확인. 글자크기는
   `hwp.set_font(Height=값)`(포인트 단위 직접 지정)로 조정하며, 현재 값은
   `hwp.HwpUnitToPoint(hwp.CharShape.Item("Height"))`로 읽는다.

6. **첨부 미리보기 이미지 용량**: `hwp.create_page_image(path, pgno=1,
   resolution=300, format="bmp")`(기본값)는 무압축이라 300dpi 한 장에 3.5MB
   가까이 나온다(96dpi로 낮춰도 마찬가지로 3.5MB대 — 직접 측정). 반면
   `format="gif"`, `resolution=60`으로 낮추면 같은 문서 기준 약 6.8KB로
   확 줄어든다(직접 측정) — 그래서 미리보기는 반드시 gif+저해상도 조합을
   쓴다.

7. **pyhwpx COM 인스턴스 생성/재생성 관련 주의(이번 사전조사 중 재현)**:
   한 파이썬 프로세스 안에서 `Hwp()`를 만들고 `quit()`한 뒤, 별다른 작업
   없이 곧바로 또 `Hwp()`를 만들어 `open()`하는 "build→quit→즉시 재오픈"
   패턴을 **한 스크립트 안에서 여러 번 연달아** 반복하면 `pywintypes.com_error
   (-2147220995)`("개체가 이미 소멸되어 사용할 수 없습니다")가 재현될 수
   있음을 확인했다(같은 변수명 재사용 여부와는 무관 — 무관한 변수명으로도
   재현됨). 반면 "문서를 만든 그 인스턴스에서 곧바로(재오픈 없이) 같은
   인스턴스로 검증 로직을 실행"하거나, "각 self-test 함수가 build→quit 후
   **한 번만** 재오픈하고, 함수 사이에 `time.sleep(2)`를 두는" 기존 이
   프로젝트의 관례대로 하면 문제가 재현되지 않았다(직접 확인). 그래서 아래
   Task들의 self-test는 이 프로젝트의 기존 관례(각 self-test 함수당 재오픈은
   최대 1회, 함수 사이 `time.sleep(2)`)를 그대로 따른다.

**주의사항 (F13에서 얻은 교훈, 전 Task 공통)**: pyhwpx self-test는 절대
병렬 실행 금지(순차 실행 + self-test 사이 `time.sleep(2)`). 한글 COM
자동화는 예고 없이 대화상자를 띄워 무인 실행을 멈출 수 있다 — 이번
배치에서 새로 쓰는 표/셀 조작, 페이지 수 확인, 이미지 생성 API들도 이
위험이 있을 수 있으니, 만약 Task 실행 중 응답없음이 발생하면 사람이
화면을 보며 확인해야 할 수 있다(발견되면 `hwp_report.py`의
`SetMessageBoxMode` 카테고리 표를 참고해 대응). 새 self-test 함수는
반드시 `if __name__ == "__main__":` 블록에 등록할 것. `chat_assistant.py`
self-test 실행은 `python chat_assistant.py --selftest`(플래그 없이 실행하면
GUI 앱이 뜬다).

---

## 파일 구조

| 파일 | 책임 |
|---|---|
| `weekly_report_tool.py` (신규) | 소스 .hwp에서 파란색 "이번주"/"다음주" 내용만 뽑아 대상 문서의 같은 칸에 이어붙이는 도구 |
| `fit_to_page_tool.py` (신규) | 문서를 1페이지에 맞추기 위해 행간→자간→글자크기를 조금씩 줄이는 도구 |
| `ignore_list.py` (신규) | 숫자검증 오탐 무시 목록을 JSON 파일로 기록/조회 (speed_tracker.py와 같은 관례) |
| `attachment_preview.py` (신규) | 첨부 문서(HWP/엑셀/PDF)의 작은 미리보기 생성 + 오래된 미리보기 정리 |
| `verify_numbers.py` (수정) | `compute_column_growth_rate()`/`compute_column_average()` 추가 |
| `verify_tool.py` (수정) | `run_verification()`이 무시 목록에 있는 값을 애초에 mismatches에서 제외 |
| `chat_assistant.py` (수정) | 4개 신규 도구 라우팅/키워드/UI 연결, "N번째는 무시해" 결정론적 분기, 첨부 시 미리보기 표시 |

---

### Task 1: weekly_report_tool.py — 소스 문서에서 파란색 내용만 뽑아 읽기

Files

- Create: `weekly_report_tool.py`

- [x] Step 1: 실패하는 테스트 작성

`weekly_report_tool.py`를 새로 만들고 맨 아래에 작성 (아직 `read_weekly_content`
함수가 없으므로 실행하면 NameError):

```python
"""weekly_report_tool.py — 여러 사람이 보낸 주간업무보고(고정 서식: 표 하나,
"이번주(...)"/"다음주(...)" 2열 헤더) 문서를 읽어, 그 표 안에 파란색으로
써넣어진 내용만 뽑아 이미 열려있는 대상 보고서의 같은 칸으로 옮겨 붙이는 도구.
verify_tool.py/table_tool.py/polish_tool.py와 나란한, 도구 하나당 파일 하나
관례를 따른다.

실측 확인(2026-09-07, Hwp(visible=False, new=True)로 직접 테스트 — 계획서
"사전조사 실측 확인 요약" 참고):
- 표 안 특정 칸의 텍스트는 커서를 그 칸으로 옮긴 뒤 hwp.get_selected_text()
  하나로 읽힌다.
- 칸 안 특정 위치의 글자색은 hwp.find(문단텍스트, direction="Forward") 후
  hwp.CharShape.Item("TextColor")로 읽으면 된다 — hwp_report.py의
  get_char_color_at()와 완전히 같은 방식이 표 안에서도 그대로 통한다.
- "이번주"/"다음주" 헤더로 표를 찾는 전용 API는 pyhwpx에 없다 — 문서를
  hwp.find(키워드, direction="Forward")로 훑으면서, 찾은 위치가 is_cell()이고
  get_cell_addr()의 행 번호가 1(헤더 행)인 첫 occurrence를 헤더로 채택한다.
- 표 하나에 헤더 1행 + 데이터 1행(2열)만 있는 고정 구조라는 전제(사용자가
  제시한 실제 양식 스크린샷 기준)를 그대로 따른다.
"""
import re


def _is_bluish(rgb: tuple[int, int, int]) -> bool:
    """RGB가 "파란색으로 써넣은 내용"으로 볼 만큼 충분히 파란지 판별한다.
    (0,0,255) 정확히 일치가 아니라 여유를 두는 이유: 사람마다 한/글에서
    고르는 "파랑"이 팔레트상 조금씩 다를 수 있어서다(순정 파랑이 아닌
    남색 계열도 있을 수 있음). B값이 충분히 밝고 R/G보다 뚜렷하게 클 때만
    파란색으로 인정한다."""
    r, g, b = rgb
    return b >= 120 and (b - max(r, g)) >= 60


def _goto_header_cell(hwp, keyword: str) -> bool:
    """문서 안에서 keyword("이번주" 또는 "다음주")가 들어간 표의 헤더 셀
    (1행)을 찾아 커서를 그 칸에 둔다. keyword가 표 밖(본문)에도 등장할 수
    있으므로, 찾은 위치가 반드시 칸 안이고 1행이어야 진짜 헤더로 인정한다 -
    아니면 다음 occurrence로 계속 넘어간다. 못 찾으면 False."""
    hwp.MoveDocBegin()
    while hwp.find(keyword, direction="Forward"):
        if hwp.is_cell():
            addr = hwp.get_cell_addr()
            row = int(re.search(r'\d+', addr).group())
            if row == 1:
                return True
    return False


def _read_blue_lines_at_cursor(hwp) -> list[str]:
    """현재 커서가 있는 칸(표 데이터 칸)의 내용을 문단(줄) 단위로 읽어, 그
    중 파란색으로 쓰인 문단만 순서대로 리스트로 반환한다. 칸이 비어있으면
    빈 리스트."""
    cell_text = hwp.get_selected_text()
    if not cell_text:
        return []
    lines = [ln for ln in cell_text.replace("\r\n", "\n").split("\n") if ln]
    blue_lines = []
    for line in lines:
        if not hwp.find(line, direction="Forward"):
            continue
        color_value = hwp.CharShape.Item("TextColor")
        rgb = (color_value & 0xFF, (color_value >> 8) & 0xFF, (color_value >> 16) & 0xFF)
        if _is_bluish(rgb):
            blue_lines.append(line)
    return blue_lines


def read_weekly_content(path: str) -> dict:
    """path(.hwp/.hwpx)를 안 보이게 열어 "이번주"/"다음주" 표에서 파란색
    내용만 뽑아 {"this_week": [...], "next_week": [...]} 형태로 반환한다.
    각 리스트는 그 칸에서 파란색으로 확인된 문단(줄) 순서 그대로다 - 파란색
    문단이 하나도 없으면 빈 리스트(병합 대상 없음이라는 뜻, 오류 아님).
    read_hwp_source()와 동일한 관례로, 열 수 없는 파일은 예외 없이 빈
    결과를 반환한다."""
    from pyhwpx import Hwp
    hwp = None
    try:
        hwp = Hwp(visible=False, new=True)
        if not hwp.open(path):
            return {"this_week": [], "next_week": []}
        result = {}
        for key, keyword in (("this_week", "이번주"), ("next_week", "다음주")):
            if _goto_header_cell(hwp, keyword):
                hwp.Run("TableColBegin")
                hwp.Run("MoveDown")
                result[key] = _read_blue_lines_at_cursor(hwp)
            else:
                result[key] = []
        return result
    except Exception:
        return {"this_week": [], "next_week": []}
    finally:
        if hwp is not None:
            hwp.quit()


def _selftest_read_weekly_content_extracts_only_blue_lines():
    """소스 문서의 "이번주"/"다음주" 데이터 칸에 파란색 실제 내용과 검정색
    안내문구가 섞여 있어도, 파란색 문단만 뽑아야 한다(실측 확인: 칸 안에서도
    find()+CharShape.Item("TextColor")로 문단별 색을 정확히 구분할 수 있음,
    모듈 docstring 참고)."""
    import os, tempfile, time
    from pyhwpx import Hwp
    import pandas as pd

    path = os.path.join(tempfile.gettempdir(), "_test_주간보고_원본.hwp")
    setup = Hwp(visible=False, new=True)
    df = pd.DataFrame({"이번주(9/1~9/5)": [""], "다음주(9/8~9/12)": [""]})
    setup.table_from_data(df, header=True, index=False, header_bold=True)

    setup.Run("TableColBegin")
    setup.Run("MoveDown")  # A2 (이번주 데이터 칸)
    setup.insert_text("- 이번주 실제 항목")
    setup.TableCellBlock()
    setup.set_font(TextColor=setup.RGBColor(0, 0, 255))
    setup.Cancel()
    setup.find("이번주 실제 항목", direction="Forward")
    setup.Cancel()  # find()로 만든 선택 - Cancel() 후 커서는 그 줄의 끝에 남음
    setup.HAction.Run("BreakPara")
    setup.insert_text("(안내문구)")  # 삽입 직후엔 직전 글자색(파랑)을 그대로 물려받음
    setup.find("(안내문구)", direction="Forward")
    setup.set_font(TextColor=setup.RGBColor(0, 0, 0))  # 검정으로 되돌려 안내문구처럼 만듦
    setup.Cancel()

    setup.Run("TableRightCell")  # B2 (다음주 데이터 칸)
    setup.insert_text("- 다음주 실제 항목")
    setup.TableCellBlock()
    setup.set_font(TextColor=setup.RGBColor(0, 0, 255))
    setup.Cancel()

    setup.save_as(path)
    setup.quit()
    time.sleep(2)

    try:
        result = read_weekly_content(path)
        assert result["this_week"] == ["- 이번주 실제 항목"], result
        assert result["next_week"] == ["- 다음주 실제 항목"], result
        print("read_weekly_content(파란색만 추출) 통과:", result)
    finally:
        os.remove(path)


if __name__ == "__main__":
    _selftest_read_weekly_content_extracts_only_blue_lines()
```

- [x] Step 2: 실패 확인

Run: `python weekly_report_tool.py`
Expected: 처음엔 함수가 이미 파일 안에 정의돼 있으므로 바로 실행하면 통과할
것처럼 보일 수 있으나, 위 코드를 그대로 붙여넣기 전(함수 정의 없이 self-test만
먼저 넣는 경우)이라면 `NameError: name 'read_weekly_content' is not defined`.
**중요**: 이 파일은 신규 생성이라 실제로는 함수와 테스트를 한 번에 작성하게
되므로, TDD 절차상 먼저 self-test 함수 부분만 파일에 작성해 `NameError`를
직접 확인한 뒤, 그 다음에 위 함수 정의들을 추가하는 순서로 진행한다.

- [x] Step 3: 통과 확인 (함수 정의 추가 후)

Run: `python weekly_report_tool.py`
Expected: `read_weekly_content(파란색만 추출) 통과: {'this_week': ['- 이번주 실제 항목'], 'next_week': ['- 다음주 실제 항목']}` 출력, exit 0

- [x] Step 4: 커밋

```bash
git add weekly_report_tool.py
git commit -m "F14: weekly_report_tool.py 추가 - 소스 문서에서 파란색 이번주/다음주 내용만 추출"
```

---

### Task 2: weekly_report_tool.py — 별도 프로세스 격리 + 대상 칸에 안전하게 이어붙이기

Files

- Modify: `weekly_report_tool.py`

- [x] Step 1: 실패하는 테스트 작성 (격리 실행)

`weekly_report_tool.py`의 `if __name__ == "__main__":` 블록 위에 추가:

```python
def _selftest_read_weekly_content_isolated_basic():
    """read_weekly_content_isolated()가 별도 프로세스를 통해서도
    read_weekly_content()와 동일한 결과를 돌려주는지 확인한다.
    source_reader.py의 read_hwp_source_isolated()와 같은 이유(같은 프로세스
    안에 이미 열려있는 대상 문서의 HwpReport COM 연결이 pyhwpx Hwp.__del__의
    CoUninitialize()로 함께 끊기는 것을 방지)로 반드시 격리 실행해야 한다."""
    import os, tempfile, time
    from pyhwpx import Hwp
    import pandas as pd

    path = os.path.join(tempfile.gettempdir(), "_test_주간보고_격리.hwp")
    setup = Hwp(visible=False, new=True)
    df = pd.DataFrame({"이번주(9/1~9/5)": [""], "다음주(9/8~9/12)": [""]})
    setup.table_from_data(df, header=True, index=False, header_bold=True)
    setup.Run("TableColBegin")
    setup.Run("MoveDown")
    setup.insert_text("이번주 항목")
    setup.TableCellBlock()
    setup.set_font(TextColor=setup.RGBColor(0, 0, 255))
    setup.Cancel()
    setup.Run("TableRightCell")
    setup.insert_text("다음주 항목")
    setup.TableCellBlock()
    setup.set_font(TextColor=setup.RGBColor(0, 0, 255))
    setup.Cancel()
    setup.save_as(path)
    setup.quit()
    time.sleep(2)

    try:
        result = read_weekly_content_isolated(path)
        assert result == {"this_week": ["이번주 항목"], "next_week": ["다음주 항목"]}, result
        print("read_weekly_content_isolated 통과:", result)
    finally:
        os.remove(path)


def _selftest_read_weekly_content_isolated_missing_file_no_crash():
    """존재하지 않는 파일도 예외 없이 빈 결과를 반환해야 한다."""
    result = read_weekly_content_isolated("존재하지_않는_주간보고.hwp")
    assert result == {"this_week": [], "next_week": []}, result
    print("read_weekly_content_isolated(파일없음) 통과")
```

- [x] Step 2: 실패 확인

Run: `python weekly_report_tool.py`
Expected: `NameError: name 'read_weekly_content_isolated' is not defined`

- [x] Step 3: 최소 구현

파일 상단 import에 추가:

```python
import json
import os
import re
import subprocess
import sys
```

`read_weekly_content()` 함수 뒤, self-test들 앞에 추가:

```python
def read_weekly_content_isolated(path: str) -> dict:
    """read_weekly_content()를 별도 프로세스에서 실행해 결과를 받아온다 -
    source_reader.py의 read_hwp_source_isolated()와 완전히 같은 이유(같은
    프로세스 안에 이미 열려있는 대상 문서의 HwpReport COM 연결이 pyhwpx
    Hwp.__del__의 CoUninitialize()로 함께 끊기는 것을 방지)로 반드시 격리
    실행해야 한다."""
    module_dir = os.path.dirname(os.path.abspath(__file__))
    script = (
        "import sys; sys.path.insert(0, sys.argv[2]); "
        "import json; from weekly_report_tool import read_weekly_content; "
        "print(json.dumps(read_weekly_content(sys.argv[1])))"
    )
    try:
        result = subprocess.run(
            [sys.executable, "-c", script, path, module_dir],
            capture_output=True, text=True, timeout=60,
        )
        if result.returncode != 0:
            return {"this_week": [], "next_week": []}
        return json.loads(result.stdout)
    except Exception:
        return {"this_week": [], "next_week": []}
```

`if __name__ == "__main__":` 블록을 아래로 갱신:

```python
if __name__ == "__main__":
    _selftest_read_weekly_content_extracts_only_blue_lines()
    _selftest_read_weekly_content_isolated_basic()
    _selftest_read_weekly_content_isolated_missing_file_no_crash()
```

- [x] Step 4: 통과 확인

Run: `python weekly_report_tool.py`
Expected: 3개 self-test 전부 통과, exit 0

- [x] Step 5: 대상 칸에 이어붙이기 - 실패하는 테스트 작성 (빈 칸)

`if __name__ == "__main__":` 블록 위에 추가:

```python
def _selftest_append_lines_to_data_cell_into_empty_cell():
    import os, tempfile, time
    from pyhwpx import Hwp
    import pandas as pd

    path = os.path.join(tempfile.gettempdir(), "_test_대상_빈칸.hwp")
    setup = Hwp(visible=False, new=True)
    df = pd.DataFrame({"이번주(9/1~9/5)": [""], "다음주(9/8~9/12)": [""]})
    setup.table_from_data(df, header=True, index=False, header_bold=True)
    setup.save_as(path)
    setup.quit()
    time.sleep(2)

    target = Hwp(visible=False, new=True)
    target.open(path)
    try:
        ok = _append_lines_to_data_cell(target, "이번주", ["새 항목1", "새 항목2"])
        assert ok is True

        assert _goto_header_cell(target, "이번주")
        target.Run("TableColBegin")
        target.Run("MoveDown")
        cell_text = target.get_selected_text()
        assert cell_text == "새 항목1\r\n새 항목2", repr(cell_text)
        print("_append_lines_to_data_cell(빈 칸) 통과:", repr(cell_text))
    finally:
        target.quit()
        os.remove(path)


def _selftest_append_lines_to_data_cell_preserves_existing_content():
    """핵심 안전규칙 회귀 테스트: 자리가 겹치면(이미 다른 사람 내용이 있으면)
    덮어쓰지 않고 그 아래에 이어붙여야 한다. TableCellBlock()+Cancel()로
    커서를 옮기면 선택 영역의 "시작점"에 남아 기존 내용 앞에 끼어드는 회귀가
    실제로 재현됐었다(모듈 docstring "사전조사 실측 확인" 참고) -
    find(마지막 줄)+Cancel() 방식만 옳게 동작한다."""
    import os, tempfile, time
    from pyhwpx import Hwp
    import pandas as pd

    path = os.path.join(tempfile.gettempdir(), "_test_대상_겹침.hwp")
    setup = Hwp(visible=False, new=True)
    df = pd.DataFrame({"이번주(9/1~9/5)": [""], "다음주(9/8~9/12)": [""]})
    setup.table_from_data(df, header=True, index=False, header_bold=True)
    setup.Run("TableColBegin")
    setup.Run("MoveDown")
    setup.insert_text("기존 내용1")
    setup.HAction.Run("BreakPara")
    setup.insert_text("기존 내용2")
    setup.save_as(path)
    setup.quit()
    time.sleep(2)

    target = Hwp(visible=False, new=True)
    target.open(path)
    try:
        ok = _append_lines_to_data_cell(target, "이번주", ["새 항목"])
        assert ok is True

        assert _goto_header_cell(target, "이번주")
        target.Run("TableColBegin")
        target.Run("MoveDown")
        cell_text = target.get_selected_text()
        assert cell_text == "기존 내용1\r\n기존 내용2\r\n새 항목", repr(cell_text)
        print("_append_lines_to_data_cell(자리 겹침, 이어붙이기) 통과:", repr(cell_text))
    finally:
        target.quit()
        os.remove(path)


def _selftest_append_lines_to_data_cell_empty_lines_is_noop():
    """lines가 비어있으면 아무 것도 안 하고 False를 반환해야 한다(병합할
    내용이 없는 소스 문서를 건너뛸 때 이 계약에 의존함, Task 3 참고)."""
    import os, tempfile, time
    from pyhwpx import Hwp
    import pandas as pd

    path = os.path.join(tempfile.gettempdir(), "_test_대상_빈리스트.hwp")
    setup = Hwp(visible=False, new=True)
    df = pd.DataFrame({"이번주(9/1~9/5)": [""], "다음주(9/8~9/12)": [""]})
    setup.table_from_data(df, header=True, index=False, header_bold=True)
    setup.save_as(path)
    setup.quit()
    time.sleep(2)

    target = Hwp(visible=False, new=True)
    target.open(path)
    try:
        ok = _append_lines_to_data_cell(target, "이번주", [])
        assert ok is False
        print("_append_lines_to_data_cell(빈 리스트, no-op) 통과")
    finally:
        target.quit()
        os.remove(path)
```

- [x] Step 6: 실패 확인

Run: `python weekly_report_tool.py`
Expected: `NameError: name '_append_lines_to_data_cell' is not defined`

- [x] Step 7: 최소 구현

`read_weekly_content_isolated()` 함수 뒤에 추가:

```python
def _append_lines_to_data_cell(hwp, keyword: str, lines: list[str]) -> bool:
    """대상 문서에서 keyword 헤더 밑 데이터 칸에 lines를 이어붙인다. 칸에
    이미 내용이 있으면 그 뒤에(자리가 겹쳐도 덮어쓰지 않고) 새 문단으로
    추가한다 - 실측 확인된 안전한 방법(모듈 docstring 참고): 기존 내용의
    마지막 줄을 find(Forward)로 다시 찾아 Cancel()로 그 줄 끝에 커서를
    남긴 뒤에만 BreakPara+insert_text로 이어붙인다. lines가 비어있으면
    아무 것도 하지 않고 False."""
    if not lines:
        return False
    if not _goto_header_cell(hwp, keyword):
        return False
    hwp.Run("TableColBegin")
    hwp.Run("MoveDown")

    existing = hwp.get_selected_text()
    if existing:
        existing_lines = [ln for ln in existing.replace("\r\n", "\n").split("\n") if ln]
        last_line = existing_lines[-1] if existing_lines else existing
        # get_selected_text() 호출 후 커서는 칸 시작 위치로 돌아와 있다 -
        # 다시 칸에 들어가 마지막 줄을 찾는다.
        if not _goto_header_cell(hwp, keyword):
            return False
        hwp.Run("TableColBegin")
        hwp.Run("MoveDown")
        hwp.find(last_line, direction="Forward")
        hwp.Cancel()  # find()로 만든 선택 - Cancel() 후 커서는 그 줄의 끝에 남음
        hwp.Run("BreakPara")

    for i, line in enumerate(lines):
        if i > 0:
            hwp.Run("BreakPara")
        hwp.insert_text(line)
    return True
```

`if __name__ == "__main__":` 블록을 아래로 갱신:

```python
if __name__ == "__main__":
    _selftest_read_weekly_content_extracts_only_blue_lines()
    _selftest_read_weekly_content_isolated_basic()
    _selftest_read_weekly_content_isolated_missing_file_no_crash()
    _selftest_append_lines_to_data_cell_into_empty_cell()
    _selftest_append_lines_to_data_cell_preserves_existing_content()
    _selftest_append_lines_to_data_cell_empty_lines_is_noop()
```

- [x] Step 8: 통과 확인

Run: `python weekly_report_tool.py`
Expected: 6개 self-test 전부 통과, exit 0

- [x] Step 9: 커밋

```bash
git add weekly_report_tool.py
git commit -m "F14: weekly_report_tool.py - 별도 프로세스 격리 읽기 + 자리 겹침 시 이어붙이기"
```

---

### Task 3: weekly_report_tool.py — merge_weekly_reports() 전체 취합 오케스트레이션

Files

- Modify: `weekly_report_tool.py`

- [x] Step 1: 실패하는 테스트 작성

`if __name__ == "__main__":` 블록 위에 추가:

```python
def _selftest_merge_weekly_reports_end_to_end():
    """전체 흐름: 소스 2개(각각 이번주/다음주에 서로 다른 파란색 항목),
    파란색 내용이 없는 소스 1개, 비-hwp 파일 1개를 섞어 첨부했을 때 -
    merged_files/no_content_files/skipped_files가 올바르게 분류되고, 대상
    문서의 "이번주" 칸(이미 다른 담당자 내용이 있는 자리 겹침 시나리오)에
    두 소스 내용이 순서대로 이어붙는지 확인한다."""
    import os, shutil, tempfile, time
    from pyhwpx import Hwp
    from hwp_report import HwpReport
    import pandas as pd

    test_dir = os.path.join(tempfile.gettempdir(), "_test_주간보고_취합")
    os.makedirs(test_dir, exist_ok=True)

    def make_source(filename, this_week_line, next_week_line):
        p = os.path.join(test_dir, filename)
        h = Hwp(visible=False, new=True)
        df = pd.DataFrame({"이번주(9/1~9/5)": [""], "다음주(9/8~9/12)": [""]})
        h.table_from_data(df, header=True, index=False, header_bold=True)
        h.Run("TableColBegin")
        h.Run("MoveDown")
        h.insert_text(this_week_line)
        h.TableCellBlock()
        h.set_font(TextColor=h.RGBColor(0, 0, 255))
        h.Cancel()
        h.Run("TableRightCell")
        h.insert_text(next_week_line)
        h.TableCellBlock()
        h.set_font(TextColor=h.RGBColor(0, 0, 255))
        h.Cancel()
        h.save_as(p)
        h.quit()
        return p

    source1 = make_source("김대리.hwp", "- 예산안 검토", "- 결과보고 작성")
    time.sleep(2)
    source2 = make_source("박주임.hwp", "- 행사 준비", "- 정산 마무리")
    time.sleep(2)

    no_content_path = os.path.join(test_dir, "빈내용.hwp")
    h_empty = Hwp(visible=False, new=True)
    df_empty = pd.DataFrame({"이번주(9/1~9/5)": [""], "다음주(9/8~9/12)": [""]})
    h_empty.table_from_data(df_empty, header=True, index=False, header_bold=True)
    h_empty.save_as(no_content_path)
    h_empty.quit()
    time.sleep(2)

    non_hwp_path = os.path.join(test_dir, "무관.txt")
    with open(non_hwp_path, "w", encoding="utf-8") as f:
        f.write("이 파일은 취합 대상이 아님")

    target_path = os.path.join(test_dir, "대상보고서.hwp")
    setup = Hwp(visible=False, new=True)
    df_target = pd.DataFrame({"이번주(9/1~9/5)": [""], "다음주(9/8~9/12)": [""]})
    setup.table_from_data(df_target, header=True, index=False, header_bold=True)
    setup.Run("TableColBegin")
    setup.Run("MoveDown")
    setup.insert_text("- 기존 담당자 항목")  # 자리 겹침 시나리오
    setup.save_as(target_path)
    setup.quit()
    time.sleep(2)

    report = None
    try:
        report = HwpReport(target_path)
        result = merge_weekly_reports(
            report, source_paths=[source1, source2, no_content_path, non_hwp_path]
        )
        assert sorted(result["merged_files"]) == sorted([source1, source2]), result
        assert result["no_content_files"] == [no_content_path], result
        assert result["skipped_files"] == [non_hwp_path], result

        assert _goto_header_cell(report.hwp, "이번주")
        report.hwp.Run("TableColBegin")
        report.hwp.Run("MoveDown")
        this_week_text = report.hwp.get_selected_text()
        assert this_week_text == "- 기존 담당자 항목\r\n- 예산안 검토\r\n- 행사 준비", repr(this_week_text)

        assert _goto_header_cell(report.hwp, "다음주")
        report.hwp.Run("TableColBegin")
        report.hwp.Run("MoveDown")
        next_week_text = report.hwp.get_selected_text()
        assert next_week_text == "- 결과보고 작성\r\n- 정산 마무리", repr(next_week_text)
        print("merge_weekly_reports(전체 흐름) 통과")
    finally:
        if report is not None:
            report.close(save=False)
        shutil.rmtree(test_dir)
```

- [x] Step 2: 실패 확인

Run: `python weekly_report_tool.py`
Expected: `NameError: name 'merge_weekly_reports' is not defined`

- [x] Step 3: 최소 구현

`_append_lines_to_data_cell()` 함수 뒤에 추가:

```python
def merge_weekly_reports(report, source_paths: list[str]) -> dict:
    """source_paths 중 .hwp/.hwpx 파일들을 각각 읽어, 파란색 내용을 대상
    문서(이미 열려있는 report: HwpReport)의 "이번주"/"다음주" 칸으로 옮겨
    붙인다. 자리가 겹치면 덮어쓰지 않고 그 아래에 이어붙인다
    (_append_lines_to_data_cell 참고). 엑셀/PDF/폴더 등 다른 형식은 이
    도구의 대상이 아니므로 조용히 건너뛴다.

    반환: {"merged_files": [...], "no_content_files": [...],
           "skipped_files": [...]}
    - merged_files: 파란색 내용이 하나라도 있어 실제로 옮겨진 소스 파일
    - no_content_files: .hwp/.hwpx이지만 파란색 내용이 하나도 없던 파일
    - skipped_files: 애초에 .hwp/.hwpx가 아니어서(폴더 포함) 건드리지 않은 경로
    """
    hwp_paths = [p for p in source_paths if os.path.isfile(p) and p.lower().endswith((".hwp", ".hwpx"))]
    merged_files, no_content_files = [], []

    for path in hwp_paths:
        content = read_weekly_content_isolated(path)
        this_week = content.get("this_week", [])
        next_week = content.get("next_week", [])
        if not this_week and not next_week:
            no_content_files.append(path)
            continue
        _append_lines_to_data_cell(report.hwp, "이번주", this_week)
        _append_lines_to_data_cell(report.hwp, "다음주", next_week)
        merged_files.append(path)

    return {
        "merged_files": merged_files,
        "no_content_files": no_content_files,
        "skipped_files": [p for p in source_paths if p not in hwp_paths],
    }
```

`if __name__ == "__main__":` 블록을 아래로 갱신:

```python
if __name__ == "__main__":
    _selftest_read_weekly_content_extracts_only_blue_lines()
    _selftest_read_weekly_content_isolated_basic()
    _selftest_read_weekly_content_isolated_missing_file_no_crash()
    _selftest_append_lines_to_data_cell_into_empty_cell()
    _selftest_append_lines_to_data_cell_preserves_existing_content()
    _selftest_append_lines_to_data_cell_empty_lines_is_noop()
    _selftest_merge_weekly_reports_end_to_end()
```

- [x] Step 4: 통과 확인

Run: `python weekly_report_tool.py`
Expected: 7개 self-test 전부 통과, exit 0

- [x] Step 5: 커밋

```bash
git add weekly_report_tool.py
git commit -m "F14: weekly_report_tool.py - merge_weekly_reports() 전체 취합 오케스트레이션 완성"
```

---

### Task 4: chat_assistant.py — merge_weekly_reports 라우팅 + UI 연결

Files

- Modify: `chat_assistant.py`

- [x] Step 1: route_intent 키워드 등록 - 실패하는 테스트 작성

`_selftest_route_intent()` 함수 안, 기존 8번 항목 뒤에 추가:

```python
    # 9) F14: 주간업무보고 취합 도구가 키워드로 잡히는지 확인
    weekly_choice = _route_by_keywords("옆 한글파일로 옮겨줘")
    assert weekly_choice == "merge_weekly_reports", weekly_choice
    print("route_intent 통과 (주간보고 취합):", weekly_choice)
```

- [x] Step 2: 실패 확인

Run: `python chat_assistant.py --selftest`
Expected: `AssertionError`(`merge_weekly_reports`를 아직 아무도 반환 안 함)

- [x] Step 3: 최소 구현

`_TOOLS` 리스트에 다섯 번째 항목 추가(`insert_numbering` 항목 뒤):

```python
    {
        "type": "function",
        "function": {
            "name": "merge_weekly_reports",
            "description": "여러 사람이 첨부한 주간업무보고 문서에서 파란색으로 쓴 내용만 뽑아, 지금 열려있는 대상 문서의 이번주/다음주 칸으로 옮겨 붙인다",
            "parameters": {"type": "object", "properties": {}, "required": []},
        },
    },
```

`_NUMBERING_KEYWORDS` 아래에 추가:

```python
_WEEKLY_MERGE_KEYWORDS = ["옆 한글파일로", "주간보고 취합", "주간업무보고 취합", "취합해"]
```

`route_intent()`(정확히는 `_route_by_keywords()`) 함수 안, `_NUMBERING_KEYWORDS`
체크 블록 뒤, `return None` 앞에 추가:

```python
    if any(keyword in cleaned_message for keyword in _WEEKLY_MERGE_KEYWORDS):
        return "merge_weekly_reports"
```

- [x] Step 4: 통과 확인

Run: `python chat_assistant.py --selftest`
Expected: 전부 통과, exit 0

- [x] Step 5: `_on_submit`에서 실제 취합 실행하도록 연결

`_on_submit()`의 `elif tool_name == "insert_numbering":` 블록 뒤에 추가:

```python
            elif tool_name == "merge_weekly_reports":
                if not self.source_paths:
                    self._log("먼저 취합할 주간업무보고 문서들을 '+'로 첨부해주세요.")
                else:
                    from weekly_report_tool import merge_weekly_reports
                    result = self._run_tool_safely(merge_weekly_reports, self.report, self.source_paths)
                    if result is not None:
                        merged_names = ", ".join(os.path.basename(p) for p in result["merged_files"])
                        lines = []
                        if result["merged_files"]:
                            lines.append(f"{len(result['merged_files'])}건 취합했어요: {merged_names}")
                        if result["no_content_files"]:
                            names = ", ".join(os.path.basename(p) for p in result["no_content_files"])
                            lines.append(f"파란색 내용이 없어 건너뜀: {names}")
                        if result["skipped_files"]:
                            names = ", ".join(os.path.basename(p) for p in result["skipped_files"])
                            lines.append(f"한글 문서가 아니라 건너뜀: {names}")
                        self._log("\n".join(lines) if lines else "취합할 내용이 없었어요.", role="success")
```

(`_run_tool_safely()`는 Task 3/4(F12 2단계)에서 이미 만들어진 공통 헬퍼로,
예외를 잡아 `self._log(f"오류가 발생했습니다 - {e}", role="error")`로
알려주고 `None`을 반환한다 — 새 코드를 추가할 필요 없음, 그대로 재사용.)

- [x] Step 6: 직접 실행으로 확인 (사람이 볼 수 없는 야간이므로, 코드 레벨로
      예외 없이 라우팅되는지만 확인 - 실제 GUI 클릭 흐름은 아침에 사람이 확인)

Run: `python chat_assistant.py --selftest`
Expected: 이미 Step 4에서 확인됨 - 별도 실행 불필요(이 Task는 팝업이 아니라
즉시 실행 흐름이라 Task 3/4(F12 2단계)의 "팝업 인스턴스화 확인" 절차가
해당되지 않음).

- [x] Step 7: 커밋

```bash
git add chat_assistant.py
git commit -m "F14: merge_weekly_reports 도구 등록 + 채팅 라우팅/실행 연결"
```

---

### Task 5: fit_to_page_tool.py — 행간→자간→글자크기 순 1페이지 맞춤 알고리즘

Files

- Create: `fit_to_page_tool.py`

- [x] Step 1: 실패하는 테스트 작성

`fit_to_page_tool.py`를 새로 만들고 아래 내용을 작성:

```python
"""fit_to_page_tool.py — 문서 전체를 1페이지에 맞추기 위해 행간→자간→
글자크기 순서로 아주 조금씩 줄여가며 페이지 수를 확인하는 도구.
verify_tool.py/table_tool.py처럼 도구 하나당 파일 하나 관례를 따른다.

실측 확인(2026-09-07): hwp.PageCount는 SelectAll() 후
hwp.set_linespacing()/hwp.set_font(Spacing=...)/hwp.set_font(Height=...)를
적용하면 바로 갱신되어 읽힌다. **자간 API는 사전 확인 내용과 실제로
달랐다** - hwp.set_para(Spacing=...)라는 API 자체가 pyhwpx에 없다
(set_para()의 실제 파라미터는 Condense 등이고 Spacing이 없음, pyhwpx
소스(core.py set_para 시그니처)로 직접 확인함). 자간은
hwp.set_font(Spacing=값)(-50~50, set_font()의 파라미터)가 맞다. 글자크기는
"축소확대%"인 Size(10~250)가 아니라, 포인트 단위로 직접 지정하는 Height
파라미터를 쓴다 - 현재 값은 hwp.HwpUnitToPoint(hwp.CharShape.Item("Height"))
로 읽는다.

칼리브레이션(42개 짧은 문단으로 만든, 1페이지를 살짝 넘는 문서 기준 직접
측정): 행간 160%→150%(2단계)만으로 1페이지로 돌아옴 - "아주 조금 넘친"
현실적인 경우엔 첫 단계(행간)만으로 대부분 해결된다는 뜻. 반면 훨씬 많이
넘치는 문서(문장을 3배로 늘린 문단 40개, 3페이지)에서는 행간을 하한(130%)
까지 줄여도, 자간을 하한(-20)까지 줄여도 1페이지가 안 되고 글자크기까지
줄여야 하는 경우도 실측으로 확인했다 - 그래서 세 단계 전부가 실제로
필요할 수 있다."""

_LINESPACING_START = 160
_LINESPACING_FLOOR = 130
_LINESPACING_STEP = 5

_SPACING_START = 0
_SPACING_FLOOR = -20
_SPACING_STEP = 3

_FONT_SIZE_FLOOR = 8.0
_FONT_SIZE_STEP = 0.5


def fit_to_one_page(
    report,
    linespacing_floor: int = _LINESPACING_FLOOR,
    spacing_floor: int = _SPACING_FLOOR,
    font_size_floor: float = _FONT_SIZE_FLOOR,
) -> dict:
    """report(HwpReport)의 문서 전체를 1페이지에 맞춘다. 이미 1페이지면
    아무 것도 안 하고 바로 성공을 반환한다. 행간을 linespacing_floor까지,
    그래도 안 되면 자간을 spacing_floor까지, 그래도 안 되면 글자크기를
    font_size_floor까지 조금씩 줄여가며 매 단계마다 PageCount를 확인한다.
    세 단계를 전부 시도해도 1페이지가 안 되면 실패로 알린다(문서를 건드린
    상태 그대로 둔다 - 자동저장 안 하므로 Ctrl+Z로 되돌릴 수 있음).
    linespacing_floor/spacing_floor/font_size_floor 파라미터는 기본값을
    쓰는 일반 호출에서는 그대로 두면 되고, self-test에서 특정 단계만
    강제로 시험하기 위해 좁혀서 넘길 수 있게 열어뒀다."""
    hwp = report.hwp
    if hwp.PageCount <= 1:
        return {"fitted": True, "method": "already_one_page", "page_count": hwp.PageCount}

    linespacing = _LINESPACING_START
    while linespacing > linespacing_floor and hwp.PageCount > 1:
        linespacing -= _LINESPACING_STEP
        hwp.SelectAll()
        hwp.set_linespacing(linespacing, method="Percent")
        if hwp.PageCount <= 1:
            hwp.Cancel()
            return {"fitted": True, "method": "linespacing", "page_count": hwp.PageCount}

    spacing = _SPACING_START
    while spacing > spacing_floor and hwp.PageCount > 1:
        spacing -= _SPACING_STEP
        hwp.SelectAll()
        hwp.set_font(Spacing=spacing)
        if hwp.PageCount <= 1:
            hwp.Cancel()
            return {"fitted": True, "method": "spacing", "page_count": hwp.PageCount}

    height = hwp.HwpUnitToPoint(hwp.CharShape.Item("Height"))
    while height > font_size_floor and hwp.PageCount > 1:
        height -= _FONT_SIZE_STEP
        hwp.SelectAll()
        hwp.set_font(Height=height)
        if hwp.PageCount <= 1:
            hwp.Cancel()
            return {"fitted": True, "method": "font_size", "page_count": hwp.PageCount}

    hwp.Cancel()
    return {"fitted": False, "method": None, "page_count": hwp.PageCount}


def _selftest_fit_to_one_page_already_one_page_is_noop():
    """이미 1페이지인 문서는 아무 것도 바꾸지 않고 바로 성공해야 한다."""
    import time
    from pyhwpx import Hwp
    from hwp_report import HwpReport
    import os, tempfile

    path = os.path.join(tempfile.gettempdir(), "_test_맞춤_이미1페이지.hwp")
    setup = Hwp(visible=False, new=True)
    setup.insert_text("한 페이지 문서입니다.")
    setup.save_as(path)
    setup.quit()
    time.sleep(2)

    report = None
    try:
        report = HwpReport(path)
        result = fit_to_one_page(report)
        assert result == {"fitted": True, "method": "already_one_page", "page_count": 1}, result
        print("fit_to_one_page(이미 1페이지) 통과:", result)
    finally:
        if report is not None:
            report.close(save=False)
        os.remove(path)


if __name__ == "__main__":
    _selftest_fit_to_one_page_already_one_page_is_noop()
```

- [x] Step 2: 실패 확인 (신규 파일이므로, 함수 정의 없이 self-test만 먼저
      작성했다면 여기서 `NameError` 확인 — 위에는 함수까지 포함해 한 번에
      제시했으나 실제 작업 시엔 self-test부터 작성해 실패를 먼저 본다)

Run: `python fit_to_page_tool.py`
Expected: 함수 정의를 아직 추가하지 않은 상태라면 `NameError: name 'fit_to_one_page' is not defined`

- [x] Step 3: 통과 확인 (함수 정의 추가 후)

Run: `python fit_to_page_tool.py`
Expected: `fit_to_one_page(이미 1페이지) 통과: {...}` 출력, exit 0

- [x] Step 4: 실제로 페이지를 줄여야 하는 경우 - 실패하는 테스트 작성

`if __name__ == "__main__":` 블록 위에 추가:

```python
def _selftest_fit_to_one_page_default_floors_uses_linespacing():
    """1페이지를 살짝 넘는 문서(42개 짧은 문단, 직접 측정으로 확인된
    경계값)는 기본 floor 설정에서 행간 조정만으로 1페이지가 돼야 한다."""
    from pyhwpx import Hwp
    from hwp_report import HwpReport
    import os, tempfile, time

    path = os.path.join(tempfile.gettempdir(), "_test_맞춤_행간.hwp")
    setup = Hwp(visible=False, new=True)
    for i in range(42):
        setup.insert_text(f"{i}번째 문단입니다. 이것은 페이지 채우기용 테스트 문장입니다.")
        setup.HAction.Run("BreakPara")
    setup.save_as(path)
    setup.quit()
    time.sleep(2)

    report = None
    try:
        report = HwpReport(path)
        assert report.hwp.PageCount == 2, report.hwp.PageCount
        result = fit_to_one_page(report)
        assert result["fitted"] is True, result
        assert result["method"] == "linespacing", result
        assert result["page_count"] == 1, result
        print("fit_to_one_page(행간 조정만으로 해결) 통과:", result)
    finally:
        if report is not None:
            report.close(save=False)
        os.remove(path)


def _selftest_fit_to_one_page_falls_through_to_font_size():
    """행간과 자간 조정 여지를 강제로 좁혀두면(linespacing_floor=158로 한
    단계만, spacing_floor=3으로 0단계) 글자크기 조정 단계까지 내려가
    성공해야 한다 - 세 단계 사이 폴백(fallthrough) 로직 자체를 검증한다."""
    from pyhwpx import Hwp
    from hwp_report import HwpReport
    import os, tempfile, time

    path = os.path.join(tempfile.gettempdir(), "_test_맞춤_글자크기.hwp")
    setup = Hwp(visible=False, new=True)
    for i in range(42):
        setup.insert_text(f"{i}번째 문단입니다. 이것은 페이지 채우기용 테스트 문장입니다.")
        setup.HAction.Run("BreakPara")
    setup.save_as(path)
    setup.quit()
    time.sleep(2)

    report = None
    try:
        report = HwpReport(path)
        result = fit_to_one_page(report, linespacing_floor=158, spacing_floor=3)
        assert result["fitted"] is True, result
        assert result["method"] == "font_size", result
        assert result["page_count"] == 1, result
        print("fit_to_one_page(글자크기까지 폴백) 통과:", result)
    finally:
        if report is not None:
            report.close(save=False)
        os.remove(path)


def _selftest_fit_to_one_page_all_floors_too_tight_fails_honestly():
    """세 단계 모두 여지를 0으로 막아두면(각 floor를 시작값과 같게) 문서를
    건드리지 못하고 fitted=False를 정직하게 반환해야 한다(무한루프 없이
    종료되는지도 함께 확인)."""
    from pyhwpx import Hwp
    from hwp_report import HwpReport
    import os, tempfile, time

    path = os.path.join(tempfile.gettempdir(), "_test_맞춤_실패.hwp")
    setup = Hwp(visible=False, new=True)
    for i in range(42):
        setup.insert_text(f"{i}번째 문단입니다. 이것은 페이지 채우기용 테스트 문장입니다.")
        setup.HAction.Run("BreakPara")
    setup.save_as(path)
    setup.quit()
    time.sleep(2)

    report = None
    try:
        report = HwpReport(path)
        result = fit_to_one_page(report, linespacing_floor=160, spacing_floor=0, font_size_floor=10.0)
        assert result == {"fitted": False, "method": None, "page_count": 2}, result
        print("fit_to_one_page(여지 없음, 정직한 실패) 통과:", result)
    finally:
        if report is not None:
            report.close(save=False)
        os.remove(path)
```

- [x] Step 5: 실패 확인

Run: `python fit_to_page_tool.py`
Expected: 아직 `if __name__ == "__main__":` 블록에 새 테스트들을 등록하지
않았으므로 이 시점에는 실행되지 않음 — 바로 아래 Step에서 등록 후 실행해
통과를 확인한다(이 세 함수는 이미 검증된 알고리즘을 그대로 호출하므로
"실패 확인"은 등록 누락 상태에서의 미실행으로 갈음한다).

- [x] Step 6: `if __name__ == "__main__":` 블록 갱신 + 통과 확인

```python
if __name__ == "__main__":
    _selftest_fit_to_one_page_already_one_page_is_noop()
    _selftest_fit_to_one_page_default_floors_uses_linespacing()
    _selftest_fit_to_one_page_falls_through_to_font_size()
    _selftest_fit_to_one_page_all_floors_too_tight_fails_honestly()
```

Run: `python fit_to_page_tool.py`
Expected: 4개 self-test 전부 통과, exit 0 (실측 확인: 이 정확한 파라미터
조합으로 이미 이 계획서 작성 과정에서 4회 전부 통과를 직접 실행으로
확인했음 — 계획서의 "사전조사 실측 확인 요약" 6번 참고)

- [x] Step 7: 커밋

```bash
git add fit_to_page_tool.py
git commit -m "F14: fit_to_page_tool.py 추가 - 행간→자간→글자크기 순 1페이지 맞춤"
```

---

### Task 6: chat_assistant.py — fit_to_one_page 라우팅 + UI 연결

Files

- Modify: `chat_assistant.py`

- [x] Step 1: route_intent 키워드 등록 - 실패하는 테스트 작성

`_selftest_route_intent()`에 추가:

```python
    # 10) F14: 한 페이지 맞춤 도구가 키워드로 잡히는지 확인
    fit_choice = _route_by_keywords("이거 한 페이지에 맞춰줘")
    assert fit_choice == "fit_to_one_page", fit_choice
    print("route_intent 통과 (한 페이지 맞춤):", fit_choice)
```

- [x] Step 2: 실패 확인

Run: `python chat_assistant.py --selftest`
Expected: `AssertionError`

- [x] Step 3: 최소 구현

`_TOOLS`에 여섯 번째 항목 추가:

```python
    {
        "type": "function",
        "function": {
            "name": "fit_to_one_page",
            "description": "지금 열려있는 문서를 행간/자간/글자크기를 조금씩 줄여 1페이지에 맞춘다",
            "parameters": {"type": "object", "properties": {}, "required": []},
        },
    },
```

`_WEEKLY_MERGE_KEYWORDS` 아래에 추가:

```python
_FIT_TO_PAGE_KEYWORDS = ["한 페이지에 맞춰", "한페이지에 맞춰", "한 장에 맞춰", "페이지 맞춤", "쪽맞춤"]
```

`_route_by_keywords()`에서 `_WEEKLY_MERGE_KEYWORDS` 체크 블록 뒤, `return None`
앞에 추가:

```python
    if any(keyword in cleaned_message for keyword in _FIT_TO_PAGE_KEYWORDS):
        return "fit_to_one_page"
```

- [x] Step 4: 통과 확인

Run: `python chat_assistant.py --selftest`
Expected: 전부 통과, exit 0

- [x] Step 5: `_on_submit`에 실행 연결

`elif tool_name == "merge_weekly_reports":` 블록 뒤에 추가:

```python
            elif tool_name == "fit_to_one_page":
                from fit_to_page_tool import fit_to_one_page
                result = self._run_tool_safely(fit_to_one_page, self.report)
                if result is not None:
                    if result["fitted"]:
                        method_label = {
                            "already_one_page": "이미 1페이지였어요",
                            "linespacing": "행간을 줄여서",
                            "spacing": "자간까지 줄여서",
                            "font_size": "글자크기까지 줄여서",
                        }[result["method"]]
                        self._log(f"{method_label} 1페이지로 맞췄어요.", role="success")
                    else:
                        self._log("행간/자간/글자크기를 다 줄여봐도 1페이지에 안 들어가요. 내용을 좀 줄여주세요.", role="error")
```

- [x] Step 6: 커밋

```bash
git add chat_assistant.py
git commit -m "F14: fit_to_one_page 도구 등록 + 채팅 라우팅/실행 연결"
```

---

### Task 7: verify_numbers.py — 컬럼 간 증감률 · 평균 계산

Files

- Modify: `verify_numbers.py`

- [x] Step 1: 실패하는 테스트 작성

`_selftest_compute_column_sums_ignores_boolean_column()` 함수 뒤에 추가:

```python
def _selftest_compute_column_growth_rate():
    rows = [{"작년": 100, "올해": 150}, {"작년": 200, "올해": 250}]
    result = compute_column_growth_rate(
        rows, source_file="원본.xlsx", sheet="Sheet1", from_col="작년", to_col="올해"
    )
    assert len(result) == 1, result
    assert result[0]["normalized"] == "33.33", result
    assert result[0]["type"] == "amount", result
    print("compute_column_growth_rate 통과:", result)


def _selftest_compute_column_growth_rate_zero_base_returns_empty():
    """분모(from_col 합계)가 0이면 0으로 나누기를 피해 빈 리스트를 반환해야 한다."""
    rows = [{"작년": 0, "올해": 150}]
    result = compute_column_growth_rate(
        rows, source_file="원본.xlsx", sheet="Sheet1", from_col="작년", to_col="올해"
    )
    assert result == [], result
    print("compute_column_growth_rate(분모 0) 통과: 빈 리스트")


def _selftest_compute_column_average():
    rows = [{"예산": 100}, {"예산": 200}, {"예산": 300}]
    result = compute_column_average(rows, source_file="원본.xlsx", sheet="Sheet1", col="예산")
    assert len(result) == 1, result
    assert result[0]["normalized"] == "200", result
    print("compute_column_average 통과:", result)


def _selftest_compute_column_average_rounds_repeating_decimal():
    """나눗셈이 딱 떨어지지 않아도(100/3 형태) 소수점 둘째자리로 반올림돼야
    한다 - 사람이 읽는 보고서 문장과 대조하기 위함(끝없는 소수는 무의미함)."""
    rows = [{"값": 30}, {"값": 30}, {"값": 40}]
    result = compute_column_average(rows, source_file="원본.xlsx", sheet="Sheet1", col="값")
    assert result[0]["normalized"] == "33.33", result
    print("compute_column_average(반복소수 반올림) 통과:", result)
```

- [x] Step 2: 실패 확인

Run: `python verify_numbers.py`
Expected: `NameError: name 'compute_column_growth_rate' is not defined`
(아직 `if __name__ == "__main__":`에 등록하지 않았다면 이 함수들은 실행되지
않으므로, 먼저 파일 맨 아래 `if __name__ == "__main__":` 블록에 새 4개
self-test 호출을 추가한 뒤 실행해야 `NameError`가 실제로 발생한다.)

- [x] Step 3: 최소 구현

`compute_column_sums()` 함수 뒤(`_selftest_compute_column_sums()` 함수 앞)에 추가:

```python
def compute_column_growth_rate(rows: list[dict], source_file: str, sheet: str,
                                from_col: str, to_col: str) -> list[dict]:
    """지정한 두 컬럼(from_col 합계 대비 to_col 합계)의 증감률(%)을 계산해
    정답 풀에 추가할 항목 하나로 반환한다. compute_column_sums()와 같은
    반환 dict 형태(type/normalized/raw/source_file/location)를 따르며, 그
    함수가 이미 계산한 컬럼별 합계를 그대로 재사용한다(같은 "숫자 컬럼
    판별" 로직을 중복 구현하지 않기 위함). from_col/to_col이 숫자 컬럼이
    아니거나(rows[0]에 없거나 일부 행에 비숫자 값이 섞여 있으면
    compute_column_sums가 애초에 그 컬럼을 대상에서 뺀다), from_col 합계가
    0이면(0으로 나누기 방지) 빈 리스트를 반환한다."""
    if not rows:
        return []
    sums = {r["raw"].removesuffix(" 합계"): Decimal(r["normalized"])
            for r in compute_column_sums(rows, source_file, sheet)}
    if from_col not in sums or to_col not in sums or sums[from_col] == 0:
        return []
    growth = (sums[to_col] - sums[from_col]) / sums[from_col] * 100
    normalized = _decimal_to_normalized_str(growth.quantize(Decimal("0.01")))
    return [{
        "type": "amount", "normalized": normalized,
        "raw": f"{from_col}→{to_col} 증감률",
        "source_file": source_file, "location": f"{sheet}!{from_col}->{to_col}(증감률)",
    }]


def compute_column_average(rows: list[dict], source_file: str, sheet: str, col: str) -> list[dict]:
    """지정한 컬럼의 평균을 계산해 정답 풀에 추가할 항목 하나로 반환한다.
    compute_column_sums()와 같은 반환 dict 형태를 따르며, col이 숫자
    컬럼이 아니면(compute_column_sums가 대상에서 뺐으면) 빈 리스트를
    반환한다."""
    if not rows:
        return []
    sums = compute_column_sums(rows, source_file, sheet)
    matching = [r for r in sums if r["raw"] == f"{col} 합계"]
    if not matching:
        return []
    total = Decimal(matching[0]["normalized"])
    average = total / len(rows)
    normalized = _decimal_to_normalized_str(average.quantize(Decimal("0.01")))
    return [{
        "type": "amount", "normalized": normalized, "raw": f"{col} 평균",
        "source_file": source_file, "location": f"{sheet}!{col}(평균)",
    }]
```

`if __name__ == "__main__":` 블록에서 `_selftest_compute_column_sums_ignores_boolean_column()`
호출 바로 뒤에 추가:

```python
    _selftest_compute_column_growth_rate()
    _selftest_compute_column_growth_rate_zero_base_returns_empty()
    _selftest_compute_column_average()
    _selftest_compute_column_average_rounds_repeating_decimal()
```

- [x] Step 4: 통과 확인

Run: `python verify_numbers.py`
Expected: 기존 self-test 전부 + 신규 4개 전부 통과, exit 0

- [x] Step 5: 커밋

```bash
git add verify_numbers.py
git commit -m "F14: verify_numbers.py - compute_column_growth_rate/compute_column_average 추가"
```

---

### Task 8: ignore_list.py — 오탐 무시 목록 JSON 영속화

Files

- Create: `ignore_list.py`

- [x] Step 1: 실패하는 테스트 작성

```python
"""ignore_list.py — 숫자검증에서 "이건 괜찮아, 무시해"로 지정한 값을 기억해
다음 검증부터 오류(빨강)로 표시하지 않게 하는 오탐 학습 도구.
speed_tracker.py와 동일한 순수 함수 + 로컬 JSON 파일 조합 관례를 따른다."""
import json
import os

DEFAULT_IGNORE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ignore_list.json")


def _load_raw_list(path: str) -> list[str]:
    if not os.path.exists(path):
        return []
    with open(path, "r", encoding="utf-8") as f:
        try:
            return json.load(f)
        except json.JSONDecodeError:
            return []


def _save_raw_list(values: list[str], path: str) -> None:
    with open(path, "w", encoding="utf-8") as f:
        json.dump(values, f, ensure_ascii=False)


def record_ignored_value(raw: str, path: str = DEFAULT_IGNORE_PATH) -> None:
    """raw 값을 무시 목록에 추가한다(중복 저장 방지). speed_tracker.py의
    record_call_speed()와 달리 최근 N건 제한은 두지 않는다 - 무시 목록은
    "한 번 안전하다고 확인된 값"의 누적 화이트리스트라서, 오래된 항목이라도
    계속 유효해야 한다(속도 기록처럼 최신 것만 의미있는 데이터가 아님)."""
    values = _load_raw_list(path)
    if raw not in values:
        values.append(raw)
        _save_raw_list(values, path)


def load_ignored_values(path: str = DEFAULT_IGNORE_PATH) -> list[str]:
    """무시 목록에 등록된 값들을 리스트로 반환한다. 파일이 없으면 빈 리스트."""
    return _load_raw_list(path)


def _selftest_record_and_load_ignored_values():
    import tempfile
    path = os.path.join(tempfile.gettempdir(), "_test_ignore_list.json")
    if os.path.exists(path):
        os.remove(path)
    try:
        assert load_ignored_values(path) == []
        record_ignored_value("9999999", path=path)
        record_ignored_value("1850000", path=path)
        assert load_ignored_values(path) == ["9999999", "1850000"]
        print("record_ignored_value + load_ignored_values 통과")
    finally:
        if os.path.exists(path):
            os.remove(path)


if __name__ == "__main__":
    _selftest_record_and_load_ignored_values()
```

- [x] Step 2: 실패 확인

이 파일은 신규 생성이라, 실제 작업 시엔 먼저 self-test 함수와
`if __name__ == "__main__":` 블록만 작성해 `python ignore_list.py`를 실행하면
`NameError: name 'load_ignored_values' is not defined`로 실패함을 먼저 확인한
뒤, 위 함수 정의들을 추가한다.

- [x] Step 3: 통과 확인 (함수 정의 추가 후)

Run: `python ignore_list.py`
Expected: `record_ignored_value + load_ignored_values 통과` 출력, exit 0

- [x] Step 4: 중복 등록 방지 + 파일 없을 때 빈 리스트 회귀 테스트 추가

`if __name__ == "__main__":` 블록 위에 추가:

```python
def _selftest_record_ignored_value_deduplicates():
    """같은 값을 두 번 기록해도 목록에는 한 번만 남아야 한다."""
    import tempfile
    path = os.path.join(tempfile.gettempdir(), "_test_ignore_list_중복.json")
    if os.path.exists(path):
        os.remove(path)
    try:
        record_ignored_value("9999999", path=path)
        record_ignored_value("9999999", path=path)
        assert load_ignored_values(path) == ["9999999"], load_ignored_values(path)
        print("record_ignored_value(중복 방지) 통과")
    finally:
        if os.path.exists(path):
            os.remove(path)


def _selftest_load_ignored_values_missing_file_returns_empty():
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_존재하지_않는_ignore_list.json")
    assert load_ignored_values(path) == []
    print("load_ignored_values(파일없음) 통과")
```

`if __name__ == "__main__":` 블록 갱신:

```python
if __name__ == "__main__":
    _selftest_record_and_load_ignored_values()
    _selftest_record_ignored_value_deduplicates()
    _selftest_load_ignored_values_missing_file_returns_empty()
```

- [x] Step 5: 통과 확인

Run: `python ignore_list.py`
Expected: 3개 self-test 전부 통과, exit 0

- [x] Step 6: `.gitignore`에 데이터 파일 추가 확인

`ignore_list.json`은 실행 중 생성되는 사용자별 데이터 파일이라
`speed_log.json`과 같은 성격이다. `.gitignore`를 확인해 `speed_log.json`이
이미 등록돼 있으면 `ignore_list.json`도 같은 방식으로 추가한다(둘 다
없다면 이 Step은 생략 - 저장소 관례를 따름).

- [x] Step 7: 커밋

```bash
git add ignore_list.py
git commit -m "F14: ignore_list.py 추가 - 숫자검증 오탐 무시 목록 JSON 영속화"
```

---

### Task 9: verify_tool.py — run_verification()이 무시 목록을 반영하도록 확장

Files

- Modify: `verify_tool.py`

- [x] Step 1: 실패하는 테스트 작성

`_selftest_run_verification()` 함수 뒤에 추가:

```python
def _selftest_run_verification_skips_ignored_mismatch():
    """무시 목록에 등록된 값은 실제로 틀렸어도(원본과 불일치) 더 이상
    오류(빨강)로 표시되지 않고 mismatch_count/summary/mismatch_items에서
    빠져야 한다."""
    test_dir = os.path.join(tempfile.gettempdir(), "_test_원본_무시목록")
    os.makedirs(test_dir, exist_ok=True)
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Sheet1"
    ws["A1"] = "예산"; ws["A2"] = 1850000
    wb.save(os.path.join(test_dir, "원본.xlsx"))

    from pyhwpx import Hwp
    report_path = os.path.join(tempfile.gettempdir(), "_test_보고서_무시목록.hwp")
    setup = Hwp(visible=False, new=True)
    setup.insert_text("예산은 185만원이며, 오타1은 9999999원, 오타2는 8888888원입니다")
    setup.save_as(report_path)
    setup.quit()

    ignore_path = os.path.join(tempfile.gettempdir(), "_test_ignore_list_검증.json")
    if os.path.exists(ignore_path):
        os.remove(ignore_path)
    from ignore_list import record_ignored_value
    record_ignored_value("9999999", path=ignore_path)  # 오타1만 무시 목록에 등록

    report = None
    try:
        report = HwpReport(report_path)
        result = run_verification(report, source_paths=[test_dir], default_year=2026,
                                   ignore_list_path=ignore_path)
        assert result["mismatch_count"] == 1, result  # 오타2만 남아야 함
        assert "9999999" not in result["summary"], result["summary"]
        assert "8888888" in result["summary"], result["summary"]
        assert "9999999" not in result["mismatch_items"], result["mismatch_items"]
        print("run_verification(무시 목록 반영) 통과:", result["summary"])
    finally:
        if report is not None:
            report.close(save=False)
        import shutil
        shutil.rmtree(test_dir)
        os.remove(report_path)
        if os.path.exists(ignore_path):
            os.remove(ignore_path)
```

- [x] Step 2: 실패 확인

Run: `python verify_tool.py`
Expected: `TypeError: run_verification() got an unexpected keyword argument 'ignore_list_path'`
(아직 파라미터가 없으므로 — 이 self-test를 `if __name__ == "__main__":`에
등록한 뒤 실행)

- [x] Step 3: 최소 구현

파일 상단 import에 추가:

```python
from ignore_list import load_ignored_values, DEFAULT_IGNORE_PATH
```

`run_verification()` 시그니처를 변경:

```python
def run_verification(report: HwpReport, source_paths: list[str], default_year: int,
                      ignore_list_path: str = DEFAULT_IGNORE_PATH) -> dict:
```

함수 docstring의 "반환 dict 계약" 문단 뒤(또는 적당한 위치)에 아래 설명을
덧붙인다:

```
    (F14) ignore_list_path에 등록된 값(raw 문자열 그대로 비교)은 원본과
    실제로 불일치해도 mismatches에서 제외한다 - "이건 괜찮아, 무시해"로
    한 번 확인한 값은 다음부터 오류로 표시하지 않는다는 오탐 학습 기능.
    기본값은 ignore_list.py의 기본 경로를 그대로 쓴다.
```

`categorized = categorize_values(report_values, answer_pool)` 다음 줄
(`mismatches = categorized["mismatches"]` 바로 뒤)에 추가:

```python
    ignored_raws = set(load_ignored_values(ignore_list_path))
    if ignored_raws:
        mismatches = [m for m in mismatches if m["raw"] not in ignored_raws]
```

- [x] Step 4: 통과 확인

Run: `python verify_tool.py`
Expected: 기존 self-test(`_selftest_run_verification`) + 신규 self-test 전부
통과, exit 0

- [x] Step 5: 커밋

```bash
git add verify_tool.py
git commit -m "F14: run_verification()에 무시 목록(ignore_list_path) 반영 - 오탐 학습"
```

---

### Task 10: chat_assistant.py — "N번째는 무시해" 결정론적 분기

Files

- Modify: `chat_assistant.py`

- [x] Step 1: parse_ignore_index() - 실패하는 테스트 작성

`_selftest_parse_goto_index()` 함수 뒤에 추가:

```python
def _selftest_parse_ignore_index():
    assert parse_ignore_index("2번째는 무시해") == 2
    assert parse_ignore_index("3번째는 그냥 괜찮아") == 3
    assert parse_ignore_index("3번째로 가줘") is None  # 무시/괜찮 키워드 없음 - goto와 구분
    assert parse_ignore_index("숫자 검증해줘") is None
    print("parse_ignore_index 통과")
```

- [x] Step 2: 실패 확인

Run: `python chat_assistant.py --selftest`
Expected: `NameError: name 'parse_ignore_index' is not defined`

- [x] Step 3: 최소 구현

`parse_goto_index()` 함수 뒤에 추가:

```python
def parse_ignore_index(text: str) -> int | None:
    """"2번째는 무시해"류 입력에서 순서 번호(1-based)를 뽑는다.
    parse_goto_index()와 같은 결정론적 정규식 방식이지만, "무시"/"괜찮"
    키워드가 함께 있어야만 매치된다 - 이 조건이 없으면 "3번째로 가줘"
    (순수 이동 요청)까지 무시 요청으로 잘못 인식하게 된다. _on_submit()은
    반드시 이 함수를 parse_goto_index()보다 먼저 확인해야 한다 -
    parse_goto_index()의 정규식(r'(\\d+)번째')은 "2번째는 무시해"에도
    매치되므로, 순서를 바꾸면 무시 요청이 이동 요청으로 잘못 처리된다."""
    m = re.search(r'(\d+)번째.*(?:무시|괜찮)', text)
    return int(m.group(1)) if m else None
```

`if __name__ == "__main__":` 블록 갱신:

```python
if __name__ == "__main__":
    import sys
    if "--selftest" in sys.argv:
        _selftest_route_intent()
        _selftest_parse_goto_index()
        _selftest_parse_ignore_index()
    else:
        app = ChatAssistant()
        app.mainloop()
```

- [x] Step 4: 통과 확인

Run: `python chat_assistant.py --selftest`
Expected: 전부 통과, exit 0

- [x] Step 5: `_on_submit`에 무시 분기 연결 (goto_index 분기보다 먼저)

파일 상단 import에 추가:

```python
from ignore_list import record_ignored_value
```

`_on_submit()`의 `goto_index = parse_goto_index(text)` 줄 **바로 앞**에 삽입
(반드시 goto_index 판단보다 먼저 실행돼야 함 - 위 docstring 근거):

```python
        # (F14) "N번째는 무시해"도 parse_goto_index()와 같은 이유로
        # route_intent()의 LLM 라우팅을 거치지 않고 결정론적으로 처리한다.
        # parse_goto_index()보다 반드시 먼저 확인해야 한다 - 그 정규식이
        # "2번째는 무시해"에도 매치되기 때문(parse_ignore_index docstring 참고).
        ignore_index = parse_ignore_index(text)
        if ignore_index is not None:
            if not self._last_mismatch_items:
                self._log("먼저 숫자 검증을 실행해주세요.")
            elif 1 <= ignore_index <= len(self._last_mismatch_items):
                target = self._last_mismatch_items[ignore_index - 1]
                record_ignored_value(target)
                self._log(f"{ignore_index}번째 항목('{target}')을 앞으로 무시할게요.", role="assistant")
            else:
                self._log(f"총 {len(self._last_mismatch_items)}건 중 {ignore_index}번째는 없어요.")
            return

```

- [x] Step 6: 커밋

```bash
git add chat_assistant.py
git commit -m "F14: 'N번째는 무시해' 결정론적 분기 추가 - 오탐 학습 UI 연결"
```

---

### Task 11: attachment_preview.py — 첨부문서 미리보기 생성 + 용량 관리

Files

- Create: `attachment_preview.py`

- [x] Step 1: 실패하는 테스트 작성 (HWP 미리보기, 격리 실행)

```python
"""attachment_preview.py — "+"로 첨부한 문서를 첨부 즉시 작은 미리보기로
보여주기 위한 도구. HWP/HWPX는 이미지 미리보기(1페이지만), 엑셀/PDF는
텍스트 미리보기(상위 몇 줄)를 만든다. source_reader.py의
read_hwp_source_isolated()와 같은 이유로, HWP 미리보기는 반드시 별도
프로세스에서 만든다(같은 프로세스 안의 다른 Hwp 인스턴스가 채팅 세션이
이미 열어둔 HwpReport의 COM 연결을 깨뜨리는 pyhwpx의 한계, source_reader.py
주석 참고).

실측 확인(2026-09-07): create_page_image(path, pgno=1, resolution=60,
format="gif")로 만든 1페이지 미리보기는 문서 한 장 기준 수 KB 수준이다
(직접 측정: 짧은 문서 1장 gif@60dpi 약 6.8KB) - bmp(기본 포맷)는 무압축이라
96dpi 한 장에도 3.5MB나 나가 첨부할 때마다 용량이 급격히 쌓이므로, 이
도구는 반드시 gif+저해상도 조합을 쓴다(사용자가 명시적으로 우려한 용량
관리 요구사항)."""
import glob
import os
import subprocess
import sys
import tempfile

_PREVIEW_DIR = os.path.join(tempfile.gettempdir(), "_hwp_attachment_previews")
_MAX_PREVIEWS = 5
_PREVIEW_RESOLUTION = 60


def _generate_hwp_preview(path: str, out_path: str) -> bool:
    """path(.hwp/.hwpx)의 1페이지를 out_path(gif)로 저장한다. 이 함수 자체는
    현재 프로세스에서 새 Hwp 인스턴스를 만들므로, 채팅 세션이 이미 다른
    Hwp를 열어둔 상태에서 절대 직접 호출하면 안 된다 - 반드시
    generate_hwp_preview_isolated()를 통해 별도 프로세스에서만 실행한다."""
    from pyhwpx import Hwp
    hwp = None
    try:
        hwp = Hwp(visible=False, new=True)
        if not hwp.open(path):
            return False
        return bool(hwp.create_page_image(out_path, pgno=1, resolution=_PREVIEW_RESOLUTION, format="gif"))
    except Exception:
        return False
    finally:
        if hwp is not None:
            hwp.quit()


def generate_hwp_preview_isolated(path: str) -> str | None:
    """_generate_hwp_preview()를 별도 프로세스에서 실행해, 성공하면 저장된
    미리보기 이미지 경로를, 실패하면 None을 반환한다. source_reader.py의
    read_hwp_source_isolated()와 동일한 subprocess 격리 패턴. 호출 성공
    시마다 _prune_old_previews()로 오래된 미리보기를 정리한다."""
    os.makedirs(_PREVIEW_DIR, exist_ok=True)
    out_path = os.path.join(_PREVIEW_DIR, f"{os.path.basename(path)}_{os.getpid()}.gif")
    module_dir = os.path.dirname(os.path.abspath(__file__))
    script = (
        "import sys; sys.path.insert(0, sys.argv[3]); "
        "from attachment_preview import _generate_hwp_preview; "
        "ok = _generate_hwp_preview(sys.argv[1], sys.argv[2]); "
        "sys.exit(0 if ok else 1)"
    )
    try:
        result = subprocess.run(
            [sys.executable, "-c", script, path, out_path, module_dir],
            capture_output=True, text=True, timeout=60,
        )
        if result.returncode != 0 or not os.path.exists(out_path):
            return None
        _prune_old_previews()
        return out_path
    except Exception:
        return None


def _prune_old_previews(keep: int = _MAX_PREVIEWS) -> None:
    """_PREVIEW_DIR에 쌓인 미리보기 이미지 중 오래된 것부터 지워 최근
    keep개만 남긴다(용량 관리 요구사항). 파일 개수가 keep 이하면 아무 것도
    안 한다."""
    if not os.path.isdir(_PREVIEW_DIR):
        return
    files = sorted(glob.glob(os.path.join(_PREVIEW_DIR, "*.gif")), key=os.path.getmtime)
    for old_file in files[:-keep] if len(files) > keep else []:
        try:
            os.remove(old_file)
        except OSError:
            pass


def _selftest_generate_hwp_preview_isolated_creates_small_file():
    """미리보기 이미지가 실제로 생성되고, gif+저해상도 조합 덕분에 충분히
    작은지(용량 관리 요구사항) 확인한다."""
    import time
    from pyhwpx import Hwp

    path = os.path.join(tempfile.gettempdir(), "_test_미리보기_원본.hwp")
    setup = Hwp(visible=False, new=True)
    setup.insert_text("미리보기 테스트 문서입니다.")
    setup.save_as(path)
    setup.quit()
    time.sleep(2)

    try:
        preview_path = generate_hwp_preview_isolated(path)
        assert preview_path is not None
        assert os.path.exists(preview_path)
        size = os.path.getsize(preview_path)
        assert size < 200_000, f"미리보기 용량이 예상보다 큼: {size} bytes"
        print("generate_hwp_preview_isolated 통과: 용량", size, "bytes")
    finally:
        os.remove(path)
        for f in glob.glob(os.path.join(_PREVIEW_DIR, "*.gif")):
            os.remove(f)


if __name__ == "__main__":
    _selftest_generate_hwp_preview_isolated_creates_small_file()
```

- [x] Step 2: 실패 확인

이 파일은 신규 생성이므로, 실제 작업 시엔 먼저 self-test 함수만 작성해
`NameError: name 'generate_hwp_preview_isolated' is not defined`를 먼저
확인한 뒤 함수 정의들을 추가한다.

- [x] Step 3: 통과 확인 (함수 정의 추가 후)

Run: `python attachment_preview.py`
Expected: `generate_hwp_preview_isolated 통과: 용량 ... bytes` (6.8KB 근방,
200KB 미만) 출력, exit 0

- [x] Step 4: 파일없음/엑셀/PDF 텍스트 미리보기 + 정리(prune) 테스트 추가

`if __name__ == "__main__":` 블록 위에 추가:

```python
def _selftest_generate_hwp_preview_isolated_missing_file_returns_none():
    result = generate_hwp_preview_isolated("존재하지_않는_미리보기.hwp")
    assert result is None, result
    print("generate_hwp_preview_isolated(파일없음) 통과")


def _selftest_generate_text_preview_excel():
    import openpyxl
    path = os.path.join(tempfile.gettempdir(), "_test_미리보기.xlsx")
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append(["항목", "금액"])
    ws.append(["인건비", 1000000])
    wb.save(path)
    try:
        preview = generate_text_preview(path, max_lines=5)
        assert "항목" in preview and "인건비" in preview, preview
        print("generate_text_preview(엑셀) 통과:", repr(preview))
    finally:
        os.remove(path)


def _selftest_generate_text_preview_pdf():
    from fpdf import FPDF
    path = os.path.join(tempfile.gettempdir(), "_test_미리보기.pdf")
    pdf = FPDF()
    pdf.add_font("Malgun", fname="C:/Windows/Fonts/malgun.ttf")
    pdf.add_page(); pdf.set_font("Malgun", size=12)
    pdf.cell(0, 10, "예산은 1,850,000원입니다")
    pdf.output(path)
    try:
        preview = generate_text_preview(path, max_lines=5)
        assert "1,850,000" in preview or "1850000" in preview, preview
        print("generate_text_preview(PDF) 통과:", repr(preview))
    finally:
        os.remove(path)


def _selftest_prune_old_previews_keeps_recent_n():
    """미리보기가 keep개보다 많이 쌓이면 오래된 것부터 지워 최근 keep개만
    남아야 한다."""
    os.makedirs(_PREVIEW_DIR, exist_ok=True)
    test_files = []
    for i in range(8):
        p = os.path.join(_PREVIEW_DIR, f"_test_prune_{i}.gif")
        with open(p, "w") as f:
            f.write("dummy")
        os.utime(p, (i, i))  # 오래된 순서대로 mtime을 강제로 다르게 설정
        test_files.append(p)
    try:
        _prune_old_previews(keep=5)
        remaining = sorted(glob.glob(os.path.join(_PREVIEW_DIR, "_test_prune_*.gif")))
        assert len(remaining) == 5, remaining
        # 가장 오래된 3개(_0,_1,_2)는 지워지고 최근 5개(_3~_7)만 남아야 함
        for i in range(3):
            assert os.path.join(_PREVIEW_DIR, f"_test_prune_{i}.gif") not in remaining
        print("_prune_old_previews(최근 5개만 유지) 통과")
    finally:
        for p in test_files:
            if os.path.exists(p):
                os.remove(p)
```

파일 상단, `_generate_hwp_preview()` 함수 앞에 `generate_text_preview()` 함수 추가:

```python
def generate_text_preview(path: str, max_lines: int = 5) -> str:
    """엑셀/PDF의 상위 max_lines줄만 텍스트로 미리보기를 만든다. 이미지이
    아니라 문자열을 그대로 반환한다(채팅 말풍선에 바로 표시할 수 있게).
    지원하지 않는 형식이거나 읽기 실패하면 빈 문자열(예외 없음, 이 모듈의
    다른 함수들과 같은 관례)."""
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext in (".xlsx", ".xls"):
            import openpyxl
            wb = openpyxl.load_workbook(path, data_only=True)
            ws = wb.active
            lines = []
            for row in ws.iter_rows(max_row=max_lines, values_only=True):
                lines.append(", ".join(str(v) for v in row if v is not None))
            return "\n".join(lines)
        if ext == ".pdf":
            import pdfplumber
            with pdfplumber.open(path) as pdf:
                text = pdf.pages[0].extract_text() or ""
            return "\n".join(text.split("\n")[:max_lines])
    except Exception:
        return ""
    return ""
```

`if __name__ == "__main__":` 블록 갱신:

```python
if __name__ == "__main__":
    _selftest_generate_hwp_preview_isolated_creates_small_file()
    _selftest_generate_hwp_preview_isolated_missing_file_returns_none()
    _selftest_generate_text_preview_excel()
    _selftest_generate_text_preview_pdf()
    _selftest_prune_old_previews_keeps_recent_n()
```

- [x] Step 5: 통과 확인

Run: `python attachment_preview.py`
Expected: 5개 self-test 전부 통과, exit 0

- [x] Step 6: 커밋

```bash
git add attachment_preview.py
git commit -m "F14: attachment_preview.py 추가 - HWP 이미지/엑셀·PDF 텍스트 미리보기 + 용량 관리"
```

---

### Task 12: chat_assistant.py — 첨부 시 자동 미리보기 표시

Files

- Modify: `chat_assistant.py`

- [x] Step 1: 미리보기 헬퍼 - 실패하는 테스트는 생략 (GUI 이미지 렌더링은
      무인 self-test로 검증하기 어려움 - F12 2단계 팝업 Task들과 같은
      이유로, 예외 없이 호출되는지만 코드 레벨로 확인하고 실제 렌더링은
      아침에 사람이 확인한다). 대신 `_build_preview_message()`라는 순수
      함수(위젯을 직접 그리지 않고, 어떤 종류의 미리보기를 보여줄지와 표시할
      텍스트/이미지경로를 결정하는 로직만 분리)를 만들어 그 판단 로직만
      결정론적으로 테스트한다.

`_selftest_route_intent()` 함수 앞(파일 하단, 클래스 정의 밖)에 추가:

```python
def _selftest_build_preview_message_routes_by_extension():
    """확장자에 따라 어떤 종류의 미리보기를 만들지 올바르게 판단하는지
    확인한다(HWP류는 이미지, 엑셀/PDF는 텍스트, 그 외는 미리보기 없음).
    이 함수는 실제 pyhwpx/openpyxl을 부르지 않고 순수하게 분기만
    검증한다 - 실제 생성은 attachment_preview.py 쪽 self-test가 담당."""
    assert _preview_kind_for(".hwp") == "image"
    assert _preview_kind_for(".hwpx") == "image"
    assert _preview_kind_for(".xlsx") == "text"
    assert _preview_kind_for(".pdf") == "text"
    assert _preview_kind_for(".txt") is None
    print("_preview_kind_for 통과")
```

- [x] Step 2: 실패 확인

Run: `python chat_assistant.py --selftest`
Expected: `NameError: name '_preview_kind_for' is not defined`

(참고: 이 self-test는 아직 `if __name__ == "__main__":`의 `--selftest`
분기에 등록돼 있지 않으므로, 먼저 그 분기에 `_selftest_build_preview_message_routes_by_extension()`
호출을 추가해야 실제로 실행되고 `NameError`가 발생한다.)

- [x] Step 3: 최소 구현

파일 상단 import에 추가:

```python
from attachment_preview import generate_hwp_preview_isolated, generate_text_preview
```

`_route_by_keywords()` 함수 앞에 추가:

```python
def _preview_kind_for(ext: str) -> str | None:
    """확장자별로 어떤 미리보기를 만들지 판단한다. ext는 점(.) 포함
    소문자(예: ".hwp"). HWP류는 이미지 미리보기, 엑셀/PDF는 텍스트
    미리보기, 그 외(이미지 자체를 첨부한 경우 등)는 미리보기를 만들지
    않는다(None)."""
    if ext in (".hwp", ".hwpx"):
        return "image"
    if ext in (".xlsx", ".xls", ".pdf"):
        return "text"
    return None
```

`ChatAssistant` 클래스 안, `_pick_source_files()` 메서드 뒤에 추가:

```python
    def _show_attachment_preview(self, path: str):
        """첨부 직후 파일 하나의 작은 미리보기를 채팅창에 바로 보여준다.
        HWP/HWPX는 이미지(별도 프로세스에서 1페이지만 렌더링, 같은
        프로세스 안의 다른 Hwp 인스턴스가 이미 열려있는 보고서의 COM
        연결을 깨뜨리는 pyhwpx 한계 때문 - attachment_preview.py 참고),
        엑셀/PDF는 텍스트(상위 5줄)를 보여준다. 실패해도 조용히
        건너뛴다(미리보기는 부가기능이라 실패가 첨부 자체를 막으면 안 됨)."""
        ext = os.path.splitext(path)[1].lower()
        kind = _preview_kind_for(ext)
        if kind is None:
            return
        try:
            if kind == "image":
                preview_path = generate_hwp_preview_isolated(path)
                if preview_path is not None:
                    card = self._log(f"미리보기: {os.path.basename(path)}", role="assistant")
                    image = ctk.CTkImage(light_image=Image.open(preview_path),
                                          dark_image=Image.open(preview_path), size=(160, 220))
                    ctk.CTkLabel(card, text="", image=image).pack(padx=8, pady=(0, 6))
            else:
                text_preview = generate_text_preview(path)
                if text_preview:
                    self._log(f"미리보기: {os.path.basename(path)}\n{text_preview}", role="assistant")
        except Exception:
            pass  # 미리보기 실패는 첨부 자체를 막지 않음 - 부가기능
```

파일 상단 import에 `from PIL import Image` 추가(customtkinter의
`CTkImage`가 PIL 이미지 객체를 요구함 - pillow는 이미 `requirements.txt`에
있어 신규 의존성 아님).

`_pick_source_files()` 메서드의 `self._log(f"📎 원본자료 {len(self.source_paths)}개 첨부됨: {names}")`
줄 뒤에 추가:

```python
        for path in self.source_paths:
            self._show_attachment_preview(path)
```

`_pick_source_folder()` 메서드의 `self._log(f"📎 원본자료(폴더) 첨부됨: {path}")`
줄 뒤에는 미리보기를 추가하지 않는다(폴더 안 파일 개수를 미리 알 수 없어
여러 개를 한꺼번에 미리보기하면 용량 관리 요구사항과 충돌할 수 있음 -
다음 라운드 재검토 대상으로 문서화만 해둠).

`if __name__ == "__main__":`의 `--selftest` 분기에 추가:

```python
        _selftest_build_preview_message_routes_by_extension()
```

- [x] Step 4: 통과 확인

Run: `python chat_assistant.py --selftest`
Expected: 전부 통과, exit 0

- [x] Step 5: 커밋

```bash
git add chat_assistant.py
git commit -m "F14: 첨부 시 자동 미리보기 표시 - HWP 이미지 / 엑셀·PDF 텍스트"
```

---

### Task 13: 통합 검증 + 계획 마무리

Files

- 없음(검증 전용 태스크)

- [x] Step 1: 신규/수정 모듈 전부 순차 self-test 재실행 (병렬 금지, 사이 간격 확보)

```bash
python weekly_report_tool.py
python fit_to_page_tool.py
python verify_numbers.py
python ignore_list.py
python verify_tool.py
python attachment_preview.py
python chat_assistant.py --selftest
```

Expected: 전부 exit 0. pyhwpx를 쓰는 모듈(weekly_report_tool.py,
fit_to_page_tool.py, verify_tool.py, attachment_preview.py) 사이에는 이전
한글 프로세스가 완전히 정리되도록 사람이 몇 초 간격을 두고 순서대로
실행할 것 — 이 계획서의 "사전조사 실측 확인 요약" 7번이 지적한 대로, 너무
빠르게 연달아 여러 번 Hwp 인스턴스를 재생성하면 `pywintypes.com_error`가
날 수 있다. 만약 재현되면 작업관리자(또는 `tasklist`/`taskkill //F //IM
Hwp.exe`)로 좀비 프로세스가 남아있는지 확인 후 정리하고 재시도한다.

- [x] Step 2: verify_numbers.py의 기존 self-test(F11/F12/F13에서 이미 있던
      것들 포함)까지 전부 회귀 없는지 최종 확인

```bash
python verify_numbers.py
```

Expected: 이 배치에서 새로 추가한 4개(compute_column_growth_rate 계열)뿐
아니라 F11/F13에서부터 있던 기존 self-test 전부(`_selftest_extract_amounts`
등)도 그대로 통과 - 새 함수 추가가 기존 로직을 건드리지 않았음을 재확인.

- [x] Step 3: git status로 의도치 않은 변경 확인

```bash
git status
git diff --stat
```

Expected: 이번 계획에서 다룬 7개 파일(`weekly_report_tool.py`,
`fit_to_page_tool.py`, `ignore_list.py`, `attachment_preview.py`,
`verify_numbers.py`, `verify_tool.py`, `chat_assistant.py`)과 테스트 중
생성됐다가 정리 안 된 임시 파일(있다면 삭제)만 보여야 한다.
`ignore_list.json`/`speed_log.json` 등 실행 중 생성되는 데이터 파일이
add되지 않았는지도 확인.

- [x] Step 4: 사람 확인이 필요한 항목 정리 (아침에 확인)

다음은 무인 self-test로 완전히 검증할 수 없어 사람이 눈으로 봐야 하는
항목이다 — 채팅창을 실제로 띄워(`python chat_assistant.py`) 아래를
확인한다:
- 주간업무보고 여러 개를 "+"로 첨부하고 "옆 한글파일로 옮겨줘"를 입력했을
  때 실제로 대상 문서의 이번주/다음주 칸에 파란색 없이(검정 기본색으로)
  깔끔하게 이어붙는지, 줄바꿈이 자연스러운지
- "한 페이지에 맞춰줘"를 실행했을 때 글자가 너무 작아지거나 자간이 너무
  좁아 읽기 불편해지지 않는지(알고리즘은 페이지 수만 확인하지 가독성은
  판단하지 않음 - floor 값이 실제로 적절한지는 사람 판단 필요)
- 첨부 시 미리보기 이미지/텍스트가 채팅창 레이아웃 안에서 잘려보이지 않는지
- "N번째는 무시해" 실행 후 실제로 다음 검증에서 그 항목이 더 이상 빨갛게
  표시되지 않는지(자동저장 없으므로 검증을 다시 돌려서 확인)

이 항목들은 계획 완료 조건이 아니라, 실행 후 사용자가 직접 확인할 후속
과제로 남긴다.

**완료 기록 (2026-09-07)**: 7개 모듈 순차 self-test 전부 통과(exit 0). 1차
실행에서 `weekly_report_tool.py`의 `_selftest_read_weekly_content_isolated_basic`
하나가 일시적으로 실패했으나, 원인은 코드 버그가 아니라 직전 테스트가 방금
`quit()`한 Hwp COM 인스턴스가 완전히 정리되기 전에 다음 테스트가 곧바로 새
인스턴스를 띄우며 겹친 타이밍성 플레이키(이 계획서 "사전조사 실측 확인
요약" 7번이 이미 경고한 `pywintypes.com_error` 계열 현상)로 확인됨 —
재실행 시 정상 재현되어 코드 수정 없이 종결. Step 4의 사람 확인 항목(주간
보고 취합/한 페이지 맞춤/미리보기 레이아웃/무시 목록 반영)은 사용자가 직접
`chat_assistant.py`를 실행해 확인하기로 함.
