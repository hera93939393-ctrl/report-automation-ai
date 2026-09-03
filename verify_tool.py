"""verify_tool.py — verify_numbers.py/source_reader.py를 엮어 "숫자검증" 도구
하나로 만든다. F12: 이미 열려있는 HwpReport 핸들과 원본자료 경로 리스트를 받는다
(F11 시절엔 이 함수가 문서를 직접 열었으나, 도구 호출마다 문서가 새로 열리는
문제가 있어 호출자(chat_assistant.py)가 한 번만 연 핸들을 넘겨주는 구조로 변경).
run_verification()은 같은 핸들에 대한 반복 호출을 전제로 설계되어 있고
(한 채팅 세션에서 "숫자 검증해줘"를 여러 번), 매 호출이 멱등적이도록
호출마다 문서 색을 리셋한 뒤 현재 대조 결과만 다시 표시한다(자세한 내용은
run_verification 함수 docstring 참고)."""
import os
import tempfile
import openpyxl

from verify_numbers import extract_values, categorize_values
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
      - match_count / unverifiable_count: int — 각각 파란색(정상)/초록색
        (대조불가) 표시된 서로 다른 값의 개수(2026-09-04 추가, 아래 참고).
      - summary: str — 채팅창에 그대로 보여줄 사람이 읽는 요약 텍스트.
      - conflicts: list — read_source_files가 찾은 원본 파일 간 불일치 목록.

    (2026-09-04, 실사용 피드백) "문서가 길어지니 전부 확인한 건지 모르겠다"는
    지적을 받아, 불일치(빨강)만 표시하던 것에 정상(파랑)/대조불가(초록)도
    추가로 표시하게 됐다. 대조불가는 answer_pool에 그 타입 자체가 없어서
    (예: 원본 엑셀엔 금액만 있고 전화번호가 없음) 애초에 비교할 수 없었던
    값이다 — 오류(빨강)와는 다른 뜻이므로 구분해서 표시한다. summary에도
    세 건수를 전부 적어, 채팅 텍스트만 봐도 "몇 개 중 몇 개를 어떻게
    확인했는지" 알 수 있게 했다.

    반복 호출 계약(F12의 핵심 전제, 2026-09-01 코드품질 리뷰로 수정됨): 이
    함수는 같은 HwpReport 핸들에 대해 한 채팅 세션 안에서 여러 번 호출되도록
    설계되어 있다 — 사용자가 "숫자 검증해줘"를 반복하거나, 원본을 고치거나
    새 원본을 첨부한 뒤 다시 검증을 요청하는 흐름이 그것이다. 이 함수는 매
    호출마다 먼저 report.reset_colors()로 문서 전체 글자색을 검정으로
    되돌린 뒤, 이번 호출에서 실제로 확인된 불일치만 다시 빨갛게 표시한다.
    그래서 매 호출은 이전 호출의 결과에 의존하지 않는 멱등적(idempotent)
    동작이 된다 — "지금 이 순간의 대조 결과"만 문서에 반영되고, 이전에
    남긴 빨간 표시가 더 이상 유효하지 않은데도 잔류하는 일이 없다. (이전엔
    아무것도 지우지 않아 재검증 후 "이상 없음"이라 답하면서도 문서엔 빨간
    글자가 그대로 남는 버그가 있었다 — 코드품질 검토에서 재현·확인됨.)
    이 계약은 Task 9의 _on_submit 통합 등 이 함수를 호출하는 쪽 어디서든
    똑같이 성립한다: 매번 새로 열지 않고 같은 핸들을 재사용해도 안전하다.

    알려진 후속 과제(이번 라운드에서 의도적으로 손대지 않음): report.get_text()가
    이미 닫힌/죽은 COM 핸들에 대해 호출되면 pywintypes.com_error가 그대로
    올라온다 — 사용자 친화적인 한글 오류 메시지로 감싸는 작업은 별도
    라운드로 미뤄둔다.
    """
    answer_pool, conflicts = read_source_files(source_paths, default_year)

    report_text = report.get_text()
    report.reset_colors()  # 재검증 시 이전 호출이 남긴 표시가 잔류하지 않도록, 매 호출 시작 시 문서 전체를 검정으로 리셋
    report_values = extract_values(report_text, default_year)
    categorized = categorize_values(report_values, answer_pool)
    mismatches = categorized["mismatches"]
    matches = categorized["matches"]
    unverifiable = categorized["unverifiable"]

    # (2026-09-04, 실사용 피드백) 원래는 "고유 raw 문자열 하나당 mark_color() 한
    # 번" 구조였다 — mark_color()가 매번 MoveDocBegin()부터 다시 시작해 문서
    # 전체를 훑기 때문에, 서로 다른 값이 여러 개면 문서를 그 개수만큼 처음부터
    # 다시 스캔하는 것처럼 보였다(사용자가 실제로 "수십 번 처음부터 끝까지
    # 훑는다"고 지적함). "위에서부터 하나씩 순서대로 빨강/파랑/초록을 칠하며
    # 내려가자"는 사용자 제안대로, extract_values가 이미 갖고 있는 문서 내
    # 위치(span)를 기준으로 전체 항목을 한 번만 정렬한 뒤, 커서를 문서 처음에
    # 한 번만 놓고 mark_next_color()(Forward 검색만 하고 되돌아가지 않음)로
    # 순서대로 칠한다 — 결과적으로 문서를 위→아래 딱 한 번만 훑는다.
    #
    # 같은 raw 문자열이 여러 번 등장하면 extract_values가 등장할 때마다 별도
    # 항목(서로 다른 span)으로 뽑아두므로, span 순서대로 처리하면 자연히 각
    # occurrence를 순서대로 하나씩 만나 정확히 칠하게 된다(값 단위로 뭉치지
    # 않음 — 그래서 여기서는 dict.fromkeys로 중복 제거하지 않는다. 채팅
    # 요약에 쓸 "고유 값 개수"는 아래에서 별도로 계산한다).
    colored_in_order = sorted(
        [(m["span"][0], m["raw"], (255, 0, 0)) for m in mismatches]
        + [(m["span"][0], m["raw"], (0, 0, 255)) for m in matches]
        + [(m["span"][0], m["raw"], (0, 128, 0)) for m in unverifiable],
        key=lambda item: item[0],
    )
    report.hwp.MoveDocBegin()
    for _pos, raw, (r, g, b) in colored_in_order:
        found = report.mark_next_color(raw, r, g, b)
        if not found:
            # 추출 순서가 실제 문서상 위치와 어긋나는 드문 경우를 대비한
            # 폴백 — 문서 처음부터 다시 찾아서라도 반드시 칠한다.
            report.hwp.MoveDocBegin()
            report.mark_next_color(raw, r, g, b)

    unique_mismatch_raw = list(dict.fromkeys(m["raw"] for m in mismatches))
    unique_match_raw = list(dict.fromkeys(m["raw"] for m in matches))
    unique_unverifiable_raw = list(dict.fromkeys(m["raw"] for m in unverifiable))

    first_type_by_raw = {}
    for m in mismatches:
        first_type_by_raw.setdefault(m["raw"], m["type"])
    lines = [f"- [{first_type_by_raw[raw]}] '{raw}' 원본에서 확인 안 됨" for raw in unique_mismatch_raw]
    for c in conflicts:
        value_desc = ", ".join(f"{v['file']}={v['normalized']}" for v in c["values"])
        lines.append(f"- ⚠ 원본자료 불일치[{c['type']}]: {c['location']} ({value_desc})")

    total_checked = len(unique_mismatch_raw) + len(unique_match_raw) + len(unique_unverifiable_raw)
    header = (
        f"총 {total_checked}건 확인 - 정상(파랑) {len(unique_match_raw)}건, "
        f"오류(빨강) {len(unique_mismatch_raw)}건, 대조불가(초록) {len(unique_unverifiable_raw)}건"
    )
    summary = header + ("\n" + "\n".join(lines) if lines else "")

    return {
        "mismatch_count": len(unique_mismatch_raw),
        "match_count": len(unique_match_raw),
        "unverifiable_count": len(unique_unverifiable_raw),
        "summary": summary,
        "conflicts": conflicts,
    }


def _selftest_run_verification():
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


def _selftest_run_verification_clears_stale_marks_on_rerun():
    """회귀테스트 — 코드품질 리뷰에서 재현된 버그: 이전 호출이 남긴 빨간 표시가
    재검증 후에도 지워지지 않던 문제. 시나리오:
    1) 원본에 없는 값으로 1차 검증 → 빨간색으로 표시됨을 확인
    2) 문서는 그대로 두고, 원본 쪽을 고쳐 그 값이 더 이상 불일치가 아니게 만듦
       (사용자가 원본을 수정한 뒤 같은 채팅 세션에서 재검증을 요청하는 F12 시나리오)
    3) 같은 HwpReport 핸들로 재검증 → 오류 0건이면서, 이제는 원본과 일치하는
       값이라 글자색이 파란색(정상 표시)으로 바뀌어 있어야 한다(2026-09-04
       갱신 — 예전엔 "표시 없음(검정)"이 기대값이었으나, 정상 값도 파란색으로
       표시하는 기능이 추가되면서 기대값이 바뀜). 둘 다 확인해 요약 텍스트만
       맞고 화면은 안 맞는 버그가 재발하면 바로 잡히도록 한다.
    """
    test_dir = os.path.join(tempfile.gettempdir(), "_test_원본_f12_rerun")
    os.makedirs(test_dir, exist_ok=True)
    source_path = os.path.join(test_dir, "원본.xlsx")
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Sheet1"
    ws["A1"] = "예산"; ws["A2"] = 1850000
    wb.save(source_path)

    from pyhwpx import Hwp
    report_path = os.path.join(tempfile.gettempdir(), "_test_보고서_f12_rerun.hwp")
    setup = Hwp(visible=False)
    setup.insert_text("예산은 9999999원입니다")
    setup.save_as(report_path)
    setup.quit()

    report = None
    try:
        report = HwpReport(report_path)  # 채팅 세션 내내 재사용되는 F12 핸들 시뮬레이션

        result1 = run_verification(report, source_paths=[test_dir], default_year=2026)
        assert result1["mismatch_count"] == 1, result1
        color_after_first = report.get_char_color_at("9999999")
        assert color_after_first == (255, 0, 0), color_after_first

        # 원본을 고쳐서(사용자가 원본자료를 수정한 상황) 더 이상 불일치가 아니게 만든다.
        # 문서 텍스트는 그대로 "9999999원"이지만, 이제 원본과 일치한다.
        wb2 = openpyxl.Workbook(); ws2 = wb2.active; ws2.title = "Sheet1"
        ws2["A1"] = "예산"; ws2["A2"] = 9999999
        wb2.save(source_path)

        result2 = run_verification(report, source_paths=[test_dir], default_year=2026)
        assert result2["mismatch_count"] == 0, result2
        assert "오류(빨강) 0건" in result2["summary"], result2["summary"]
        color_after_second = report.get_char_color_at("9999999")
        assert color_after_second == (0, 0, 255), color_after_second  # 이제 원본과 일치 → 파란색(정상)
        print("run_verification 재검증 회귀테스트 통과: 재검증 후 이전 빨간표시가 사라지고 파란색(정상)으로 바뀜")
    finally:
        if report is not None:
            report.close(save=False)
        import shutil
        shutil.rmtree(test_dir)
        os.remove(report_path)


def _selftest_run_verification_colors_all_three_categories_in_one_pass():
    """(2026-09-04, 실사용 피드백) "왜 값 하나마다 문서를 처음부터 끝까지
    다시 훑냐"는 지적을 받아 mark_color()를 값마다 반복 호출하던 것을,
    문서 내 위치(span) 순서로 정렬한 뒤 커서를 한 번만 문서 처음에 두고
    mark_next_color()로 순서대로 칠하는 방식으로 바꿨다 — 정상(파랑)/
    오류(빨강)/대조불가(초록) 세 종류가 한 문서에 섞여 있을 때도 각각
    정확한 위치에 정확한 색으로 칠해지는지 확인한다(리팩터링 전
    mark_color() 기반 구현에서 이미 확인했던 것과 같은 시나리오 —
    구현 방식만 바뀌었지 결과는 같아야 한다)."""
    test_dir = os.path.join(tempfile.gettempdir(), "_test_삼색_순서")
    os.makedirs(test_dir, exist_ok=True)
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Sheet1"
    ws["A1"] = "예산"; ws["A2"] = 1850000
    wb.save(os.path.join(test_dir, "원본.xlsx"))

    from pyhwpx import Hwp
    report_path = os.path.join(tempfile.gettempdir(), "_test_보고서_삼색순서.hwp")
    setup = Hwp(visible=False, new=True)
    setup.insert_text(
        "예산은 185만원이며, 오타는 9999999원이고, "
        "담당자 연락처는 031-1234-5678입니다"
    )
    setup.save_as(report_path)
    setup.quit()

    report = None
    try:
        report = HwpReport(report_path)
        result = run_verification(report, source_paths=[test_dir], default_year=2026)
        assert result["match_count"] == 1, result
        assert result["mismatch_count"] == 1, result
        assert result["unverifiable_count"] == 1, result

        assert report.get_char_color_at("185만원") == (0, 0, 255), "정상(파랑) 표시 안 됨"
        assert report.get_char_color_at("9999999") == (255, 0, 0), "오류(빨강) 표시 안 됨"
        assert report.get_char_color_at("031-1234-5678") == (0, 128, 0), "대조불가(초록) 표시 안 됨"
        print("run_verification(세 카테고리 한 번에 색칠) 통과:", result["summary"])
    finally:
        if report is not None:
            report.close(save=False)
        import shutil
        shutil.rmtree(test_dir)
        os.remove(report_path)


if __name__ == "__main__":
    _selftest_run_verification()
    _selftest_run_verification_clears_stale_marks_on_rerun()
    _selftest_run_verification_colors_all_three_categories_in_one_pass()
