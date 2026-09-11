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

from verify_numbers import (extract_values, categorize_values, check_weekday_consistency,
                            build_table_contexts, attach_table_context)
from source_reader import read_source_files
from hwp_report import HwpReport
from ignore_list import load_ignored_values, DEFAULT_IGNORE_PATH


def run_verification(report: HwpReport, source_paths: list[str], default_year: int,
                      ignore_list_path: str = DEFAULT_IGNORE_PATH) -> dict:
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
      - ambiguous_count: int — 회색(확인 필요) 표시된 서로 다른 값의 개수
        (2026-09-08 추가, 아래 참고).
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

    (2026-09-08, 실제 fixture 테스트로 발견) categorize_values()가 이제
    matches/mismatches/unverifiable 외에 ambiguous(문맥상 어느 원본 항목을
    가리키는지 특정할 수 없는 값)도 반환한다 — 같은 타입의 값이 원본에
    여럿 있는데 보고서 문맥에 그중 어느 것의 라벨(항목명)도 안 걸릴 때다.
    "다른 항목의 값과 우연히 같은 숫자"를 잘못 파랑/빨강으로 단정하지 않고
    회색 "확인 필요"로 사람에게 넘긴다. 회색은 (128,128,128)로 표시한다.

    알려진 후속 과제(이번 라운드에서 의도적으로 손대지 않음): report.get_text()가
    이미 닫힌/죽은 COM 핸들에 대해 호출되면 pywintypes.com_error가 그대로
    올라온다 — 사용자 친화적인 한글 오류 메시지로 감싸는 작업은 별도
    라운드로 미뤄둔다.
    """
    answer_pool, conflicts = read_source_files(source_paths, default_year)
    report.reset_colors()  # 재검증 시 이전 호출이 남긴 표시가 잔류하지 않도록, 매 호출 시작 시 문서 전체를 검정으로 리셋

    if not answer_pool:
        # (2026-09-04, 실사용 피드백) 원본자료를 첨부했는데 실제로 읽을 수
        # 있는 데이터가 하나도 없으면(예: 지원 안 되는 형식만 첨부한 경우 —
        # .hwp/.hwpx는 source_reader.py의 _READERS에 아예 없음), 그대로
        # 진행하면 categorize_values가 모든 값에 candidates=[]를 판정해
        # 문서의 모든 값이 "대조불가(초록)"로만 칠해진다 — 사용자가 실제로
        # 이 상황을 겪고 "검증을 안 한 거냐"고 물었다. 실제로는 비교할
        # 원본 자체가 없어서 검증이 의미 없는 상태이므로, "N건 대조불가"로
        # 뭉뚱그리지 않고 원인을 명확히 알려준다.
        return {
            "mismatch_count": 0,
            "match_count": 0,
            "unverifiable_count": 0,
            "ambiguous_count": 0,
            "summary": "원본자료에서 읽을 수 있는 데이터가 없어요. 지원 형식(엑셀 .xlsx/.xls, PDF .pdf)인지 확인해주세요.",
            "conflicts": conflicts,
        }

    report_text = report.get_text()
    report_values = extract_values(report_text, default_year)

    # (F14) 표 안에서 뽑힌 값에는 그 칸의 행/열 머리말을 붙여준다. 평문
    # 텍스트(GetTextFile("TEXT",""))는 표를 셀 하나당 한 줄로 평평하게
    # 흘려보내 행/열 정보를 통째로 잃어버리므로, 표 격자를 따로 읽어
    # (get_table_grids) 그 평문에 다시 맞춘다. 맥락이 붙은 값만
    # categorize_values에서 "출처가 맞는 원본 값"으로 후보가 좁혀진다 —
    # 표 밖의 값과 격자를 못 맞춘 표는 맥락이 안 붙어 기존 동작 그대로다.
    # get_table_grids()는 실패해도 예외 대신 빈 목록을 돌려주므로, 표를
    # 못 읽는 문서에서도 검증 자체는 예전처럼 끝까지 진행된다.
    table_contexts = build_table_contexts(report_text, report.get_table_grids())
    attach_table_context(report_values, table_contexts)

    categorized = categorize_values(report_values, answer_pool)
    mismatches = categorized["mismatches"]
    matches = categorized["matches"]
    unverifiable = categorized["unverifiable"]
    ambiguous = categorized["ambiguous"]

    # (F14) ignore_list_path에 등록된 값(raw 문자열 그대로 비교)은 원본과
    # 실제로 불일치해도 mismatches에서 제외한다 - "이건 괜찮아, 무시해"로
    # 한 번 확인한 값은 다음부터 오류로 표시하지 않는다는 오탐 학습 기능.
    ignored_raws = set(load_ignored_values(ignore_list_path))
    if ignored_raws:
        mismatches = [m for m in mismatches if m["raw"] not in ignored_raws]

    # (F13) 날짜 뒤 괄호에 적힌 요일이 실제 요일과 맞는지도 확인한다. 이
    # 값들은 extract_dates()가 이미 "date" 타입으로 뽑아 matches/mismatches/
    # unverifiable 중 하나에 이미 들어가 있으므로(같은 span), 별도 항목을
    # 새로 추가하지 않고 해당 span의 색만 빨강으로 덮어쓴다 - 그래야 같은
    # 문서 위치를 두 번 칠하려다 커서 위치가 꼬이는 문제(mark_next_color가
    # 다음 occurrence를 잘못 찾는 것)가 생기지 않는다.
    weekday_mismatches = check_weekday_consistency(report_text)
    weekday_mismatch_spans = {w["span"] for w in weekday_mismatches}

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
    def _color_override(entry, base_color):
        # 요일불일치가 확인된 span은 원래 판정(파랑/빨강/초록)과 무관하게
        # 화면에 빨강으로 보이게 한다 - "확인이 필요한 항목"이라는 신호는
        # 값 일치 여부와 요일 표기 오류 둘 다 동등하게 취급한다.
        return (255, 0, 0) if entry["span"] in weekday_mismatch_spans else base_color

    colored_in_order = sorted(
        [(m["span"][0], m["raw"], _color_override(m, (255, 0, 0))) for m in mismatches]
        + [(m["span"][0], m["raw"], _color_override(m, (0, 0, 255))) for m in matches]
        + [(m["span"][0], m["raw"], _color_override(m, (0, 128, 0))) for m in unverifiable]
        + [(m["span"][0], m["raw"], _color_override(m, (128, 128, 128))) for m in ambiguous],
        key=lambda item: item[0],
    )
    # (F13) 화면에 실제로 빨갛게 표시된 값들을, 표시된 순서(span 순서) 그대로
    # 뽑아둔다 - "N번째로 가줘" 이동 기능(chat_assistant.py)이 이 순서를
    # 그대로 신뢰하고 인덱싱한다. colored_in_order에서 뽑으므로 요일불일치로
    # 빨강 덮어쓰기된 항목도 자연히 포함된다.
    mismatch_items = list(dict.fromkeys(
        raw for _pos, raw, color in colored_in_order if color == (255, 0, 0)
    ))

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
    unique_ambiguous_raw = list(dict.fromkeys(m["raw"] for m in ambiguous))

    first_type_by_raw = {}
    for m in mismatches:
        first_type_by_raw.setdefault(m["raw"], m["type"])
    lines = [f"- [{first_type_by_raw[raw]}] '{raw}' 원본에서 확인 안 됨" for raw in unique_mismatch_raw]
    for w in weekday_mismatches:
        lines.append(f"- [요일불일치] '{w['raw']}' 실제로는 {w['actual_weekday']}요일")
    for c in conflicts:
        value_desc = ", ".join(f"{v['file']}={v['normalized']}" for v in c["values"])
        lines.append(f"- ⚠ 원본자료 불일치[{c['type']}]: {c['location']} ({value_desc})")
    ambiguous_first_type_by_raw = {}
    for m in ambiguous:
        ambiguous_first_type_by_raw.setdefault(m["raw"], m["type"])
    for raw in unique_ambiguous_raw:
        lines.append(f"- [{ambiguous_first_type_by_raw[raw]}] '{raw}' 원본의 어느 항목인지 문맥상 불명확 - 확인 필요")

    total_checked = (len(unique_mismatch_raw) + len(unique_match_raw)
                     + len(unique_unverifiable_raw) + len(unique_ambiguous_raw))
    header = (
        f"총 {total_checked}건 확인 - 정상(파랑) {len(unique_match_raw)}건, "
        f"오류(빨강) {len(unique_mismatch_raw)}건, 대조불가(초록) {len(unique_unverifiable_raw)}건, "
        f"확인 필요(회색) {len(unique_ambiguous_raw)}건"
    )
    summary = header + ("\n" + "\n".join(lines) if lines else "")

    return {
        "mismatch_count": len(unique_mismatch_raw),
        "match_count": len(unique_match_raw),
        "unverifiable_count": len(unique_unverifiable_raw),
        "ambiguous_count": len(unique_ambiguous_raw),
        "summary": summary,
        "conflicts": conflicts,
        "mismatch_items": mismatch_items,
    }


def _selftest_run_verification():
    test_dir = os.path.join(tempfile.gettempdir(), "_test_원본_f12")
    os.makedirs(test_dir, exist_ok=True)
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Sheet1"
    ws["A1"] = "예산"; ws["A2"] = 1850000
    wb.save(os.path.join(test_dir, "원본.xlsx"))

    from pyhwpx import Hwp
    report_path = os.path.join(tempfile.gettempdir(), "_test_보고서_f12.hwp")
    # (2026-09-04, 실사용 세션 중 실제 재현·발견) new=True 없이 Hwp()를 만들면
    # 이미 떠 있는 다른 한글 프로세스에 그대로 접속(재사용)해버려서, 그
    # 프로세스가 열어둔 다른 문서 내용에 이 테스트 문장이 이어붙여진 채로
    # 저장될 수 있다(실제로 다른 테스트와 동시에 돌 때 "존재하지 않는
    # 숫자가 불일치로 잡힌다" 형태로 재현됨 — source_reader.py의 같은
    # 수정 참고). new=True로 항상 독립된 새 프로세스를 쓴다.
    setup = Hwp(visible=False, new=True)
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
    setup.insert_text("예산은 185만원이며, 첫오류는 9999999원, 둘째오류는 8888888원입니다")
    setup.save_as(report_path)
    setup.quit()

    ignore_path = os.path.join(tempfile.gettempdir(), "_test_ignore_list_검증.json")
    if os.path.exists(ignore_path):
        os.remove(ignore_path)
    from ignore_list import record_ignored_value
    # (2026-09-07, 실측으로 정정) 무시 목록엔 raw 원문 그대로(단위 포함) 저장돼야
    # verify_tool.py의 필터(m["raw"] not in ignored_raws)와 실제로 일치한다 -
    # 실사용 경로(chat_assistant.py의 parse_ignore_index 연동)도 항상
    # self._last_mismatch_items[i](예: "9999999원")를 그대로 저장하므로,
    # 여기서 단위 없는 "9999999"만 저장하면 필터가 매치되지 않는다.
    record_ignored_value("9999999원", path=ignore_path)  # 첫오류만 무시 목록에 등록

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
    setup = Hwp(visible=False, new=True)  # 다른 프로세스 재사용 방지 — 위 _selftest_run_verification 참고
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


def _selftest_run_verification_connects_by_label_and_marks_ambiguous_gray():
    """(2026-09-08, 실제 한글 문서로 재현·수정 확인 — U02/U03) 원본에 같은
    타입(amount) 후보가 둘 이상(참여 인원=21, 지원 인원=40) 있을 때:
    - "참여 인원 21명"은 문맥의 라벨 "참여 인원"으로 21 후보와 연결돼
      일치(파랑)여야 한다.
    - "지원 인원 33명"은 라벨 "지원 인원"으로 40 후보와 연결돼 33≠40이라
      불일치(빨강)여야 한다.
    - "회의 횟수는 55회"는 어느 라벨과도 안 걸려 원본과의 연결이 불명확하므로
      55가 원본에 없다는 이유만으로 빨강으로 단정하지 말고 확인
      필요(회색)로 남아야 한다 — 수정 전에는 이 값도 무조건 빨강이었다."""
    test_dir = os.path.join(tempfile.gettempdir(), "_test_라벨연결_확인필요")
    os.makedirs(test_dir, exist_ok=True)
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "실적"
    ws.append(["항목", "값"])
    ws.append(["참여 인원", 21])
    ws.append(["지원 인원", 40])
    wb.save(os.path.join(test_dir, "원본.xlsx"))

    from pyhwpx import Hwp
    report_path = os.path.join(tempfile.gettempdir(), "_test_보고서_라벨연결_확인필요.hwp")
    setup = Hwp(visible=False, new=True)
    setup.insert_text("참여 인원 21명이며, 지원 인원 33명이고, 회의 횟수는 55회입니다")
    setup.save_as(report_path)
    setup.quit()

    report = None
    try:
        report = HwpReport(report_path)
        result = run_verification(report, source_paths=[test_dir], default_year=2026)
        assert result["match_count"] == 1, result
        assert result["mismatch_count"] == 1, result
        assert result["ambiguous_count"] == 1, result
        assert "확인 필요(회색) 1건" in result["summary"], result["summary"]
        assert "'55' 원본의 어느 항목인지 문맥상 불명확" in result["summary"], result["summary"]

        assert report.get_char_color_at("21") == (0, 0, 255), "라벨 연결된 일치가 파랑이 아님"
        assert report.get_char_color_at("33") == (255, 0, 0), "라벨 연결된 불일치가 빨강이 아님"
        assert report.get_char_color_at("55") == (128, 128, 128), "문맥 불명확 값이 회색이 아님"
        print("run_verification(라벨 연결 + 확인필요 회색) 통과:", result["summary"])
    finally:
        if report is not None:
            report.close(save=False)
        import shutil
        shutil.rmtree(test_dir)
        os.remove(report_path)


def _selftest_run_verification_empty_answer_pool_gives_clear_message():
    """(2026-09-04, 실사용 피드백) 사용자가 실제로 겪은 상황: 원본자료로
    지원 안 되는 형식(당시엔 .hwp)만 첨부해서 answer_pool이 완전히
    비었는데, 문서의 모든 값이 "대조불가(초록)"로 나와서 "검증을 안 한
    거냐"고 물었다. 이제는 answer_pool이 비어 있으면 색칠/분류를 진행하지
    않고, 원인을 명확히 알려주는 메시지를 즉시 반환해야 한다."""
    test_dir = os.path.join(tempfile.gettempdir(), "_test_원본없음")
    os.makedirs(test_dir, exist_ok=True)
    # 이 폴더 안에는 _READERS가 읽을 수 있는 형식(.xlsx/.xls/.pdf/.hwp/.hwpx)이
    # 하나도 없다 — 순수 텍스트 파일만 있어서 read_source_files가 조용히
    # 건너뛰고 answer_pool이 빈 채로 돌아온다.
    with open(os.path.join(test_dir, "메모.txt"), "w", encoding="utf-8") as f:
        f.write("이건 원본자료로 못 읽는 형식입니다")

    from pyhwpx import Hwp
    report_path = os.path.join(tempfile.gettempdir(), "_test_보고서_원본없음.hwp")
    setup = Hwp(visible=False, new=True)
    setup.insert_text("예산은 1,850,000원입니다")
    setup.save_as(report_path)
    setup.quit()

    report = None
    try:
        report = HwpReport(report_path)
        result = run_verification(report, source_paths=[test_dir], default_year=2026)
        assert result["match_count"] == 0, result
        assert result["mismatch_count"] == 0, result
        assert result["unverifiable_count"] == 0, result
        assert "읽을 수 있는 데이터가 없어요" in result["summary"], result["summary"]
        print("run_verification(원본 데이터 없음, 명확한 안내) 통과:", result["summary"])
    finally:
        if report is not None:
            report.close(save=False)
        import shutil
        shutil.rmtree(test_dir)
        os.remove(report_path)


def _selftest_run_verification_flags_weekday_mismatch():
    """(F13) 날짜 뒤 괄호에 적힌 요일이 실제 요일과 다르면, 그 값이 원래
    파랑(정상)이든 초록(대조불가)이든 상관없이 빨간색으로 표시되고
    summary에 '[요일불일치]' 줄이 추가되는지 확인한다. 2026-09-07의 실제
    요일은 월요일(datetime.date(2026,9,7).weekday()==0)이므로, 문서에는
    일부러 틀린 "화"로 적어 불일치를 재현한다."""
    import datetime
    assert datetime.date(2026, 9, 7).weekday() == 0, "전제 확인: 2026-09-07은 월요일"

    test_dir = os.path.join(tempfile.gettempdir(), "_test_원본_요일불일치")
    os.makedirs(test_dir, exist_ok=True)
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Sheet1"
    ws["A1"] = "예산"; ws["A2"] = 1850000  # date 타입은 원본에 없음 -> 대조불가(초록) 대상
    wb.save(os.path.join(test_dir, "원본.xlsx"))

    from pyhwpx import Hwp
    report_path = os.path.join(tempfile.gettempdir(), "_test_보고서_요일불일치.hwp")
    setup = Hwp(visible=False, new=True)
    setup.insert_text("회의는 '26.9.7(화)에 진행하며, 예산은 185만원입니다")
    setup.save_as(report_path)
    setup.quit()

    report = None
    try:
        report = HwpReport(report_path)
        result = run_verification(report, source_paths=[test_dir], default_year=2026)
        assert "[요일불일치] ''26.9.7(화)' 실제로는 월요일" in result["summary"], result["summary"]
        assert report.get_char_color_at("'26.9.7(화)") == (255, 0, 0), (
            "요일 불일치 항목이 빨간색으로 표시되지 않음"
        )
        print("run_verification(요일불일치 감지) 통과:", result["summary"])
    finally:
        if report is not None:
            report.close(save=False)
        import shutil
        shutil.rmtree(test_dir)
        os.remove(report_path)


def _selftest_run_verification_mismatch_items_in_span_order():
    """(F13) mismatch_items가 문서에 실제로 빨갛게 표시된 순서(span 순서)
    그대로 나오는지 확인한다. 원본에 없는 값 두 개를 문서에 순서대로
    배치해서, 추출 순서(타입별로 묶임)가 아니라 진짜 문서 위치 순서를
    따르는지가 드러나게 한다."""
    test_dir = os.path.join(tempfile.gettempdir(), "_test_원본_다중오탐")
    os.makedirs(test_dir, exist_ok=True)
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Sheet1"
    ws["A1"] = "예산"; ws["A2"] = 1850000
    wb.save(os.path.join(test_dir, "원본.xlsx"))

    from pyhwpx import Hwp
    report_path = os.path.join(tempfile.gettempdir(), "_test_보고서_다중오탐.hwp")
    setup = Hwp(visible=False, new=True)
    setup.insert_text("예산은 8888888원이고, 그 다음은 7777777원입니다")
    setup.save_as(report_path)
    setup.quit()

    report = None
    try:
        report = HwpReport(report_path)
        result = run_verification(report, source_paths=[test_dir], default_year=2026)
        assert result["mismatch_items"] == ["8888888원", "7777777원"], result["mismatch_items"]
        print("run_verification(mismatch_items 순서) 통과:", result["mismatch_items"])
    finally:
        if report is not None:
            report.close(save=False)
        import shutil
        shutil.rmtree(test_dir)
        os.remove(report_path)


def _build_hwp_table(hwp, rows: list, merge_first_column: bool = False) -> None:
    """진짜 한글 표를 만들어 rows 내용을 채운다(테스트 픽스처용).
    merge_first_column=True면 1열의 데이터 행 두 칸을 세로 병합한다 —
    글자가 이미 들어있는 두 칸을 합치는 것이라, 한글이 두 문단을 한 칸에
    몰아넣는 실제 병합 동작(table_to_df는 값 복제, 평문은 두 줄)이 그대로
    재현된다."""
    hwp.create_table(rows=len(rows), cols=len(rows[0]))
    hwp.get_into_nth_table(0)
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            if val:
                hwp.insert_text(val)
            if not (r == len(rows) - 1 and c == len(row) - 1):
                hwp.TableRightCell()
    if merge_first_column:
        hwp.get_into_nth_table(0)
        for _ in range(len(rows[0])):  # 머리말 행을 지나 1열 첫 데이터 칸으로
            hwp.TableRightCell()
        hwp.TableCellBlock()
        hwp.TableCellBlockExtend()
        hwp.TableLowerCell()
        hwp.TableMergeCell()
    hwp.MoveDocEnd()


def _selftest_run_verification_table_cell_wrong_year_is_flagged():
    """(F14, 이 기능이 존재하는 이유 — 진짜 .hwp와 진짜 .xlsx로 확인)
    원본의 정기점검 실적은 2024년 1,694 / 2025년 1,827인데, 보고서 표의
    "2024년" 칸에 2025년 값(1,827)을 잘못 옮겨적은 상황이다.

    이 기능이 없으면 2024년 칸의 오류를 특정할 수 없다. 이제는 각 칸이
    자기 열(2024년/2025년)의 원본 값하고만 대조되므로, 2024년 칸은
    빨강(오류), 2025년 칸은 파랑(정상)이어야 한다. 같은 글자("1,827")가
    문서에 두 번 나오므로 색을 문서 순서대로 읽어 [빨강, 파랑]인지 확인한다."""
    test_dir = os.path.join(tempfile.gettempdir(), "_test_원본_표맥락")
    os.makedirs(test_dir, exist_ok=True)
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "실적"
    ws.append(["분류", "지표명", "2024", "2025"])
    ws.append(["점검", "정기점검 실적", 1694, 1827])
    wb.save(os.path.join(test_dir, "원본.xlsx"))

    from pyhwpx import Hwp
    report_path = os.path.join(tempfile.gettempdir(), "_test_보고서_표맥락.hwp")
    setup = Hwp(visible=False, new=True)
    setup.insert_text("2025년 주요 실적\n")
    _build_hwp_table(setup, [
        ["구분", "2024년", "2025년"],
        ["정기점검 실적", "1,827", "1,827"],  # 2024년 칸이 틀렸다(진짜 값은 1,694)
    ])
    setup.save_as(report_path)
    setup.quit()

    report = None
    try:
        report = HwpReport(report_path)
        run_verification(report, source_paths=[test_dir], default_year=2026)
        colors = report.get_char_colors("1,827")
        assert colors == [(255, 0, 0), (0, 0, 255)], (
            f"2024년 칸은 빨강, 2025년 칸은 파랑이어야 하는데 {colors}. "
            f"둘 다 같은 색이면 표 맥락이 안 붙은 것(기능 이전 동작)이다."
        )
        print("run_verification(표 2024년 칸의 잘못된 연도 값 적발) 통과:", colors)
    finally:
        if report is not None:
            report.close(save=False)
        import shutil
        shutil.rmtree(test_dir)
        os.remove(report_path)


def _selftest_run_verification_table_with_merged_cell_keeps_alignment():
    """(F14) 병합된 셀이 있는 진짜 한글 표에서도 줄 정렬이 밀리지 않아야 한다.
    table_to_df는 병합 칸의 값을 걸친 행마다 복제해 넣지만 평문 텍스트에는
    그 내용이 한 번만(여기서는 두 문단이라 두 줄로) 나오므로, 처리하지
    않으면 그 뒤의 모든 칸이 한 줄씩 밀려 엉뚱한 열 머리말이 붙는다.

    표 마지막 행의 "2024년" 칸에 2025년 값(600,000)을 적어두고, 그 값이
    빨강으로 잡히는지 본다 — 정렬이 밀렸다면 이 칸의 열 머리말이
    "2025년"으로 잘못 읽혀 파랑(정상)이 되어버린다. 즉 이 단언은 병합
    처리가 실제로 동작했는지를 결과로 확인한다."""
    test_dir = os.path.join(tempfile.gettempdir(), "_test_원본_병합표")
    os.makedirs(test_dir, exist_ok=True)
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "실적"
    ws.append(["구분", "2024", "2025"])
    ws.append(["점검 예산", 500000, 600000])
    wb.save(os.path.join(test_dir, "원본.xlsx"))

    from pyhwpx import Hwp
    report_path = os.path.join(tempfile.gettempdir(), "_test_보고서_병합표.hwp")
    setup = Hwp(visible=False, new=True)
    setup.insert_text("병합 표 확인\n")
    _build_hwp_table(setup, [
        ["구분", "2024년", "2025년"],
        ["점검", "500,000", "600,000"],   # 정상
        ["예산", "600,000", "600,000"],   # 2024년 칸이 틀렸다(진짜 값은 500,000)
    ], merge_first_column=True)
    setup.save_as(report_path)
    setup.quit()

    report = None
    try:
        report = HwpReport(report_path)
        run_verification(report, source_paths=[test_dir], default_year=2026)
        # 문서에 "600,000"은 세 번 나온다: 1행 2025년(정상, 파랑),
        # 2행 2024년(오류, 빨강), 2행 2025년(정상, 파랑).
        colors = report.get_char_colors("600,000")
        assert colors == [(0, 0, 255), (255, 0, 0), (0, 0, 255)], (
            f"병합 표에서 열 머리말 정렬이 어긋났다: {colors}"
        )
        # 정상 값은 그대로 파랑이어야 한다(병합이 앞쪽 행까지 망치지 않았는지).
        assert report.get_char_colors("500,000") == [(0, 0, 255)], \
            report.get_char_colors("500,000")
        print("run_verification(병합 셀 있는 표의 정렬 유지) 통과:", colors)
    finally:
        if report is not None:
            report.close(save=False)
        import shutil
        shutil.rmtree(test_dir)
        os.remove(report_path)


def _selftest_run_verification_table_with_tall_source_keeps_old_behavior():
    """(F14 회귀 방지 — 실사용에서 가장 중요한 보장) 보고서에 표가 있어도,
    원본이 가로형(연도별 열)이 아닌 보통 세로형 엑셀이면 판정이 표 맥락
    기능 이전과 똑같아야 한다.

    이 경우 원본 값에는 열/행 출처 정보가 아예 없으므로, "모르는 것은 막지
    않는다"는 원칙에 따라 후보가 하나도 걸러지지 않는다 — 즉 표 안 값도
    예전처럼 값이 맞으면 파랑, 틀리면 빨강이어야 한다. 여기서 회색이나
    초록이 나오면 멀쩡히 쓰던 검증이 조용히 무력화된 것이므로 반드시
    잡아야 한다."""
    test_dir = os.path.join(tempfile.gettempdir(), "_test_원본_세로형")
    os.makedirs(test_dir, exist_ok=True)
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Sheet1"
    ws.append(["항목", "값"])            # 가로형이 아님(2열, 연도 없음)
    ws.append(["예산", 1850000])
    wb.save(os.path.join(test_dir, "원본.xlsx"))

    from pyhwpx import Hwp
    report_path = os.path.join(tempfile.gettempdir(), "_test_보고서_세로형원본.hwp")
    setup = Hwp(visible=False, new=True)
    setup.insert_text("예산 현황\n")
    _build_hwp_table(setup, [
        ["구분", "금액"],
        ["예산", "1,850,000"],   # 원본과 일치 -> 파랑이어야 한다
        ["기타", "7,777,777"],   # 원본에 없음 -> 빨강이어야 한다
    ])
    setup.save_as(report_path)
    setup.quit()

    report = None
    try:
        report = HwpReport(report_path)
        run_verification(report, source_paths=[test_dir], default_year=2026)
        assert report.get_char_colors("1,850,000") == [(0, 0, 255)], (
            "가로형이 아닌 원본인데 표 안의 정상 값이 파랑이 아니다 — "
            f"{report.get_char_colors('1,850,000')} (기존 동작이 깨졌다)"
        )
        assert report.get_char_colors("7,777,777") == [(255, 0, 0)], \
            report.get_char_colors("7,777,777")
        print("run_verification(가로형 아닌 원본 + 표 = 기존 동작 유지) 통과")
    finally:
        if report is not None:
            report.close(save=False)
        import shutil
        shutil.rmtree(test_dir)
        os.remove(report_path)


def _selftest_run_verification_table_is_not_more_lenient_than_prose():
    """(2026-09-11, 두 기능 통합을 정면으로 겨냥한 테스트) 표 맥락은 후보를
    "좁히는" 장치일 뿐이므로, 좁힐 근거가 없을 때 표 안의 값이 평문보다
    관대하게 판정되면 안 된다.

    원본이 라벨 없는(숫자만 있는 행) 2열 시트라, master의 문맥 연결 규칙은
    "어느 항목인지 못 좁혔다"며 회색으로 남긴다. 같은 숫자를 표에 적었다고
    해서 파랑/빨강으로 단정하면, master가 이미 없앤 가짜 파랑·가짜 빨강이
    표 안에서만 되살아난다. 문장과 표 둘 다 회색이어야 한다.

    (이 픽스처는 표 맥락 기능을 처음 만들 때 "기존 동작 유지" 테스트로
    쓰였고 그때 기대값은 파랑/빨강이었다. 그 뒤 master에 문맥 연결 수정이
    들어가면서, 표가 없는 평문에서도 이 원본에 대한 판정이 회색으로
    바뀌었다 — 즉 기준선 자체가 옮겨갔다. 위의
    _selftest_run_verification_table_with_tall_source_keeps_old_behavior가
    라벨이 있는 원본으로 파랑/빨강 쪽을 따로 지킨다.)"""
    test_dir = os.path.join(tempfile.gettempdir(), "_test_원본_라벨없음")
    os.makedirs(test_dir, exist_ok=True)
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Sheet1"
    ws.append(["예산", "인원"])          # 항목명이 머리행에만 있고 데이터 행엔 없음
    ws.append([1850000, 342])
    wb.save(os.path.join(test_dir, "원본.xlsx"))

    from pyhwpx import Hwp
    report_path = os.path.join(tempfile.gettempdir(), "_test_보고서_라벨없는원본.hwp")
    setup = Hwp(visible=False, new=True)
    setup.insert_text("예산은 1,850,000원입니다\n")   # 평문 — 회색이 기준선
    _build_hwp_table(setup, [
        ["구분", "금액"],
        ["예산", "1,850,000"],
    ])
    setup.save_as(report_path)
    setup.quit()

    report = None
    try:
        report = HwpReport(report_path)
        run_verification(report, source_paths=[test_dir], default_year=2026)
        colors = report.get_char_colors("1,850,000")
        assert colors == [(128, 128, 128), (128, 128, 128)], (
            f"평문과 표의 판정이 갈렸다: {colors} — 표 안의 값만 관대하게 "
            f"판정되면 master의 문맥 연결 보호가 표 안에서 풀린 것이다."
        )
        print("run_verification(표가 평문보다 관대해지지 않음) 통과:", colors)
    finally:
        if report is not None:
            report.close(save=False)
        import shutil
        shutil.rmtree(test_dir)
        os.remove(report_path)


def _selftest_run_verification_routes_table_and_prose_in_one_document():
    """(2026-09-11, 두 기능 통합을 정면으로 겨냥한 테스트) 한 문서 안에 표와
    평문이 같이 있을 때, 각 값이 자기에게 맞는 판정 경로로 가야 한다.

    - 평문 "정기점검 실적은 1,694건" → 표 맥락이 없으므로 master의 라벨
      연결 경로. 문맥에 원본 라벨("정기점검 실적")이 걸려 1,694와 비교돼
      파랑이어야 한다.
    - 표의 "2024년" 칸 1,900 → 표 맥락 경로. 2024년 열의 원본 값(1,694)
      하고만 비교돼 빨강이어야 한다.
    - 표의 "2025년" 칸 1,827 → 같은 경로로 2025년 값과 일치해 파랑.

    셋이 동시에 맞아야 통과한다 — 한쪽 경로가 다른 쪽을 덮어쓰면(예: 표
    맥락 dict가 평문 문맥 문자열을 밀어내면) 반드시 깨진다."""
    test_dir = os.path.join(tempfile.gettempdir(), "_test_원본_표와평문")
    os.makedirs(test_dir, exist_ok=True)
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "실적"
    ws.append(["지표명", "2024", "2025"])
    ws.append(["정기점검 실적", 1694, 1827])
    wb.save(os.path.join(test_dir, "원본.xlsx"))

    from pyhwpx import Hwp
    report_path = os.path.join(tempfile.gettempdir(), "_test_보고서_표와평문.hwp")
    setup = Hwp(visible=False, new=True)
    setup.insert_text("점검 실적 보고\n")
    setup.insert_text("정기점검 실적은 1,694건입니다\n")   # 평문 → 라벨 연결 경로
    _build_hwp_table(setup, [
        ["구분", "2024년", "2025년"],
        ["정기점검 실적", "1,900", "1,827"],   # 2024년 칸이 틀렸다
    ])
    setup.save_as(report_path)
    setup.quit()

    report = None
    try:
        report = HwpReport(report_path)
        run_verification(report, source_paths=[test_dir], default_year=2026)
        assert report.get_char_colors("1,694") == [(0, 0, 255)], (
            "평문 값이 라벨 연결로 파랑이 되지 않았다 — "
            f"{report.get_char_colors('1,694')}"
        )
        assert report.get_char_colors("1,900") == [(255, 0, 0)], (
            "표의 2024년 칸 오류가 빨강으로 잡히지 않았다 — "
            f"{report.get_char_colors('1,900')}"
        )
        assert report.get_char_colors("1,827") == [(0, 0, 255)], (
            f"표의 2025년 칸 정상 값이 파랑이 아니다 — "
            f"{report.get_char_colors('1,827')}"
        )
        print("run_verification(한 문서 안 표/평문 경로 분기) 통과")
    finally:
        if report is not None:
            report.close(save=False)
        import shutil
        shutil.rmtree(test_dir)
        os.remove(report_path)


if __name__ == "__main__":
    _selftest_run_verification()
    _selftest_run_verification_skips_ignored_mismatch()
    _selftest_run_verification_clears_stale_marks_on_rerun()
    _selftest_run_verification_colors_all_three_categories_in_one_pass()
    _selftest_run_verification_connects_by_label_and_marks_ambiguous_gray()
    _selftest_run_verification_empty_answer_pool_gives_clear_message()
    _selftest_run_verification_flags_weekday_mismatch()
    _selftest_run_verification_mismatch_items_in_span_order()
    # (F14) 표 맥락 + 통합 검증
    _selftest_run_verification_table_cell_wrong_year_is_flagged()
    _selftest_run_verification_table_with_merged_cell_keeps_alignment()
    _selftest_run_verification_table_with_tall_source_keeps_old_behavior()
    _selftest_run_verification_table_is_not_more_lenient_than_prose()
    _selftest_run_verification_routes_table_and_prose_in_one_document()
