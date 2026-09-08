"""source_reader.py — 엑셀/한글/PDF 원본데이터를 읽어 verify_numbers의 정답 풀 형식으로 변환"""
import os
import re
import sys
import json
import datetime
import subprocess
from decimal import Decimal

import openpyxl
from verify_numbers import extract_values, compute_column_sums, _decimal_to_normalized_str


def read_excel_source(path: str, default_year: int) -> list[dict]:
    """엑셀 파일의 모든 시트, 모든 셀 값을 정답 풀 항목으로 변환한다.
    셀 값이 순수 숫자면 amount로, 날짜 타입이면 date로, 문자열이면 extract_values로
    4종을 뽑는다. 추가로 각 시트의 숫자 컬럼(1행이 헤더라고 가정) 합계도 파생값으로
    포함한다. 손상되었거나 읽을 수 없는 파일은 예외 없이 빈 리스트를 반환한다
    (PRD 12-5, read_pdf_source와 동일한 관례).

    (2026-09-04, 실사용 피드백으로 발견·수정) 원래는 무조건 1행을 헤더로
    가정했다 — 공공기관 실적표에서 흔한 "제목 행 → 출처/각주 행 → 빈 행 →
    진짜 헤더 행" 구조에서는 이 가정이 깨진다. 1행(제목, 셀 하나만 채워짐)을
    헤더로 잘못 삼으면, 진짜 헤더 행("2023","2024","2025","2026" 같은 연도
    라벨)이 "데이터"로 취급돼 그 라벨 자체가 금액 값처럼 정답 풀에 들어가는
    문제가 실제로 재현됐다 — 사용자 문서의 "('23)","('24)","('25)" 같은
    연도 약칭도 각각 23/24/25라는 숫자로 추출되는데, 이게 "2023"(4자리)과
    달라서 진짜 데이터가 아닌데도 "오류(빨강)"로 잘못 표시됐다. 이제는 셀이
    2칸 이상 채워진 첫 번째 행을 진짜 헤더로 찾는다(제목/각주/빈 행은 보통
    셀 1개 이하만 채워져 있음) — 그런 행을 못 찾으면(정말 한 칸짜리 헤더뿐인
    시트 등) 기존처럼 1행을 헤더로 쓴다(하위호환).

    알려진 한계(문서화만 하고 이번엔 해결하지 않음): 수식 셀은 data_only=True로
    열어도 실제 Excel에서 한 번도 저장된 적 없으면 캐시된 값이 없어 None으로
    읽힐 수 있다(합계 등 파생값이 원본에 있어도 못 읽는 경우 발생 가능).
    시간만 있는 셀(datetime.time, 날짜 없이 시각만 서식 지정된 셀)도 date/amount/str
    어디에도 안 걸려 조용히 누락된다 — 이 도구가 다루는 원본(주로 금액·날짜 중심의
    공공기관 계획서/실적표)에서는 드문 케이스라 이번 라운드는 해결하지 않고
    한계로만 남겨둔다.
    """
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception:
        return []
    results = []
    for ws in wb.worksheets:
        rows_as_dicts = []
        all_rows = list(ws.iter_rows(values_only=False))

        header = None
        header_row_idx = None
        for idx, row in enumerate(all_rows):
            non_empty_count = sum(1 for c in row if c.value is not None)
            if non_empty_count >= 2:  # 제목/각주/빈 행은 보통 셀 0~1개만 채워짐
                header = [c.value for c in row]
                header_row_idx = idx
                break
        if header is None and all_rows:
            # 여러 칸짜리 행을 못 찾았으면(정말 한 칸짜리 헤더뿐인 시트 등)
            # 기존 동작대로 1행을 그대로 헤더로 쓴다.
            header = [c.value for c in all_rows[0]]
            header_row_idx = 0

        for row in (all_rows[header_row_idx + 1:] if header_row_idx is not None else []):
            row_dict = {}
            # (2026-09-08 추가, D01/D05 문맥 연결 요건) 이 행의 문자열 셀 값을
            # 전부 모아 "라벨"로 삼는다 — "항목/값" 형태 표처럼 각 행이 한
            # 지표를 가리키는 경우, 항목명 셀("참여 인원" 등)이 그 행의 숫자·
            # 날짜 값의 문맥을 나타낸다. verify_numbers.categorize_values()가
            # 같은 타입 후보가 여럿일 때 보고서 문맥과 이 라벨의 겹침으로
            # 항목을 연결하는 데 쓴다(라벨이 없으면 그 값은 문맥 연결에서
            # 항상 빠지고, 후보가 하나뿐일 때만 비교 대상이 됨 — 안전한 기본값).
            row_label = " ".join(
                str(cell.value).strip() for cell in row
                if isinstance(cell.value, str) and cell.value.strip()
            )
            for col_name, cell in zip(header, row):
                value = cell.value
                if value is None:
                    continue
                if isinstance(value, (datetime.date, datetime.datetime)):
                    # (2026-08-30 재검토 후 추가) 엑셀 날짜 타입 셀 — extract_dates와
                    # 같은 형식(YYYY-MM-DD)으로 정규화해야 정답 풀에서 날짜로 인식된다.
                    results.append({
                        "type": "date", "normalized": value.strftime("%Y-%m-%d"),
                        "raw": str(value), "source_file": path, "location": f"{ws.title}!{cell.coordinate}",
                        "label": row_label,
                    })
                elif isinstance(value, (int, float)) and not isinstance(value, bool):
                    # (2026-08-30 Task 5/6 검토에서 미리 반영) bool은 int의 서브클래스라
                    # 별도 제외 필요. 정규화도 float 대신 Decimal 기반 공용 헬퍼를 써서
                    # 부동소수점 오차·과학적 표기법 문제를 처음부터 피한다.
                    results.append({
                        "type": "amount", "normalized": _decimal_to_normalized_str(Decimal(str(value))),
                        "raw": str(value), "source_file": path, "location": f"{ws.title}!{cell.coordinate}",
                        "label": row_label,
                    })
                    row_dict[col_name] = value
                elif isinstance(value, str):
                    for v in extract_values(value, default_year):
                        v["source_file"] = path
                        v["location"] = f"{ws.title}!{cell.coordinate}"
                        v["label"] = row_label
                        del v["span"]
                        results.append(v)
            if row_dict:
                rows_as_dicts.append(row_dict)
        results.extend(compute_column_sums(rows_as_dicts, source_file=path, sheet=ws.title))
    return results


def _selftest_read_excel_source():
    test_path = "_test_원본.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "실적"
    ws.append(["예산", "인원"])
    ws.append([500000, 10])
    ws.append([300000, 20])
    wb.save(test_path)
    try:
        result = read_excel_source(test_path, default_year=2026)
        amounts = sorted(r["normalized"] for r in result if r["type"] == "amount")
        assert "500000" in amounts and "300000" in amounts and "800000" in amounts, amounts
        assert all(r["source_file"] == test_path for r in result)
        print("read_excel_source 통과:", len(result), "건")
    finally:
        os.remove(test_path)


def _selftest_read_excel_source_attaches_row_label():
    """(2026-09-08 추가) "항목/값" 형태 표에서, 값 셀의 label에 같은 행의
    항목명 셀 텍스트가 담겨야 한다 — verify_numbers.categorize_values()가
    이 label로 보고서 문맥과 항목을 연결한다."""
    test_path = "_test_원본_라벨.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "실적"
    ws.append(["항목", "값"])
    ws.append(["참여 인원", 21])
    ws.append(["지원 인원", 40])
    wb.save(test_path)
    try:
        result = read_excel_source(test_path, default_year=2026)
        by_value = {r["normalized"]: r["label"] for r in result if r["type"] == "amount"
                    and r["normalized"] in ("21", "40")}
        assert by_value == {"21": "참여 인원", "40": "지원 인원"}, by_value
        print("read_excel_source_attaches_row_label 통과:", by_value)
    finally:
        os.remove(test_path)


def _selftest_read_excel_source_date_cell():
    """엑셀의 날짜 타입 셀(datetime.date)이 date 항목으로 정확히 인식되어야 한다."""
    test_path = "_test_원본_날짜.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "일정"
    ws.append(["착수일", "예산"])
    ws.append([datetime.date(2026, 9, 7), 500000])
    wb.save(test_path)
    try:
        result = read_excel_source(test_path, default_year=2026)
        dates = [r for r in result if r["type"] == "date"]
        assert len(dates) == 1 and dates[0]["normalized"] == "2026-09-07", dates
        print("read_excel_source_date_cell 통과:", dates)
    finally:
        os.remove(test_path)


def _selftest_read_excel_source_corrupted_file_no_crash():
    """손상된/엑셀이 아닌 파일도 예외 없이 빈 리스트를 반환해야 한다."""
    test_path = "_test_손상됨.xlsx"
    with open(test_path, "w") as f:
        f.write("이건 엑셀 파일이 아님")
    try:
        result = read_excel_source(test_path, default_year=2026)
        assert result == [], result
        print("read_excel_source_corrupted_file_no_crash 통과:", result)
    finally:
        os.remove(test_path)


def _selftest_read_excel_source_title_row_before_real_header():
    """(2026-09-04, 실사용 피드백으로 발견한 실제 버그 재현) "제목 행 → 각주 행
    → 빈 행 → 진짜 헤더 행" 구조(공공기관 실적표에서 흔함)에서, 1행(제목)을
    헤더로 잘못 삼지 않고 진짜 헤더 행(셀 2개 이상 채워진 첫 행)을 찾아야
    한다. 잘못되면 진짜 헤더 행의 연도 라벨("2023","2024" 등)이 데이터 값처럼
    정답 풀에 섞여 들어간다 — 실사용 문서에서 실제 재현된 버그."""
    test_path = "_test_제목행버그.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "요약"
    ws.append(["2025년 핵심 추이 요약"])  # 제목 행(셀 1개)
    ws.append(["출처: ...hwp"])  # 각주 행(셀 1개)
    ws.append([])  # 빈 행
    ws.append(["분류", "지표명", "2023", "2024", "2025"])  # 진짜 헤더 행(셀 5개)
    ws.append(["점검", "정기점검 실적", 1595, 1694, 1827])
    wb.save(test_path)
    try:
        result = read_excel_source(test_path, default_year=2026)
        normalized_amounts = {r["normalized"] for r in result if r["type"] == "amount"}
        assert {"1595", "1694", "1827"} <= normalized_amounts, normalized_amounts
        # 진짜 헤더 행의 연도 라벨이 데이터 값으로 섞여 들어가면 안 된다 —
        # "2023"/"2024"가 컬럼합계(1595/1694 그대로, 데이터가 한 행뿐이라
        # 합계=그 값)로 나오는 건 정상이지만, 헤더 셀 자체(location이 "!2023"
        # 같은 셀 주소 형식)가 별도 항목으로 잡히면 안 된다.
        header_cell_locations = [r["location"] for r in result if r["location"].endswith(("!2023", "!2024", "!2025"))]
        assert header_cell_locations == [], header_cell_locations
        print("read_excel_source_title_row_before_real_header 통과:", result)
    finally:
        os.remove(test_path)


def read_hwp_source(path: str, default_year: int) -> list[dict]:
    """한글 문서를 안 보이게(visible=False) 열어서 전체 텍스트를 읽고,
    4종 값을 뽑아 정답 풀 항목으로 변환한다. 원본 참고용으로만 열기 때문에
    화면에 띄우지 않는다 (사용자가 실제로 편집 중인 보고서와는 별개의 인스턴스).
    손상되었거나 열 수 없는 파일은 예외 없이 빈 리스트를 반환한다 (PRD 12-5,
    read_excel_source/read_pdf_source와 동일한 관례). hwp.quit()은 어떤 예외가
    나든 반드시 호출되도록 try/finally로 감싼다 — 그렇지 않으면 보이지 않는
    Hwp.exe 프로세스가 누적되어 남는 문제가 실제로 재현된 적 있다.

    (2026-09-01 F12 Task10 최종검증 중 발견 — 치명적, 수정됨) new=True 없이
    Hwp(visible=False)만 쓰면, pyhwpx는 이미 실행 중인 한글 프로세스가 있을 때
    그 프로세스에 그대로 접속(재사용)해버린다(hwp_report.py의 HwpReport가 이미
    문서화한, 이 프로젝트에서 여러 번 확인된 동작). F12에서 채팅 프로그램이
    HwpReport로 보고서를 이미 열어둔 상태(new=True로 독립 실행됨)에서
    run_verification이 이 함수를 호출하면, 위 hwp가 그 "이미 열려있는 보고서"의
    프로세스에 붙어버리고, finally의 hwp.quit()이 그 공유 프로세스 전체를
    종료시켜 사용자가 보고 있던 보고서 창까지 함께 꺼져버리는 게 실제로
    재현됐다(get_text() 등 이후 호출이 "개체가 열려 있지 않거나 등록되지
    않았습니다" COM 오류로 실패). new=True를 추가해 이 함수가 항상 독립된
    새 프로세스에서만 동작하도록 한다.
    """
    from pyhwpx import Hwp
    hwp = None
    try:
        hwp = Hwp(visible=False, new=True)
        if not hwp.open(path):
            return []
        text = hwp.GetTextFile("TEXT", "")
        if not text:
            return []
    except Exception:
        return []
    finally:
        if hwp is not None:
            hwp.quit()

    results = extract_values(text, default_year)
    for r in results:
        start, end = r["span"]
        before = text[max(0, start - 10):start]
        matched = text[start:end]
        after = text[end:end + 10]
        r["source_file"] = path
        r["location"] = f"{before}**{matched}**{after}"
        del r["span"]
    return results


def _selftest_read_hwp_source():
    from pyhwpx import Hwp
    test_path = os.path.abspath("_test_원본.hwp")
    # (2026-09-04, 실사용 세션 중 실제 재현·발견) new=True가 빠지면, 이미 떠
    # 있는 다른 한글 프로세스가 있을 때 pyhwpx가 그 프로세스에 그대로 접속
    # (재사용)해버린다 — 그 상태에서 insert_text()는 그 기존 문서 내용
    # 뒤에 테스트 문장을 이어붙이고, save_as()는 그 오염된 전체 내용을
    # 테스트 파일로 저장해버려, 이 테스트가 기대하는 "값 3종만 있는 깨끗한
    # 문서"라는 전제가 깨진다(직접 재현: 다른 테스트가 남긴 큰 문서가 섞여
    # amount만 수백 건 나오는 것으로 확인됨). read_hwp_source() 자신은 이미
    # new=True를 쓰고 있었는데, 정작 이 테스트의 픽스처 설정에서만 빠져있던
    # 것 — HwpReport 클래스가 이미 강조한 것과 같은 원칙을 테스트 설정에도
    # 그대로 적용한다.
    hwp = Hwp(visible=False, new=True)
    hwp.insert_text("예산은 1,850,000원이며 회의는 2026-09-07 14:00~16:00 진행")
    hwp.save_as(test_path)
    hwp.quit()
    try:
        result = read_hwp_source(test_path, default_year=2026)
        types = sorted(r["type"] for r in result)
        assert types == ["amount", "date", "time"], types
        print("read_hwp_source 통과:", result)
    finally:
        os.remove(test_path)


def _selftest_read_hwp_source_no_match():
    """숫자/날짜/시간/전화번호가 전혀 없는 문서는 빈 리스트를 반환해야 한다."""
    from pyhwpx import Hwp
    test_path = os.path.abspath("_test_원본_빈값.hwp")
    hwp = Hwp(visible=False, new=True)  # 다른 프로세스 재사용 방지 — 위 _selftest_read_hwp_source 참고
    hwp.insert_text("이 문서에는 특별한 값이 없습니다")
    hwp.save_as(test_path)
    hwp.quit()
    try:
        result = read_hwp_source(test_path, default_year=2026)
        assert result == [], result
        print("read_hwp_source_no_match 통과:", result)
    finally:
        os.remove(test_path)


def _selftest_read_hwp_source_missing_file_no_crash():
    """존재하지 않는 파일은 예외 없이 빈 리스트를 반환해야 한다 (프로세스 누수 방지 포함)."""
    result = read_hwp_source(os.path.abspath("_존재하지_않는_파일.hwp"), default_year=2026)
    assert result == [], result
    print("read_hwp_source_missing_file_no_crash 통과:", result)


def read_hwp_source_isolated(path: str, default_year: int) -> list[dict]:
    """read_hwp_source()를 완전히 별도의 파이썬 프로세스에서 실행해 결과를
    받아온다(2026-09-04, 실사용 피드백 — 사용자가 "별도 프로세스로 격리하면
    .hwp도 원본자료로 쓸 수 있지 않냐"고 직접 제안함).

    같은 프로세스 안에서 read_hwp_source()를 직접 호출하면, pyhwpx의
    Hwp.__del__이 그 Hwp 인스턴스를 정리할 때 pythoncom.CoUninitialize()를
    무조건 호출하는데, 이게 같은 프로세스 안의 "다른" Hwp 인스턴스(F12가
    채팅 세션 내내 열어두고 있는, 사용자가 실제로 편집 중인 보고서
    HwpReport)의 COM 연결까지 함께 끊어버리는 게 실측 확인된 pyhwpx 자체의
    한계다(_READERS 근처의 기존 주석 참고, 2026-09-01 F12 Task10에서 처음
    발견) — 그래서 .hwp/.hwpx가 이번까지 _READERS에서 빠져 있었다.

    완전히 다른 프로세스에서 실행하면 이 문제를 피할 수 있다 — 별도 프로세스는
    자기만의 COM 아파트를 가지므로, 그 프로세스가 끝나며 CoUninitialize를
    호출해도 원래 프로세스(채팅 세션)의 COM 연결과는 아무 상관이 없다.
    subprocess로 `python -c "..."` 형태로 read_hwp_source를 호출하고,
    표준출력으로 JSON을 받아 파싱한다(read_hwp_source의 반환값은 문자열
    필드만 가진 dict 리스트라 JSON 직렬화에 문제가 없다). 어떤 이유로든
    실패하면(타임아웃, 프로세스 오류, JSON 파싱 실패 등) 이 모듈의 다른
    리더들과 같은 관례대로 예외 없이 빈 리스트를 반환한다.
    """
    module_dir = os.path.dirname(os.path.abspath(__file__))
    script = (
        "import sys; sys.path.insert(0, sys.argv[3]); "
        "import json; from source_reader import read_hwp_source; "
        "print(json.dumps(read_hwp_source(sys.argv[1], int(sys.argv[2]))))"
    )
    try:
        result = subprocess.run(
            [sys.executable, "-c", script, path, str(default_year), module_dir],
            capture_output=True, text=True, timeout=60,
        )
        if result.returncode != 0:
            return []
        return json.loads(result.stdout)
    except Exception:
        return []


def _selftest_read_hwp_source_isolated_basic():
    """read_hwp_source_isolated()가 별도 프로세스를 통해서도 read_hwp_source()와
    동일한 결과(4종 값 추출)를 돌려주는지 확인한다."""
    from pyhwpx import Hwp
    test_path = os.path.abspath("_test_원본_격리.hwp")
    setup = Hwp(visible=False, new=True)
    setup.insert_text("예산은 1,850,000원이며 회의는 2026-09-07 14:00~16:00 진행")
    setup.save_as(test_path)
    setup.quit()
    try:
        result = read_hwp_source_isolated(test_path, default_year=2026)
        types = sorted(r["type"] for r in result)
        assert types == ["amount", "date", "time"], types
        print("read_hwp_source_isolated 통과:", result)
    finally:
        os.remove(test_path)


def _selftest_read_hwp_source_isolated_missing_file_no_crash():
    """존재하지 않는 파일도 예외 없이 빈 리스트를 반환해야 한다."""
    result = read_hwp_source_isolated(os.path.abspath("_존재하지_않는_파일_격리.hwp"), default_year=2026)
    assert result == [], result
    print("read_hwp_source_isolated_missing_file_no_crash 통과:", result)


def _selftest_read_hwp_source_isolated_does_not_break_live_hwp_instance():
    """(2026-09-04, 가장 중요한 회귀 테스트) read_hwp_source_isolated()가
    원래 버그(2026-09-01 F12 Task10에서 발견)를 실제로 피하는지 확인한다 —
    "채팅 세션이 계속 열어두고 있는, 사용자가 편집 중인 문서"를 흉내낸 살아있는
    Hwp 인스턴스를 하나 열어둔 채로, 별도의 .hwp 원본자료를
    read_hwp_source_isolated()로 읽고, 그 후에도 살아있던 인스턴스가 여전히
    정상 동작하는지(get_text 호출이 COM 오류 없이 성공하는지) 검증한다.

    원래 버그(read_hwp_source를 같은 프로세스에서 직접 호출)는 이 정확한
    시나리오에서 살아있던 인스턴스의 COM 연결을 끊어버렸었다 — 격리된
    버전은 완전히 다른 프로세스에서 실행되므로 이 문제가 재현되지 않아야
    한다."""
    from pyhwpx import Hwp

    live_path = os.path.abspath("_test_살아있는문서.hwp")
    live_setup = Hwp(visible=False, new=True)
    live_setup.insert_text("사용자가 지금 편집 중인 문서입니다")
    live_setup.save_as(live_path)
    live_setup.quit()

    source_path = os.path.abspath("_test_원본자료_격리.hwp")
    source_setup = Hwp(visible=False, new=True)
    source_setup.insert_text("원본 예산은 1,850,000원입니다")
    source_setup.save_as(source_path)
    source_setup.quit()

    live_instance = None
    try:
        live_instance = Hwp(visible=False, new=True)
        assert live_instance.open(live_path), "살아있는 문서를 열지 못함"
        before_text = live_instance.GetTextFile("TEXT", "")
        assert "편집 중인 문서" in before_text, before_text

        # 살아있는 인스턴스가 열려있는 채로, 별도 프로세스로 원본자료를 읽는다.
        result = read_hwp_source_isolated(source_path, default_year=2026)
        assert any(r["normalized"] == "1850000" for r in result), result

        # 원래 버그라면 여기서 COM 오류("개체가 열려 있지 않거나 등록되지
        # 않았습니다" 등)가 났을 것이다 — 격리됐으므로 정상 동작해야 한다.
        after_text = live_instance.GetTextFile("TEXT", "")
        assert after_text == before_text, (
            "살아있던 Hwp 인스턴스의 연결이 깨짐(원래 버그가 재현됨)", after_text
        )
        print("read_hwp_source_isolated_does_not_break_live_hwp_instance 통과: "
              "살아있는 문서 연결이 그대로 유지됨")
    finally:
        if live_instance is not None:
            live_instance.quit()
        os.remove(live_path)
        os.remove(source_path)


def read_pdf_source(path: str, default_year: int) -> list[dict]:
    """디지털 PDF에서 텍스트를 추출해 정답 풀 항목으로 변환한다. 표 형태 데이터는
    별도 표 파싱(extract_tables) 없이 흐르는 텍스트로 추출하는데, 칸이 잘
    구분된 단순한 표는 이 방식으로도 값이 잘 잡히지만, 여러 줄에 걸친 셀이나
    칸 구분이 애매한 복잡한 표는 놓칠 수 있다(알려진 한계, 이번엔 해결 안 함).
    파일이 없거나, 텍스트를 전혀 추출할 수 없는 스캔 이미지 PDF인 경우
    예외를 던지지 않고 빈 리스트를 반환한다 (PRD 12-2: OCR 비목표, 조용히 건너뜀).
    """
    import pdfplumber
    if not os.path.exists(path):
        return []
    results = []
    try:
        with pdfplumber.open(path) as pdf:
            for page_num, page in enumerate(pdf.pages, start=1):
                text = page.extract_text() or ""
                if not text.strip():
                    continue  # 텍스트 없음 = 스캔 이미지로 추정, 건너뜀
                for v in extract_values(text, default_year):
                    v["source_file"] = path
                    v["location"] = f"{page_num}페이지"
                    del v["span"]
                    results.append(v)
    except Exception:
        return []  # 손상되었거나 읽을 수 없는 PDF는 조용히 건너뜀 (PRD 12-5)
    return results


def _selftest_read_pdf_source_missing_file():
    result = read_pdf_source("존재하지_않는_파일.pdf", default_year=2026)
    assert result == [], result
    print("read_pdf_source(없는 파일) 통과: 빈 리스트, 오류 없이 건너뜀")


def _selftest_read_pdf_source_multipage_skips_blank():
    """여러 페이지 중 빈 페이지(스캔 이미지 가정)가 껴 있어도, 나머지 페이지는
    정상적으로 처리되어야 한다."""
    from fpdf import FPDF
    test_path = "_test_원본_다중페이지.pdf"
    pdf = FPDF()
    # 참고: fpdf2 기본 코어 폰트("Helvetica")는 Latin-1만 지원해 한글이
    # FPDFUnicodeEncodingException으로 실패한다(이 환경에서 실측 확인).
    # Windows 기본 제공 맑은 고딕(TTF)을 등록해 한글 테스트 문자열을 그대로 쓴다.
    pdf.add_font("Malgun", fname="C:/Windows/Fonts/malgun.ttf")
    pdf.add_page(); pdf.set_font("Malgun", size=12)
    pdf.cell(0, 10, "예산은 1,850,000원입니다")
    pdf.add_page()  # 빈 페이지 (텍스트 없음)
    pdf.add_page(); pdf.set_font("Malgun", size=12)
    pdf.cell(0, 10, "인원은 342명입니다")
    pdf.output(test_path)
    try:
        result = read_pdf_source(test_path, default_year=2026)
        locations = sorted(r["location"] for r in result)
        assert locations == ["1페이지", "3페이지"], locations  # 2페이지(빈 페이지)는 건너뜀
        print("read_pdf_source_multipage_skips_blank 통과:", result)
    finally:
        os.remove(test_path)


def _selftest_read_pdf_source_simple_table():
    """칸이 잘 구분된 단순 표 형태 데이터도 텍스트 추출로 값이 잡혀야 한다."""
    from fpdf import FPDF
    test_path = "_test_원본_표.pdf"
    pdf = FPDF()
    # 위 _selftest_read_pdf_source_multipage_skips_blank와 동일한 이유로
    # 한글 렌더링을 위해 맑은 고딕(TTF)을 등록한다.
    pdf.add_font("Malgun", fname="C:/Windows/Fonts/malgun.ttf")
    pdf.add_page(); pdf.set_font("Malgun", size=12)
    pdf.cell(60, 10, "항목", border=1); pdf.cell(60, 10, "금액", border=1); pdf.ln()
    pdf.cell(60, 10, "예산", border=1); pdf.cell(60, 10, "1850000", border=1); pdf.ln()
    pdf.output(test_path)
    try:
        result = read_pdf_source(test_path, default_year=2026)
        amounts = [r["normalized"] for r in result if r["type"] == "amount"]
        assert "1850000" in amounts, amounts
        print("read_pdf_source_simple_table 통과:", result)
    finally:
        os.remove(test_path)


# (2026-09-01 F12 Task10 최종검증 중 발견 — 치명적, 임시 조치로 .hwp/.hwpx 제외)
# read_hwp_source는 원본 파일을 읽으려고 임시로 자기만의 Hwp() 인스턴스를 새로
# 열었다가 다 읽으면 닫는데, pyhwpx의 Hwp.__del__이 어떤 Hwp 객체든 상관없이
# pythoncom.CoUninitialize()를 무조건 호출하는 것으로 확인됨(pyhwpx 소스
# core.py 1008-1014행) — 이게 같은 파이썬 프로세스/스레드 안에 있는 "다른"
# Hwp 인스턴스(F12에서는 채팅 프로그램이 계속 열어두고 있는 보고서 문서,
# HwpReport)의 COM 연결까지 깨뜨려버리는 게 최소 재현으로 확인됐다 — new=True
# 로 독립 인스턴스를 만들어도, quit() 후 제대로 정리해도 마찬가지였고,
# pythoncom.CoInitialize()를 다시 호출해 복구를 시도해도 안 됐다(마샬링된
# 인터페이스 프록시 자체가 죽어버리는 것으로 보임 — 이건 pyhwpx 자체의
# 한계이지 이 프로젝트 코드의 버그가 아님). 결과: 원본자료로 .hwp 파일을
# 첨부하고 검증을 돌리면, 사용자가 보고 있던 보고서 문서 창의 연결이
# 조용히 끊겨버리는 심각한 문제가 있었다.
#
# (2026-09-04 갱신) 위 문제의 근본 해결책으로 남겨뒀던 "별도 프로세스로
# 격리해서 읽기"를 사용자 제안으로 이번에 실제로 구현했다 — read_hwp_source
# 자체가 아니라 read_hwp_source_isolated(완전히 새 파이썬 프로세스에서
# read_hwp_source를 실행)를 _READERS에 연결한다. 별도 프로세스는 자기만의
# COM 아파트를 쓰므로, 그 프로세스가 끝나며 CoUninitialize를 호출해도
# 채팅 세션의 HwpReport와는 완전히 무관하다 — 직접 재현 테스트로 확인함
# (아래 _selftest_read_hwp_source_isolated_does_not_break_live_hwp_instance
# 참고, 원래 버그를 그대로 재현하는 시나리오로 검증).
_READERS = {".xlsx": read_excel_source, ".xls": read_excel_source,
            ".pdf": read_pdf_source,
            ".hwp": read_hwp_source_isolated, ".hwpx": read_hwp_source_isolated}


_CELL_ADDRESS_PATTERN = re.compile(r'^[A-Za-z]{1,3}\d+$')


def read_source_files(paths: list[str], default_year: int) -> tuple[list[dict], list[dict]]:
    """경로 리스트(개별 파일과 폴더가 섞여 있어도 됨)를 읽어 정답 풀과 충돌
    목록을 만든다. F12에서 "+" 버튼으로 파일 여러 개를 개별 선택하거나
    폴더를 선택하는 두 경우를 하나의 함수로 통일해서 처리하기 위함이다
    (2026-08-31 이전까지는 read_source_folder가 폴더 하나만 받았음).

    폴더 경로가 섞여 있으면 그 폴더 바로 아래(하위 폴더 재귀 없음, 기존
    read_source_folder와 동일한 한계) 파일들을 각각 개별 파일처럼 펼쳐서 읽는다.
    처리 가능한 확장자만 실제로 읽고 나머지는 조용히 건너뛴다. 서로 다른 파일의
    "같은 시트!같은 셀 주소"에 다른 값이 있으면 충돌로 간주해 별도 리스트로
    반환한다 (PRD 12-5).

    반환: (정답_풀, 충돌_목록)
    충돌_목록 항목 형식: {"location": "Sheet1!A2", "type": "amount", "values": [
        {"file": "초안.xlsx", "normalized": "1800000"},
        {"file": "최종.xlsx", "normalized": "1850000"}]}

    (2026-08-31 코드품질 검토 후 명시) 충돌은 (location, type) 단위로 탐지하므로,
    같은 location이 서로 다른 타입의 충돌 항목으로 두 번 이상 나타날 수 있다 —
    예를 들어 한 셀에 금액도 날짜도 잘못된 값이 있으면 "Sheet1!A2"가 amount
    충돌 하나, date 충돌 하나로 각각 별도 항목이 된다. 즉 location만으로는
    충돌_목록 항목을 유일하게 식별할 수 없다.

    (2026-08-30 구현 중 실제 테스트로 발견 — 중요) 최초 버전은 "!" 개수만 세어
    "Sheet1!A2"(진짜 셀 주소)와 "Sheet1!예산(합계)"(compute_column_sums가 만드는
    파생 컬럼합계 라벨, Task 6)을 구분하지 못했다 — 둘 다 "!"가 정확히 1개라서
    똑같이 충돌 탐지 대상으로 잡혔다. 실제로 이 테스트 데이터(시트당 데이터 행이
    1개뿐이라 컬럼합계가 그 한 셀 값과 같음)로 돌려보니, 파일마다 컬럼합계 값도
    달라서 "Sheet1!A2" 충돌 1건과 "Sheet1!예산(합계)" 충돌 1건, 총 2건이 나와
    "같은 셀 주소" 하나만 충돌로 잡는다는 모델링 결정과 어긋났다. "!" 뒤쪽이
    "A2"처럼 열문자+행번호로 된 진짜 셀 주소 형식인지까지 확인해서, 파생 라벨은
    애초에 충돌 탐지 후보에서 제외한다 (범위를 넓히는 게 아니라, "같은 셀 주소"라는
    원래 정의를 정확히 지키기 위한 수정).

    (2026-09-01 코드품질 검토 후 수정) 폴더를 펼치는 os.listdir 호출은 권한 문제,
    또는 파일 선택과 실제 호출 사이에 폴더가 삭제되는 등의 이유로 실패할 수 있다.
    이 모듈의 다른 리더들(read_excel_source 등)과 같은 관례로, 그런 폴더는
    예외를 던지지 않고 조용히 건너뛴다. 또한 같은 경로가 리스트에 두 번 들어오거나
    (다중 선택 파일 대화상자에서 실수로 같은 파일을 두 번 고르는 경우 등), 같은
    파일이 상대경로/절대경로로 각각 한 번씩 들어오면, 펼쳐진 경로 목록을 절대경로
    기준으로 중복 제거해 같은 파일이 두 번 읽혀 정답 풀이 조용히 부풀려지는 것을
    막는다.
    """
    expanded_paths = []
    for path in paths:
        if os.path.isdir(path):
            try:
                expanded_paths.extend(os.path.join(path, name) for name in os.listdir(path))
            except OSError:
                continue  # 권한 문제 등으로 폴더를 열 수 없으면 조용히 건너뜀 (기존 리더들과 같은 관례)
        else:
            expanded_paths.append(path)

    # 같은 파일이 두 번 들어오거나(중복 선택) 상대/절대경로로 각각 한 번씩
    # 들어와도, 실제로는 같은 파일이면 한 번만 읽도록 절대경로 기준으로
    # 중복 제거한다(순서는 유지).
    expanded_paths = list(dict.fromkeys(os.path.normpath(os.path.abspath(p)) for p in expanded_paths))

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
        # (2026-08-31 코드품질 검토 후 수정됨 — 중요) 키를 location만으로 잡으면,
        # 같은 셀에서 금액+날짜처럼 여러 타입이 같이 나올 때(예: "1,850,000원
        # 2026-09-07"이 든 셀) 파일별로 나중 타입이 앞 타입을 덮어써서, 날짜와
        # 금액을 서로 비교하는 말도 안 되는 "충돌"이 나오거나 진짜 금액 불일치가
        # 가려지는 문제가 있었다. (location, type) 튜플로 키를 잡아 타입별로
        # 따로 그룹핑해야 한다.
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


def _selftest_read_source_folder_conflict():
    os.makedirs("_test_원본폴더", exist_ok=True)
    wb1 = openpyxl.Workbook(); ws1 = wb1.active; ws1.title = "Sheet1"
    ws1["A1"] = "예산"; ws1["A2"] = 1800000
    wb1.save("_test_원본폴더/초안.xlsx")
    wb2 = openpyxl.Workbook(); ws2 = wb2.active; ws2.title = "Sheet1"
    ws2["A1"] = "예산"; ws2["A2"] = 1850000
    wb2.save("_test_원본폴더/최종.xlsx")
    with open("_test_원본폴더/무관한파일.txt", "w") as f:
        f.write("이건 읽으면 안 되는 파일")
    try:
        pool, conflicts = read_source_folder("_test_원본폴더", default_year=2026)
        assert len(conflicts) == 1, conflicts
        assert conflicts[0]["location"] == "Sheet1!A2", conflicts
        print("read_source_folder 통과: 충돌", conflicts)
    finally:
        import shutil
        shutil.rmtree("_test_원본폴더")


def _selftest_read_source_folder_multi_type_cell_no_cross_type_conflict():
    """같은 셀에서 금액+날짜 등 여러 타입이 같이 나올 때, 타입을 섞어서 엉뚱하게
    비교하면 안 된다 — 실제 금액 불일치만 정확히 잡아야 한다."""
    os.makedirs("_test_원본폴더2", exist_ok=True)
    wb1 = openpyxl.Workbook(); ws1 = wb1.active; ws1.title = "Sheet1"
    ws1["A1"] = "메모"; ws1["A2"] = "예산은 1,850,000원이며 회의는 2026-09-07 진행"
    wb1.save("_test_원본폴더2/초안.xlsx")
    wb2 = openpyxl.Workbook(); ws2 = wb2.active; ws2.title = "Sheet1"
    ws2["A1"] = "메모"; ws2["A2"] = 1900000  # 금액만 다름(진짜 불일치), 날짜는 없음
    wb2.save("_test_원본폴더2/최종.xlsx")
    try:
        pool, conflicts = read_source_folder("_test_원본폴더2", default_year=2026)
        amount_conflicts = [c for c in conflicts if any(
            v["normalized"] in ("1850000", "1900000") for v in c["values"])]
        assert len(amount_conflicts) == 1, conflicts
        values = sorted(v["normalized"] for v in amount_conflicts[0]["values"])
        assert values == ["1850000", "1900000"], values  # 날짜와 뒤섞이지 않아야 함
        print("read_source_folder_multi_type_cell_no_cross_type_conflict 통과:", conflicts)
    finally:
        import shutil
        shutil.rmtree("_test_원본폴더2")


def _selftest_read_source_files_mixed_list():
    """폴더 경로 하나와 개별 파일 경로 하나가 섞인 리스트를 줘도 둘 다 읽어서
    합쳐지는지 확인한다 (F12: "+" 버튼으로 파일 여러 개 또는 폴더를 자유롭게
    섞어 첨부할 수 있어야 하므로)."""
    os.makedirs("_test_원본_혼합", exist_ok=True)
    wb1 = openpyxl.Workbook(); ws1 = wb1.active; ws1.title = "Sheet1"
    ws1["A1"] = "예산"; ws1["A2"] = 1850000
    wb1.save("_test_원본_혼합/폴더안파일.xlsx")

    # (구현 중 실제 테스트로 발견) B열을 쓰는 이유: 두 파일 모두 A열을 쓰면
    # "예산"(파일1)과 "인원"(파일2)이라는 서로 무관한 값인데도 우연히 같은
    # 셀 주소(Sheet1!A2)에 놓여, 이 혼합 리스트 테스트의 의도(두 출처가 각각
    # 정상적으로 읽혀 합쳐지는지 확인)와 무관하게 read_source_files의 진짜
    # 충돌 탐지 로직(같은 위치·다른 값 → 충돌, read_source_folder 시절부터
    # 있던 의도된 동작)이 걸려 conflicts가 비지 않게 된다. 서로 다른 열을 써서
    # 이 테스트가 우연한 셀 충돌이 아니라 원래 검증하려던 것(혼합 리스트 병합)만
    # 확인하도록 한다.
    wb2 = openpyxl.Workbook(); ws2 = wb2.active; ws2.title = "Sheet1"
    ws2["B1"] = "인원"; ws2["B2"] = 12
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


def _selftest_read_source_files_folder_listdir_permission_error_skipped():
    """(2026-09-01 코드품질 검토 반영) 폴더를 펼치는 os.listdir이 권한 문제 등으로
    실패해도(PermissionError/OSError) 예외를 던지지 않고 그 폴더만 조용히
    건너뛰어야 하며, 같은 호출에 함께 들어온 다른 정상 경로는 영향받지 않고
    정상적으로 읽혀야 한다. 실제 권한 없는 폴더는 이식 가능하게 재현하기
    어려워, 실재하는 두 폴더 중 하나에 대해서만 os.listdir이 예외를 던지도록
    mock으로 시뮬레이션한다."""
    from unittest.mock import patch
    os.makedirs("_test_원본_정상폴더", exist_ok=True)
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Sheet1"
    ws["A1"] = "예산"; ws["A2"] = 1850000
    wb.save("_test_원본_정상폴더/원본.xlsx")

    os.makedirs("_test_원본_권한없음폴더", exist_ok=True)
    blocked_abs = os.path.abspath("_test_원본_권한없음폴더")
    real_listdir = os.listdir

    def fake_listdir(path):
        if os.path.abspath(path) == blocked_abs:
            raise PermissionError("접근 거부(시뮬레이션)")
        return real_listdir(path)

    try:
        with patch("os.listdir", side_effect=fake_listdir):
            pool, conflicts = read_source_files(
                ["_test_원본_권한없음폴더", "_test_원본_정상폴더"], default_year=2026
            )
        normalized_values = {item["normalized"] for item in pool}
        assert "1850000" in normalized_values, pool
        print("read_source_files_folder_listdir_permission_error_skipped 통과:", normalized_values)
    finally:
        import shutil
        shutil.rmtree("_test_원본_정상폴더")
        shutil.rmtree("_test_원본_권한없음폴더")


def _selftest_read_source_files_dedup_duplicate_path():
    """(2026-09-01 코드품질 검토 반영) 같은 파일 경로가 리스트에 두 번 들어오면
    (다중 선택 파일 대화상자에서 실수로 같은 파일을 두 번 고르는 경우 등) 정답
    풀이 두 배로 부풀려지면 안 된다. 같은 파일을 상대경로/절대경로로 각각 한 번씩
    넣어도(경로 문자열은 다르지만 같은 파일) 결과는 동일해야 한다."""
    test_path = "_test_원본_중복.xlsx"
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Sheet1"
    ws["A1"] = "예산"; ws["A2"] = 1850000
    wb.save(test_path)
    try:
        pool_once, _ = read_source_files([test_path], default_year=2026)
        pool_same_twice, _ = read_source_files([test_path, test_path], default_year=2026)
        pool_relative_and_absolute, _ = read_source_files(
            [test_path, os.path.abspath(test_path)], default_year=2026
        )
        assert len(pool_same_twice) == len(pool_once), (len(pool_same_twice), len(pool_once))
        assert len(pool_relative_and_absolute) == len(pool_once), (
            len(pool_relative_and_absolute), len(pool_once))
        print("read_source_files_dedup_duplicate_path 통과:", len(pool_once), "건 (중복 제거됨, 배로 부풀지 않음)")
    finally:
        os.remove(test_path)


if __name__ == "__main__":
    _selftest_read_excel_source()
    _selftest_read_excel_source_attaches_row_label()
    _selftest_read_excel_source_date_cell()
    _selftest_read_excel_source_corrupted_file_no_crash()
    _selftest_read_excel_source_title_row_before_real_header()
    _selftest_read_hwp_source()
    _selftest_read_hwp_source_no_match()
    _selftest_read_hwp_source_missing_file_no_crash()
    _selftest_read_hwp_source_isolated_basic()
    _selftest_read_hwp_source_isolated_missing_file_no_crash()
    _selftest_read_hwp_source_isolated_does_not_break_live_hwp_instance()
    _selftest_read_pdf_source_missing_file()
    _selftest_read_pdf_source_multipage_skips_blank()
    _selftest_read_pdf_source_simple_table()
    _selftest_read_source_folder_conflict()
    _selftest_read_source_folder_multi_type_cell_no_cross_type_conflict()
    _selftest_read_source_files_mixed_list()
    _selftest_read_source_folder_still_works()
    _selftest_read_source_files_folder_listdir_permission_error_skipped()
    _selftest_read_source_files_dedup_duplicate_path()
