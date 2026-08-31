"""source_reader.py — 엑셀/한글/PDF 원본데이터를 읽어 verify_numbers의 정답 풀 형식으로 변환"""
import os
import re
import datetime
from decimal import Decimal

import openpyxl
from verify_numbers import extract_values, compute_column_sums, _decimal_to_normalized_str


def read_excel_source(path: str, default_year: int) -> list[dict]:
    """엑셀 파일의 모든 시트, 모든 셀 값을 정답 풀 항목으로 변환한다.
    셀 값이 순수 숫자면 amount로, 날짜 타입이면 date로, 문자열이면 extract_values로
    4종을 뽑는다. 추가로 각 시트의 숫자 컬럼(1행이 헤더라고 가정) 합계도 파생값으로
    포함한다. 손상되었거나 읽을 수 없는 파일은 예외 없이 빈 리스트를 반환한다
    (PRD 12-5, read_pdf_source와 동일한 관례).

    알려진 한계(문서화만 하고 이번엔 해결하지 않음): 수식 셀은 data_only=True로
    열어도 실제 Excel에서 한 번도 저장된 적 없으면 캐시된 값이 없어 None으로
    읽힐 수 있다(합계 등 파생값이 원본에 있어도 못 읽는 경우 발생 가능). 병합된
    제목 행이 실제 헤더보다 위에 있는 레이아웃도 1행=헤더 가정과 어긋날 수 있다.
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
        header = None
        for row in ws.iter_rows(values_only=False):
            if header is None:
                header = [c.value for c in row]
                continue
            row_dict = {}
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
                    })
                elif isinstance(value, (int, float)) and not isinstance(value, bool):
                    # (2026-08-30 Task 5/6 검토에서 미리 반영) bool은 int의 서브클래스라
                    # 별도 제외 필요. 정규화도 float 대신 Decimal 기반 공용 헬퍼를 써서
                    # 부동소수점 오차·과학적 표기법 문제를 처음부터 피한다.
                    results.append({
                        "type": "amount", "normalized": _decimal_to_normalized_str(Decimal(str(value))),
                        "raw": str(value), "source_file": path, "location": f"{ws.title}!{cell.coordinate}",
                    })
                    row_dict[col_name] = value
                elif isinstance(value, str):
                    for v in extract_values(value, default_year):
                        v["source_file"] = path
                        v["location"] = f"{ws.title}!{cell.coordinate}"
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


def read_hwp_source(path: str, default_year: int) -> list[dict]:
    """한글 문서를 안 보이게(visible=False) 열어서 전체 텍스트를 읽고,
    4종 값을 뽑아 정답 풀 항목으로 변환한다. 원본 참고용으로만 열기 때문에
    화면에 띄우지 않는다 (사용자가 실제로 편집 중인 보고서와는 별개의 인스턴스).
    손상되었거나 열 수 없는 파일은 예외 없이 빈 리스트를 반환한다 (PRD 12-5,
    read_excel_source/read_pdf_source와 동일한 관례). hwp.quit()은 어떤 예외가
    나든 반드시 호출되도록 try/finally로 감싼다 — 그렇지 않으면 보이지 않는
    Hwp.exe 프로세스가 누적되어 남는 문제가 실제로 재현된 적 있다.
    """
    from pyhwpx import Hwp
    hwp = None
    try:
        hwp = Hwp(visible=False)
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
    hwp = Hwp(visible=False)
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
    hwp = Hwp(visible=False)
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


_READERS = {".xlsx": read_excel_source, ".xls": read_excel_source,
            ".hwp": read_hwp_source, ".hwpx": read_hwp_source,
            ".pdf": read_pdf_source}


_CELL_ADDRESS_PATTERN = re.compile(r'^[A-Za-z]{1,3}\d+$')


def read_source_folder(folder_path: str, default_year: int) -> tuple[list[dict], list[dict]]:
    """폴더 안 모든 파일을 읽되, 처리 가능한 확장자만 실제로 읽고 나머지는 조용히
    건너뛴다. 서로 다른 엑셀 파일의 "같은 시트!같은 셀 주소"에 다른 값이 있으면
    충돌로 간주해 별도 리스트로 반환한다 (PRD 12-5).

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
    """
    pool = []
    for name in os.listdir(folder_path):
        ext = os.path.splitext(name)[1].lower()
        reader = _READERS.get(ext)
        if reader is None:
            continue  # 이미지, 워드, 알 수 없는 형식 등 → 조용히 건너뜀
        full_path = os.path.join(folder_path, name)
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


if __name__ == "__main__":
    _selftest_read_excel_source()
    _selftest_read_excel_source_date_cell()
    _selftest_read_excel_source_corrupted_file_no_crash()
    _selftest_read_hwp_source()
    _selftest_read_hwp_source_no_match()
    _selftest_read_hwp_source_missing_file_no_crash()
    _selftest_read_pdf_source_missing_file()
    _selftest_read_pdf_source_multipage_skips_blank()
    _selftest_read_pdf_source_simple_table()
    _selftest_read_source_folder_conflict()
    _selftest_read_source_folder_multi_type_cell_no_cross_type_conflict()
