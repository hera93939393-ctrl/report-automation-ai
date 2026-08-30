"""source_reader.py — 엑셀/한글/PDF 원본데이터를 읽어 verify_numbers의 정답 풀 형식으로 변환"""
import os
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
