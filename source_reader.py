"""source_reader.py — 엑셀/한글/PDF 원본데이터를 읽어 verify_numbers의 정답 풀 형식으로 변환"""
import os
from decimal import Decimal

import openpyxl
from verify_numbers import extract_values, compute_column_sums, _decimal_to_normalized_str


def read_excel_source(path: str, default_year: int) -> list[dict]:
    """엑셀 파일의 모든 시트, 모든 셀 값을 정답 풀 항목으로 변환한다.
    셀 값이 순수 숫자면 amount로, 문자열이면 extract_values로 4종을 뽑는다.
    추가로 각 시트의 숫자 컬럼(1행이 헤더라고 가정) 합계도 파생값으로 포함한다.
    """
    wb = openpyxl.load_workbook(path, data_only=True)
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
                if isinstance(value, (int, float)) and not isinstance(value, bool):
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


if __name__ == "__main__":
    _selftest_read_excel_source()
