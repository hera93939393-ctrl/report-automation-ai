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
