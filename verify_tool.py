"""verify_tool.py — Task 1~11의 모듈을 엮어 "숫자검증" 도구 하나로 만든다.
chat_assistant.py가 Ollama 도구호출로 실행할 최종 진입점."""
import os
import tempfile
import openpyxl

from verify_numbers import extract_values, compare_values
from source_reader import read_excel_source, read_hwp_source, read_pdf_source, read_source_folder
from hwp_report import HwpReport


def run_verification(report_path: str, source_path: str, default_year: int) -> dict:
    """채팅창이 호출하는 최종 도구 함수.
    1) report_path를 pyhwpx로 열어 화면에 띄운다 (사용자가 보는 그 창)
    2) source_path(파일 또는 폴더)에서 정답 풀을 만든다
    3) 보고서 텍스트에서 4종 값을 뽑아 대조한다
    4) 불일치 항목을 빨간색으로 표시한다 (자동저장 안 함)
    5) 채팅창에 보여줄 요약 텍스트와 충돌 목록을 반환한다
    """
    ext = os.path.splitext(source_path)[1].lower()
    if os.path.isdir(source_path):
        answer_pool, conflicts = read_source_folder(source_path, default_year)
    elif ext in (".xlsx", ".xls"):
        answer_pool, conflicts = read_excel_source(source_path, default_year), []
    elif ext in (".hwp", ".hwpx"):
        answer_pool, conflicts = read_hwp_source(source_path, default_year), []
    elif ext == ".pdf":
        answer_pool, conflicts = read_pdf_source(source_path, default_year), []
    else:
        return {"mismatch_count": 0, "summary": f"지원하지 않는 원본 형식: {source_path}", "conflicts": []}

    report = HwpReport(report_path)
    report_text = report.get_text()
    report_values = extract_values(report_text, default_year)
    mismatches = compare_values(report_values, answer_pool)

    for mismatch in mismatches:
        report.mark_red(mismatch["raw"])

    lines = [f"- [{m['type']}] '{m['raw']}' 원본에서 확인 안 됨" for m in mismatches]
    for c in conflicts:
        value_desc = ", ".join(f"{v['file']}={v['normalized']}" for v in c["values"])
        # (2026-08-31) 같은 location에 타입이 다른 충돌이 각각 별도 항목으로 올 수 있으므로
        # (Task 10 참고) type을 같이 표시해야 두 줄이 나와도 사용자가 구분할 수 있다.
        lines.append(f"- ⚠ 원본자료 불일치[{c['type']}]: {c['location']} ({value_desc})")

    summary = (f"{len(mismatches)}건 확인 필요\n" + "\n".join(lines)
               if mismatches or conflicts else "이상 없음, 모두 원본과 일치합니다")

    return {"mismatch_count": len(mismatches), "summary": summary, "conflicts": conflicts,
            "_report_handle": report}  # UI가 문서를 닫지 않고 계속 보여주기 위해 핸들 유지


def _selftest_run_verification():
    # (2026-08-31) hwp_report.py의 _selftest_open_and_mark_red()와 동일한 이유로,
    # 이 워크트리 폴더 안에 테스트 파일을 만들면 한글이 자체 보안 모듈 경고
    # ("...접근하려는 시도...")를 띄우며 무한 대기하는 현상이 간헐적으로 재현된 적
    # 있어, 무인 테스트 실행이 이 문제를 겪지 않도록 시스템 임시폴더를 쓴다.
    test_dir = os.path.join(tempfile.gettempdir(), "_test_원본_f11")
    os.makedirs(test_dir, exist_ok=True)
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Sheet1"
    ws["A1"] = "예산"; ws["A2"] = 1850000
    wb.save(os.path.join(test_dir, "원본.xlsx"))

    from pyhwpx import Hwp
    report_path = os.path.join(tempfile.gettempdir(), "_test_보고서_f11.hwp")
    setup = Hwp(visible=False)
    setup.insert_text("예산은 185만원이며, 오타는 9999999원입니다")
    setup.save_as(report_path)
    setup.quit()

    result = None
    try:
        result = run_verification(report_path=report_path, source_path=test_dir, default_year=2026)
        assert result["mismatch_count"] == 1, result
        assert "9999999" in result["summary"], result
        print("run_verification 통과:", result["summary"])
    finally:
        # (2026-08-31) run_verification()은 UI가 문서를 계속 보여줄 수 있도록 의도적으로
        # HwpReport를 닫지 않고 result["_report_handle"]로 넘겨준다 — 주어진 셀프테스트
        # 원안은 이를 닫지 않아 Hwp.exe 프로세스가 열린 채로 남는 문제가 있었으므로,
        # 테스트 스스로 여기서 명시적으로 닫는다(안 그러면 아래 os.remove(report_path)도
        # 파일이 한글에 열려 있어 실패한다).
        if result is not None and result.get("_report_handle") is not None:
            result["_report_handle"].close(save=False)
        import shutil
        shutil.rmtree(test_dir)
        os.remove(report_path)


if __name__ == "__main__":
    _selftest_run_verification()
