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

    반환 dict 계약 (Task 13~15 등 이 함수를 직접 호출하는 후속 코드가 알아야 함):
      - mismatch_count: int — 서로 다른(중복 제거된) 불일치 값의 개수. summary
        첫 줄 숫자와는 항상 일치하지만, 그 아래 나열된 줄 수와는 conflicts가
        비어있을 때만 일치한다 — conflicts가 있으면 그 충돌 항목들이 추가 줄로
        더 붙어서(아래 conflicts 참고) 줄 수가 이 숫자보다 많아질 수 있다.
      - summary: str — 채팅창에 그대로 보여줄 사람이 읽는 요약 텍스트.
      - conflicts: list — read_source_folder가 찾은 원본 파일 간 불일치 목록
        (단일 파일/폴더가 아닌 경우 항상 빈 리스트).
      - _report_handle: HwpReport | None — 열린 보고서 문서 핸들. 호출자가 문서를
        계속 화면에 띄워두려면 이 핸들을 들고 있다가 다 쓴 뒤 .close(save=False)로
        닫아야 한다(자동저장 없음). 문서를 아예 열지 않은 이른 반환 경로
        (지원하지 않는 원본 형식, 존재하지 않는 report_path)에서는 None이다 —
        이 두 이른 반환 경로도 이 함수 계약의 정식 일부이며, 코드를 다시 읽지
        않고도 알 수 있어야 한다.
    """
    if not os.path.exists(report_path):
        # (2026-08-31 코드품질 리뷰 반영) HwpReport 생성 전에 저렴하게 걸러낸다 —
        # 존재하지 않는 report_path는 비프로그래머 사용자의 가장 흔한 실수(오타)이고,
        # 이 체크는 Hwp 프로세스를 아예 띄우지 않고 빠르게 끝난다. 이 체크를 통과하지
        # 못하는 나머지 경우(파일은 있으나 손상되어 못 여는 등)는 아래 HwpReport(...)
        # 생성이 예외를 던지도록 그대로 둔다 — 여기서 광범위한 예외처리로 감싸지 않는다.
        return {"mismatch_count": 0, "summary": f"보고서 파일을 찾을 수 없습니다: {report_path}",
                "conflicts": [], "_report_handle": None}

    if os.path.isdir(source_path):
        answer_pool, conflicts = read_source_folder(source_path, default_year)
    elif (ext := os.path.splitext(source_path)[1].lower()) in (".xlsx", ".xls"):
        answer_pool, conflicts = read_excel_source(source_path, default_year), []
    elif ext in (".hwp", ".hwpx"):
        answer_pool, conflicts = read_hwp_source(source_path, default_year), []
    elif ext == ".pdf":
        answer_pool, conflicts = read_pdf_source(source_path, default_year), []
    else:
        return {"mismatch_count": 0, "summary": f"지원하지 않는 원본 형식: {source_path}",
                "conflicts": [], "_report_handle": None}

    report = HwpReport(report_path)
    report_text = report.get_text()
    report_values = extract_values(report_text, default_year)
    mismatches = compare_values(report_values, answer_pool)

    # (2026-08-31 코드품질 리뷰 반영) 같은 잘못된 값이 여러 곳(본문+요약표 등)에
    # 등장하면 compare_values가 raw는 같고 span만 다른 mismatch dict를 여러 개
    # 돌려줄 수 있다. mark_red() 자체가 이미 문서 전체를 훑어 모든 occurrence를
    # 한 번에 칠하므로, 같은 raw로 두 번 부르면 문서 전체를 다시 훑는 중복 COM
    # 순회일 뿐이다 — dict.fromkeys로 raw를 중복 제거한 뒤 한 번씩만 호출한다.
    unique_raw_values = list(dict.fromkeys(m["raw"] for m in mismatches))
    for raw in unique_raw_values:
        report.mark_red(raw)

    # 요약에 표시되는 "확인 필요" 건수·목록도 같은 이유로 raw 기준 중복 제거한다 —
    # 같은 잘못된 값이 목록에 두 번 나오면 "서로 다른 문제 2건"처럼 보여 혼란만
    # 준다. mismatch_count와 summary 첫 줄 숫자를 이 중복 제거된 개수로 통일한다.
    # 단, conflicts 항목은 이 숫자에 포함되지 않고 목록 뒤에 추가로 붙으므로
    # (아래 참고), conflicts가 있으면 첫 줄 숫자보다 실제 나열되는 줄 수가 더
    # 많아진다 — 이건 의도된 것이다(원본 파일 간 불일치는 "불일치 값 개수"와
    # 성격이 달라 같은 숫자에 합산하지 않기로 함).
    first_type_by_raw = {}
    for m in mismatches:
        first_type_by_raw.setdefault(m["raw"], m["type"])
    lines = [f"- [{first_type_by_raw[raw]}] '{raw}' 원본에서 확인 안 됨" for raw in unique_raw_values]
    for c in conflicts:
        value_desc = ", ".join(f"{v['file']}={v['normalized']}" for v in c["values"])
        # (2026-08-31) 같은 location에 타입이 다른 충돌이 각각 별도 항목으로 올 수 있으므로
        # (Task 10 참고) type을 같이 표시해야 두 줄이 나와도 사용자가 구분할 수 있다.
        lines.append(f"- ⚠ 원본자료 불일치[{c['type']}]: {c['location']} ({value_desc})")

    summary = (f"{len(unique_raw_values)}건 확인 필요\n" + "\n".join(lines)
               if mismatches or conflicts else "이상 없음, 모두 원본과 일치합니다")

    return {"mismatch_count": len(unique_raw_values), "summary": summary, "conflicts": conflicts,
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


def _selftest_missing_report_path():
    """(2026-08-31 코드품질 리뷰 반영 회귀테스트) report_path가 존재하지 않을 때
    거짓 "이상 없음"이 아니라 "찾을 수 없습니다" 안내가 반환되는지 확인한다.
    os.path.exists() 가드가 HwpReport 생성보다 먼저 반환해야 하므로, 이 테스트는
    Hwp.exe 프로세스를 전혀 띄우지 않아야 한다 — tasklist로 실행 전후 Hwp.exe
    프로세스 수가 그대로인지까지 확인해 그 사실을 직접 검증한다.
    """
    import subprocess

    def count_hwp_processes():
        out = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq Hwp.exe"],
            capture_output=True, text=True,
        ).stdout
        return out.count("Hwp.exe")

    before = count_hwp_processes()
    result = run_verification(
        report_path="_이런파일없음.hwp", source_path="_이런폴더도없음", default_year=2026,
    )
    after = count_hwp_processes()

    assert result["mismatch_count"] == 0, result
    assert result["_report_handle"] is None, result
    assert "찾을 수 없습니다" in result["summary"], result
    assert after == before, f"Hwp.exe 프로세스가 새로 뜸 (before={before}, after={after})"
    print("run_verification(존재하지 않는 report_path) 통과:", result["summary"],
          f"(Hwp.exe 프로세스 변화 없음: {before} -> {after})")


if __name__ == "__main__":
    _selftest_missing_report_path()
    _selftest_run_verification()
