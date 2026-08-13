# -*- coding: utf-8 -*-
"""
한글(HWP) 스크립트 매크로(F10)용 브릿지.
- 한글 안의 매크로(자바스크립트)가 이 스크립트를 커맨드라인으로 호출한다.
- 입력/출력 모두 UTF-16(유니코드) 텍스트 파일로 주고받는다.
  (한글 매크로의 FileSystemObject가 CreateTextFile(path, true, true)/
   OpenTextFile(path, 1, false, -1)로 UTF-16을 안정적으로 읽고 쓸 수 있어,
   클립보드나 커맨드라인 인자보다 한글 인코딩이 깨질 위험이 없다.)

사용법: python macro_bridge.py <mode> <in_path> <out_path>
  mode: gaejoshik | meeting | law | review
"""
import sys

import ai_writer
import law_search
import format_checker


def run_gaejoshik(text: str) -> str:
    return ai_writer.to_gaejoshik(text)


def run_meeting(text: str) -> str:
    return ai_writer.summarize_meeting(text)


def run_law(text: str) -> str:
    results = law_search.search_law(text, display=5)
    if not results:
        return "(관련 법령을 찾지 못했습니다)"
    return "\n".join(f"- {r['법령명']} ({r['공포일자']}, {r['소관부처']})" for r in results)


def run_review(text: str) -> str:
    fmt = format_checker.run_all_checks(text)
    spell = ai_writer.check_spelling_llm(text)
    lines = [
        f"항목번호 오류: {fmt['numbering_issues']}",
        f"계산 오류: {fmt['arithmetic_issues']}",
        f"맞춤법 검토(참고용):\n{spell}",
    ]
    return "\n".join(lines)


MODES = {
    "gaejoshik": run_gaejoshik,
    "meeting": run_meeting,
    "law": run_law,
    "review": run_review,
}


def main():
    if len(sys.argv) != 4:
        print("사용법: python macro_bridge.py <mode> <in_path> <out_path>", file=sys.stderr)
        sys.exit(1)

    mode, in_path, out_path = sys.argv[1], sys.argv[2], sys.argv[3]
    fn = MODES.get(mode)
    if fn is None:
        _write_result(out_path, f"[오류] 알 수 없는 모드: {mode}")
        sys.exit(1)

    with open(in_path, "r", encoding="utf-16") as f:
        text = f.read()

    try:
        result = fn(text)
    except Exception as e:
        result = f"[오류] {type(e).__name__}: {e}"

    _write_result(out_path, result)


def _write_result(out_path: str, text: str):
    with open(out_path, "w", encoding="utf-16") as f:
        f.write(text)


if __name__ == "__main__":
    main()
