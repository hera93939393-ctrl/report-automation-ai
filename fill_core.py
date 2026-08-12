# -*- coding: utf-8 -*-
"""
채우기 엔진 (핵심 로직)
- 서식 뼈대(계획안_서식.hwp)를 열어 토큰을 실제 값으로 바꾸고 새 파일로 저장한다.
- 서식/도형/표 모양은 그대로 유지된다. (토큰 글자만 교체)
GUI 등에서 fill_report(values, out_path) 를 호출해서 사용.
"""
import win32com.client as win32
import os

TEMPLATE = r"D:\보고서자동화\계획안_서식.hwp"

# 토큰 목록 (GUI가 참고)
FIELDS = [
    ("제목",   "제목 (끝의 ' 계획(안)'은 자동으로 붙어있음)"),
    ("목적",   "추진목적"),
    ("일자",   "회의 일자 (예: 9월 7일(수))"),
    ("시간",   "회의 시간 (예: 14:00~16:00)"),
    ("소요",   "소요시간 (예: 2시간)"),
    ("방법",   "회의 방법"),
    ("대상",   "대상"),
    ("내용1",  "내용 항목 1"),
    ("내용2",  "내용 항목 2"),
    ("내용3",  "내용 항목 3"),
    ("내용4",  "내용 항목 4"),
    ("일정설명", "세부일정 한 줄 설명"),
]


def fill_report(values: dict, out_path: str):
    """values: {'제목': '...', '목적': '...', ...}  ->  out_path 에 저장"""
    if not os.path.exists(TEMPLATE):
        raise FileNotFoundError(f"템플릿이 없습니다: {TEMPLATE} (먼저 make_template.py 실행)")

    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
    hwp.XHwpWindows.Item(0).Visible = False
    try:
        hwp.Open(TEMPLATE, "HWP", "forceopen:true")

        def replace_all(find, repl):
            hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
            o = hwp.HParameterSet.HFindReplace
            o.FindString = find
            o.ReplaceString = repl if repl is not None else ""
            o.IgnoreMessage = 1
            o.ReplaceMode = 1
            o.Direction = 0
            return hwp.HAction.Execute("AllReplace", o.HSet)

        for key, _label in FIELDS:
            token = "{{%s}}" % key
            val = values.get(key, "")
            replace_all(token, val)

        hwp.SaveAs(out_path, "HWP", "")
    finally:
        hwp.Quit()
    return out_path


if __name__ == "__main__":
    # 테스트용 샘플 값
    sample = {
        "제목": "학교급식 위생점검 하반기 추진",
        "목적": "하반기 위생점검 계획 공유 및 현장 의견 수렴",
        "일자": "7월 22일(화)",
        "시간": "10:00~11:30",
        "소요": "1시간 30분",
        "방법": "대면회의 (본관 3층 대회의실)",
        "대상": "각 학교 영양사, 위생점검단, 교육청 담당자",
        "내용1": "상반기 위생점검 결과 총평 및 우수사례 공유",
        "내용2": "하반기 중점 점검 항목 안내",
        "내용3": "위생점검 체크리스트 개정사항 설명",
        "내용4": "현장 애로사항 청취 및 질의응답",
        "일정설명": "(1부) 결과 공유 및 계획 안내 / (2부) 질의응답 및 토론",
    }
    out = r"D:\보고서자동화\_테스트결과.hwp"
    fill_report(sample, out)
    print("saved:", out)
