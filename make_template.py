# -*- coding: utf-8 -*-
"""
서식 뼈대(템플릿) 생성기
- 원본 계획(안) 파일을 열어, 매번 바뀌는 내용 부분을 짧은 토큰({{...}})으로 치환한다.
- 서식/도형/표 모양은 절대 건드리지 않는다. (텍스트만 토큰으로 교체)
- 결과: 계획안_서식.hwp  (앞으로 이 파일을 재사용)
"""
import win32com.client as win32
import os, io

SRC = r"C:\Users\lakka\OneDrive\바탕 화면\예시 계획(안).hwp"   # 원본 (읽기 전용으로만 사용)
OUT = r"D:\보고서자동화\계획안_서식.hwp"                        # 생성할 템플릿
LOG = r"D:\보고서자동화\_make_template_log.txt"

log = io.open(LOG, "w", encoding="utf-8")
def L(*a): log.write(" ".join(str(x) for x in a) + "\n")

# (찾을 원문, 바꿀 토큰) — 긴 문장은 숨은 제어문자로 실패할 수 있어 결과를 로그로 확인
REPLACEMENTS = [
    ("공공급식 공급업체 전수점검 중간 간담회 개최", "{{제목}}"),
    ("상반기 전수점검 결과 공유 및 전수점검 운영에 대한 의견 수렴", "{{목적}}"),
    ("9월 7일(수)", "{{일자}}"),
    ("14:00~16:00", "{{시간}}"),
    ("(2시간)", "({{소요}})"),
    ("화상회의 프로그램(ZOOM)을 활용한 온라인 회의", "{{방법}}"),
    ("지역본부 담당자, 급식관리단, 학부모점검단 및 유관기관 담당자", "{{대상}}"),
    ("상반기 전수점검 결과 공유 및 하반기 추진 계획 안내", "{{내용1}}"),
    ("지역별 점검 특이사례 및 급식관리단·학부모점검단 건의사항 공유", "{{내용2}}"),
    ("급식관리단 대상 전수점검 모바일 시스템 사용 방법 안내 및 시연", "{{내용3}}"),
    ("학부모점검단 대상 제로웨이스트 관련 강연 및 체험활동 실시", "{{내용4}}"),
    ("(1부) 전수점검 결과 및 사례 공유 / (2부) 제로웨이스트 관련 강연 및 활동", "{{일정설명}}"),
]

hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
hwp.XHwpWindows.Item(0).Visible = False
hwp.Open(SRC, "HWP", "forceopen:true")

def replace_all(find, repl):
    hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
    o = hwp.HParameterSet.HFindReplace
    o.FindString = find
    o.ReplaceString = repl
    o.IgnoreMessage = 1
    o.ReplaceMode = 1
    o.Direction = 0
    return hwp.HAction.Execute("AllReplace", o.HSet)

fail = []
for find, tok in REPLACEMENTS:
    r = replace_all(find, tok)
    L(("OK " if r else "FAIL "), tok, "<-", repr(find))
    if not r:
        fail.append((find, tok))

hwp.SaveAs(OUT, "HWP", "")
txt = hwp.GetTextFile("TEXT", "")
present = [tok for _, tok in REPLACEMENTS if tok in txt]
L("\nTOKENS PRESENT IN TEMPLATE:", present)
L("FAILED:", [t for _, t in fail])
hwp.Quit()
log.close()
print("done; failed =", len(fail))
