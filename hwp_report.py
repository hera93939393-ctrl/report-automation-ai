"""hwp_report.py — 채팅 도구가 직접 여는 한글 보고서 문서를 다루는 얇은 pyhwpx 래퍼"""
import os

from pyhwpx import Hwp


class HwpReport:
    """채팅 도구가 pyhwpx로 직접 여는 보고서 문서. 사용자가 한글 아이콘으로
    따로 열지 않게 해서, F9에서 확인된 COM 재연결 제약(PRD 12-3)을 피한다.
    """

    def __init__(self, path: str):
        # new=True로 반드시 새 한글 프로세스를 띄운다. pyhwpx의 Hwp()는 기본값
        # (new=False)에서 이미 떠 있는 한/글 프로세스가 있으면 그 프로세스에
        # 그대로 접속(재사용)해버리는데, 이는 클래스 docstring이 밝힌 설계
        # 의도("COM 재연결 제약을 피한다")와 정반대로 동작한다. 특히 사용자가
        # 따로 작업 중이던 한글 창을 자동화가 그대로 붙잡아버릴 위험이 있어
        # new=True가 맞다.
        #
        # 별개로: 이 폴더(worktree 경로) 안의 파일을 hwp.open()으로 열 때,
        # 한글이 자체 보안 모듈 경고("...접근하려는 시도(파일의 손상 또는
        # 유출의 위험 등)가 있습니다")를 띄우며 무한 대기하는 현상이 실제로
        # 재현되었다. new=True 여부와 무관하게 나타났고(즉 인스턴스 재사용
        # 문제가 원인이 아님), "접근 허용"을 사람이 눌러주면 즉시 정상
        # 진행되었다 — visible=True로 사용자가 직접 보게 만든 이 설계와
        # 일치하는 동작이다. 코드로 완전히 억제하는 방법은 찾지 못했으므로
        # (이 프로젝트 히스토리에 이미 기록된 COM 자동화 환경 불안정성의
        # 연장선으로 보임), 무인 실행 시 이 대화상자가 뜨면 사람이 "접근
        # 허용"을 눌러줘야 진행된다는 점을 알아두어야 한다.
        self.hwp = Hwp(visible=True, new=True)  # 사용자가 직접 봐야 하므로 visible=True
        self.hwp.open(path)
        self.path = path

    def get_text(self) -> str:
        return self.hwp.GetTextFile("TEXT", "")

    def mark_red(self, target_text: str) -> bool:
        """문서 안에서 target_text를 찾아 글자색을 빨간색(255,0,0)으로 바꾼다.
        찾지 못하면 False를 반환한다. 자동저장은 하지 않는다.
        """
        found = self.hwp.find(target_text, direction="AllDoc")
        if not found:
            return False
        self.hwp.set_font(TextColor=self.hwp.RGBColor(255, 0, 0))
        return True

    def get_char_color_at(self, target_text: str):
        """target_text 위치의 현재 글자색을 (R,G,B) 튜플로 반환한다 (테스트 검증용).

        pyhwpx의 find()는 찾은 뒤 커서를 그 텍스트의 끝으로 이동시키므로,
        direction="Forward"(기본값)로 다시 찾으면 이미 지나친 유일한 occurrence를
        다시 찾지 못한다. AllDoc으로 찾아야 문서 처음부터 다시 감싸 탐색해 확실히 찾는다.

        중요: mark_red()가 set_font()로 글자색을 바꾼 직후에는 그때 찾았던 블록이
        여전히 선택된 상태로 남아있다. 그 선택이 남아있는 채로 바로 다시 find()를
        호출하면, RepeatFind가 이미 선택되어 있는 occurrence를 기준으로 헷갈려서
        (실제 문서 색은 바뀌었음에도) 색을 읽어오는 결과가 이전 색(검정)으로
        보이는 현상이 재현 확인됨. Cancel()로 선택을 먼저 해제해 커서를 빈
        상태로 만든 뒤 find()를 호출해야 새로 칠해진 색을 안정적으로 읽는다.
        """
        self.hwp.Cancel()
        self.hwp.find(target_text, direction="AllDoc")
        color_value = self.hwp.CharShape.Item("TextColor")
        return (color_value & 0xFF, (color_value >> 8) & 0xFF, (color_value >> 16) & 0xFF)

    def close(self, save: bool):
        self.hwp.quit(save=save)


def _selftest_open_and_mark_red():
    from pyhwpx import Hwp
    test_path = os.path.abspath("_test_보고서.hwp")
    setup = Hwp(visible=False, new=True)
    setup.insert_text("예산은 9999999원이며 정상입니다")
    setup.save_as(test_path)
    setup.quit()
    try:
        report = HwpReport(test_path)
        text = report.get_text()
        assert "9999999" in text, text
        report.mark_red("9999999")
        color_at_match = report.get_char_color_at("9999999")
        assert color_at_match == (255, 0, 0), color_at_match
        report.close(save=False)
        print("HwpReport 통과: 텍스트 읽기 + 빨간색 표시 확인")
    finally:
        os.remove(test_path)


if __name__ == "__main__":
    _selftest_open_and_mark_red()
