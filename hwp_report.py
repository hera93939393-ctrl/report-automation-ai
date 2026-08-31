"""hwp_report.py — 채팅 도구가 직접 여는 한글 보고서 문서를 다루는 얇은 pyhwpx 래퍼"""
import os

from pyhwpx import Hwp


class HwpReport:
    """채팅 도구가 pyhwpx로 직접 여는 보고서 문서. 사용자가 한글 아이콘으로
    따로 열지 않게 해서, F9에서 확인된 COM 재연결 제약(PRD 12-3)을 피한다.

    참고: 이 클래스는 의도적으로 메서드 호출을 try/except로 감싸지 않는다.
    source_reader.py의 read_hwp_source()와 달리, 여기서는 오류가 조용히
    삼켜지지 않고 화면을 보고 있는 사람에게 그대로 드러나야 한다
    (visible=True로 사람이 직접 지켜보는 도구라는 설계에 맞는 선택).
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
        # 주의: 위 Hwp() 생성이 성공한 뒤 아래 open(path)이 예외를 던지면, 방금 띄운
        # 한글 프로세스를 가리키는 핸들이 호출자에게 없어 코드로 닫을 수 없다 —
        # visible=True로 사람이 지켜보는 도구이므로, 그 경우 사람이 직접 창을
        # 닫아줘야 하는 것을 받아들인 트레이드오프다(지금 단계에서 고칠 문제 아님).
        self.hwp.open(path)
        self.path = path

    def get_text(self) -> str:
        return self.hwp.GetTextFile("TEXT", "")

    def mark_red(self, target_text: str) -> bool:
        """문서 안에서 target_text의 모든 occurrence를 찾아 글자색을 빨간색
        (255,0,0)으로 바꾼다. 같은 잘못된 값이 표와 요약 문장 등 여러 곳에
        중복 등장하는 경우가 흔해서, 한 곳만 바꾸면 나머지가 안 바뀐 채로
        남는다 — 그래서 문서 전체를 훑어 전부 바꾼다.
        문서 어디에도 없으면 False를 반환한다. 자동저장은 하지 않는다.

        구현 메모: 커서를 문서 처음으로 옮긴 뒤 direction="Forward"로 반복
        탐색한다. "AllDoc"을 루프 안에서 쓰면 문서 끝에 닿았을 때 처음으로
        다시 감싸 돌아가 버려(pyhwpx core.py의 find() 문서 참고) 같은
        occurrence를 무한히 다시 찾아 무한 루프가 된다. "Forward"는 문서
        끝에서 자연히 멈추므로 이 루프에 맞다. pyhwpx 자체도 동일한 패턴을
        set_field_by_bracket()에서 쓴다(MoveDocBegin() 후 while self.find(...)).
        """
        self.hwp.MoveDocBegin()
        found_any = False
        while self.hwp.find(target_text, direction="Forward"):
            self.hwp.set_font(TextColor=self.hwp.RGBColor(255, 0, 0))
            found_any = True
        return found_any

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

    def get_char_colors(self, target_text: str) -> list:
        """target_text의 모든 occurrence 위치의 글자색을 문서 순서대로
        (R,G,B) 튜플 리스트로 반환한다 (테스트 검증용 — mark_red()가 정말
        모든 occurrence를 바꿨는지 하나씩 확인하기 위함).

        mark_red()와 동일하게 커서를 문서 처음으로 옮긴 뒤 direction="Forward"로
        반복 탐색한다. get_char_color_at()의 Cancel() 관련 주의사항은 "색을
        바꾼 직후 바로 같은 위치를 다시 찾을 때"에 해당하는 것이고, 여기서는
        찾은 뒤 쓰기(set_font) 없이 읽기만 하고 다음 occurrence로 넘어가므로
        그 문제가 재현되지 않는다 — 다만 호출 시작 시점에 이전 호출이 남긴
        선택 상태가 있을 수 있으니 시작 전에 한 번 Cancel()로 정리한다.
        """
        self.hwp.Cancel()
        self.hwp.MoveDocBegin()
        colors = []
        while self.hwp.find(target_text, direction="Forward"):
            color_value = self.hwp.CharShape.Item("TextColor")
            colors.append((color_value & 0xFF, (color_value >> 8) & 0xFF, (color_value >> 16) & 0xFF))
        return colors

    def close(self, save: bool):
        """한글 문서를 닫는다.

        save=True이면 pyhwpx의 quit(save=True)를 호출하는데, 이는 내부적으로
        save()를 실행해 self.path의 원본 파일을 그 자리에서 덮어쓴다 —
        "다른 이름으로 저장"이 아니라 원본을 직접 덮어쓰며, 백업을 남기지
        않는다. 이 모듈 전체의 핵심 설계가 "자동저장 없음"이므로, save=True를
        넘기는 유일한 지점인 이 메서드의 동작은 분명히 해둘 필요가 있다.
        save=False이면 변경사항을 저장하지 않고 닫는다(자동저장 없음, 기본 사용 경로).
        """
        self.hwp.quit(save=save)


def _selftest_open_and_mark_red():
    import tempfile
    from pyhwpx import Hwp
    # (2026-08-31) 이 워크트리 폴더 안에 테스트 파일을 만들면 한글이 자체 보안
    # 모듈 경고("...접근하려는 시도...")를 띄우며 무한 대기하는 현상이 간헐적으로
    # 재현됐다(직접 확인: 워크트리 밖 임시폴더에서는 재현 안 됨) — 이 폴더
    # 경로 자체가 원인으로 보여, 무인 테스트 실행이 이 문제를 안 겪도록 시스템
    # 임시폴더에 테스트 파일을 만든다.
    test_path = os.path.join(tempfile.gettempdir(), "_test_보고서.hwp")
    setup = Hwp(visible=False, new=True)
    setup.insert_text("예산은 9999999원이며, 표에도 9999999원이 있습니다")
    setup.save_as(test_path)
    setup.quit()
    report = None
    try:
        try:
            report = HwpReport(test_path)
            text = report.get_text()
            assert "9999999" in text, text
            found = report.mark_red("9999999")
            assert found is True, found
            colors = report.get_char_colors("9999999")
            assert len(colors) == 2, colors
            assert all(c == (255, 0, 0) for c in colors), colors
            print("HwpReport 통과: 텍스트 읽기 + 모든 occurrence 빨간색 표시 확인")
        finally:
            if report is not None:
                report.close(save=False)
    finally:
        os.remove(test_path)


if __name__ == "__main__":
    _selftest_open_and_mark_red()
