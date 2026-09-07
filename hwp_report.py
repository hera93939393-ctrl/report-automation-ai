"""hwp_report.py — 채팅 도구가 직접 여는 한글 보고서 문서를 다루는 얇은 pyhwpx 래퍼"""
import os
import time

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
        # (2026-08-31 코드품질 리뷰 반영) pyhwpx의 Hwp.open()은 성공하면 True, 실패하면
        # False를 반환할 뿐 흔한 실사용자 실수(존재하지 않는 경로, 오타, 손상/형식오류
        # 파일)에 대해 반드시 예외를 던지지는 않는다. 반환값을 무시하면 문서가 실제로는
        # 열리지 않았는데도(빈 화면) get_text()가 빈 문자열을 돌려주고, 이를 바탕으로 한
        # 대조 로직이 "이상 없음"이라는 거짓 결과를 내는 최악의 실패 모드로 이어진다
        # (verify_tool.py run_verification 참고). 이 클래스의 docstring이 명시한 설계
        # 의도("오류가 조용히 삼켜지지 않고 화면을 보고 있는 사람에게 그대로 드러나야
        # 한다")에 맞춰, open() 실패를 명시적으로 감지해 예외로 드러낸다.
        #
        # 방금 띄운 한글 프로세스가 open() 실패로 빈/실패 상태로 남는 문제(위 옛 주석이
        # 지적했던 누수)는 여기서 self.hwp.quit()으로 닫아 정리한다 — quit()이 다시
        # 예외를 던지는 극단적 상황까지 감안해 원래의 open-실패 예외가 항상 사용자에게
        # 보이도록 try/except로 감싼다(quit() 실패로 원래 원인이 가려지지 않게).
        if not self.hwp.open(path):
            try:
                self.hwp.quit()
            except Exception:
                pass
            raise FileNotFoundError(f"한글 문서를 열 수 없습니다: {path}")
        self.path = path
        # 변경내용 추적 on/off 상태 — IsTrackChange 프로퍼티가 이 PC의 한글
        # 버전(빌드 11.0.0.2129)에서 AttributeError가 나서 한글에 직접
        # 물어볼 수 없어(2026-09-05 실측 확인), 파이썬 쪽에서 직접 관리한다.
        self._track_changes_enabled = False

    def get_text(self) -> str:
        # GetTextFile()은 문서에 지금까지 단 한 글자도 삽입된 적 없는(정말로
        # 아무 내용도 없는) 문서에서는 빈 문자열("")이 아니라 파이썬 None을
        # 반환한다(직접 테스트로 확인됨). 이 클래스와 verify_tool.py의
        # extract_values() 등 모든 호출자가 get_text()는 항상 str을 반환한다고
        # 가정하므로, None을 여기서 빈 문자열로 정규화한다.
        text = self.hwp.GetTextFile("TEXT", "")
        return text if text is not None else ""

    def mark_color(self, target_text: str, r: int, g: int, b: int) -> bool:
        """문서 안에서 target_text의 모든 occurrence를 찾아 글자색을 (r,g,b)로
        바꾼다. 같은 값이 표와 요약 문장 등 여러 곳에 중복 등장하는 경우가
        흔해서, 한 곳만 바꾸면 나머지가 안 바뀐 채로 남는다 — 그래서 문서
        전체를 훑어 전부 바꾼다. 문서 어디에도 없으면 False를 반환한다.
        자동저장은 하지 않는다.

        (2026-09-04, 실사용 피드백) 원래 이름은 mark_red()였고 빨간색
        고정이었다 — 사용자가 "확인했지만 정상(파랑)"/"확인은 됐지만
        원본에 대조할 항목이 없음(초록)"도 문서에 표시해달라고 요청해
        색을 파라미터로 받도록 일반화했다. mark_red()는 이제 이 함수를
        RGB(255,0,0)으로 호출하는 얇은 래퍼로 남겨 기존 호출자(self-test
        포함)와의 하위호환을 유지한다.

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
            self.hwp.set_font(TextColor=self.hwp.RGBColor(r, g, b))
            found_any = True
        return found_any

    def mark_next_color(self, target_text: str, r: int, g: int, b: int) -> bool:
        """현재 커서 위치부터 앞으로(Forward)만 target_text의 다음 occurrence
        하나를 찾아 그 글자색을 (r,g,b)로 바꾼다. mark_color()와 다른 점 둘:
        (1) 문서 처음으로 되돌아가지 않고 지금 커서 위치부터 이어서 찾는다,
        (2) 모든 occurrence가 아니라 딱 하나만 바꾼다. 못 찾으면 False.

        (2026-09-04, 실사용 피드백) run_verification()이 값(고유 raw 문자열)
        하나마다 mark_color()로 문서 전체를 처음부터 다시 훑던 것이, 값이
        많을 땐 문서를 수십 번 처음부터 끝까지 스캔하는 것처럼 보여 느리고
        어지럽다는 지적을 받았다 — "위에서부터 하나씩 순서대로 칠하면서
        내려가자"는 사용자 제안을 그대로 구현한 저수준 빌딩블록이다.
        호출자(verify_tool.py)가 문서 위치(span) 순서대로 정렬한 값 목록을
        커서를 한 번만 문서 처음으로 옮긴 뒤 이 메서드로 순서대로 호출하면,
        전체적으로 문서를 위→아래 딱 한 번만 훑으며 색칠하게 된다."""
        found = self.hwp.find(target_text, direction="Forward")
        if found:
            self.hwp.set_font(TextColor=self.hwp.RGBColor(r, g, b))
        return found

    def mark_red(self, target_text: str) -> bool:
        """mark_color(target_text, 255, 0, 0)의 얇은 래퍼 — "오류(원본과 불일치)"
        표시용으로 기존 호출자들이 계속 이 이름을 쓴다."""
        return self.mark_color(target_text, 255, 0, 0)

    def reset_colors(self) -> None:
        """문서 전체의 글자색을 기본값(검정, RGB 0,0,0)으로 초기화한다.

        mark_red()가 남긴 빨간 표시는 아무 것도 자동으로 지워주지 않는다 —
        run_verification()이 같은 문서 핸들에 대해 여러 번(예: 사용자가 원본을
        고치거나 새 원본을 첨부한 뒤 재검증) 반복 호출될 수 있게 되면서
        (F12), 이전 호출에서 남긴 빨간 표시가 이번엔 더 이상 틀리지 않은
        값인데도 그대로 남아있는 문제가 실제로 재현됨(코드품질 검토에서
        확인) — "이상 없음"이라고 채팅창은 답하는데 문서엔 빨간 글자가
        남아있는 모순이 생김. 매 검증 실행 전에 문서 전체를 검정으로
        되돌린 뒤 이번 실행에서 실제로 확인된 값만 다시 빨갛게 표시하면,
        호출할 때마다 항상 "지금 진짜로 틀린 것만" 빨갛게 남는 상태가
        보장된다(멱등성).

        구현 메모: SelectAll() + set_font(TextColor=...) + Cancel() 조합은
        이 프로젝트가 새로 지어낸 패턴이 아니라, pyhwpx 자체가 core.py에서
        문서 전체에 스타일을 적용할 때 쓰는 것과 동일한 패턴이다(예:
        get_used_style_dict(), remove_unused_styles()가 SelectAll() 후
        작업하고 Cancel()로 선택을 해제함). mark_red()가 이미 find()로 찾은
        블록에 set_font(TextColor=...)를 적용하는 것을 검증했으므로, 그
        대상이 find()의 블록이든 SelectAll()의 전체 선택이든 set_font()
        동작 자체는 같다.

        알려진 한계: 사용자가 이 도구와 무관하게 직접 다른 색으로 칠해둔
        글자가 있었다면 그것도 함께 검정으로 초기화된다 — 이 도구의
        범위(숫자검증 표시)를 벗어나는 시나리오라 이번 라운드에서는
        받아들이는 트레이드오프로 문서화만 해둔다.
        """
        self.hwp.SelectAll()
        self.hwp.set_font(TextColor=self.hwp.RGBColor(0, 0, 0))
        self.hwp.Cancel()

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

    def get_window_handle(self) -> int:
        """이 문서를 보여주는 실제 Win32 창의 HWND(창 핸들)를 반환한다
        (창 배치 등 win32gui 함수에 바로 넘겨 쓸 수 있음).

        pyhwpx 자신도 get_title() 등에서 이 프로퍼티(hwp.XHwpWindows.
        Active_XHwpWindow.WindowHandle)로 얻은 핸들을 win32gui에 그대로
        넘겨 쓰고 있음을 소스에서 확인함 — 창 제목으로 FindWindow하는 것보다
        안전하다(여러 한글 창이 동시에 떠 있어도 이 인스턴스가 연 그 창을
        정확히 특정 가능, 제목 문자열은 파일명에 따라 바뀌어 검색 기준으로
        불안정함).
        """
        return self.hwp.XHwpWindows.Active_XHwpWindow.WindowHandle

    def close(self, save: bool):
        """한글 문서를 닫는다.

        save=True이면 pyhwpx의 quit(save=True)를 호출하는데, 이는 내부적으로
        save()를 실행해 self.path의 원본 파일을 그 자리에서 덮어쓴다 —
        "다른 이름으로 저장"이 아니라 원본을 직접 덮어쓰며, 백업을 남기지
        않는다. 이 모듈 전체의 핵심 설계가 "자동저장 없음"이므로, save=True를
        넘기는 유일한 지점인 이 메서드의 동작은 분명히 해둘 필요가 있다.
        save=False이면 변경사항을 저장하지 않고 닫는다(자동저장 없음, 기본 사용 경로).

        (2026-09-06, 실사용/실측 확인) 변경내용 추적을 켠 문서를 닫을 때,
        한글이 "변경 추적 기능이 사용된 문서입니다. ... 선택한 파일
        형식으로 저장할까요?"라는 Yes/No 확인창을 띄운다 — 무인 실행
        환경에서는 이 창에 답할 사람이 없어 그대로 멈춘다(실제로 재현·확인
        됨). pyhwpx가 노출하는 `SetMessageBoxMode`로 이 창이 뜨면 자동으로
        "아니오"(이 형식으로 저장하지 않음)를 선택하게 만들어 무인 실행도
        멈추지 않게 한다 — (2026-09-06 정정) 이 대화상자는 버튼이 "저장(Y)/
        취소(N)"인 **예/아니오형** 팝업이라 pyhwpx 공식 문서(set_message_box_mode
        docstring)의 카테고리 5에 해당하는 `0x20000`(아니오 자동누르기)을
        써야 한다 — 처음엔 카테고리 2(확인/취소형)의 `0x20`을 잘못 썼다가
        실제로 대화상자가 안 막히는 것을 실측으로 확인하고 정정함. 끝나면
        반드시 `0xF0000`(예/아니오 옵션 해제, 기본값)으로 되돌려 이후 다른
        대화상자의 정상 동작(사람이 보는 채팅창 흐름 등)에 영향이 남지
        않게 한다.
        """
        self.hwp.SetMessageBoxMode(0x20000)
        try:
            self.hwp.quit(save=save)
        finally:
            self.hwp.SetMessageBoxMode(0xF0000)

    def enable_track_changes(self) -> None:
        """한글의 "변경내용 추적" 모드를 켠다. MenuExTrackChange는 토글
        액션이라 이미 켜진 상태에서 또 부르면 꺼져버리므로, 파이썬이 직접
        관리하는 self._track_changes_enabled 플래그로 아직 꺼져있을 때만
        실제로 토글한다.

        (2026-09-06, 실측 확인) 추적이 켜진 상태에서 편집(insert_text 등)을
        하면, 그 편집 도중 한글이 "변경 추적 기능이 사용된 문서입니다...
        선택한 파일 형식으로 저장할까요?" Yes/No 확인창을 예고 없이 띄우는
        현상이 재현됨(close() 시점뿐 아니라 편집 중에도 발생) — 무인 실행
        환경에서는 답할 사람이 없어 그대로 멈춘다. 추적을 켜는 시점부터
        SetMessageBoxMode(0x20000)("예/아니오" 팝업의 "아니오" 자동누르기,
        pyhwpx set_message_box_mode 문서의 카테고리 5 — close()의 정정
        경위 참고)으로 이런 확인창에 자동 "아니오"를 응답하게 해서, 추적이
        켜져있는 동안 어떤 편집 도중에 떠도 멈추지 않게 한다
        (disable_track_changes()가 짝을 맞춰 기본값으로 되돌림 — close()도
        별도로 자체 방어를 하지만, 그건 추적을 끄지 않고 바로 닫는 흔한
        경로에 대한 보험이지 이 시점의 방어를 대신하지 않는다)."""
        if not self._track_changes_enabled:
            self.hwp.SetMessageBoxMode(0x20000)
            self.hwp.HAction.Run("MenuExTrackChange")
            self._track_changes_enabled = True

    def disable_track_changes(self) -> None:
        """변경내용 추적 모드를 끈다. enable_track_changes()와 대칭으로,
        켜져 있을 때만 토글하고, enable_track_changes()가 켜둔
        SetMessageBoxMode도 기본값(0xF0000)으로 되돌린다."""
        if self._track_changes_enabled:
            self.hwp.HAction.Run("MenuExTrackChange")
            self._track_changes_enabled = False
            self.hwp.SetMessageBoxMode(0xF0000)

    def accept_all_changes(self) -> None:
        """추적된 모든 변경사항을 일괄 승인(적용)한다. pyhwpx가 이미
        TrackChangeApplyAll 액션을 hwp.TrackChangeApplyAll()로 래핑해뒀다
        (site-packages/pyhwpx/run_methods.py 4863번째 줄 확인됨) — 그대로
        호출한다. 개별 항목 탐색/승인(TrackChangeNext/Prev)은 특정 상황
        (숨김창)에서 응답없음이 실측 확인되어(2026-09-05) 이번 라운드에서
        의도적으로 쓰지 않는다 — 다음 라운드 과제."""
        self.hwp.TrackChangeApplyAll()

    def reject_all_changes(self) -> None:
        """추적된 모든 변경사항을 일괄 거부(취소)한다. accept_all_changes()와
        같은 이유로 hwp.TrackChangeCancelAll()을 그대로 호출한다."""
        self.hwp.TrackChangeCancelAll()


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


def _selftest_get_window_handle():
    """get_window_handle()이 실제 win32gui 함수에 바로 쓸 수 있는 정수 HWND를
    돌려주는지 확인한다 (창 배치 기능의 전제 조건)."""
    import tempfile
    import win32gui
    from pyhwpx import Hwp

    test_path = os.path.join(tempfile.gettempdir(), "_test_핸들.hwp")
    setup = Hwp(visible=False, new=True)
    setup.insert_text("핸들 테스트")
    setup.save_as(test_path)
    setup.quit()

    report = None
    try:
        report = HwpReport(test_path)
        hwnd = report.get_window_handle()
        assert isinstance(hwnd, int) and hwnd > 0, hwnd
        assert win32gui.IsWindow(hwnd), f"win32gui가 인식하지 못하는 핸들: {hwnd}"
        title = win32gui.GetWindowText(hwnd)
        assert "한글" in title or "핸들 테스트" in title or title, title
        print("HwpReport.get_window_handle() 통과:", hwnd, repr(title))
    finally:
        if report is not None:
            report.close(save=False)
        os.remove(test_path)


def _selftest_track_changes_toggle_guard():
    """enable_track_changes()를 두 번 연달아 호출해도 실제로는 한 번만
    토글되는지 확인한다. MenuExTrackChange는 토글 액션이라 상태 가드 없이
    두 번 부르면 다시 꺼져버리는데, IsTrackChange 프로퍼티는 이 PC의 한글
    버전에서 AttributeError가 나서(2026-09-05 실측 확인) 한글에 직접 물어볼
    수 없다 — 그래서 파이썬이 직접 관리하는 self._track_changes_enabled
    플래그가 기대대로 바뀌는지를 직접 확인한다(실제 추적 동작 자체의
    증명은 accept_all/reject_all이 추가되는 Task 2의 self-test에서 한다)."""
    import tempfile
    from pyhwpx import Hwp

    test_path = os.path.join(tempfile.gettempdir(), "_test_추적토글.hwp")
    setup = Hwp(visible=False, new=True)
    setup.insert_text("토글 테스트")
    setup.save_as(test_path)
    setup.quit()

    report = None
    try:
        report = HwpReport(test_path)
        assert report._track_changes_enabled is False, "초기 상태는 꺼짐이어야 함"

        report.enable_track_changes()
        assert report._track_changes_enabled is True, "enable 후 켜짐 상태여야 함"

        report.enable_track_changes()  # 두 번째 호출 - 가드가 없으면 여기서 다시 꺼짐
        assert report._track_changes_enabled is True, (
            "두 번째 enable 호출 후에도 여전히 켜짐 상태여야 함(가드 동작 확인)"
        )

        report.disable_track_changes()
        assert report._track_changes_enabled is False, "disable 후 꺼짐 상태여야 함"

        report.disable_track_changes()  # 두 번째 호출 - 가드가 없으면 여기서 다시 켜짐
        assert report._track_changes_enabled is False, (
            "두 번째 disable 호출 후에도 여전히 꺼짐 상태여야 함(가드 동작 확인)"
        )
        print("HwpReport 변경내용추적 토글 가드 통과")
    finally:
        if report is not None:
            report.close(save=False)
        os.remove(test_path)


def _selftest_reject_all_changes_reverts_tracked_insertion():
    """핵심 가치 증명: 추적모드를 켠 상태에서 삽입한 텍스트는 "변경 이력"으로
    기록되고, 그 이력을 거부하면 실제로 되돌려진다는 것을 확인한다.

    (2026-09-06, 실측 확인·정정) 추적 중 아직 승인/거부되지 않은 상태에서는
    get_text()(GetTextFile)가 삽입한 부분만이 아니라 문서 전체를 빈
    문자열로 반환한다는 것을 실측으로 확인함(원본 "원본 문장입니다."까지
    같이 안 보임) — 그래서 "삽입 직후, 거부 전"에 get_text()로 삽입이
    반영됐는지 확인하는 중간 검증은 애초에 성립하지 않는 잘못된 전제였다.
    이 중간 assert는 제거하고, reject_all_changes() 이후의 최종 상태만
    확인한다(이것만으로도 "추적 중 삽입 → 거부 → 되돌려짐"이라는 핵심
    가치는 충분히 증명됨)."""
    import tempfile
    from pyhwpx import Hwp

    test_path = os.path.join(tempfile.gettempdir(), "_test_추적거부.hwp")
    setup = Hwp(visible=False, new=True)
    setup.insert_text("원본 문장입니다.")
    setup.save_as(test_path)
    setup.quit()

    report = None
    try:
        report = HwpReport(test_path)
        report.enable_track_changes()
        report.hwp.MoveDocEnd()
        report.hwp.insert_text("추가로 삽입된 문장입니다.")

        report.reject_all_changes()
        text_after_reject = report.get_text()
        assert "추가로 삽입된 문장입니다." not in text_after_reject, (
            "추적모드에서 삽입한 텍스트가 거부 후에도 남아있음 - "
            "추적이 실제로 기록되지 않은 것으로 보임"
        )
        assert "원본 문장입니다." in text_after_reject, "원본 텍스트까지 사라짐"
        print("HwpReport 변경내용추적 거부(reject_all_changes) 통과: 추적 중 삽입이 거부로 되돌려짐")
    finally:
        if report is not None:
            report.close(save=False)
        os.remove(test_path)


def _selftest_enable_track_changes_twice_still_tracks():
    """enable_track_changes()를 두 번 호출해도(가드 동작) 여전히 추적 중이어서,
    그 뒤 삽입한 텍스트가 거부로 되돌려지는지 확인한다 — Task 1의 플래그
    기반 테스트를 실제 한글 동작으로 재확인한다(가드가 없었다면 두 번째
    enable 호출이 추적을 도로 꺼버려 이 테스트가 실패했을 것)."""
    import tempfile
    from pyhwpx import Hwp

    test_path = os.path.join(tempfile.gettempdir(), "_test_추적이중enable.hwp")
    setup = Hwp(visible=False, new=True)
    setup.insert_text("원본 문장입니다.")
    setup.save_as(test_path)
    setup.quit()

    report = None
    try:
        report = HwpReport(test_path)
        report.enable_track_changes()
        report.enable_track_changes()  # 두 번째 호출 - 가드가 없으면 여기서 추적이 꺼짐
        report.hwp.MoveDocEnd()
        report.hwp.insert_text("추가로 삽입된 문장입니다.")

        report.reject_all_changes()
        text_after_reject = report.get_text()
        assert "추가로 삽입된 문장입니다." not in text_after_reject, (
            "이중 enable 호출 후에도 추적이 켜진 상태여야 하는데, 삽입이 "
            "거부로 안 지워짐 - 가드 실패로 추적이 꺼졌던 것으로 보임"
        )
        print("HwpReport 이중 enable 후에도 추적 유지 확인 통과")
    finally:
        if report is not None:
            report.close(save=False)
        os.remove(test_path)


def _selftest_reject_all_changes_no_op_without_tracking():
    """추적모드를 켜지 않은 상태(일반 편집)에서 삽입한 텍스트는 "변경 이력"이
    아니므로, reject_all_changes()를 호출해도 사라지면 안 된다."""
    import tempfile
    from pyhwpx import Hwp

    test_path = os.path.join(tempfile.gettempdir(), "_test_추적없음거부.hwp")
    setup = Hwp(visible=False, new=True)
    setup.insert_text("원본 문장입니다.")
    setup.save_as(test_path)
    setup.quit()

    report = None
    try:
        report = HwpReport(test_path)
        # enable_track_changes()를 호출하지 않음 - 추적 없이 일반 삽입
        report.hwp.MoveDocEnd()
        report.hwp.insert_text("추적 없이 삽입된 문장입니다.")
        assert "추적 없이 삽입된 문장입니다." in report.get_text()

        report.reject_all_changes()
        text_after_reject = report.get_text()
        assert "추적 없이 삽입된 문장입니다." in text_after_reject, (
            "추적모드가 꺼진 상태에서 삽입한 일반 텍스트가 reject_all_changes() "
            "호출만으로 사라짐 - 예상치 못한 부작용"
        )
        print("HwpReport 추적 없을 때 reject_all_changes() no-op 확인 통과")
    finally:
        if report is not None:
            report.close(save=False)
        os.remove(test_path)


if __name__ == "__main__":
    # (2026-09-06, 실측 확인) 한글 COM 인스턴스를 연달아 너무 빠르게 만들면
    # "서버 실행이 실패했습니다"(pywintypes.com_error)가 간헐적으로 재현됨
    # — 각 self-test 사이에 짧은 대기를 둬서 이전 인스턴스가 완전히 정리될
    # 시간을 준다.
    _selftest_open_and_mark_red()
    time.sleep(2)
    _selftest_get_window_handle()
    time.sleep(2)
    _selftest_track_changes_toggle_guard()
    time.sleep(2)
    _selftest_reject_all_changes_reverts_tracked_insertion()
    time.sleep(2)
    _selftest_enable_track_changes_twice_still_tracks()
    time.sleep(2)
    _selftest_reject_all_changes_no_op_without_tracking()
