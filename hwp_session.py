"""hwp_session.py — 한글 인스턴스를 만들고 문서를 여는 공용 진입점.

이 프로젝트에서 Hwp()를 만드는 곳이 여러 파일(hwp_report, source_reader,
weekly_report_tool, attachment_preview)에 흩어져 있고, 각각 "반드시 new=True"
라는 원칙을 주석으로만 지키고 있었다. 여기 한 곳에 모아 두 가지를 강제한다.

1. 새 프로세스(new=True)만 만든다 — 이미 떠 있는 사용자의 한글 창을 붙잡는
   사고(hwp-report-tool 메모 2026-09-04 항목)를 구조적으로 막는다.
2. 보안모듈(FilePathCheckerModule) 등록 결과를 확인한다 — pyhwpx는 Hwp()
   생성 시 RegisterModule을 자동 호출하지만 반환값을 버린다. 등록이 실패하면
   파일 접근 때마다 "접근하려는 시도(파일의 손상 또는 유출의 위험 등)가
   있습니다" 대화상자가 떠서 무인 실행이 멈추므로, 실패를 최소한 눈에
   보이게 한다(예외로 막지는 않음 — visible=True 창이면 사람이 '접근 허용'을
   눌러 계속 진행할 수 있다).

(2026-10-01 실측) 이 PC에서는 pyhwpx 1.7.2가 동봉한 32비트 DLL이
HKCU\\Software\\HNC\\HwpAutomation\\Modules에 이미 등록돼 있고, 한글 2020
(32비트)과 비트 수가 맞으며, RegisterModule이 True를 돌려준다. 임시폴더·
점(.)폴더·.worktrees 경로 모두 대화상자 없이 0.1초 만에 열렸다. 즉
hwp-report-tool 메모(2026-09-16)의 "보안모듈 등록" 과제는 별도 DLL 다운로드
없이 이미 충족된 상태였고, 남은 일은 이 확인을 코드에 넣는 것뿐이었다.

열기 옵션(HwpAutomation.pdf, pyhwpx open() docstring에서 확인):
- forceopen:true — 읽기 전용으로 열어야 할 때 대화상자를 띄우지 않는다.
- suspendpassword:true — 암호 걸린 파일이면 묻지 않고 그냥 실패(False)로.
  무인(숨김) 읽기에만 쓴다. 사람이 보는 창(HwpReport)에서는 암호를 직접
  입력할 수 있어야 하므로 이 옵션을 넣지 않는다.
- versionwarning은 기본값이 FALSE(경고 안 띄움)라 따로 지정하지 않는다."""
import sys


def _warn_stderr(message: str) -> None:
    """기본 경고 출력. stdout이 아니라 stderr로 — source_reader/attachment_preview의
    격리 subprocess는 stdout을 JSON 결과로 읽기 때문에 stdout에 섞이면 안 된다."""
    print(message, file=sys.stderr)


SECURITY_MODULE_TYPE = "FilePathCheckDLL"
SECURITY_MODULE_ID = "FilePathCheckerModule"

INTERACTIVE_OPEN_ARG = "forceopen:true"
UNATTENDED_OPEN_ARG = "forceopen:true;suspendpassword:true"


def ensure_security_module(hwp, warn=_warn_stderr) -> bool:
    """보안모듈을 (다시) 등록하고 결과를 돌려준다. 실패하면 warn()으로 알린다.
    hwp는 pyhwpx.Hwp 인스턴스(원본 COM 객체는 hwp.hwp)."""
    try:
        registered = bool(hwp.hwp.RegisterModule(SECURITY_MODULE_TYPE, SECURITY_MODULE_ID))
    except Exception as e:
        registered = False
        warn(f"[hwp_session] RegisterModule 호출 실패: {e}")
    if not registered:
        warn("[hwp_session] 보안모듈이 등록되지 않았습니다 — 파일을 열 때마다 "
             "한글의 접근 확인 대화상자가 뜰 수 있습니다(pyhwpx의 "
             "FilePathCheckerModule.dll 레지스트리 등록 상태를 확인하세요).")
    return registered


def new_hwp(visible: bool, warn=_warn_stderr):
    """항상 새 한글 프로세스를 만들고 보안모듈 등록을 확인해서 돌려준다."""
    from pyhwpx import Hwp
    hwp = Hwp(visible=visible, new=True)
    ensure_security_module(hwp, warn=warn)
    return hwp


def open_document(hwp, path: str, unattended: bool) -> bool:
    """문서를 연다. unattended=True(숨김 읽기)면 암호 파일도 묻지 않고 실패
    처리해서 무인 실행이 멈추지 않게 한다. 반환값은 pyhwpx open()과 같다."""
    arg = UNATTENDED_OPEN_ARG if unattended else INTERACTIVE_OPEN_ARG
    return hwp.open(path, arg=arg)


# ---------------------------------------------------------------- self-tests

class _FakeCom:
    def __init__(self, result):
        self._result = result
        self.calls = []

    def RegisterModule(self, module_type, module_data):
        self.calls.append((module_type, module_data))
        if isinstance(self._result, Exception):
            raise self._result
        return self._result


class _FakeHwp:
    def __init__(self, register_result):
        self.hwp = _FakeCom(register_result)
        self.opened = None

    def open(self, path, format="", arg=""):
        self.opened = (path, arg)
        return True


def _selftest_ensure_security_module_reports_failure():
    """등록 실패(False)와 호출 예외 모두 warn으로 드러나고 False를 돌려준다."""
    warnings = []
    ok = ensure_security_module(_FakeHwp(False), warn=warnings.append)
    assert ok is False and len(warnings) == 1, (ok, warnings)
    warnings.clear()
    ok = ensure_security_module(_FakeHwp(RuntimeError("COM 오류")), warn=warnings.append)
    assert ok is False and len(warnings) == 2, (ok, warnings)
    warnings.clear()
    fake = _FakeHwp(True)
    ok = ensure_security_module(fake, warn=warnings.append)
    assert ok is True and warnings == [], (ok, warnings)
    assert fake.hwp.calls == [(SECURITY_MODULE_TYPE, SECURITY_MODULE_ID)], fake.hwp.calls
    print("ensure_security_module(성공/실패/예외 보고) 통과")


def _selftest_open_document_passes_open_args():
    """숨김 읽기와 사람이 보는 창에 각각 맞는 열기 옵션이 전달된다."""
    fake = _FakeHwp(True)
    open_document(fake, "a.hwp", unattended=True)
    assert fake.opened == ("a.hwp", UNATTENDED_OPEN_ARG), fake.opened
    open_document(fake, "b.hwp", unattended=False)
    assert fake.opened == ("b.hwp", INTERACTIVE_OPEN_ARG), fake.opened
    assert "suspendpassword" not in INTERACTIVE_OPEN_ARG
    print("open_document(열기 옵션 전달) 통과")


def _selftest_live_new_hwp_opens_without_dialog():
    """실제 한글: 새 숨김 프로세스에서 보안모듈이 True로 등록되고, 점(.)폴더
    경로의 문서가 대화상자 없이 곧바로 열린다(멈추면 이 테스트가 끝나지
    않는 것으로 드러남). 다른 한글 창이 떠 있어도 new=True라 안전하다."""
    import os, shutil, tempfile, time
    folder = os.path.join(tempfile.gettempdir(), ".hwp_session_probe")
    os.makedirs(folder, exist_ok=True)
    path = os.path.join(folder, "probe.hwp")
    warnings = []
    hwp = new_hwp(visible=False, warn=warnings.append)
    try:
        assert warnings == [], warnings
        hwp.insert_text("hwp_session 프로브")
        assert hwp.save_as(path)
        started = time.time()
        assert open_document(hwp, path, unattended=True)
        elapsed = time.time() - started
        assert elapsed < 5, f"열기에 {elapsed:.1f}초 — 대화상자가 떴을 가능성"
        print(f"new_hwp/open_document(실제 한글, {elapsed:.2f}초) 통과")
    finally:
        hwp.quit()
        shutil.rmtree(folder, ignore_errors=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    _selftest_ensure_security_module_reports_failure()
    _selftest_open_document_passes_open_args()
    _selftest_live_new_hwp_opens_without_dialog()
