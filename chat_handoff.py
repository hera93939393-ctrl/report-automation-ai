"""chat_handoff.py — 한글 툴바 버튼(F10 매크로)에서 채팅창으로 선택 텍스트를 넘긴다.

툴바 버튼 연결 방식 2번(2026-10-01 사용자 결정, 플랜 10절): 버튼은 사용자가
띄운 한글 창 안에서 돌고, 채팅 앱은 자기 전용 한글 프로세스로 문서를 연다 —
둘은 다른 문서 핸들이라 버튼이 문서를 직접 고치게 하지 않는다. 버튼은
"채팅창을 앞으로 가져오고 선택 텍스트를 입력창에 채우는" 일만 한다.

흐름:
  채팅 앱 시작 → register_chat_window(hwnd)로 창 핸들을 등록 파일에 적음
  버튼 → 매크로가 선택 텍스트를 UTF-16 임시파일로 저장 → 이 파일을 CLI로 실행
       → send_to_chat(): 받은편지함 파일에 텍스트를 쓰고 채팅창을 앞으로
  채팅 앱 → 주기적으로 take_inbox()로 받은편지함을 비워 입력창에 채움

채팅창을 앞으로 가져오는 일은 채팅 앱이 아니라 이 CLI가 한다 — 윈도우는
배경 프로세스의 SetForegroundWindow를 막지만, "지금 앞에 있는 프로세스(한글)가
실행한 프로세스"에는 허용한다(SetForegroundWindow 문서의 허용 조건).

사용법(매크로가 호출): python chat_handoff.py <in_path> <out_path> [doc_path]
  in_path: 선택 텍스트(UTF-16), out_path: 결과 안내문(UTF-16, 실패 시 매크로가 팝업으로 보여줌)"""
import json
import os
import sys
import tempfile

REGISTRY_NAME = "hangul_butler_chat.json"
INBOX_NAME = "hangul_butler_inbox.json"


def _default_dir() -> str:
    """등록·받은편지함 파일 위치. TEMP가 아니라 LOCALAPPDATA를 쓴다 — 이 PC는
    ESTsoft가 TEMP를 바꿔 두어(CLAUDE.md 2절), 한글이 실행한 파이썬과 채팅 앱이
    서로 다른 TEMP를 보면 버튼이 채팅창을 영영 못 찾는다."""
    base = os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()
    d = os.path.join(base, "HangulButler")
    os.makedirs(d, exist_ok=True)
    return d


def _paths(base_dir=None):
    base = base_dir or _default_dir()
    return os.path.join(base, REGISTRY_NAME), os.path.join(base, INBOX_NAME)


def _write_json_atomic(path: str, data: dict) -> None:
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)
    os.replace(tmp, path)


def register_chat_window(hwnd: int, base_dir=None) -> None:
    registry, _ = _paths(base_dir)
    _write_json_atomic(registry, {"hwnd": int(hwnd), "pid": os.getpid()})


def unregister_chat_window(base_dir=None) -> None:
    registry, _ = _paths(base_dir)
    try:
        with open(registry, encoding="utf-8") as f:
            if json.load(f).get("pid") != os.getpid():
                return  # 다른 채팅 앱 인스턴스가 나중에 등록한 것이면 건드리지 않는다
        os.remove(registry)
    except (OSError, ValueError):
        pass


def _window_alive(hwnd: int) -> bool:
    import win32gui
    return bool(win32gui.IsWindow(hwnd))


def _bring_to_front(hwnd: int) -> None:
    import win32con
    import win32gui
    if win32gui.IsIconic(hwnd):
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    try:
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        pass  # 포그라운드 권한이 없으면 실패할 수 있다 - 채팅창은 항상 위(topmost)라 보이기는 한다


def send_to_chat(text: str, doc_path: str = "", base_dir=None,
                 window_alive=_window_alive, focus=_bring_to_front) -> dict:
    """선택 텍스트를 채팅창 받은편지함에 넣고 채팅창을 앞으로 가져온다.
    채팅 앱이 안 떠 있으면 {"ok": False, "reason": ...}."""
    registry, inbox = _paths(base_dir)
    try:
        with open(registry, encoding="utf-8") as f:
            hwnd = int(json.load(f)["hwnd"])
    except (OSError, ValueError, KeyError):
        return {"ok": False, "reason": "한글집사 채팅창이 켜져 있지 않아요. 먼저 채팅창을 실행해 주세요."}
    if not window_alive(hwnd):
        return {"ok": False, "reason": "한글집사 채팅창이 켜져 있지 않아요. 먼저 채팅창을 실행해 주세요."}
    _write_json_atomic(inbox, {"text": text, "doc_path": doc_path})
    focus(hwnd)
    return {"ok": True}


def take_inbox(base_dir=None):
    """받은편지함에 온 것이 있으면 꺼내고(파일 삭제) 돌려준다. 없으면 None."""
    _, inbox = _paths(base_dir)
    try:
        with open(inbox, encoding="utf-8") as f:
            data = json.load(f)
    except (OSError, ValueError):
        return None
    try:
        os.remove(inbox)
    except OSError:
        pass
    return data


def same_document(a: str, b: str) -> bool:
    """두 경로가 같은 문서인지(대소문자·슬래시 차이 무시). 한쪽이라도 비면 판단하지 않는다(True)."""
    if not a or not b:
        return True
    return os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))


def main(argv, base_dir=None) -> int:
    if len(argv) not in (3, 4):
        print("사용법: python chat_handoff.py <in_path> <out_path> [doc_path]", file=sys.stderr)
        return 1
    in_path, out_path = argv[1], argv[2]
    doc_path = argv[3] if len(argv) == 4 else ""
    try:
        with open(in_path, encoding="utf-16") as f:
            text = f.read()
    except OSError:
        text = ""
    result = send_to_chat(text.strip(), doc_path, base_dir=base_dir)
    with open(out_path, "w", encoding="utf-16") as f:
        f.write("" if result["ok"] else result["reason"])
    return 0 if result["ok"] else 2


# ---------------------------------------------------------------- self-tests (한글·서버 불필요)

def _test_dir():
    import atexit
    import shutil
    d = tempfile.mkdtemp(prefix="_test_handoff_")
    atexit.register(shutil.rmtree, d, ignore_errors=True)
    return d


def _selftest_send_without_chat_app_fails_honestly():
    d = _test_dir()
    r = send_to_chat("문장", base_dir=d, window_alive=lambda h: True, focus=lambda h: None)
    assert r["ok"] is False and "켜져 있지 않아요" in r["reason"], r
    assert take_inbox(d) is None
    print("핸드오프: 채팅 앱이 없으면 정직하게 실패 통과")


def _selftest_stale_registration_is_treated_as_not_running():
    d = _test_dir()
    register_chat_window(12345, base_dir=d)
    focused = []
    r = send_to_chat("문장", base_dir=d, window_alive=lambda h: False, focus=focused.append)
    assert r["ok"] is False and focused == [] and take_inbox(d) is None, (r, focused)
    print("핸드오프: 창이 이미 닫힌 등록은 '안 켜짐'으로 처리 통과")


def _selftest_roundtrip_fills_inbox_and_focuses_once():
    d = _test_dir()
    register_chat_window(777, base_dir=d)
    focused = []
    text = "심사 절차를 표준화한다.\n① 서류심사 — 「공공급식」 'eaT'"
    r = send_to_chat(text, doc_path="C:/문서/계획(안).hwp", base_dir=d,
                     window_alive=lambda h: True, focus=focused.append)
    assert r == {"ok": True} and focused == [777], (r, focused)
    got = take_inbox(d)
    assert got == {"text": text, "doc_path": "C:/문서/계획(안).hwp"}, got
    assert take_inbox(d) is None  # 한 번 꺼내면 비어야 한다(같은 텍스트가 두 번 채워지지 않게)
    print("핸드오프: 받은편지함 왕복(유니코드 보존) + 한 번만 꺼내짐 통과")


def _selftest_unregister_only_removes_own_registration():
    d = _test_dir()
    registry, _ = _paths(d)
    _write_json_atomic(registry, {"hwnd": 1, "pid": os.getpid() + 1})  # 다른 인스턴스가 등록한 것
    unregister_chat_window(d)
    assert os.path.exists(registry)
    register_chat_window(2, base_dir=d)
    unregister_chat_window(d)
    assert not os.path.exists(registry)
    print("핸드오프: 자기 등록만 지움 통과")


def _selftest_same_document():
    assert same_document("C:/문서/A.hwp", "c:\\문서\\a.hwp")
    assert not same_document("C:/문서/A.hwp", "C:/문서/B.hwp")
    assert same_document("", "C:/문서/B.hwp")  # 매크로가 경로를 못 넘기면 판단 보류
    print("핸드오프: 같은 문서 판정 통과")


def _selftest_cli_writes_reason_when_chat_not_running():
    d = _test_dir()
    in_path, out_path = os.path.join(d, "in.txt"), os.path.join(d, "out.txt")
    with open(in_path, "w", encoding="utf-16") as f:
        f.write("문장")
    code = main(["chat_handoff.py", in_path, out_path], base_dir=d)
    with open(out_path, encoding="utf-16") as f:
        msg = f.read()
    assert code == 2 and "켜져 있지 않아요" in msg, (code, msg)
    print("핸드오프: CLI가 실패 안내문을 UTF-16으로 남김 통과")


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] != "--selftest":
        sys.exit(main(sys.argv))
    sys.stdout.reconfigure(encoding="utf-8")
    _selftest_send_without_chat_app_fails_honestly()
    _selftest_stale_registration_is_treated_as_not_running()
    _selftest_roundtrip_fills_inbox_and_focuses_once()
    _selftest_unregister_only_removes_own_registration()
    _selftest_same_document()
    _selftest_cli_writes_reason_when_chat_not_running()
