"""run_selftests.py — 모든 모듈의 self-test를 한 번에 돌리는 회귀 실행기(PRD 16-5 "평가셋" 1단계).

각 모듈을 **별도 프로세스로, 순차로** 실행한다. 이유: pyhwpx self-test를 동시에
여러 개 돌리면 서로 다른 프로세스의 한글 상태가 섞여 거짓 실패가 난다
(hwp-report-tool 메모 2026-09-03/09-04). 한글 COM을 쓰는 모듈 사이에는 짧은
대기를 둔다(연달아 인스턴스를 만들면 "서버 실행이 실패했습니다"가 간헐적으로
난다 — hwp_report.py __main__ 주석).

분류:
  pure   — 한글·서버 없이 도는 것(초 단위)
  com    — 한글 COM 필요(새 프로세스 new=True만 쓰므로 사용자의 한글 창은 안전)
  server — 홈서버 LLM이 켜져 있어야 통과하는 것(꺼져 있으면 건너뜀으로 표시)

사용:
  python run_selftests.py            # pure + com (server는 서버가 응답할 때만)
  python run_selftests.py --pure     # 빠른 확인
  python run_selftests.py --only doc_sections section_edit_tool
결과는 .tmp/selftest-runs/<시각>.log 에도 남긴다. 실패가 하나라도 있으면 종료코드 1."""
import argparse
import os
import subprocess
import sys
import time
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))

MODULES = {
    # name: (category, argv 추가 인자, 타임아웃 초)
    "verify_numbers": ("pure", [], 120),
    "ignore_list": ("pure", [], 60),
    "privacy_guard": ("pure", [], 60),
    "speed_tracker": ("pure", [], 60),
    "routing_graph": ("pure", [], 60),
    "agent_loop": ("pure", [], 120),
    "loop_runner": ("pure", [], 120),
    "archive_index": ("pure", [], 120),
    "archive_qa_tool": ("pure", [], 120),
    "attachment_qa_tool": ("pure", [], 120),
    "chat_assistant": ("pure", ["--selftest"], 120),
    "chat_handoff": ("pure", ["--selftest"], 60),
    "window_layout": ("com", [], 120),
    "hwp_session": ("com", [], 180),
    "hwp_report": ("com", [], 600),
    "doc_sections": ("com", [], 300),
    "section_edit_tool": ("com", [], 300),
    "attachments": ("com", [], 300),
    "attachment_preview": ("com", [], 300),
    "source_reader": ("com", [], 600),
    "table_tool": ("com", [], 600),
    "numbering_tool": ("com", [], 600),
    "chart_tool": ("com", [], 600),
    "fit_to_page_tool": ("com", [], 600),
    "weekly_report_tool": ("com", [], 600),
    "verify_tool": ("com", [], 900),
    "polish_tool": ("server", [], 300),
}


def server_alive(url: str = "http://100.74.107.35:11434/api/tags", timeout: float = 3) -> bool:
    try:
        with urllib.request.urlopen(url, timeout=timeout) as r:
            return r.status == 200
    except Exception:
        return False


def _hwp_pids() -> set:
    try:
        out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq Hwp.exe", "/FO", "CSV", "/NH"],
                             capture_output=True, text=True).stdout
    except Exception:
        return set()
    pids = set()
    for line in out.splitlines():
        parts = [p.strip('"') for p in line.split('","')]
        if len(parts) >= 2 and parts[1].isdigit():
            pids.add(int(parts[1]))
    return pids


def _kill_new_hwp(before: set) -> int:
    """모듈이 남기고 간 한글 프로세스만 정리(실행 전에 있던 사용자의 창은 건드리지 않음)."""
    killed = 0
    for pid in _hwp_pids() - before:
        subprocess.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True)
        killed += 1
    return killed


def run_module(name: str, extra: list, timeout: int, log) -> tuple:
    before = _hwp_pids()
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    started = time.time()
    try:
        proc = subprocess.run([sys.executable, os.path.join(HERE, f"{name}.py"), *extra],
                              capture_output=True, text=True, encoding="utf-8", errors="replace",
                              timeout=timeout, env=env, cwd=HERE)
        status = "PASS" if proc.returncode == 0 else "FAIL"
        output = (proc.stdout or "") + (proc.stderr or "")
    except subprocess.TimeoutExpired as e:
        status = "TIMEOUT"
        output = (e.stdout or "") if isinstance(e.stdout, str) else ""
    elapsed = time.time() - started
    stray = _kill_new_hwp(before)
    log.write(f"\n===== {name} [{status}] {elapsed:.1f}s stray_hwp_killed={stray}\n{output}\n")
    log.flush()
    tail = [ln for ln in output.strip().splitlines() if "MERG NOT subset" not in ln][-1:] or [""]
    return status, elapsed, tail[0][:120]


def main(argv):
    ap = argparse.ArgumentParser()
    ap.add_argument("--pure", action="store_true")
    ap.add_argument("--com", action="store_true")
    ap.add_argument("--server", action="store_true")
    ap.add_argument("--only", nargs="*")
    a = ap.parse_args(argv)
    cats = set()
    if a.pure: cats.add("pure")
    if a.com: cats.add("com")
    if a.server: cats.add("server")
    if not cats and not a.only:
        cats = {"pure", "com", "server"}

    alive = server_alive()
    selected = [(n, *spec) for n, spec in MODULES.items()
                if (a.only and n in a.only) or (not a.only and spec[0] in cats)]
    os.makedirs(os.path.join(HERE, ".tmp", "selftest-runs"), exist_ok=True)
    log_path = os.path.join(HERE, ".tmp", "selftest-runs", time.strftime("%Y%m%d-%H%M%S") + ".log")
    rows = []
    with open(log_path, "w", encoding="utf-8") as log:
        log.write(f"server_alive={alive}\n")
        for name, cat, extra, timeout in selected:
            if cat == "server" and not alive:
                rows.append((name, cat, "SKIP(서버 꺼짐)", 0.0, ""))
                print(f"{name:22s} {cat:6s} SKIP(서버 꺼짐)")
                continue
            print(f"{name:22s} {cat:6s} 실행 중…", end="", flush=True)
            status, elapsed, tail = run_module(name, extra, timeout, log)
            rows.append((name, cat, status, elapsed, tail))
            print(f"\r{name:22s} {cat:6s} {status:8s} {elapsed:6.1f}s  {tail}")
            if cat == "com":
                time.sleep(3)
    failed = [r for r in rows if r[2] in ("FAIL", "TIMEOUT")]
    print(f"\n합계: {len(rows)}개 모듈, 통과 {sum(1 for r in rows if r[2] == 'PASS')}, "
          f"실패 {len(failed)}, 건너뜀 {sum(1 for r in rows if r[2].startswith('SKIP'))}  (로그: {log_path})")
    for name, cat, status, _, tail in failed:
        print(f"  ✗ {name}: {status} — {tail}")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.exit(main(sys.argv[1:]))
