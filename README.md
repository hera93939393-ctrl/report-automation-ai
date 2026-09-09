# AI Agent 파헤치기 — 나만의 에이전트 하네스 과제 제출

모두의연구소 "AI Agent 파헤치기" 6강 프로젝트 제출용 브랜치입니다. 이 브랜치는 원래 있던 개인 업무자동화 프로젝트(report-automation-ai, F1~F13) 위에서 작업했고, 이 과제와 무관한 F1~F10 기능 파일은 제외했습니다.

## 이 과제의 신규 결과물

- [agent_loop.py](./agent_loop.py) — 모델 → 도구 요청 해석 → 검사·실행 → 결과 반환 → 다음 판단을 반복하는 실제 에이전트 루프 (R02)
- [code_fix_tool.py](./code_fix_tool.py) + [coding_fixture/](./coding_fixture/) — 코딩 시나리오: 결함 제안 → 사용자 승인 → 적용 → 테스트 (R04/R05). 승인은 모델이 스스로 부르는 도구가 아니라 사용자 채팅 메시지를 직접 파싱해서만 결정합니다(`chat_assistant.py`의 `parse_approval_response`).
- `chat_assistant.py`(특히 `_on_submit`, `_build_tool_dispatch`, `_handle_tool_result`) — 위 두 가지를 기존 프로그램의 채팅 화면에 실제로 연결한 부분(D02: 기존 프로그램 통합)

## D01 사용자 이야기(숫자 검증) 구현

한글 문서의 수치를 엑셀 원본과 문맥 기반으로 대조해 색상으로 표시하는 기능입니다.

- [verify_numbers.py](./verify_numbers.py) — 값 추출·문맥 연결·대조 로직
- [hwp_report.py](./hwp_report.py) — 한글 문서 색상 표시
- [source_reader.py](./source_reader.py) — 엑셀 등 원본 자료 읽기
- [verify_tool.py](./verify_tool.py) — 위를 묶어 실행하는 진입점
- [file_answer.py](./file_answer.py) + [_selftest_file_answer.py](./_selftest_file_answer.py) — 작은 텍스트 파일을 실제 모델로 읽고 근거 있게 답하는 도구(R03/A01)

## 설계 문서와 벤치마크

- [harness-design-kit/](./harness-design-kit/) — PRD·설계 결정(DECISIONS)·도구 계약(INTERFACES)·작업 기록(IMPLEMENTATION_PLAN)·완료 조건과 실제 증거(ACCEPTANCE) 전체
- [harness-lab/](./harness-lab/) — 고정 10문항(Terminal-Bench Pro 로컬 이식판) 벤치마크 하네스와 실행 결과. [EXPERIMENT_REPORT.md](./harness-lab/EXPERIMENT_REPORT.md)에 기준(Ollama CPU)·개선(vLLM GPU) 비교와 결론을 정리했습니다.

## 그 외 포함된 파일 (이 과제의 채점 대상은 아님)

`chat_assistant.py`는 F1~F13이 이미 통합된 화면이라, 정상적으로 import되어 실행되려면 다음 지원 모듈이 함께 있어야 합니다: `attachment_preview.py`, `ignore_list.py`, `privacy_guard.py`, `window_layout.py`, `table_tool.py`, `numbering_tool.py`, `polish_tool.py`, `speed_tracker.py`, `weekly_report_tool.py`, `fit_to_page_tool.py`. 이 파일들은 과제 이전부터 있던 기존 프로그램 기능이며, 이번 과제의 신규 구현물이 아닙니다.

## 실행 방법

```bash
pip install -r requirements.txt
python chat_assistant.py
```

Ollama(`qwen3.5:2b`)가 로컬에 떠 있어야 하며, 한글(HWP)이 설치되어 있어야 D01 시나리오를 실제로 확인할 수 있습니다. 벤치마크 재현 방법은 [harness-lab/README_lab.md](./harness-lab/README_lab.md) 참고.
