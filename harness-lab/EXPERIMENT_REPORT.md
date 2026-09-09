# 하네스 구현과 실험 보고서
빈칸은 실행 후 채웁니다. 아래는 결과 예시가 아니라 제출 양식입니다.

## 제품과 구현
- 사용자 시나리오 / 완성한 PRD 위치: D01(한글 문서 수치를 엑셀 원본과 문맥 기반으로 대조 검증) + D05-코딩(지정 결함 수정, 승인 후 적용). `harness-design-kit/PRD.md`, `DECISIONS.md`, `INTERFACES.md`, `ACCEPTANCE.md`
- 구현 언어와 실행 방법: Python 3.13. 기존 한글 자동화 프로그램(`chat_assistant.py`, pyhwpx)의 채팅창에서 실행 — `python chat_assistant.py`
- 직접 작성하거나 변경한 부분 / 참고한 부분: `agent_loop.py`(반복 루프)·`code_fix_tool.py`(코딩 시나리오)는 신규 작성. `chat_assistant.py`의 1회성 `route_intent()`를 제거하고 `run_agent_loop` 연결. 이 harness_lab(과제 제공 참고 구현)은 원본을 그대로 두고 `VLLMProvider`(providers.py)와 `--base-url`(local_agent.py/bench.py) 지원만 patch로 추가함 — zip 다운로드 그대로라 git 커밋 이력은 없음.
- 모델·도구 반복, 오류·종료, 권한, 세션의 코드 위치: 반복/오류/한도 = `agent_loop.py:run_agent_loop`. 승인 전용 실행(`apply_fix_and_test`)은 모델 도구 스키마에서 의도적으로 제외 = `code_fix_tool.py`. 세션 결선·승인 파싱 = `chat_assistant.py:_on_submit`, `parse_approval_response`.
- Python 예시와 다른 설계 결정 및 이유: 코드 수정 승인은 모델이 스스로 호출하는 도구가 아니라, 사용자가 채팅창에 입력한 문장을 직접 파싱해서만 결정한다(`parse_approval_response`) — 모델이 자기 자신의 변경을 승인할 수 있는 경로를 원천 차단하기 위함.

## 평가 조건
- 원본 commit / subset ID / manifest SHA256: suite_id=`terminal-bench-pro-local-port-v1`, revision=`874af409da6aafebccbf3bc5bb41a2fa4d78784d` (jobs/baseline/manifest.json 및 jobs/baseline-vllm/manifest.json에 문항별 source_files SHA256 전체 기록)
- 제공자와 모델의 실제 이름: 기준=Ollama `qwen3.5:2b`(로컬 CPU) / 개선(시도)=vLLM `cyankiwi/Qwen3.5-4B-AWQ-4bit`(Colab GPU, ngrok 터널 경유)
- 코드 commit 또는 소스 SHA256 (기준 / 개선): harness-lab은 git 저장소가 아닌 zip 다운로드본 - 커밋 해시 없음. 기준/개선 모두 동일한 patch본(providers.py에 VLLMProvider 추가, local_agent.py/bench.py에 vllm/--base-url 분기 추가) 사용. [OPEN: 제출 전 harness-lab 폴더 전체를 소스 사본으로 동봉하는 것으로 대체]
- 실행 환경 / 아키텍처 / Python 버전 / 로컬 이식판 버전: Windows 11 Pro, 기준=i5-10210U(GPU 없음)/16GB RAM 로컬 실행, 개선=Google Colab GPU 런타임(Tesla T4, 15360MiB, Driver 580.82.07, CUDA 13.0) + ngrok 터널 경유 원격 호출(vLLM 0.28.0, `--enforce-eager --max-model-len 32768 --tool-call-parser qwen3_xml`). port_version 1.0.0(manifest.json)
- 의존성 잠금 파일, 로컬 실행 조건과 재현 제약: [OPEN: uv.lock/requirements 경로 확인 후 기입]
- 문항 수 10 / easy 2 / medium 4 / hard 4 / 문항당 반복 수: 1
- 모델 호출·도구·시간·출력 한도: max_steps=40, max_seconds=300, command_timeout=10 (jobs/baseline/requested-config.json과 동일 설정을 개선 실행에도 사용)
- 기준 실행 폴더 / 개선 실행 폴더: `jobs/baseline` (완료, 유효) / `jobs/baseline-vllm`(1차 시도, 인프라 중단으로 무효 - 아래 실패 분석 참고) → `jobs/improved-vllm`(2차 시도, 완료·유효, 최종 채택)
- 설치·문제 준비·채점 경로 확인 결과 (개발 진단은 학생 점수에서 제외): 기준·개선(무효화된 시도) 모두 설치·준비 단계 자체는 정상 진행됨(채점기 진단 오류 없음) — 두 실행의 실패는 전부 에이전트 쪽 원인(기준=속도, 개선 시도=인프라 중단)으로 분류

## 결과
모든 문제와 모든 시도를 포함한 CSV를 첨부합니다. 실패·오류·미완료를 구분합니다.

| 측정 | 통과/전체 시도 | 쉬움 | 중간 | 어려움 | 실패 | 실행 오류 | 미완료 | 시간 | 토큰 | 비용 |
|---|---|---|---|---|---|---|---|---|---|---|
| 기준 | 0/10 | 0/2 | 0/4 | 0/4 | 0 | 0 | 10 (전부 AgentIncomplete) | 총 2688.4초(평균 268.8초, jobs/baseline/scoreboard/trials.csv) | 알 수 없음(미기록) | 알 수 없음 |
| 개선(1차 시도, 무효) | 0/10 (인프라 중단으로 재측정) | 0/2 | 0/4 | 0/4 | 0 | 0 | 10 (그중 7건은 Colab 연결 끊김에 따른 즉시 HTTPStatusError, 1건은 9시간+ 뒤 ReadTimeout, 2건만 진짜 시도) | 총 34009.3초(대부분 절전 중 경과 시간, jobs/baseline-vllm/scoreboard/trials.csv) | 알 수 없음(미기록) | 알 수 없음 |
| 개선(2차, 유효) | 0/10 | 0/2 | 0/4 | 0/4 | 0 | 0 | 10 (9건은 실제 GPU 시도, 1건은 실행 도중 ngrok 터널 종료로 HTTPStatusError) | 총 2063.9초(평균 206.4초, jobs/improved-vllm/scoreboard/trials.csv) | 알 수 없음(미기록) | 알 수 없음 |

관측하지 않은 시간·토큰·비용은 '알 수 없음'으로 씁니다. 별도 가격표로 비용을 계산했다면 공식 청구값과 구별하고 단가·날짜·계산 방법을 적습니다.

## 실패 분석과 개선 가설
- 대표 실패의 요청 / 관찰한 도구 실행 / 채점 결과: 기준 10문항 전부 `AgentIncomplete`, elapsed_seconds 133~330초로 `max_seconds=300` 근처에서 시간을 소진하거나 근접해 미완료로 종료. 문항별 실제 model_calls/tool_calls 수는 jobs/baseline/*/events.jsonl 단위 확인 필요 - [OPEN]
- 추정 원인과 이를 뒷받침하는 기록: GPU 없이 CPU에서 qwen3.5:2b를 추론하면 model_call 1회당 응답 시간이 길어, 정해진 max_steps/max_seconds 안에 도구를 여러 번 호출해 문제를 풀 반복 횟수 자체를 확보하지 못한다는 가설(모델 능력 부족이 아니라 하드웨어 속도 병목). 부분 근거: jobs/baseline-vllm의 3번째 문항(detect-corrupted-blockchain-transaction)은 GPU 경로에서 실제로 model_calls=6·tool_calls=5까지 진행함 - 같은 문항이 기준에서는 이 정도 반복에 도달했는지 개별 확인 필요 - [OPEN]
- 변경한 한 가지 요소: 추론 backend를 Ollama(로컬 CPU)에서 vLLM(Colab GPU, ngrok 경유)으로 교체. 문항·순서·max_steps/max_seconds/command_timeout은 기준과 동일하게 유지.
- 예상한 영향: 반복당 응답 시간이 줄어, 같은 시간·반복 한도 안에서 실제로 도구를 여러 번 호출해 완주하는 문항이 늘어날 것으로 예상.
- 실제 변화 (좋아진 문제와 나빠진 문제 모두): 2026-09-09 유효 재실행(jobs/improved-vllm) 완료. **통과 문항 수는 0/10→0/10으로 변화 없음**(개선 없음). 그러나 문항별 model_calls/tool_calls 비교(jobs/baseline vs jobs/improved-vllm)에서, 같은 300초 예산 안에 여러 문항이 실제로 더 많은 반복을 수행했음을 확인: `advanced-json-to-rfc4180-csv-converter`(model_calls 1→6), `implement-depgraph-dependency-resolver`(9→16, tool_calls 11→17), `implement-lz77-file-compressor`(tool_calls 6→14), `python-sudoku-solver-backtracking`(5→15), `recover-encrypted-db-credentials`(5→20). 즉 "GPU가 반복 횟수를 늘려준다"는 가설의 앞부분은 실측으로 지지됨. 다만 늘어난 반복이 검증 통과로 이어지지는 않았고(`python-sudoku-solver-backtracking`은 오히려 checks_passed 5/22→2/22로 더 나빠짐), 실패 원인의 다수가 기준의 "시간 한도 도달"·"ValueError" 등 다양한 사유에서 `ReadTimeout`(개별 요청 응답 지연) 쪽으로 쏠림 - 대화가 길어질수록(반복이 늘수록) 각 요청의 컨텍스트도 커져 vLLM/ngrok 경유 개별 요청이 오래 걸리다 타임아웃됐을 가능성. `implement-nonogram-puzzle-solver` 1건은 실행 도중 ngrok 무료 세션이 끊기며 즉시 HTTPStatusError(model_calls=1) - 모델 능력과 무관한 인프라 실패로 별도 표시.
- 동일하게 유지한 조건 / 바뀐 조건: 문항 구성·순서·max_steps=40/max_seconds=300/command_timeout=10·모델 계열(Qwen3.5)은 기준과 동일. provider(Ollama→vLLM)·모델 크기(2b→4B AWQ 양자화)·실행 위치(로컬 CPU→Colab 원격 GPU, ngrok 경유)가 함께 바뀜 - 모델 크기까지 같이 바뀐 점, 그리고 무료 ngrok 터널의 자체 불안정성(세션 종료·응답 지연)이 섞여 들어간 점은 "동일 조건 재평가"의 한계로 남는다.
- 결론과 다음 실험: 이 실험에서 "느린 CPU가 병목"이라는 가설은 절반만 맞았다 - GPU로 바꾸자 실제로 더 많은 반복을 소화했지만(가설이 예측한 그대로), 그게 문제 해결로 이어지지 않아 **모델 자체의 문제 해결 능력(추론/코드 정확도)이 속도보다 더 근본적인 병목**이라는 결론에 무게가 실린다. 다음 실험 후보: (1) 같은 조건에서 더 크거나 다른 계열의 모델로 비교, (2) 요청 단위 타임아웃(현재 120초)을 늘려 ReadTimeout이 진짜 모델 지연 때문인지 네트워크(ngrok) 문제인지 분리, (3) ngrok 대신 안정적인 터널(고정 도메인)이나 로컬 GPU로 재현해 무료 터널 변수를 제거.

점수가 내려가거나 그대로여도 결과를 기록합니다. 이 10문항에서의 차이를 전체 성능 우월성으로 일반화하지 않습니다. 같은 문항을 반복 개선에 사용한 개발 실험이라는 한계를 설명합니다.

## 제출 확인
- [x] 실행 가능한 소스와 잠금 파일, 실행 안내 (harness_lab README + 기존 프로그램)
- [x] PRD·결정 기록·인터페이스·완료 조건 (harness-design-kit 폴더)
- [x] 기준의 설정, run-metadata.json, 원본 trial result와 실행 기록 (jobs/baseline)
- [x] 개선의 설정, run-metadata.json, 원본 trial result와 실행 기록 (jobs/improved-vllm - jobs/baseline-vllm은 무효 시도 기록으로 함께 보관)
- [x] 두 실행의 10문항 결과 CSV·JSON·HTML (jobs/baseline, jobs/improved-vllm)
- [x] 개선 가설, 구현 변경과 관찰 결과 (위 실패 분석 참고)
- [ ] 키·토큰·개인 자료 제외 (최종 제출 직전 재확인 필요 - 특히 ngrok URL, Colab 노트북 링크)

## 제출용 사본에서 제외한 것

- `jobs/{baseline,baseline-vllm,improved-vllm}/source/benchmark/upstream/` 전체(원본 문제의 `solution/solve.sh`, `tests/test_outputs.py` 등) — 실행 당시 harness_lab이 참조용으로 복사해 둔 원본 upstream 문제/정답/채점 코드라 제외. 재현 시 필요하면 `suite_id: terminal-bench-pro-local-port-v1`, `repo: alibaba/terminal-bench-pro`, `revision: 874af409da6aafebccbf3bc5bb41a2fa4d78784d`(harness-lab/benchmark/tasks.json 및 각 jobs/*/manifest.json에 문항별 원본 파일 SHA256 전체 기록됨)로 다시 내려받으면 된다.
- 각 시행 폴더의 `verifier/test_outputs.py`(원본 채점 테스트 코드 사본) — 같은 이유로 제외. `verifier/results.xml`(채점 결과)·`stdout.txt`·`stderr.txt`(실행 로그)는 실행 증거이므로 그대로 유지.
- 에이전트가 각 시행의 `workspace/` 안에 직접 작성한 파일(예: 일부 시행의 `solution.sh`라는 이름의 자체 시도 스크립트)은 원본 정답과 무관한 에이전트 실행 결과물이라 그대로 유지. upstream 원본 `solve.sh`와 diff로 대조해 다른 내용임을 확인함.
- `.venv/`, `__pycache__/`, `.pytest_cache/`, `.benchmark-cache/` — 재설치로 복원 가능한 환경 산출물이라 제외.

## 로컬 이식과 검증 범위

- 난이도 출처: 고정 upstream task.toml (easy 2 / medium 4 / hard 4)
- 실행 방식: local-port, Docker 미사용
- OS·CPU·메모리와 동시 실행 프로그램: Windows 11 Pro, i5-10210U(GPU 없음), 16GB RAM (기준 실행 환경). 개선 시도는 로컬 PC에서 harness_lab.bench를 실행하고 실제 추론은 Colab GPU 런타임에 ngrok 터널로 위임하는 구조 - [OPEN: 동시 실행 중이던 다른 프로그램 기록 필요]
- 원본 대비 변경: 작업 경로, 현재 Python 실행 파일, 설치 보일러플레이트 제외, 기대값 파일의 평가자 전용 이동 (manifest.json의 local_port_changes 참고 - recover-encrypted-db-credentials의 정답 파일을 평가자 전용 캐시로 분리)
- 채점기 제약 또는 원본 검사 오류가 결과에 미친 영향: manifest.json의 limitations 참고 - GoBoard 채점기는 mypy/flake8 실행파일이 필요하며 두 도구의 실패를 assert하지 않음(통과가 lint/typecheck 성공을 보장하지 않음). recover 과제의 setup은 os.urandom 기반 무작위 데이터를 매 실행마다 새로 생성함.
- 다른 운영체제에서 실제 실행했는지: 아니오(Windows에서만 실행)

공식 컨테이너 점수나 전체 벤치마크 점수로 표현하지 않습니다. 기준·개선의 소스 사본과 환경 조건을 함께 제출합니다.
