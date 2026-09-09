# 구현 계획
## 시작 조건
D01~D09의 최초 범위와 계약·완료 조건을 정하고 학생이 구현을 요청했다. 첫 수직 구현은 기존 프로그램에서 작은 텍스트 파일을 읽고 질문에 답하는 흐름으로 시작한다.

## 1. 첫 수직 구현
사용자 요청 → 작은 텍스트 파일 읽기 → 실제 Ollama 모델에 파일 내용과 질문 전달 → 답변 표시를 기존 프로그램의 첨부·채팅 흐름에 연결한다. 파일은 `.txt`/`.md` 1개, 1 MiB 이하로 제한하며 자동 저장하지 않는다.
- 선택 작업/도구: `read_small_text_file` + 파일 질문 답변
- 필요한 파일·의존성과 실행 방법: 기존 Python 의존성과 Ollama 로컬 API를 재사용하고, `python chat_assistant.py`로 실행
- 실제 모델 사용 전 모의 응답으로 확인할 부분: 파일 읽기 성공/실패, 빈 파일, 지원하지 않는 확장자, 크기 초과, 모델 실패 시 상태 전이와 성공 오기록 방지
- 완료 증거: 민감하지 않은 작은 fixture로 파일의 사실을 묻고, 답변·읽기 단계·실패 기록을 ACCEPTANCE의 사용자 시나리오에 남김

## 2. 범용 작업의 품질
원문과 결과 대조, 인자 오류, 경로 이탈, 반복 한도를 검증한다. 실패를 보고 계약/코드를 수정한다.

## 3. 코딩 작업과 권한
실습용 작은 결함 코드/테스트를 준비한다. 변경 제안 → 내용 표시 → 승인/거절 → 적용 → 지정 테스트 → 결과를 연결한다. 거절 시 파일 불변도 확인한다.

## 4. 제품 기능
선택한 세션 정책과 상태 표시를 구현한다. 추가 provider, 영속 저장, 취소 등은 DECISIONS의 범위에 따라 작업을 추가하거나 후속으로 남긴다.

## 5. 고정 10문항 비교 실험
harness-lab의 benchmark/tasks.json과 로컬 평가 연결 방식을 확인한다. 기준 실행 후 실패 기록을 바탕으로 가설 하나를 선택하고 개선한 뒤 같은 10문항을 재측정한다. 모델·예산·반복 수·원본 버전을 유지하고 차이가 있으면 기록한다. 원본 결과·CSV·HTML과 구현 소스를 함께 제출한다. 채점기 진단과 모의 검사는 학생 성능 점수가 아니다.

## 각 단계의 기록
| 단계 | 관련 요구·검증 ID | 구현할 변경 | 확인한 결과 | 남은 문제 |
|---|---|---|---|---|
| 0(선행, 재설계 전) | D05 | `verify_numbers.categorize_values()`에 문맥/라벨 연결 + 4번째 판정 `ambiguous`(회색) 추가 | 커밋 74d0c2b. ACCEPTANCE.md U01~U05·U07 실제 fixture로 PASS(U07은 "타입 단위로 좁게 구현" 제한 있음, 후속 논의 필요로 명시) | U06(재검증 부분갱신)·U08(실패경로 기록) NOT_RUN |
| 0.5(선행, 진행중이던 것) | R03 첫 수직구현 | `file_answer.py`(read_small_text_file+Ollama/vLLM fallback 질문답변) 작성, `chat_assistant.py`에 하드코딩 분기로 임시 연결 | file_answer 자체 self-test 6개 통과. 단 route_intent()를 거치지 않는 임시 분기라 "모델이 도구를 해석"하는 진짜 하네스 구조가 아님 | 1단계에서 agent_loop 정식 도구로 재연결 필요 |
| 1a. 반복 루프 인프라 | R02 | `agent_loop.py` 신규 - Ollama 우선/vLLM 폴백 호출, 도구 요청 검사(필수인자)·실행·결과를 모델에 반환하는 반복(최대 8회), 반복한도/provider실패는 failed로 정직히 보고 | self-test 10개 통과(모의 모델로 정상/여러도구/인자누락/도구예외/미등록도구/반복한도/provider실패 전부 재현·확인). 실제 Ollama로 tool_calls 응답 구조(arguments가 이미 dict, id 없음)를 직접 호출해 실측 확인 | vLLM 쪽 도구 호출 응답 형식(OpenAI 표준 가정)은 미검증 - 실제 Colab vLLM으로 도구 호출 검증 필요(A07과 함께) |
| 1b. 코딩 시나리오 도구 | R04, D05-코딩 | `coding_fixture/split_cost.py`(결함)+`test_split_cost.py`, `code_fix_tool.py`(read_fixture_code/propose_code_fix는 모델 도구, apply_fix_and_test는 승인 자기결정 방지를 위해 의도적으로 모델 도구에서 제외 - INTERFACES.md 참고) | self-test 8개 통과. A03(승인→적용→PASS)·A04(거절→파일불변) 대응 테스트 통과 | chat_assistant.py에 아직 연결 안 됨(다음 단계) |
| 1c. chat_assistant.py 실제 연결 | R01~R05, D02, D06 | `route_intent()`(1회성) 삭제, `_on_submit()`이 `agent_loop.run_agent_loop()`로 9개 도구(기존 6 + read_small_text_file/read_fixture_code/propose_code_fix)를 실행. 도구별 결과 렌더링(`_handle_tool_result`)과 도구 dispatch 구성(`_build_tool_dispatch`)을 분리. 코드 승인/거절은 `parse_approval_response()`로 사용자 메시지만 보고 결정(모델 도구 아님, R05). 모델이 도구를 하나도 안 부르면 기존 `_route_by_keywords` 안전망 + "텍스트파일 1개면 file_answer.answer_file_question" 폴백 유지(F11 실측 도구호출 성공률 ~33% 대응) | self-test 전부(chat_assistant --selftest, agent_loop, code_fix_tool) 통과. **실제 Ollama로 9개 도구 스키마 전체를 놓고 "숫자 검증해줘" 라우팅 재확인 - 새 도구 3개 추가에도 verify_numbers를 정확히 선택함(trace=['verify_numbers'], status=completed)** | 코딩 시나리오(read_fixture_code→propose_code_fix→승인) 전체를 실제 Ollama로 엔드투엔드 실행한 기록은 아직 없음(부품별로는 각각 실측 확인됨 - agent_loop의 실제 tool_calls 구조, code_fix_tool의 실제 적용/거절). vLLM 폴백 경로도 여전히 미검증(A07과 함께 후속) |

설계 변경은 PRD와 인터페이스, 완료 조건에도 반영한다. 코드가 먼저 바뀌어 명세와 충돌했다면 어느 쪽이 사용자 의도에 맞는지 판단한다.
