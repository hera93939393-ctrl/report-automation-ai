# PRD 16-5 ③단계 — 부품 1·2·3 + "구간 편집" 도구 (2026-10-01 야간 자율 작업)

브랜치: `feature/harness-step3-sections` (master 미병합 — 아침에 확인 후 병합)

## 1. 만든 것

| 파일 | 내용 | 테스트 |
|---|---|---|
| `doc_sections.py` | 부품 1 `read_outline`, 2 `read_section`, 3 `write_section` + `section_at_cursor`, `format_outline` | 순수 함수 1건 + 실제 한글 2건(읽기, 4가지 쓰기 모드·표 보호·id 갱신·커서 보존) 통과 |
| `section_edit_tool.py` | `propose_section_edit`(LLM 제안, 문서 안 건드림) / `apply_section_edit` / `revert_section_edit` | 가짜 LLM 2건 + 실제 한글 1건(제안→적용→되돌리기→재적용, 선택 경로, 미접속, 표 안 거부) 통과 |
| `routing_graph.py` | 도구 `edit_section` 등록 + 키워드 안전망(`_EDIT_SECTION_KEYWORDS`), 우선순위 조정 | 기존 케이스 + 신규 5건 통과 |
| `chat_assistant.py` | `_show_section_edit_card`: 전/후 카드 + [적용][취소], 적용 뒤 [되돌리기] | Tk 숨김창 + 실제 한글 + 가짜 LLM 스모크(적용/되돌리기/취소/대상없음/서버미접속) 통과 |

라우팅 우선순위(키워드 안전망): 검증 > 문체 지목(공문서/다듬어/매끄럽게/격식있게) > **구간 편집**(줄여/늘려/보완/반영/추가해/수정해/다시 써/요약해/구체적으로/간결하게/이 문단/이 구간/이 부분…) > 일반 작성(써줘/작성해/바꿔줘/고쳐줘…→공문서체) > 표 > 그래프 > 번호 > 주간보고 > 쪽맞춤.

## 2. 실측으로 확인한 pyhwpx 동작(설계 근거, doc_sections.py docstring에도 기록)

- 본문 list 0 문단은 `set_pos(0, i, 0)` 실패까지 순회, `select_text(i,0,i,-1,0)`+`get_selected_text()`로 읽기. 표가 앵커된 문단은 셀 내용이 `\r\n`으로 이어진 한 문단.
- 표 `GetAnchorPos(0)` = (0, para, 0). 다중 문단 교체(`select_text(s,0,e,-1,0)`+`insert_text`) 뒤 뒤쪽 문단 번호·표 앵커가 당겨짐 → **쓰기 뒤 구간 id 무효, write_section이 새 outline 반환**.
- `set_pos(0,i,-1)`+`BreakPara()`+`insert_text`로 문단 뒤 끼워넣기.
- **읽기 부품은 커서를 보존해야 함** — 처음엔 read_outline이 커서를 문서 처음으로 옮겨 "커서 구간 없음"이 났음(수정: `_restore_pos`).
- **변경추적을 켠 채 교체하면** 삭제 원문이 문단에 남아 다음 읽기가 "원문+수정문"을 읽고 `get_text()`는 승인 전 빈 문자열 → 편집 도구는 추적을 기본으로 켜지 않고 자체 되돌리기(`revert_section_edit`) 제공. 필요 시 `track_changes=True`.
- 테스트 픽스처: 문서를 만든 `Hwp`(setup) 참조를 테스트 끝까지 살려 둬야 함(`Hwp.__del__`의 CoUninitialize가 HwpReport 연결을 끊어 "개체가 열려 있지 않거나…" 재현).

## 3. 아침에 확인할 것(서버 켠 뒤)

1. 실제 `qwen3.5:9b`로 `python section_edit_tool.py`는 그대로(가짜 LLM) 통과하고, **실사용은 채팅창에서**: 보고서(테스트용 가짜 문서 권장) 열고 → 문단에 커서 → "이 문단 좀 줄여줘" → 카드에서 [적용]/[되돌리기].
2. 9B가 규칙(⑤ 제목 줄 다시 쓰지 않기, ④ 사실 지어내지 않기)을 얼마나 지키는지 3~5개 지시로 확인. 자주 어기면 `_SYSTEM_PROMPT`에 예시(few-shot) 추가.
3. 라우팅: LLM 도구호출이 `edit_section`을 고르는지(서버 켜진 상태) — 키워드 없이 "요 문단 반쯤으로" 같은 표현.

## 4. 알려진 한계 / 다음 단계로 넘긴 것

- 구간 안에 표가 있으면 통째 편집 거부(문단 선택 요청 안내). 표 편집은 별도 부품.
- 항목기호 없는 문서는 문단별 구간(폴백) — 커서 문단만 편집 대상.
- 글머리 기호 계열(□ ○ - ·)과 편람 8단계를 한 순서로 섞어 단계 판정 — 문서에 따라 어긋날 수 있음(같은 기호=같은 단계만 보장).
- 선택 텍스트 경로는 polish_tool과 같은 `find()` 재선택 방식이라 같은 문장이 여러 번 나오면 첫 occurrence를 고칠 수 있음.
- ④단계(부품 4·6 재포장, 첨부문서 즉석 QA), ⑤단계(루프 그래프·승인 interrupt·스레드/스트리밍)는 미착수.

## 5. 같은 밤에 이어서 한 ④단계 (같은 브랜치)

| 파일 | 내용 | 테스트 |
|---|---|---|
| `attachments.py` | 부품 4 `read_attachment(s)`: hwp/hwpx(별도 프로세스, 무인 열기), xlsx(전 시트, 시트당 500행), pdf(쪽), txt/md/csv → `parts[{label,text}]` + `"[파일명 · 위치]"` 머리말 텍스트, 글자 예산 초과 시 잘림 표시, 폴더 펼치기·중복 제거, 실패 이유 문자열(OneDrive 온디맨드 안내 포함) | text/xlsx/pdf/오류/잘림/폴더 1건 + hwp 격리(살아있는 Hwp 인스턴스 보존 확인) 1건 통과 |
| `attachment_qa_tool.py` | `ask_attachments(paths, question)`: 첨부만 근거, 출처 인용 강제, 없으면 "찾지 못했습니다", `num_ctx=12288` 명시, 예산(14,000자) 초과 시 질문과 단어 겹침 점수로 부분 선택 | select_context 1건 + 경로 5종 1건 통과 |
| `routing_graph.py` | 도구 `ask_attachment` + 키워드(첨부에서/회의결과에/몇 건이/얼마라고…), 우선순위 편집 > 첨부QA > 일반 작성 | 5건 추가 통과 |
| `chat_assistant.py` | `_answer_from_attachments`: 첨부 없으면 안내, 답변 + (잘림/못 읽은 파일) 안내 | Tk 스모크(첨부없음/정상/못읽은 파일) 통과 |

부품 6(polish 재포장)은 ⑥단계에서 유사 문단 예시(few-shot)가 생길 때 같이 하는 게 맞아 미룸(지금 바꿀 실익 없음).

아침 확인 추가: 서버 켠 뒤 채팅에서 "+"로 가짜 회의결과(.txt나 .hwp)를 첨부하고 "첨부에서 … 찾아줘"로 실제 9B 답변·출처 형식 확인. `num_ctx=12288`이 서버 GPU에서 무리 없는지(Ollama 로그/응답 속도) 확인 — 무리면 `attachment_qa_tool._NUM_CTX`를 8192로.

## 6. ⑤단계 준비 — `agent_loop.py` (같은 밤, 채팅 미연결)

PRD 16-3 그래프를 랭그래프 1.2로 그대로 구현한 뼈대. planner/tools/context_provider/apply_write를 주입하는 구조라 서버·한글 없이 가짜로 검증했다(4건 통과):
- interrupt로 멈춤(`{"type":"approve_write", before, after, unverified_numbers}`) → `resume(approved=True)`로 같은 자리에서 재개 → apply 콜백 실행 → 출처 붙은 답변. 승인 전엔 문서(가짜 dict) 불변.
- 거부 시 문서 유지 + 제안문 속 "근거 없는 숫자"(원문·문맥·이전 도구 결과에 없는 숫자) 표시.
- `answer` 도구로 조기 종료, 도구 호출 상한(`MAX_TOOL_CALLS`) 도달 시 안내 후 종료.
- 같은 도구 연속 2회 실패 → 재계획 1회 → 또 실패 시 되묻기 종료.
- 분기 함수는 순수 판정만(상태 변경 금지), 안내문은 answer 노드가 붙임(랭그래프 관례).
- `OllamaPlanner`(계획: `format="json"`으로 `{"steps": [...]}`, 도구 선택: 정식 tool-calling, think=False)는 **서버 켠 뒤 검증 대상** — 요청문 10개로 계획 JSON 일관성·도구호출 성공률을 재고, 흔들리면 레시피 기반 고정 계획으로 대체(16-3 설계대로).
- 채팅 연결 시 할 일: 승인 카드에서 `resume()` 호출, LLM/그래프 실행은 스레드로, HWP 도구 호출은 메인 스레드(COM 아파트)로 마샬링 — 이 부분이 남은 사전 확인 항목.

## 7. ⑤단계 — 루프를 채팅에 연결 (서버 꺼진 상태에서 할 수 있는 부분 완료)

- `loop_runner.py`: `MainThreadBridge`(작업 스레드→메인 스레드 큐, Tk `after(50)`로 pump; 메인 스레드에서 부르면 즉시 실행해 교착 방지), `build_tools`(read_outline/read_section/write_section(제안만)/read_attachment/polish — 한글 호출은 전부 bridge.run), `build_context_provider`(선택 텍스트·커서 구간·목차·첨부 목록), `build_apply_write`, `LoopSession`(stream_mode="updates"로 노드별 진행 문구, `__interrupt__` 청크 → 승인 콜백, `resume()`은 `Command(resume=)`).
- `chat_assistant.py`: 라우팅에 안 걸린 요청은 서버가 켜져 있으면 `_start_loop()`로(꺼져 있으면 종전 되묻기). 진행 말풍선 한 개를 계속 갱신, 승인 카드 `_show_loop_approval_card`([적용][취소] → resume), 완료/오류 시 입력창 복구. `LOOP_PLANNER_FACTORY`로 계획자 교체 가능(테스트는 가짜).
- 검증: loop_runner self-test 2건(다리, 세션) + 통합 스모크(숨김 Tk + 실제 한글 + 가짜 계획자: 계획자는 작업 스레드에서, COM은 메인 스레드에서 실행됨을 확인 / 승인→반영→완료 답변 / 취소 + 근거 없는 숫자 경고 / 계획자 예외 → 복구) 통과.
- 설계상 선택: 토큰 스트리밍 대신 "계획: A → B", "도구 실행: read_section" 같은 노드 진행 문구로 '멈춘 게 아님'을 보여준다. 토큰 스트리밍은 polish 같은 단일 생성 호출에만 의미가 있어 후순위.
- 서버 켠 뒤 확인: `OllamaPlanner.plan`(format=json)이 `{"steps": [...]}`를 일관되게 내는지 요청문 10개로, `choose`의 도구호출 성공률(F11 때 2b는 33%였음 — 9b는 라우팅에서 이미 정식 지원 확인). 흔들리면 16-3대로 레시피 기반 고정 계획으로 전환.

## 8. ⑥단계 1차 — 과거 문서 아카이브 검색(서버 없이 되는 부분)

- `archive_index.py`: 폴더 등록(`archive_folders.json`, gitignore) → 부품 4로 읽어 700자 덩어리(문단 경계 존중, 긴 문단은 120자 겹침)로 SQLite(`archive_index.sqlite`, gitignore)에 저장 → 한글 2글자 n-gram + 영숫자 토큰 BM25 검색(`search_archive(query, k, year, name_contains)`). 증분 색인(mtime), 삭제 파일 정리, 읽기 실패 기록, 파일명에서 연도 추정. CLI: `--add-folder / --reindex / --search / --stats`.
- `archive_qa_tool.py`: `ask_archive` = 증분 색인 → 검색 → `attachment_qa_tool.answer_from_context`(출처 강제, 없으면 없다고). 폴더 미등록/결과 없음/미접속 구분.
- 라우팅: `ask_archive`(작년·재작년·지난해·예전·과거 문서·N년 파일에서·당시…), "첨부" 명시 시 첨부 QA 우선, 편집 키워드가 있으면 편집 우선. **검증 키워드 "점검"이 "현장점검/정기점검"에 오탐되는 문제 발견 → denylist 추가**(이 사용자 업무에선 사업명이라 중요).
- 채팅: `_answer_from_archive`가 작업 스레드에서 색인·검색·LLM을 돌리고 진행 문구("색인 중: 파일명")를 보여줌. 루프 도구 레지스트리에도 `search_archive` 추가.
- 검증: archive_index self-test(토큰화·청킹·증분·연도 필터·수정·삭제·실패), archive_qa self-test, 라우팅, 채팅 스모크(작업 스레드 실행·성공/실패 경로), 루프 스모크 회귀 통과.
- 한계·다음: 키워드 검색이라 동의어("제재"↔"행정처분")는 못 찾음 → 서버 임베딩 점수 하이브리드로. 사업 카드(사업별 요약, 사용자가 고칠 수 있는 파일)와 "동료" 대화 프롬프트는 서버 LLM 필요. .hwp 첫 색인은 파일당 수 초(새 한글 프로세스) — 아침에 실제 폴더로 체감 시간 확인.

## 9. 회귀 실행기 `run_selftests.py` (평가셋 1단계)

- `python run_selftests.py --pure`(11개 모듈, 약 1~2분) / `--com`(14개, 순차·격리, 수십 분) / 전체. server 분류(polish_tool)는 서버 응답 시만.
- 모듈 하나가 남긴 한글 프로세스만 정리하고 실행 전부터 떠 있던 창은 보존. 로그는 `.tmp/selftest-runs/`.
- 다음 단계(평가셋 2단계): "실사용에서 걸린 케이스는 무조건 평가셋행" 규칙대로, 시나리오 폴더 + 정답 JSON + 채점 스크립트(`D:
eport-verify-agent`의 `eval/` 방식)로 확장. LLM이 끼는 시나리오는 서버가 있어야 하므로 그때.

## 10. 결정이 필요한 것 — 툴바 버튼 → 패널 연결(원래 표의 4번)

지금 구조는 채팅 앱이 **자기 한글 프로세스(new=True)로 문서를 연다**(사용자가 따로 띄운 한글 창을 자동화가 붙잡는 사고를 막기 위한 원칙). 반면 F10 툴바 버튼(매크로)은 **사용자가 띄운 한글 창 안에서** 실행된다. 즉 둘은 다른 프로세스·다른 문서 핸들이라, 버튼이 선택 텍스트를 채팅 앱에 넘겨도 채팅 앱은 그 문서를 고칠 수 없다. 선택지:
1. **채팅 앱이 사용자의 한글 창에 붙는 모드 추가**(new=False로 실행 중 인스턴스에 연결). F9 때 겪은 COM 재연결 제약(PRD 12-3)과 "다른 창을 붙잡는" 위험을 다시 감수해야 함. 대신 "한글에서 열고 버튼 누르면 바로"가 됨(사용자 선호 방식).
2. **지금처럼 앱이 문서를 열고, 버튼은 "채팅창 앞으로 가져오기 + 선택 텍스트 입력창에 채우기"만** 담당. 안전하지만 두 창의 문서가 같은지 사용자가 신경 써야 함.
3. 버튼을 포기하고 채팅창의 "보고서 파일 선택"을 유지(현상 유지).
→ **2번으로 결정·구현(2026-10-01)**. 근거: 앱이 도는 동안 업무와 무관한 한글 창이 여러 개 떠 있는 게 실제로 관찰돼, 1번은 엉뚱한 창에 붙을 위험이 큼.
- `chat_handoff.py`: 채팅 앱이 시작할 때 창 핸들을 `%LOCALAPPDATA%\HangulButler\`에 등록 → 매크로가 실행한 CLI가 받은편지함에 선택 텍스트(+문서 경로)를 쓰고 채팅창을 앞으로 → 앱이 0.5초마다 받은편지함을 비워 입력창에 채움(바로 실행은 안 함). TEMP 대신 LOCALAPPDATA인 이유: ESTsoft가 TEMP를 바꿔 둬서 한글이 띄운 파이썬과 앱이 다른 TEMP를 볼 수 있음. 포그라운드 전환은 앱이 아니라 CLI가 함(포그라운드 프로세스가 실행한 프로세스만 SetForegroundWindow 허용).
- 문서가 다르면(매크로가 넘긴 경로 ≠ 앱이 연 문서) 채팅에 경고, 앱이 문서를 안 열었으면 열라고 안내, 작업 중이면 다시 누르라고 안내.
- `hwp_macro_ai.js`에 `OnScriptMacro_한글집사에게물어보기` 추가. 문서 경로는 `Path` 전역을 try로 읽음(못 읽는 버전이면 빈 값 → 비교 생략).
- 검증: chat_handoff self-test 6개(pure 회귀에 편입), 실제 앱 + CLI 왕복(가짜 텍스트, 받은편지함이 1.5초 안에 비워짐).
- **남은 것(사용자 수동)**: 한글에서 매크로 등록(Alt+Shift+H → 이름 → Alt+Shift+L 코드 편집에 붙여넣기) 후 툴바 아이콘 추가, 실제 버튼으로 확인. `Path` 전역이 읽히는지(문서 다를 때 경고가 뜨는지)도 그때 확인.

## 11. 전체 회귀 결과(2026-10-01 오후) 및 알려진 간헐 실패

- `run_selftests.py --pure` 11/11 통과, `--com` 14/14 통과(세션 재시작으로 두 번에 나눠 실행). server 분류(polish_tool)는 서버 꺼져 있어 SKIP.
- `verify_tool`의 두 테스트(`_selftest_run_verification_llm_rescues_paraphrased_row`, `_…_llm_falls_back_to_ambiguous_when_unclear`)는 실제 서버 LLM이 있어야 의미가 있어, 서버가 꺼져 있으면 **건너뜀**으로 바꿈(결과의 `llm_skipped_for_server`로 판정). 서버를 켜고 다시 돌리면 본검증이 된다.
- **간헐 실패 1건 관찰**: `_selftest_run_verification_connects_by_label_and_marks_ambiguous_gray`가 세 번 중 한 번 "총 0건 확인"(문서 텍스트가 비어 읽힘)으로 실패했고, 그대로 재실행하면 통과. 코드 변경과 무관(첫 실행 통과). 원인 후보: pyhwpx `Hwp.__del__`이 무조건 `CoUninitialize()`를 부르는데 verify_tool의 테스트들은 사이에 대기 없이 연달아 새 인스턴스를 만든다(hwp_report.py는 같은 이유로 테스트 사이 2초 대기). 다음에 또 보이면 verify_tool `__main__`에도 테스트 간 `time.sleep(2)`를 넣고, 픽스처의 setup/HwpReport 수명을 점검할 것. "실사용에서 걸린 건 평가셋행" 규칙상 재현 조건이 잡히면 전용 테스트로 고정.
