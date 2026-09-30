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
