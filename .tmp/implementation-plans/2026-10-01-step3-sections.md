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
