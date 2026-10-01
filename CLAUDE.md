# 보고서자동화 — 모든 AI 에이전트가 지킬 프로젝트 규칙

이 파일은 Claude Code뿐 아니라 Orca 등으로 붙는 어떤 코딩 에이전트도 먼저 읽어야 한다.
설계 전체는 `PRD.md`(특히 16절 하네스 재설계), 진행 기록은 `.tmp/implementation-plans/`.

## 1. 절대 규칙 — 실제 업무 문서를 열지 않는다
- 사용자의 실제 보고서(.hwp/.hwpx), 실적표(.xlsx), 회의록, 미리보기 이미지는 **읽지도 열지도 않는다.**
  에이전트가 읽는 내용은 외부 모델 서버로 전송되는데, 이 문서들에는 개인정보가 있어 외부 반출이 금지다.
  (이 프로젝트가 "완전 로컬 LLM"인 이유 그 자체.)
- 디버깅·재현·검증은 **테스트 코드가 직접 만드는 가짜 문서**로만 한다(`_make_probe_report`, `_test*.hwp` 등).
- `.gitignore`에 있는 실제 문서·키 파일(`law_api_oc.txt`, `archive_folders.json`, `archive_index.sqlite`)은 커밋 금지.
- 실제 문서로 확인이 필요하면 **사용자에게 직접 돌려 보고 결과만 말로 알려 달라고** 요청한다.

## 2. 한글(HWP) 자동화 규칙
- 한글 인스턴스는 반드시 `hwp_session.new_hwp(visible)`로 만든다(항상 새 프로세스 `new=True`).
  pyhwpx `Hwp()`를 직접 부르지 않는다 — 기본값은 사용자가 띄워 둔 한글 창을 붙잡아 버린다.
- 문서 열기는 `hwp_session.open_document(hwp, path, unattended=…)`(보안모듈 확인·무인 열기 옵션 포함).
- **COM 테스트는 절대 병렬로 돌리지 않는다.** 동시에 돌리면 서로 다른 프로세스의 문서 상태가 섞여 거짓 실패가 난다.
  검증은 `python run_selftests.py --pure`(한글·서버 불필요, 1~2분) → 한글 관련 수정이면 `--com`(순차, 수십 분) 순서로.
  여러 에이전트가 동시에 작업할 때는 **한 에이전트만** `--com`을 실행한다.
- 같은 프로세스에서 Hwp 인스턴스를 둘 이상 살려 두지 않는다(`Hwp.__del__`이 CoUninitialize로 남의 연결을 끊음).
  다른 문서를 읽어야 하면 `attachments.py`/`source_reader.py`처럼 **subprocess 격리**로.
- 읽기 부품은 커서 위치를 보존한다. 쓰기 뒤에는 문단 번호가 밀리므로 구간 id를 다시 읽는다(`doc_sections.py`).
- 편집 도구는 변경추적을 기본으로 켜지 않는다(켜면 삭제 원문이 남아 다음 읽기가 오염됨) — propose/apply/revert 방식.
- 사용자의 한글 창(테스트 시작 전부터 떠 있던 Hwp 프로세스)은 절대 종료하지 않는다. 테스트가 남긴 프로세스만 정리.
- 이 PC의 `tempfile.gettempdir()`은 `C:\Users\Public\Documents\ESTsoft\CreatorTemp`다(ESTsoft가 TEMP를 바꿔 둠).

## 3. LLM 규칙
- 모델·호스트는 `ollama_client.py` 한 곳에만(`SERVER_HOST`, `ROUTING_MODEL`, `GENERATION_MODEL`). 호출부에 모델명을 직접 쓰지 않는다.
- `qwen3.5:9b`는 사고형 모델이라 짧은 답/도구호출/구조화 출력에는 `think=False`를 반드시 준다(안 주면 빈 답).
- 서버(홈서버, Tailscale 사설망)가 꺼져 있을 수 있다. 연결 실패는 빈 응답과 **구분해서** 보고한다(`error="server_unreachable"`).
- 외부 클라우드 AI(OpenAI/Claude API 등)를 프로그램 런타임에 붙이지 않는다. 유일한 외부 통신은 법령 검색 키워드뿐.
- 새 LLM 기능은 가짜 클라이언트(`_FakeClient`)로 self-test를 쓰고, 실제 모델 품질은 사용자가 서버를 켠 뒤 확인한다.

## 4. 코드·테스트 관례
- 모듈마다 `if __name__ == "__main__":`에 self-test(`_selftest_*`)를 둔다. 순수 테스트를 먼저, COM 테스트를 뒤에.
- 테스트 픽스처에서 문서를 만든 `Hwp` 객체(setup)는 테스트가 끝날 때까지 참조를 살려 둔다.
- 도구 결과 계약: 실패는 예외 대신 `{"applied"/"ok": False, "reason": …}`로 정직하게. 문서는 자동 저장하지 않는다.
- 커밋 메시지는 한국어, 바뀐 이유와 실측 근거를 적는다. 실사용에서 걸린 버그는 평가셋(self-test)에 영구 편입한다.
- 파이썬 패치를 셸 heredoc으로 넘길 때 `\n` 같은 백슬래시가 풀릴 수 있다 — 소스에 이스케이프가 필요한 수정은 Edit/Write 도구로.
- 답변·문서·주석은 한국어.

## 5. 현재 로드맵(PRD 16-5)
①②③④ 완료, ⑤(루프→채팅 연결) 서버 검증만 남음, ⑥ 1차(BM25 아카이브 검색) 완료 — 임베딩 하이브리드·사업 카드·동료 대화 모드는 서버 필요.
브랜치 `feature/harness-step3-sections`(master 미병합). 결정 대기: 툴바 버튼 연결 방식(플랜 10절).
