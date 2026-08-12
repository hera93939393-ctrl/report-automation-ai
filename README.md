# 보고서자동화 도구 — AI 문장 생성 확장 (Main Quest 2)

공공기관 계획(안)·보고서 작성 업무에 AI를 적용하는 프로젝트. 서식·표를 자동으로 채워주는 기존 도구에, 자연어(텍스트·음성)를 공문서 개조식 문장으로 바꾸고, 법령·이미지를 자동으로 붙이고, 최종 오류를 체크하는 로컬 AI 파이프라인을 추가했다.

- 문제 정의서: [문제정의서.md](./문제정의서.md)
- 상세 스펙(PRD, 시도·한계 기록 포함): [PRD.md](./PRD.md)

## 진행 상태

| 기능 | 내용 | 파일 | 상태 |
|---|---|---|---|
| F1 | 자연어 → 공문서 개조식 문장 생성 | [ai_writer.py](./ai_writer.py) | ✅ |
| F2 | 회의록 요약 (참석자/논의사항/결정사항) | [ai_writer.py](./ai_writer.py) | ✅ |
| F3 | 음성 인식(STT) | [stt.py](./stt.py) | ✅ |
| F4 | 매뉴얼 이미지 검색 (임베딩) | [image_search.py](./image_search.py) | ✅ |
| F5 | 법령 검색 (국가법령정보센터 API) | [law_search.py](./law_search.py) | ✅ |
| F6 | 엑셀 값 추출 → 문장화 | [excel_extractor.py](./excel_extractor.py) | ✅ |
| F7 | 최종 검토(항목번호·계산·맞춤법) | [format_checker.py](./format_checker.py), [ai_writer.py](./ai_writer.py) | ✅ |
| F8 | 홍보 포스터 이미지 생성 | [poster_generator.py](./poster_generator.py) | ✅ |
| 통합 GUI | F1~F8을 버튼/입력창으로 실행하는 도우미 창 | [assistant_panel.py](./assistant_panel.py) | ✅ |
| F9 | 한글 등 실제 문서에서 단축키로 바로 쓰기 | [hotkey_assistant.py](./hotkey_assistant.py) | ✅ (F1/F5/F7 검증, F2 준검증 — [PRD 참고](./PRD.md)) |

## 실행 방법

### 0. 사전 준비
1. [Ollama](https://ollama.com) 설치 (Windows: `winget install Ollama.Ollama`)
2. 모델 다운로드: `ollama pull exaone3.5:7.8b`
3. 파이썬 패키지 설치:
   ```bash
   pip install -r requirements.txt
   pip install --index-url https://download.pytorch.org/whl/cpu torch  # CPU 전용 빌드
   pip install pywin32  # 한글 서식 채우기 기능용
   ```
4. 한컴오피스(한글) 설치되어 있어야 함 (서식 채우기 기능용)
5. 법령 검색(F5)을 쓰려면 [open.law.go.kr](https://open.law.go.kr) 가입 후 발급받은 OC 값을 환경변수로 설정:
   ```powershell
   $env:LAW_API_OC = "발급받은OC값"
   ```

### 1. 서식 자동 채우기 (기존 v1)
`계획안_작성도구.bat` 더블클릭 → 입력창에 내용 입력 → [한글파일 만들기] 클릭. 자세한 사용법은 [사용법.txt](./사용법.txt) 참고.

### 2. 통합 도우미 창 (F1~F8을 버튼/텍스트로)
```bash
python assistant_panel.py
```
버튼을 눌러 각 기능을 실행하고, 결과를 복사해 문서에 붙여넣는 방식.

### 3. 단축키로 한글 등 문서에서 바로 쓰기 (F9)
`단축키도우미_시작.bat` 더블클릭 → 켜둔 채로 한글/워드/메모장 등에서 작업. 문장을 드래그로 선택한 뒤:

| 단축키 | 기능 |
|---|---|
| `Ctrl+Shift+G` | F1 문장을 공문서 개조식으로 변환 (원문 아래 줄에 결과 추가) |
| `Ctrl+Shift+M` | F2 선택한 녹취/텍스트를 회의록으로 요약 |
| `Ctrl+Shift+L` | F5 선택한 키워드로 관련 법령 검색 |
| `Ctrl+Shift+R` | F7 선택한(또는 전체) 문서를 검토 (알림창으로 표시, 문서는 안 바뀜) |

F3(음성)·F4(이미지검색)·F6(엑셀)·F8(포스터)은 파일 선택이 필요해 이번엔 단축키 대신 통합 도우미 창(2번)에서 사용한다.

### 4. 각 기능 개별 테스트
각 파일을 직접 실행하면 자체 테스트가 돌아간다 (`python ai_writer.py`, `python stt.py <음성파일>`, `python format_checker.py`, `python excel_extractor.py`, `python image_search.py`, `python law_search.py`, `python poster_generator.py`).

## 예시

```
[F1] 입력: 이번에 간담회 하는 이유는 주민들이 요즘 시끄럽다고 민원을 많이 넣어서 그거 들어보려고
     출력: 주민 민원 증가 및 소음 문제 청취를 위한 간담회 실시 예정

[F3+F2] 회의 음성 → STT → "참석자: 김주무관(예산 담당), 이준무관(장소 담당) / 논의사항: ..."

[F6] 엑셀(참가인원 342, 예산 1,850,000) → "참가인원 342명, 전월대비 증가율 12%, 예산 집행액 185만 원 확인됨"

[F7] "3) 장소: 회의실" → 번호 순서 오류(기대: 2, 실제: 3) 자동 검출
     "합계: 900,000" (실제 합 800,000) → 합계 불일치 자동 검출
```

## 모델 선정 이유

다루는 문서에 개인정보가 포함되어 외부 클라우드 AI(ChatGPT, Claude 등)를 사용할 수 없다. 이 제약으로 후보가 로컬 실행 가능한 오픈소스 모델로 좁혀졌다.

| 기능 | 모델 | 이유 |
|---|---|---|
| 텍스트 생성(F1,F2) | `exaone3.5:7.8b` (Ollama) | 한국어 학습 비중 높음, Ollama 정식 등록으로 받기 쉬움. Qwen2.5·HyperCLOVA X SEED와 비교 검토했으나 시간상 EXAONE으로 확정([PRD](./PRD.md#8-1-llm-모델-선정-테스트-계획-phase-0-세부)) |
| 음성인식(F3) | `faster-whisper` (small) | CPU 전용 노트북에 적합, 컴파일 불필요 |
| 이미지 검색(F4) | `google/siglip2-base-patch16-224` | 사진+문장 동시 학습, 한국어 질의 지원, 노트북급 크기 |
| 이미지 생성(F8) | `stabilityai/sd-turbo` | 1~2단계로 생성 가능해 CPU에서도 상대적으로 빠름 |

## 재현성(Reproducibility)에 대한 메모

이 프로젝트가 다루는 로컬 LLM·Diffusion 모델은 **다른 컴퓨터에서 실행해도 완전히 동일한 문장·이미지가 나오는 것을 보장하지 않는다** (같은 씨앗을 고정해도 하드웨어·라이브러리 버전에 따라 계산 결과가 달라질 수 있음 — 3강에서 다룬 내용과 같은 이유).

대신 다음을 보장한다:
- **같은 모델**: 코드에 모델 이름이 고정돼 있어 누가 실행해도 동일한 모델 가중치를 내려받는다.
- **같은 라이브러리 버전**: [requirements.txt](./requirements.txt)로 버전을 고정했다.
- **완전히 확정적인 부분은 100% 동일하게 재현됨**: 항목번호·계산 검증(F7), 엑셀 값 추출(F6), 법령 검색 결과(F5)는 AI가 아니라 규칙/API 기반이라 항상 같은 결과가 나온다.
- **AI 생성 부분(F1,F2,F3,F4,F8)은 "같은 과정으로 비슷한 품질의 결과"까지만 재현됨** — 이는 결함이 아니라 이 종류의 모델이 갖는 본질적 특성이다.

## 개인정보 처리 방침

- 모든 AI 처리는 노트북에서 완전히 로컬로 실행되며, 문서 내용을 외부로 전송하지 않는다.
- 유일한 외부 통신은 법령 검색(F5)의 검색 키워드 전송뿐이며, 문서 원본은 전송하지 않는다.
- 실제 업무 문서(개인정보 포함 가능)는 이 저장소에 포함하지 않는다.
