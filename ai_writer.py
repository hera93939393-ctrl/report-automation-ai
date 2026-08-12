# -*- coding: utf-8 -*-
"""
AI 문장 생성기 (F1)
- Ollama 로컬 LLM을 호출해 자연어 문장을 공문서 개조식으로 변환한다.
- 마크다운 등 불필요한 서식 기호를 제거해 순수 텍스트로 정리한다.
- 인터넷으로 아무것도 전송하지 않음 (localhost 로컬 호출만).
"""
import re
import requests

OLLAMA_URL = "http://localhost:11434/api/generate"
MODEL = "exaone3.5:7.8b"

INSTRUCTION = (
    "다음 문장을 한국 공공기관 보고서의 개조식으로 바꿔줘. "
    "개조식 규칙: 완전한 문장 종결어미(-습니다,-합니다)를 쓰지 말고 "
    "명사형이나 '-함','-음','-예정' 등으로 끝낼 것. "
    "하나의 목적/이유에 속한 짧은 요소들은 항목을 나누지 말고 '및'으로 이어서 한 줄로 쓸 것 "
    "(예: '주민 의견 수렴 및 소음 관련 민원 증가 대응'). "
    "일정/장소/대상처럼 서로 다른 항목일 때만 '- '로 줄을 나눠 쓸 것. "
    "마크다운 굵게(**) 쓰지 말 것. 결과만 출력해:"
)


def clean_markdown(text: str) -> str:
    """AI 출력에서 마크다운 서식 기호를 제거해 순수 텍스트로 만든다.
    지시문으로 막아도 가끔 새어나오므로(실제 관찰됨), 코드에서 무조건 제거한다."""
    text = re.sub(r"\*\*(.*?)\*\*", r"\1", text)   # **굵게**
    text = re.sub(r"__(.*?)__", r"\1", text)         # __굵게__
    text = re.sub(r"`(.*?)`", r"\1", text)           # `코드`
    text = re.sub(r"^#+\s*", "", text, flags=re.MULTILINE)  # # 헤더
    text = re.sub(r"\*+", "", text)                  # 짝 안 맞는 별표 전부 제거
    text = re.sub(r"[ \t]{2,}", " ", text)           # 중복 공백 정리
    return text.strip()


def to_gaejoshik(sentence: str) -> str:
    """자연어 문장을 공문서 개조식으로 변환한다."""
    body = {"model": MODEL, "prompt": f"{INSTRUCTION} {sentence}", "stream": False}
    res = requests.post(OLLAMA_URL, json=body, timeout=300)
    res.raise_for_status()
    raw = res.json()["response"]
    return clean_markdown(raw)


SPELLCHECK_PROMPT = (
    "다음 문장에서 맞춤법이나 띄어쓰기, 없는 단어(오탈자)가 있으면 찾아서 "
    "'틀린 표현 -> 올바른 표현' 형식으로만 나열해줘. 문제 없으면 '문제 없음'이라고만 답해:"
)


MEETING_SUMMARY_PROMPT = (
    "다음은 회의 녹취를 그대로 옮긴 글이야. 이걸 회의록 형식으로 정리해줘. "
    "'참석자', '논의사항', '결정사항' 세 항목으로 나누고, 각 항목은 개조식('- '로 시작, 명사형/'-함' 등으로 끝맺음)으로 써줘. "
    "녹취에 사람 이름이나 직책이 한 번이라도 언급되면 반드시 참석자로 나열해줘 — 이름이 하나도 없을 때만 '확인 필요'라고 써줘. "
    "참석자는 '성명 직위' 형태로만 쓰고 '님' 같은 존칭은 붙이지 마 (형식 예시: 'OOO 주무관' — 이 예시 자체는 실제 참석자가 아니니 절대 결과에 넣지 말고, 녹취에 실제로 등장한 이름만 써). "
    "논의사항은 각 안건마다 실제로 어떤 상태였는지(예: 확정/미확정, 완료/진행중) 구체적으로 반영하고, 일반적인 말로 뭉치지 마. "
    "녹취에 없는 내용은 지어내지 말 것. 마크다운 굵게(**) 쓰지 말 것. 결과만 출력해:"
)


STT_FLAG_PROMPT = (
    "다음은 음성인식(STT)으로 자동 변환된 텍스트야. 발음이 비슷해서 잘못 인식됐을 가능성이 있는 단어를 찾아줘 "
    "(예: 실제로 없는 직급/단어, 문맥상 어색한 단어). 원문을 고치지 말고, 의심되는 부분만 "
    "'의심표현 -> 교정후보' 형식으로, 한 줄에 하나씩 나열해줘. 없으면 '없음'이라고만 답해:"
)


def find_suspicious_spans(text: str):
    """STT 결과에서 의심되는 표현을 찾아 원문 속 위치와 함께 반환한다 (원문은 바꾸지 않음).
    반환: [{"start": 시작위치, "end": 끝위치, "원문": ..., "교정후보": ...}, ...]
    GUI에서는 이 위치 정보로 해당 구간을 빨간색으로 표시하면 된다."""
    body = {"model": MODEL, "prompt": f"{STT_FLAG_PROMPT} {text}", "stream": False}
    res = requests.post(OLLAMA_URL, json=body, timeout=300)
    res.raise_for_status()
    raw = clean_markdown(res.json()["response"])

    spans = []
    for line in raw.splitlines():
        line = line.strip().lstrip("- ").strip()
        if "->" not in line:
            continue
        original, suggestion = [p.strip() for p in line.split("->", 1)]
        pos = text.find(original)
        if pos == -1:
            continue  # 원문에 실제로 없는 표현이면(모델이 지어낸 것) 무시
        spans.append({
            "start": pos, "end": pos + len(original),
            "원문": original, "교정후보": suggestion,
        })
    return spans


def render_with_red_terminal(text: str, spans: list) -> str:
    """터미널에서 의심 구간을 빨간색으로 보여주는 데모용 함수 (ANSI 색상 코드 사용).
    실제 GUI(tkinter)에서는 spans의 start/end로 Text 위젯에 빨간 태그를 적용하면 된다."""
    spans = sorted(spans, key=lambda s: s["start"])
    result = ""
    cursor = 0
    for s in spans:
        result += text[cursor:s["start"]]
        result += f"\033[91m{text[s['start']:s['end']]}\033[0m(→{s['교정후보']}?)"
        cursor = s["end"]
    result += text[cursor:]
    return result


def summarize_meeting(transcript: str) -> str:
    """회의 녹취 전문(F3 STT 결과)을 참석자/논의사항/결정사항으로 정리한다 (F2, F1 재사용)."""
    body = {"model": MODEL, "prompt": f"{MEETING_SUMMARY_PROMPT} {transcript}", "stream": False}
    res = requests.post(OLLAMA_URL, json=body, timeout=300)
    res.raise_for_status()
    raw = res.json()["response"]
    return clean_markdown(raw)


def check_spelling_llm(text: str) -> str:
    """로컬 LLM으로 맞춤법을 검토한다 (100% 로컬, 외부 전송 없음).
    전문 맞춤법 검사기(부산대/네이버 등)는 외부 서버로 문장을 보내야 해서 쓸 수 없어 대체함.
    정확도는 전문 검사기보다 낮고, 사족(불필요한 스타일 제안)이 섞일 수 있음 — 참고용으로 사용."""
    body = {"model": MODEL, "prompt": f"{SPELLCHECK_PROMPT} {text}", "stream": False}
    res = requests.post(OLLAMA_URL, json=body, timeout=300)
    res.raise_for_status()
    raw = res.json()["response"]
    return clean_markdown(raw)


if __name__ == "__main__":
    samples = [
        "이번에 간담회 하는 이유는 주민들이 요즘 시끄럽다고 민원을 많이 넣어서 그거 들어보려고",
        "예산은 대략 200만원 정도 쓸 것 같아",
    ]
    for s in samples:
        print("입력:", s)
        print("출력:", to_gaejoshik(s))
        print("---")
