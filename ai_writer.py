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


if __name__ == "__main__":
    samples = [
        "이번에 간담회 하는 이유는 주민들이 요즘 시끄럽다고 민원을 많이 넣어서 그거 들어보려고",
        "예산은 대략 200만원 정도 쓸 것 같아",
    ]
    for s in samples:
        print("입력:", s)
        print("출력:", to_gaejoshik(s))
        print("---")
