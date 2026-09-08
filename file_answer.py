"""Small text-file reader and question-answering flow for the first vertical slice."""
import os

import ollama
import requests


MAX_FILE_BYTES = 1024 * 1024
SUPPORTED_EXTENSIONS = {".txt", ".md"}
DEFAULT_MODEL = "qwen3.5:2b"
VLLM_BASE_URL = os.getenv(
    "VLLM_BASE_URL",
    "https://surcharge-condense-bootie.ngrok-free.dev/v1",
).rstrip("/")
VLLM_MODEL = os.getenv("VLLM_MODEL", "cyankiwi/Qwen3.5-4B-AWQ-4bit")


def read_small_text_file(path: str) -> dict[str, str]:
    """Read one bounded text fixture without modifying or persisting it."""
    extension = os.path.splitext(path)[1].lower()
    if extension not in SUPPORTED_EXTENSIONS:
        raise ValueError("지원하지 않는 파일 형식입니다. .txt 또는 .md만 사용할 수 있어요.")
    if not os.path.isfile(path):
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {path}")
    if os.path.getsize(path) > MAX_FILE_BYTES:
        raise ValueError("파일 크기가 1 MiB를 초과합니다.")

    with open(path, "r", encoding="utf-8") as handle:
        text = handle.read()
    return {"file_name": os.path.basename(path), "text": text}


def build_question_prompt(file_text: str, question: str) -> str:
    return (
        "다음 파일 내용만 근거로 질문에 답해줘. "
        "파일에 없는 내용은 추측하지 말고 확인 필요라고 답해.\n\n"
        f"[파일 내용]\n{file_text}\n\n"
        f"[질문]\n{question}"
    )


def answer_file_question(path: str, question: str, model: str = DEFAULT_MODEL) -> str:
    """Read a small file and ask Ollama, falling back to OpenAI-compatible vLLM."""
    fixture = read_small_text_file(path)
    messages = [{"role": "user", "content": build_question_prompt(fixture["text"], question)}]

    ollama_error = None
    try:
        response = ollama.chat(model=model, messages=messages)
        answer = response.get("message", {}).get("content", "").strip()
        if answer:
            return answer
        ollama_error = RuntimeError("Ollama 응답이 비어 있습니다.")
    except Exception as error:
        ollama_error = error

    try:
        response = requests.post(
            f"{VLLM_BASE_URL}/chat/completions",
            json={
                "model": VLLM_MODEL,
                "messages": messages,
                "stream": False,
                "chat_template_kwargs": {"enable_thinking": False},
            },
            headers={"ngrok-skip-browser-warning": "1"},
            timeout=120,
        )
        response.raise_for_status()
        answer = response.json().get("choices", [{}])[0].get("message", {}).get("content", "").strip()
        if answer:
            return answer
        raise RuntimeError("vLLM 응답이 비어 있습니다.")
    except Exception as vllm_error:
        raise RuntimeError(
            f"Ollama와 vLLM 모두 답변하지 못했습니다. "
            f"Ollama: {ollama_error}; vLLM: {vllm_error}"
        ) from vllm_error