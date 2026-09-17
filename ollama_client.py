"""ollama_client.py — 홈서버(Ollama)에 접속하기 위한 공용 클라이언트.

라우팅(routing_graph.py)이 쓰는 모델은 qwen3.5:9b, 호스트는 Tailscale
사설망을 통한 서버(desktop-o819qui, 100.74.107.35)다. 이 노트북에는
Ollama 자체가 설치되어 있지 않다 — LLM 호출은 전부 이 서버를 거친다
(2026-09-17 서버 구축 완료, 접속 테스트 확인됨)."""
import ollama

SERVER_HOST = "http://100.74.107.35:11434"
ROUTING_MODEL = "qwen3.5:9b"


def get_client() -> ollama.Client:
    """서버의 Ollama에 접속하는 클라이언트를 새로 만든다."""
    return ollama.Client(host=SERVER_HOST)
