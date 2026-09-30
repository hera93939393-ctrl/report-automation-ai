"""ollama_client.py — 홈서버(Ollama)에 접속하기 위한 공용 클라이언트.

라우팅(routing_graph.py)이 쓰는 모델은 qwen3.5:9b, 호스트는 Tailscale
사설망을 통한 서버(desktop-o819qui, 100.74.107.35)다. 이 노트북에는
Ollama 자체가 설치되어 있지 않다 — LLM 호출은 전부 이 서버를 거친다
(2026-09-17 서버 구축 완료, 접속 테스트 확인됨)."""
import ollama

SERVER_HOST = "http://100.74.107.35:11434"
ROUTING_MODEL = "qwen3.5:9b"
# (2026-10-01) 문장 생성(polish_tool 등)도 같은 서버 모델을 쓴다. 모델명을
# 여기 한 곳에만 두는 이유: 나중에 더 큰 모델로 올릴 때 호출부를 뒤지지 않고
# 이 상수만 바꾸면 되게 하려는 것. 새 LLM 호출부는 반드시 이 파일을 거칠 것.
GENERATION_MODEL = "qwen3.5:9b"


def get_client(timeout: float | None = None) -> ollama.Client:
    """서버의 Ollama에 접속하는 클라이언트를 새로 만든다. timeout을 주면
    그 초만큼(요청 하나당) 기다리고, 안 주면 라이브러리 기본값을 쓴다."""
    if timeout is None:
        return ollama.Client(host=SERVER_HOST)
    return ollama.Client(host=SERVER_HOST, timeout=timeout)
