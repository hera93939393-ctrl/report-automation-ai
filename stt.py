# -*- coding: utf-8 -*-
"""
음성 인식 (F3) — faster-whisper (로컬 실행, 외부 전송 없음)
- 음성 메모/회의 녹음 파일을 텍스트로 변환한다.
- 모델은 최초 실행 시 한 번만 다운로드되며(공개 모델 파일), 이후 인식 처리는 완전히 로컬에서 이루어진다.
"""
from faster_whisper import WhisperModel

MODEL_SIZE = "small"  # 노트북(CPU, GPU 없음) 기준 속도/정확도 균형점

_model = None


def _get_model():
    global _model
    if _model is None:
        _model = WhisperModel(MODEL_SIZE, device="cpu", compute_type="int8")
    return _model


def transcribe(audio_path: str, language: str = "ko") -> str:
    """음성 파일을 텍스트로 변환한다."""
    model = _get_model()
    segments, info = model.transcribe(audio_path, language=language)
    return " ".join(seg.text.strip() for seg in segments)


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("사용법: python stt.py <음성파일경로>")
    else:
        text = transcribe(sys.argv[1])
        print("인식 결과:", text)
