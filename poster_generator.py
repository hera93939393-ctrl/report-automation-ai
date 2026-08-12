# -*- coding: utf-8 -*-
"""
홍보 포스터 이미지 생성 (F8) — SD-Turbo (로컬 실행)
- 행사 정보(한국어) -> LLM이 영어 이미지 묘사 프롬프트로 변환 -> Diffusion이 포스터 이미지 생성
- 개인정보와 무관한 홍보용 이미지 생성 목적이라 사용 범위 제약 없음.
- 노트북(GPU 없음)에서 계산량이 가장 크므로, 빠른 SD-Turbo(1~4단계)를 사용.
"""
import torch
from diffusers import AutoPipelineForText2Image
from ai_writer import to_gaejoshik, OLLAMA_URL, MODEL, clean_markdown
import requests

SD_MODEL = "stabilityai/sd-turbo"

_pipe = None

IMAGE_PROMPT_INSTRUCTION = (
    "다음 행사 정보를 보고, 이미지 생성 AI에게 줄 영어 그림 묘사 프롬프트를 한 문장으로 만들어줘. "
    "포스터/플라이어 스타일, 밝고 친근한 느낌으로. 설명 없이 영어 프롬프트 문장만 출력해:"
)


def _load_pipeline():
    global _pipe
    if _pipe is None:
        _pipe = AutoPipelineForText2Image.from_pretrained(SD_MODEL, torch_dtype=torch.float32)
        _pipe.to("cpu")
    return _pipe


def to_image_prompt(event_info_korean: str) -> str:
    """한국어 행사 정보를 영어 이미지 생성 프롬프트로 바꾼다 (로컬 LLM 재사용)."""
    body = {"model": MODEL, "prompt": f"{IMAGE_PROMPT_INSTRUCTION} {event_info_korean}", "stream": False}
    res = requests.post(OLLAMA_URL, json=body, timeout=300)
    res.raise_for_status()
    return clean_markdown(res.json()["response"]).strip()


def generate_poster_image(prompt: str, output_path: str = "_테스트포스터.png"):
    """영어 프롬프트로 포스터 이미지를 생성해 저장한다."""
    pipe = _load_pipeline()
    image = pipe(prompt=prompt, num_inference_steps=2, guidance_scale=0.0).images[0]
    image.save(output_path)
    return output_path


if __name__ == "__main__":
    event_info = "주민 소음 민원 관련 간담회를 다음 주 화요일 오후 2시 주민센터 회의실에서 개최합니다."

    print("영어 이미지 프롬프트 생성 중...")
    prompt = to_image_prompt(event_info)
    print("프롬프트:", prompt)

    print("포스터 이미지 생성 중... (CPU라 시간이 걸릴 수 있음)")
    path = generate_poster_image(prompt)
    print("저장 완료:", path)
