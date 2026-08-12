# -*- coding: utf-8 -*-
"""
매뉴얼 이미지 검색 (F4) — SigLIP2 임베딩 검색 (4강에서 배운 기법과 동일한 방식)
- 사진 폴더를 미리 색인(임베딩+정규화)해서 index.npz에 저장한다.
- 검색 시 질의 문장을 같은 모델로 임베딩하고, 색인과 내적을 계산해 상위 후보를 반환한다.
- 모든 처리는 노트북에서 로컬로 이루어지며, 이미지·문장을 외부로 전송하지 않는다.
- 벡터DB 없이 배열 하나 + 내적으로 처리 (사진 수천 장 규모 기준, 4강 근거).
"""
import os
import numpy as np
import torch
from PIL import Image
from transformers import AutoProcessor, AutoModel

MODEL_NAME = "google/siglip2-base-patch16-224"

_model = None
_processor = None


def _load_model():
    global _model, _processor
    if _model is None:
        _processor = AutoProcessor.from_pretrained(MODEL_NAME)
        _model = AutoModel.from_pretrained(MODEL_NAME)
        _model.eval()
    return _model, _processor


def _normalize(vec: np.ndarray) -> np.ndarray:
    """임베딩 길이를 1로 맞춘다 (내적이 곧 코사인 유사도가 되도록)."""
    return vec / np.linalg.norm(vec, axis=-1, keepdims=True)


def build_index(image_dir: str, index_path: str = "manual_image_index.npz"):
    """이미지 폴더 전체를 임베딩으로 바꿔 index_path에 저장한다. (사진이 늘 때만 다시 실행)"""
    model, processor = _load_model()
    filenames = [f for f in os.listdir(image_dir) if f.lower().endswith((".png", ".jpg", ".jpeg"))]

    embeddings = []
    for fname in filenames:
        image = Image.open(os.path.join(image_dir, fname)).convert("RGB")
        inputs = processor(images=image, return_tensors="pt")
        with torch.no_grad():
            out = model.get_image_features(**inputs)
            feat = (out.pooler_output if hasattr(out, "pooler_output") else out).numpy()[0]
        embeddings.append(_normalize(feat))

    np.savez(index_path, embeddings=np.array(embeddings), filenames=np.array(filenames))
    return len(filenames)


def search(query: str, index_path: str = "manual_image_index.npz", top_k: int = 5):
    """질의 문장으로 색인된 이미지 중 상위 top_k를 찾아 반환한다.
    반환: [{"파일명": ..., "유사도": ...}, ...] (유사도 높은 순)
    """
    model, processor = _load_model()
    data = np.load(index_path)
    embeddings, filenames = data["embeddings"], data["filenames"]

    inputs = processor(text=[query], return_tensors="pt", padding="max_length")
    with torch.no_grad():
        out = model.get_text_features(**inputs)
        text_feat = (out.pooler_output if hasattr(out, "pooler_output") else out).numpy()[0]
    text_feat = _normalize(text_feat)

    scores = embeddings @ text_feat  # 정규화된 벡터끼리의 내적 = 코사인 유사도
    order = np.argsort(-scores)[:top_k]

    return [{"파일명": str(filenames[i]), "유사도": float(scores[i])} for i in order]


if __name__ == "__main__":
    from PIL import ImageDraw

    test_dir = "_테스트매뉴얼이미지"
    os.makedirs(test_dir, exist_ok=True)

    # 테스트용 간단한 이미지 3장 생성 (실제 매뉴얼 이미지 대신)
    specs = [
        ("소화기.png", (220, 30, 30), "FIRE"),
        ("안전모.png", (230, 200, 30), "HELMET"),
        ("배전반.png", (30, 100, 220), "PANEL"),
    ]
    for fname, color, label in specs:
        img = Image.new("RGB", (224, 224), color)
        draw = ImageDraw.Draw(img)
        draw.text((60, 100), label, fill=(255, 255, 255))
        img.save(os.path.join(test_dir, fname))

    print("색인 생성 중...")
    n = build_index(test_dir)
    print(f"{n}장 색인 완료")

    for q in ["소화기 위치 안내", "안전모 착용 안내", "존재하지 않는 물건"]:
        print(f"\n질의: {q}")
        for r in search(q, top_k=3):
            print(" ", r)
