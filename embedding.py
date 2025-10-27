# -*- coding: utf-8 -*-
import sys
from pathlib import Path
import json
import numpy as np
import torch
from transformers import AutoTokenizer, AutoModel

# --------------------------------------------------------
# 0. 경로 설정
# --------------------------------------------------------
sys.path.append(str(Path(__file__).resolve().parents[1]))  # src
sys.path.append(str(Path(__file__).resolve().parents[2]))  # project

BASE_DIR = Path(__file__).resolve().parents[2]
PARSED_DIR = BASE_DIR / "parsed"
EMBED_DIR = BASE_DIR / "embeddings"
EMBED_DIR.mkdir(exist_ok=True)

MODEL_NAME = "BAAI/bge-m3"
tokenizer = AutoTokenizer.from_pretrained(MODEL_NAME)
model = AutoModel.from_pretrained(MODEL_NAME)

# --------------------------------------------------------
# 1. 텍스트 → 임베딩 변환 함수
# --------------------------------------------------------
def get_embedding(text: str):
    inputs = tokenizer(text, return_tensors="pt", truncation=True, padding=True)
    with torch.no_grad():
        embeddings = model(**inputs).last_hidden_state[:, 0, :]  # CLS 토큰
    return embeddings[0].cpu().numpy().astype("float32")

# --------------------------------------------------------
# 2. 임베딩 실행 함수 (main.py에서 호출 가능)
# --------------------------------------------------------
def run_embedding(input_path: Path = None):
    """
    파싱된 JSON 파일을 불러와 임베딩을 수행하고
    .npy / .json 파일로 저장
    """
    if input_path is None:
        input_path = PARSED_DIR / "parsed_docx.json"

    if not input_path.exists():
        raise FileNotFoundError(f"❌ 파싱된 JSON 파일을 찾을 수 없습니다: {input_path}")

    with open(input_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    if isinstance(data, list):
        documents = data
    elif isinstance(data, dict):
        documents = data.get("documents", [])
    else:
        raise ValueError("지원하지 않는 JSON 구조입니다.")

    if not documents:
        raise ValueError("❌ 문서 내용이 비어 있습니다.")

    print(f"📘 문장 개수: {len(documents)}개 → 임베딩 생성 중...")
    vectors = np.array([get_embedding(doc) for doc in documents])
    print("✅ 임베딩 shape:", vectors.shape)

    np.save(EMBED_DIR / "doc_embeddings.npy", vectors)

    embedding_data = [
        {"text": doc, "vector": vec.tolist()}
        for doc, vec in zip(documents, vectors)
    ]
    with open(EMBED_DIR / "doc_embeddings.json", "w", encoding="utf-8") as f:
        json.dump(embedding_data, f, ensure_ascii=False, indent=2)

    print(f"✅ 임베딩 저장 완료 → {EMBED_DIR / 'docx_embeddings.npy'}")
    print(f"✅ 텍스트 매핑 저장 완료 → {EMBED_DIR / 'docx_embeddings.json'}")

    return vectors


# --------------------------------------------------------
# 3. 단독 실행 가능하도록 유지
# --------------------------------------------------------
if __name__ == "__main__":
    run_embedding()
