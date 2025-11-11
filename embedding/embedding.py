# -*- coding: utf-8 -*-
import json
import argparse
from pathlib import Path
import numpy as np
from sentence_transformers import SentenceTransformer

# -----------------------------
# CLI
# -----------------------------
parser = argparse.ArgumentParser(description="Embed JSON with dragonkue/BGE-m3-ko and save near the input.")
parser.add_argument("--in", dest="input_path", required=True, help="입력 JSON 경로")
parser.add_argument("--outdir", dest="out_dir", default=None, help="결과 저장 폴더(기본: 입력과 동일 폴더)")
parser.add_argument("--batch-size", dest="batch_size", type=int, default=32, help="배치 크기 (기본 32)")
args = parser.parse_args()

INPUT_PATH = Path(args.input_path).resolve()
OUT_DIR = Path(args.out_dir).resolve() if args.out_dir else INPUT_PATH.parent
OUT_DIR.mkdir(parents=True, exist_ok=True)

prefix = INPUT_PATH.stem
NPY_PATH = OUT_DIR / f"{prefix}_embeddings.npy"
JSON_PATH = OUT_DIR / f"{prefix}_embeddings.json"

MODEL_NAME = "dragonkue/BGE-m3-ko"

# -----------------------------
# JSON 로더
# -----------------------------
def load_texts(path: Path):
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    texts = []
    if isinstance(data, list):
        for it in data:
            if isinstance(it, str):
                t = it.strip()
                if t:
                    texts.append(t)
            elif isinstance(it, dict):
                t = str(it.get("text", "")).strip()
                if t:
                    texts.append(t)
    elif isinstance(data, dict):
        docs = data.get("documents", [])
        for it in docs:
            if isinstance(it, str):
                t = it.strip()
                if t:
                    texts.append(t)
            elif isinstance(it, dict):
                t = str(it.get("text", "")).strip()
                if t:
                    texts.append(t)
    else:
        raise ValueError("지원하지 않는 JSON 구조입니다 (list 또는 dict).")

    if not texts:
        raise ValueError("❌ 문서 내용이 비어 있습니다.")
    return texts

# -----------------------------
# 메인
# -----------------------------
def main():
    if not INPUT_PATH.exists():
        raise FileNotFoundError(f"❌ 입력 파일을 찾을 수 없습니다: {INPUT_PATH}")

    texts = load_texts(INPUT_PATH)
    print(f"📘 문장 개수: {len(texts)} → 임베딩 생성 중... (model={MODEL_NAME})")

    model = SentenceTransformer(MODEL_NAME)
    try:
        model.max_seq_length = 512
    except Exception:
        pass

    vectors = model.encode(
        texts,
        batch_size=args.batch_size,
        convert_to_numpy=True,
        normalize_embeddings=True,
        show_progress_bar=True
    ).astype("float32")

    print("✅ 임베딩 shape:", vectors.shape)

    np.save(NPY_PATH, vectors)
    mapping = [{"text": t, "vector": v.tolist()} for t, v in zip(texts, vectors)]
    with open(JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(mapping, f, ensure_ascii=False, indent=2)

    print("💾 저장 완료")
    print(f"  • NPY : {NPY_PATH}")
    print(f"  • JSON: {JSON_PATH}")

if __name__ == "__main__":
    main()

#uv run python embeddingkorea2.py --in "C:/Users/gkseh/hansung/pz1023/ls/futurebattery_chunk.json"
