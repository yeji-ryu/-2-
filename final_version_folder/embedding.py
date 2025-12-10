# -*- coding: utf-8 -*-
import json
from pathlib import Path
from typing import List, Tuple, Optional

import argparse
import numpy as np
from sentence_transformers import SentenceTransformer

# -----------------------------
# 공통 설정
# -----------------------------
MODEL_NAME = "dragonkue/BGE-m3-ko"

# 전역 모델 캐시 (서버에서 여러 번 호출해도 1번만 로드)
_model: Optional[SentenceTransformer] = None


# -----------------------------
# 모델 로더 + 임베딩 함수
# -----------------------------
def get_model(model_name: str = MODEL_NAME) -> SentenceTransformer:
    """전역 캐시를 사용하는 SentenceTransformer 로더"""
    global _model
    if _model is None:
        _model = SentenceTransformer(model_name)
        try:
            _model.max_seq_length = 512
        except Exception:
            pass
    return _model


def embed_texts(texts: List[str], batch_size: int = 32) -> np.ndarray:
    """
    문자열 리스트를 받아 BGE-m3-ko 임베딩을 반환.
    서버/다른 스크립트에서 직접 사용 가능.
    """
    if not texts:
        raise ValueError("❌ embed_texts: 빈 텍스트 리스트입니다.")
    model = get_model()
    vectors = model.encode(
        texts,
        batch_size=batch_size,
        convert_to_numpy=True,
        normalize_embeddings=True,
        show_progress_bar=False,  # 서버에서는 False, CLI에서는 True로 바꿀 수도 있음
    ).astype("float32")
    return vectors


# -----------------------------
# JSON 로더
# -----------------------------
def load_texts(path: Path) -> List[str]:
    """
    JSON 구조:
      - ["문장1", "문장2", ...]
      - [{"text": "..."}, ...]
      - {"documents": ["..."]} 또는 {"documents":[{"text":"..."}]}
    를 지원.
    """
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    texts: List[str] = []
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
# JSON 파일 → 임베딩 (서버/스크립트용)
# -----------------------------
def embed_json_file(
    json_path: str | Path,
    out_dir: str | Path | None = None,
    batch_size: int = 32,
    save: bool = True,
) -> Tuple[np.ndarray, List[str], Path, Path]:
    """
    JSON 파일을 읽어 임베딩을 생성하고, 필요하면 NPY/JSON으로 저장.

    return:
      vectors, texts, npy_path, json_out_path
    """
    in_path = Path(json_path).resolve()
    if not in_path.exists():
        raise FileNotFoundError(f"❌ 입력 파일을 찾을 수 없습니다: {in_path}")

    texts = load_texts(in_path)
    print(f"📘 문장 개수: {len(texts)} → 임베딩 생성 중... (model={MODEL_NAME})")

    vectors = embed_texts(texts, batch_size=batch_size)
    print("✅ 임베딩 shape:", vectors.shape)

    # 저장 경로 설정
    if out_dir is None:
        out_dir = in_path.parent
    out_dir = Path(out_dir).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    prefix = in_path.stem
    npy_path = out_dir / f"{prefix}_embeddings.npy"
    json_out_path = out_dir / f"{prefix}_embeddings.json"

    if save:
        np.save(npy_path, vectors)
        mapping = [{"text": t, "vector": v.tolist()} for t, v in zip(texts, vectors)]
        with open(json_out_path, "w", encoding="utf-8") as f:
            json.dump(mapping, f, ensure_ascii=False, indent=2)

        print("💾 저장 완료")
        print(f"  • NPY : {npy_path}")
        print(f"  • JSON: {json_out_path}")

    return vectors, texts, npy_path, json_out_path


# -----------------------------
# CLI 진입점
# -----------------------------
def main():
    parser = argparse.ArgumentParser(
        description="Embed JSON with dragonkue/BGE-m3-ko and save near the input."
    )
    parser.add_argument("--in", dest="input_path", required=True, help="입력 JSON 경로")
    parser.add_argument("--outdir", dest="out_dir", default=None, help="결과 저장 폴더(기본: 입력과 동일 폴더)")
    parser.add_argument("--batch-size", dest="batch_size", type=int, default=32, help="배치 크기 (기본 32)")
    args = parser.parse_args()

    embed_json_file(
        json_path=args.input_path,
        out_dir=args.out_dir,
        batch_size=args.batch_size,
        save=True,
    )


if __name__ == "__main__":
    main()

# 사용 예시:
# uv run python embedding.py --in "C:/Users/gkseh/hansung/pz1023/ls/futurebattery_chunk.json"
