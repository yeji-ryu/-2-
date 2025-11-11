import json
import numpy as np
import faiss
from pathlib import Path

# -------------------------
# 1. 경로 설정
# -------------------------
BASE_DIR = Path(__file__).resolve().parents[2]
EMBED_DIR = BASE_DIR / "embeddings"
VSTORE_DIR = BASE_DIR / "vectorstores"
VSTORE_DIR.mkdir(exist_ok=True)

EMBED_PATH = EMBED_DIR / "docx_embeddings.npy"
TEXT_PATH = EMBED_DIR / "docx_embeddings.json"

INDEX_PATH = VSTORE_DIR / "index.faiss"
DOCS_PATH = VSTORE_DIR / "docs.json"

# -------------------------
# 2. 임베딩 & 원문 로드
# -------------------------
print(f"📂 임베딩 불러오는 중: {EMBED_PATH}")
vectors = np.load(EMBED_PATH).astype("float32")

with open(TEXT_PATH, "r", encoding="utf-8") as f:
    docs = json.load(f)

print(f"📘 임베딩 shape: {vectors.shape}")
print(f"🧾 문장 개수: {len(docs)}")

# -------------------------
# 3. 벡터 정규화 (Cosine Similarity용)
# -------------------------
faiss.normalize_L2(vectors)
print("✅ 벡터 정규화 완료 (L2 Norm 적용)")

# -------------------------
# 4. FAISS 인덱스 생성 (Inner Product = Cosine Similarity)
# -------------------------
index = faiss.IndexFlatIP(vectors.shape[1])
index.add(vectors)

# -------------------------
# 5. 인덱스 및 문장 저장
# -------------------------
faiss.write_index(index, str(INDEX_PATH))
with open(DOCS_PATH, "w", encoding="utf-8") as f:
    json.dump(docs, f, ensure_ascii=False, indent=2)

print(f"✅ FAISS 인덱스 저장 완료 → {INDEX_PATH}")
print(f"✅ 문장 매핑 저장 완료 → {DOCS_PATH}")
print("🚀 벡터스토어 구축 완료 (Cosine Similarity 기반)")
