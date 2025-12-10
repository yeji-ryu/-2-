import argparse
import numpy as np
import uuid
import time
from qdrant_client import QdrantClient, models

# ===== 설정 =====
VECTOR_DIM = 1024
QDRANT_PATH = "./qdrant_storage"
COLLECTION_NAME = "my_document_store"

def ensure_collection(client: QdrantClient, dim: int):
    if not client.collection_exists(collection_name=COLLECTION_NAME):
        print(f"[init] 컬렉션 '{COLLECTION_NAME}'이 없어 새로 생성합니다.")
        client.create_collection(
            collection_name=COLLECTION_NAME,
            vectors_config=models.VectorParams(size=dim, distance=models.Distance.COSINE),
        )
        client.create_payload_index(
            collection_name=COLLECTION_NAME,
            field_name="file_id",
            field_schema=models.PayloadSchemaType.KEYWORD,
        )

def append_npy_only(client: QdrantClient, npy_path: str, file_id: str | None = None):
    print(f"[load] NPY: {npy_path}")
    V = np.load(npy_path)
    if V.ndim != 2 or V.shape[1] != VECTOR_DIM:
        raise ValueError(f"NPY shape {V.shape} 가 맞지 않습니다. (N, {VECTOR_DIM}) 필요")

    fid = file_id or f"doc-{uuid.uuid4()}"
    print(f"[meta] file_id = {fid} (vectors={len(V)})")

    points = []
    for i in range(V.shape[0]):
        points.append(
            models.PointStruct(
                id=str(uuid.uuid4()),
                vector=V[i].astype(np.float32).tolist(),
                payload={"file_id": fid}
            )
        )

    t0 = time.time()
    client.upsert(collection_name=COLLECTION_NAME, points=points, wait=True)
    dt = time.time() - t0
    print(f"[upsert] {len(points)} vectors upsert 완료 ({dt:.2f}s)")
    return fid, len(points)

def main():
    p = argparse.ArgumentParser(description="Append NPY-only embeddings into Qdrant.")
    p.add_argument("--npy", required=True, help="NPY 파일 경로 (shape: [N, 1024])")
    p.add_argument("--file-id", default=None, help="문서 식별자 (자동 생성 가능)")
    args = p.parse_args()

    client = QdrantClient(path=QDRANT_PATH)
    try:
        ensure_collection(client, VECTOR_DIM)
        fid, n = append_npy_only(client, args.npy, args.file_id)
        print(f"[done] file_id='{fid}' 로 {n}개의 벡터가 추가되었습니다.")
    finally:
        client.close()

if __name__ == "__main__":
    main()

#python qdrant_dense.py --npy "C:\Users\LG\zonaibattery_chunk_embeddings.npy" --file-id zonaibattery