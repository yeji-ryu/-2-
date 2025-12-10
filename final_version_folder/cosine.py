# -*- coding: utf-8 -*-
import numpy as np
from pathlib import Path
from qdrant_client import QdrantClient, models

# ----------------------------
# 설정 (rag_pipeline_v18.py와 동일)
# ----------------------------
VECTOR_DIM = 1024
DENSE_COLLECTION = "my_document_store"
QDRANT_PATH = "./qdrant_storage"


# -------------------------------------
# Util: L2 normalization + cosine
# -------------------------------------
def l2norm(X: np.ndarray) -> np.ndarray:
    return X / (np.linalg.norm(X, axis=1, keepdims=True) + 1e-12)


def cosine_matrix(A: np.ndarray, B: np.ndarray) -> np.ndarray:
    A = l2norm(A)
    B = l2norm(B)
    return A @ B.T


# -------------------------------------
# Qdrant에서 해당 문서(file_id)의 Dense 벡터 가져오기
# -------------------------------------
def fetch_vectors(client: QdrantClient, file_id: str) -> np.ndarray:
    vecs = []
    next_offset = None

    flt = models.Filter(
        must=[models.FieldCondition(
            key="file_id", match=models.MatchValue(value=file_id)
        )]
    )

    while True:
        points, next_offset = client.scroll(
            collection_name=DENSE_COLLECTION,
            scroll_filter=flt,
            with_vectors=True,
            with_payload=False,
            limit=2048,
            offset=next_offset
        )

        for p in points:
            vecs.append(np.asarray(p.vector, dtype=np.float32))

        if not next_offset:
            break

    if not vecs:
        return np.zeros((0, VECTOR_DIM), dtype=np.float32)

    return np.vstack(vecs)


# -------------------------------------
# rag_pipeline_v18 에서 호출하는 코사인 유사도 메인 함수
#   - Qdrant로부터 문서별 벡터를 가져와
#   - 쿼리 벡터와의 코사인 유사도를 문서 단위 점수로 계산
#   - 문장 매칭 정보는 포함하지 않음 (문장 매칭은 rag_pipeline에서 재계산)
# -------------------------------------
def cosine_rank_with_qdrant(query_vectors, query_chunks):
    """
    query_vectors: (Q_chunks, 1024) 벡터
    query_chunks : 쿼리 문장의 텍스트 리스트 (여기서는 DocSim 계산에는 사용 안 함)
    return: (정렬 리스트, 같은 리스트)
    """
    client = QdrantClient(path=QDRANT_PATH)

    # -------------------
    # Qdrant에 있는 모든 file_id 수집
    # -------------------
    file_ids = set()
    next_offset = None

    while True:
        points, next_offset = client.scroll(
            collection_name=DENSE_COLLECTION,
            with_payload=True,
            with_vectors=False,
            limit=2048,
            offset=next_offset
        )
        for p in points:
            payload = p.payload or {}
            fid = payload.get("file_id")
            if fid:
                file_ids.add(fid)

        if not next_offset:
            break

    # -------------------
    # 각 문서와 쿼리 간 코사인 점수 계산
    # -------------------
    results = []

    for fid in sorted(file_ids):
        C = fetch_vectors(client, fid)

        if C.size == 0 or query_vectors.size == 0:
            score = 0.0
        else:
            S = cosine_matrix(query_vectors, C)  # (Q x C)
            # 각 쿼리 문장 기준으로 가장 비슷한 청크의 점수
            best_vals = np.max(S, axis=1)
            # 그 평균을 문서 단위 점수로 사용
            score = float(np.mean(best_vals))

        results.append({
            "file_id": fid,
            "DocSim": score,
        })

    client.close()

    # 유사도 기준 정렬
    sorted_results = sorted(results, key=lambda r: r["DocSim"], reverse=True)

    # rag_pipeline_v18.py 는 두 개 반환 받도록 설계
    return sorted_results, sorted_results
