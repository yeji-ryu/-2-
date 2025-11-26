# cosine.py (정리본)

import numpy as np
from qdrant_client import QdrantClient, models
from typing import List, Dict, Tuple

# ====== Qdrant / 차원 등 환경 ======
QDRANT_PATH = "./qdrant_storage"
DOC_COLLECTION = "my_document_store"
VECTOR_DIM = 1024

# ====== 랭킹/점수 파라미터 ======
TOPK = 10
MAX_QUERY_CHUNKS = 2000
MAX_CAND_CHUNKS  = 2000
SEED = 42

# 디바이싱(템플릿 제거)용 PCA
PCA_SAMPLE_PER_DOC = 400
PCA_COMPONENTS      = 24

# 보수적 점수화
PERCENTILE = 0.90     # 상위 10%만 평균
TRIM_RATIO = 0.15     # 상/하위 15% 컷

# 커버리지 임계
COVER_TAU_HIGH  = 0.88
COVER_TAU_VHIGH = 0.93


# ================= 유틸 =================
def l2n(m):  # row-wise normalize
    d = np.linalg.norm(m, axis=1, keepdims=True) + 1e-12
    return m / d

def sample_rows(m, cap, seed=SEED):
    if cap is None or m.shape[0] <= cap:
        return m
    idx = np.random.default_rng(seed).choice(m.shape[0], size=cap, replace=False)
    return m[idx]

def scroll_file_ids(client: QdrantClient) -> List[str]:
    s, nxt = set(), None
    while True:
        pts, nxt = client.scroll(
            DOC_COLLECTION,
            with_vectors=False,
            with_payload=True,
            limit=2048,
            offset=nxt,
        )
        for p in pts:
            fid = p.payload.get("file_id")
            if isinstance(fid, str):
                s.add(fid)
        if not nxt:
            break
    return sorted(s)

def fetch_vectors_for_file(client: QdrantClient, fid: str, cap: int) -> np.ndarray:
    vecs, nxt = [], None
    flt = models.Filter(
        must=[models.FieldCondition(key="file_id", match=models.MatchValue(value=fid))]
    )
    while True:
        pts, nxt = client.scroll(
            DOC_COLLECTION,
            scroll_filter=flt,
            with_vectors=True,
            with_payload=False,
            limit=1024,
            offset=nxt,
        )
        for p in pts:
            vecs.append(np.asarray(p.vector, dtype=np.float32))
        if not nxt:
            break

    if not vecs:
        return np.zeros((0, VECTOR_DIM), dtype=np.float32)
    M = np.vstack(vecs)
    return sample_rows(M, cap)

def fit_pca_template_basis(client: QdrantClient) -> np.ndarray:
    """코퍼스 전체에서 공통 성분(템플릿) 주성분 k개 추정"""
    file_ids = scroll_file_ids(client)
    bag = []
    for fid in file_ids:
        V = fetch_vectors_for_file(client, fid, PCA_SAMPLE_PER_DOC)
        if V.size == 0:
            continue
        bag.append(l2n(V))
    if not bag:
        return np.zeros((VECTOR_DIM, 0), dtype=np.float32)

    X = np.vstack(bag)  # [N, D]
    U, S, VT = np.linalg.svd(X, full_matrices=False)
    B = VT[:PCA_COMPONENTS].T   # [D, k]
    return B.astype(np.float32)

def remove_template(X: np.ndarray, B: np.ndarray) -> np.ndarray:
    """X에서 주성분 기저 B로의 성분을 제거 (정사영 제거)"""
    if X.size == 0 or B.size == 0:
        return l2n(X.astype(np.float32))
    X = X.astype(np.float32)
    Xn = l2n(X)
    C = Xn @ B              # [N, k]
    X_res = Xn - C @ B.T    # [N, D]
    return l2n(X_res)

def rowmax_percentile_mean(sims: np.ndarray, pct: float) -> float:
    """각 행의 최대치 벡터 → 상위 pct 구간만 평균"""
    mx = sims.max(axis=1)  # [N]
    k = int(max(1, np.ceil(len(mx) * pct)))
    part = np.partition(mx, -k)[-k:]
    return float(part.mean())

def rowmax_trimmed_mean(sims: np.ndarray, trim_ratio: float) -> float:
    """각 행의 최대치 벡터 → 상하위 일부 제거한 트림 평균"""
    mx = np.sort(sims.max(axis=1))
    n = len(mx)
    t = int(n * trim_ratio)
    sl = mx[t: n - t] if n - 2 * t > 0 else mx
    if len(sl) > 0:
        return float(sl.mean())
    return float(mx.mean()) if n > 0 else 0.0

def coverage_at_tau(A: np.ndarray, B: np.ndarray, tau: float) -> float:
    """A의 각 청크가 B의 어느 청크랑이라도 tau 이상으로 붙는 비율(%)"""
    if A.size == 0 or B.size == 0:
        return 0.0
    mx = (A @ B.T).max(axis=1)
    return float((mx >= tau).mean() * 100.0)


# ================= 메인 로직 =================

def cosine_rank_with_qdrant(query_vecs: np.ndarray) -> Tuple[List[Dict], List[Dict]]:
    """
    서버/스크립트에서 직접 호출할 수 있는 코사인 랭킹 함수.

    query_vecs : (N, VECTOR_DIM) float32 ndarray (정규화 안 돼 있어도 됨)
    return     : (topN 리스트, 전체 결과 리스트)
    """
    if query_vecs.ndim != 2 or query_vecs.shape[1] != VECTOR_DIM:
        raise ValueError(f"query_vecs shape {query_vecs.shape}가 잘못되었습니다.")

    client = QdrantClient(path=QDRANT_PATH)
    try:
        # 1) 템플릿 기저 추정
        basis = fit_pca_template_basis(client)

        # 2) 쿼리 템플릿 제거 + 샘플링
        Q = remove_template(query_vecs.astype(np.float32), basis)
        Q = sample_rows(Q, MAX_QUERY_CHUNKS)

        file_ids = scroll_file_ids(client)
        results: List[Dict] = []

        for fid in file_ids:
            C = fetch_vectors_for_file(client, fid, MAX_CAND_CHUNKS)
            C = remove_template(C, basis)

            if C.size == 0:
                continue

            sims = Q @ C.T  # [|Q|, |C|]

            pm_q2c = rowmax_percentile_mean(sims, PERCENTILE)
            pm_c2q = rowmax_percentile_mean(sims.T, PERCENTILE)
            tm_q2c = rowmax_trimmed_mean(sims, TRIM_RATIO)
            tm_c2q = rowmax_trimmed_mean(sims.T, TRIM_RATIO)

            docsim = 0.5 * ((pm_q2c + pm_c2q) / 2.0 + (tm_q2c + tm_c2q) / 2.0)

            cov85_q2c = coverage_at_tau(Q, C, COVER_TAU_HIGH)
            cov85_c2q = coverage_at_tau(C, Q, COVER_TAU_HIGH)
            cov90_q2c = coverage_at_tau(Q, C, COVER_TAU_VHIGH)
            cov90_c2q = coverage_at_tau(C, Q, COVER_TAU_VHIGH)

            results.append({
                "file_id": fid,
                "DocSim": docsim,
                "PctMean_Q2C": pm_q2c,
                "PctMean_C2Q": pm_c2q,
                "TrimMean_Q2C": tm_q2c,
                "TrimMean_C2Q": tm_c2q,
                "Cov85_Q2C_pct": cov85_q2c,
                "Cov85_C2Q_pct": cov85_c2q,
                "Cov90_Q2C_pct": cov90_q2c,
                "Cov90_C2Q_pct": cov90_c2q,
                "Q_chunks": int(Q.shape[0]),
                "C_chunks": int(C.shape[0]),
            })

        results.sort(key=lambda r: r["DocSim"], reverse=True)
        topn = results[:TOPK]
        return topn, results
    finally:
        client.close()


def rank_strict(query_npy_path: str):
    # 1) 쿼리 임베딩 로드
    Q = np.load(query_npy_path).astype(np.float32)
    assert Q.ndim == 2 and Q.shape[1] == VECTOR_DIM

    # 2) 공통 랭킹 함수 호출
    topn, results = cosine_rank_with_qdrant(Q)

    print(f"\n🔎 Content-sim (PCA-debiased & strict) — Top-{TOPK} for '{query_npy_path}':")
    for i, r in enumerate(topn, 1):
        print(
            f"{i:>2}. {r['file_id']:20s}  "
            f"DocSim={r['DocSim']:.4f}  "
            f"PctMean_Q2C={r['PctMean_Q2C']:.4f}  "
            f"PctMean_C2Q={r['PctMean_C2Q']:.4f}  "
            f"Trim_Q2C={r['TrimMean_Q2C']:.4f}  "
            f"Trim_C2Q={r['TrimMean_C2Q']:.4f}  "
            f"Cov85_Q2C={r['Cov85_Q2C_pct']:.1f}%  "
            f"Cov85_C2Q={r['Cov85_C2Q_pct']:.1f}%  "
            f"Cov90_Q2C={r['Cov90_Q2C_pct']:.1f}%  "
            f"Cov90_C2Q={r['Cov90_C2Q_pct']:.1f}%  "
            f"[Q={r['Q_chunks']}, C={r['C_chunks']}]"
        )
    return topn, results


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("사용법: python cosine.py <새문서_npy경로>")
        print("예시: python cosine.py ./embedding_result/greenfuture_chunk_embeddings.npy")
        raise SystemExit(1)
    rank_strict(sys.argv[1])
