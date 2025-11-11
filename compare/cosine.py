# rank_full_doc2doc_strict.py
import numpy as np
from qdrant_client import QdrantClient, models

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
PCA_SAMPLE_PER_DOC = 400     # (↑) 문서당 샘플 2배로 늘려 공통성분 추정 안정화
PCA_COMPONENTS      = 24      # (↑) 12 → 24: 템플릿 성분을 더 제거하되 의미 손실은 방지

# 보수적 점수화
PERCENTILE = 0.90            # (↑) 상위 10%만 평균: 형식적 저유사 구간 영향 축소
TRIM_RATIO = 0.15            # (↑) 상/하위 15% 컷: outlier 및 잔여 템플릿 흔적 제거

# 커버리지 임계
COVER_TAU_HIGH  = 0.88       # (↑) “강한” 판정 강화
COVER_TAU_VHIGH = 0.93       # (↑) “복제급” 매우 보수적으로


# ================= 유틸 =================
def l2n(m):  # row-wise normalize
    d = np.linalg.norm(m, axis=1, keepdims=True) + 1e-12
    return m / d

def sample_rows(m, cap, seed=SEED):
    if cap is None or m.shape[0] <= cap:
        return m
    idx = np.random.default_rng(seed).choice(m.shape[0], size=cap, replace=False)
    return m[idx]

def scroll_file_ids(client):
    s, nxt = set(), None
    while True:
        pts, nxt = client.scroll(DOC_COLLECTION, with_vectors=False, with_payload=True, limit=2048, offset=nxt)
        for p in pts:
            fid = p.payload.get("file_id")
            if isinstance(fid, str): s.add(fid)
        if not nxt: break
    return sorted(s)

def fetch_vectors_for_file(client, fid, cap):
    vecs, nxt = [], None
    flt = models.Filter(must=[models.FieldCondition(key="file_id", match=models.MatchValue(value=fid))])
    while True:
        pts, nxt = client.scroll(DOC_COLLECTION, scroll_filter=flt, with_vectors=True, with_payload=False, limit=1024, offset=nxt)
        for p in pts:
            vecs.append(np.asarray(p.vector, dtype=np.float32))
        if not nxt: break
    if not vecs:
        return np.zeros((0, VECTOR_DIM), dtype=np.float32)
    M = np.vstack(vecs)
    return sample_rows(M, cap)

def fit_pca_template_basis(client) -> np.ndarray:
    """코퍼스 전체에서 공통 성분(템플릿) 주성분 k개 추정"""
    file_ids = scroll_file_ids(client)
    bag = []
    for fid in file_ids:
        V = fetch_vectors_for_file(client, fid, PCA_SAMPLE_PER_DOC)
        if V.size == 0: continue
        bag.append(l2n(V))
    if not bag:
        return np.zeros((VECTOR_DIM, 0), dtype=np.float32)

    X = np.vstack(bag)  # [N, D]
    # SVD: X = U S VT, 주성분은 V의 행(또는 VT의 열)
    # 메모리 절약을 위해 randomized SVD 대체가 필요할 수도 있으나 규모가 크지 않으면 OK
    U, S, VT = np.linalg.svd(X, full_matrices=False)
    B = VT[:PCA_COMPONENTS].T   # [D, k]
    return B.astype(np.float32)

def remove_template(X: np.ndarray, B: np.ndarray) -> np.ndarray:
    """X에서 주성분 기저 B로의 성분을 제거 (정사영 제거)"""
    if X.size == 0 or B.size == 0:
        return l2n(X.astype(np.float32))
    X = X.astype(np.float32)
    Xn = l2n(X)
    # 정사영: Xn - (Xn·B)B^T
    C = Xn @ B              # [N, k]
    X_res = Xn - C @ B.T    # [N, D]
    return l2n(X_res)

def rowmax_percentile_mean(sims, pct: float) -> float:
    """각 행의 최대치 벡터 → 상위 pct 구간만 평균 (포화 완화)"""
    mx = sims.max(axis=1)  # [N]
    k = int(max(1, np.ceil(len(mx) * pct)))
    part = np.partition(mx, -k)[-k:]
    return float(part.mean())

def rowmax_trimmed_mean(sims, trim_ratio: float) -> float:
    """각 행의 최대치 벡터 → 상하위 일부 제거한 트림 평균 (극단치 완화)"""
    mx = np.sort(sims.max(axis=1))  # 오름차순
    n = len(mx)
    t = int(n * trim_ratio)
    sl = mx[t: n - t] if n - 2*t > 0 else mx
    return float(sl.mean()) if len(sl) > 0 else float(mx.mean()) if n > 0 else 0.0

def coverage_at_tau(A, B, tau):
    mx = (A @ B.T).max(axis=1)
    return float((mx >= tau).mean() * 100.0)

# ================= 메인 로직 =================
def rank_strict(query_npy_path: str):
    client = QdrantClient(path=QDRANT_PATH)

    # 1) 코퍼스 템플릿 기저 추정 (한 번만 학습, 이후 재사용 가능)
    basis = fit_pca_template_basis(client)  # [D, k]

    # 2) 쿼리 로드 → 템플릿 제거 → 샘플링
    Q = np.load(query_npy_path).astype(np.float32)
    assert Q.ndim == 2 and Q.shape[1] == VECTOR_DIM
    Q = remove_template(Q, basis)
    Q = sample_rows(Q, MAX_QUERY_CHUNKS)

    file_ids = scroll_file_ids(client)
    results = []
    for fid in file_ids:
        C = fetch_vectors_for_file(client, fid, MAX_CAND_CHUNKS)
        C = remove_template(C, basis)

        # 전면 비교 (정규화된 상태)
        sims = Q @ C.T  # [|Q|, |C|]

        # 보수적 집계: 퍼센타일 평균 + 트림 평균 (양방향)
        pm_q2c = rowmax_percentile_mean(sims, PERCENTILE)
        pm_c2q = rowmax_percentile_mean(sims.T, PERCENTILE)
        tm_q2c = rowmax_trimmed_mean(sims, TRIM_RATIO)
        tm_c2q = rowmax_trimmed_mean(sims.T, TRIM_RATIO)

        # 최종 점수: (퍼센타일 평균과 트림 평균의 양방향 평균)
        docsim = 0.5 * ((pm_q2c + pm_c2q) / 2.0 + (tm_q2c + tm_c2q) / 2.0)

        cov85_q2c = coverage_at_tau(Q, C, COVER_TAU_HIGH)
        cov85_c2q = coverage_at_tau(C, Q, COVER_TAU_HIGH)
        cov90_q2c = coverage_at_tau(Q, C, COVER_TAU_VHIGH)
        cov90_c2q = coverage_at_tau(C, Q, COVER_TAU_VHIGH)

        results.append({
            "file_id": fid,
            "DocSim": docsim,                 # 0~1, 포화 덜함
            "PctMean_Q→C": pm_q2c,
            "PctMean_C→Q": pm_c2q,
            "TrimMean_Q→C": tm_q2c,
            "TrimMean_C→Q": tm_c2q,
            "Cov@0.85_Q→C(%)": cov85_q2c,
            "Cov@0.85_C→Q(%)": cov85_c2q,
            "Cov@0.90_Q→C(%)": cov90_q2c,
            "Cov@0.90_C→Q(%)": cov90_c2q,
            "Q_chunks": int(Q.shape[0]),
            "C_chunks": int(C.shape[0]),
        })

    client.close()

    # 정렬: DocSim 기준
    results.sort(key=lambda r: r["DocSim"], reverse=True)
    topn = results[:TOPK]

    print(f"\n🔎 Content-sim (PCA-debiased & strict) — Top-{TOPK} for '{query_npy_path}':")
    for i, r in enumerate(topn, 1):
        print(
            f"{i:>2}. {r['file_id']:20s}  "
            f"DocSim={r['DocSim']:.4f}  "
            f"PctMean_Q→C={r['PctMean_Q→C']:.4f}  "
            f"PctMean_C→Q={r['PctMean_C→Q']:.4f}  "
            f"Trim_Q→C={r['TrimMean_Q→C']:.4f}  "
            f"Trim_C→Q={r['TrimMean_C→Q']:.4f}  "
            f"Cov85_Q→C={r['Cov@0.85_Q→C(%)']:.1f}%  "
            f"Cov85_C→Q={r['Cov@0.85_C→Q(%)']:.1f}%  "
            f"Cov90_Q→C={r['Cov@0.90_Q→C(%)']:.1f}%  "
            f"Cov90_C→Q={r['Cov@0.90_C→Q(%)']:.1f}%  "
            f"[Q={r['Q_chunks']}, C={r['C_chunks']}]"
        )
    return topn, results

if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("사용법: python rank_full_doc2doc_strict.py <새문서_npy경로>")
        print("예시: python rank_full_doc2doc_strict.py ./embedding_result/greenfuture_chunk_embeddings.npy")
        raise SystemExit(1)
    rank_strict(sys.argv[1])

#python rank_full_doc2doc_strict.py <새문서_npy경로>
