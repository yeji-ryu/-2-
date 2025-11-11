# bm25_search_doc_vecstore.py
# 이전 CLI와 동일하게: search-doc --query_json ... --chunk2doc ... --mode ... --M ... --topk ...
# 단, BM25 점수는 로컬 CSR이 아니라 Qdrant(임베디드) sparse 컬렉션에서 가져옴.

import json, re, argparse
from pathlib import Path
from collections import Counter, defaultdict
from qdrant_client import QdrantClient, models

# ---------- tokenizers ----------
_KOREAN_RE = re.compile(r"[가-힣]")

def _tok_simple(s: str):
    s = re.sub(r"[^0-9a-zA-Z가-힣]+", " ", s)
    return [t for t in s.strip().split() if t]

def _tok_okt(s: str):
    from konlpy.tag import Okt
    okt = Okt()
    return [t for t in okt.nouns(s) if len(t) >= 2]

def _tok_mecab(s: str):
    from mecab import MeCab
    m = MeCab()
    return [t for t in m.nouns(s) if len(t) >= 2]

def make_tokenizer(name: str):
    name = (name or "simple").lower()
    base = _tok_simple
    if name == "okt":   base = _tok_okt
    if name == "mecab": base = _tok_mecab

    def _tok(s: str):
        if base is _tok_simple:        return _tok_simple(s)
        if _KOREAN_RE.search(s or ""):
            try:                       return base(s)
            except Exception as e:
                print(f"[Tokenizer:{name}] 실패 → simple 대체: {e}")
        return _tok_simple(s)
    return _tok

# ---------- loaders ----------
def load_vocab(p: Path) -> dict:
    return json.loads(p.read_text(encoding="utf-8"))

def load_chunks(p: Path) -> list[str]:
    obj = json.loads(p.read_text(encoding="utf-8"))
    if isinstance(obj, list):
        out=[]
        for it in obj:
            if isinstance(it, str) and it.strip():
                out.append(it.strip())
            elif isinstance(it, dict) and isinstance(it.get("text"), str) and it["text"].strip():
                out.append(it["text"].strip())
        if out: return out
    # 재귀 탐색
    texts=[]
    def rec(o):
        if isinstance(o, dict):
            for k,v in o.items():
                if k.lower()=="text" and isinstance(v,str) and v.strip():
                    texts.append(v.strip())
                else:
                    rec(v)
        elif isinstance(o, list):
            for x in o: rec(x)
    rec(obj)
    if not texts:
        raise ValueError("JSON에서 text를 찾지 못했습니다.")
    return texts

def load_chunk2doc(p: Path) -> dict[int, str]:
    obj = json.loads(p.read_text(encoding="utf-8"))
    out={}
    for m in obj:
        out[int(m["row"])] = str(m["doc_id"])
    return out

# ---------- helpers ----------
def make_sparse_query(chunks: list[str], vocab: dict, tokenize) -> tuple[list[int], list[float]]:
    toks=[]
    for s in chunks: toks.extend(tokenize(s))
    toks=[t for t in toks if t in vocab]
    if not toks: return [], []
    tf = Counter(toks)
    idx = [vocab[t] for t in tf.keys()]
    val = [float(tf[t]) for t in tf.keys()]  # 필요시 전부 1.0 가능
    return idx, val

def aggregate(scores_by_doc: dict[str, list[float]], mode: str, M: int) -> list[tuple[str, float]]:
    ranked=[]
    for doc_id, arr in scores_by_doc.items():
        arr.sort(reverse=True)
        if mode=="sum":
            s = sum(arr)
        elif mode=="max":
            s = arr[0]
        else:  # topMsum
            s = sum(arr[:max(1, M)])
        ranked.append((doc_id, s))
    ranked.sort(key=lambda x:x[1], reverse=True)
    return ranked

# ---------- main ----------
def main():
    ap = argparse.ArgumentParser(description="BM25 search-doc (Qdrant 임베디드 사용)")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p_sd = sub.add_parser("search-doc", help="원문/청크 JSON을 쿼리로 사용하여 문서 Top-K")
    p_sd.add_argument("--qdrant-path", required=True, help="임베디드 저장 경로")
    p_sd.add_argument("--bm25-collection", required=True, help="BM25 스파스 컬렉션명 (예: bm25_chunks)")
    p_sd.add_argument("--sparse-name", default="bm25", help="스파스 벡터 이름 (기본: bm25)")
    p_sd.add_argument("--vocab", required=True, help="bm25_vocab.json (term->id)")
    p_sd.add_argument("--query_json", required=True, help="['...', ...] or [{'text':'...'}, ...]")
    p_sd.add_argument("--chunk2doc", required=True, help="row→doc_id 매핑 JSON")
    p_sd.add_argument("--mode", choices=["sum","max","topMsum"], default="topMsum")
    p_sd.add_argument("--M", type=int, default=3)
    p_sd.add_argument("--topk", type=int, default=3)
    p_sd.add_argument("--tokenizer", choices=["simple","okt","mecab"], default="simple")

    args = ap.parse_args()

    if args.cmd == "search-doc":
        tokenize = make_tokenizer(args.tokenizer)
        client = QdrantClient(path=args.qdrant_path)
        vocab = load_vocab(Path(args.vocab))
        chunks = load_chunks(Path(args.query_json))
        row2doc = load_chunk2doc(Path(args.chunk2doc))

        # 1) 스파스 쿼리
        idx, val = make_sparse_query(chunks, vocab, tokenize)
        if not idx:
            raise SystemExit("쿼리 토큰이 vocab에 없음")

        # 2) 컬렉션 전체 개수만큼 검색(이 스크립트는 별도 limit 파라미터 없음: 이전 CLI와 동일 컨셉)
        total = client.count(collection_name=args.bm25_collection, exact=True).count

        hits = client.search(
            collection_name=args.bm25_collection,
            query_vector=models.NamedSparseVector(
                name=args.sparse_name,
                vector=models.SparseVector(indices=idx, values=val),
            ),
            query_filter=models.Filter(
                must=[models.FieldCondition(key="type", match=models.MatchValue(value="chunk"))]
            ),
            with_payload=True,
            limit=total,  # 전체 검색
        )

        # 3) 문서 단위 집계 (payload.row → row2doc → doc_id, 또는 payload.doc_id 직접 사용)
        buckets = defaultdict(list)
        for h in hits:
            doc_id = h.payload.get("doc_id")
            if not doc_id:
                # 업서트 시 row를 payload에 넣어두었다면 이를 통해 매핑
                r = h.payload.get("row")
                if r is not None:
                    doc_id = row2doc.get(int(r))
            if not doc_id:
                # 마지막 안전장치: 문서 미상은 스킵
                continue
            buckets[doc_id].append(h.score)

        ranked = aggregate(buckets, args.mode, args.M)[:args.topk]

        print(f"\n[문서 Top-{args.topk}] (mode={args.mode}, M={args.M})")
        for doc_id, s in ranked:
            print(f"{s:12.6f}\t{doc_id}")

        client.close()

if __name__ == "__main__":
    main()

#uv run python compare.py search-doc --qdrant-path "C:\Users\gkseh\hansung\pz1023\qdrant_storage" --bm25-collection "bm25_chunks" --sparse-name "bm25" --vocab "C:\Users\gkseh\hansung\pz1023\bm25_test\bm25_vocab.json" --query_json "C:\Users\gkseh\hansung\pz1023\ls\flexwavepay.json" --chunk2doc "C:\Users\gkseh\hansung\pz1023\bm25_test\bm25_chunk2doc.json" --mode topMsum --M 3 --topk 3 --tokenizer mecab
