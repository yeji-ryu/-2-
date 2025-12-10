# -*- coding: utf-8 -*-

from rank_bm25 import BM25Okapi
from collections import Counter, defaultdict
from pathlib import Path
from typing import List, Tuple, Dict
import numpy as np
import json
from scipy.sparse import csr_matrix, save_npz, load_npz
import argparse
import re
import unicodedata

# ------------------ pretty JSON ------------------
def write_pretty_json(path: str | Path, obj) -> None:
    Path(path).write_text(
        json.dumps(obj, ensure_ascii=False, indent=2, separators=(", ", ": ")),
        encoding="utf-8"
    )

def normalize_text(s: str) -> str:
    """NFC 정규화 + zero-width space 제거"""
    return unicodedata.normalize("NFC", s).replace("\u200b", "")

# =============== Tokenizers ===============
_KOREAN_RE = re.compile(r"[가-힣]")

def _tokenize_simple(text: str) -> List[str]:
    text = re.sub(r"[^0-9a-zA-Z가-힣]+", " ", text)
    return [t for t in text.strip().split() if t]

def _tokenize_okt_nouns(text: str) -> List[str]:
    from konlpy.tag import Okt
    okt = Okt()
    return [t for t in okt.nouns(text) if len(t) >= 2]

def _tokenize_mecab_nouns(text: str) -> List[str]:
    from mecab import MeCab
    m = MeCab()
    return [t for t in m.nouns(text) if len(t) >= 2]

def make_tokenizer(name: str):
    name = (name or "simple").lower()
    if name == "okt":
        base = _tokenize_okt_nouns
    elif name == "mecab":
        base = _tokenize_mecab_nouns
    else:
        base = _tokenize_simple

    def _tok(text: str) -> List[str]:
        text = normalize_text(text)
        if name in ("okt", "mecab") and _KOREAN_RE.search(text):
            try:
                return base(text)
            except Exception as e:
                print(f"[Tokenizer:{name}] 로드/토크나이즈 실패, simple로 대체 → {e}")
                return _tokenize_simple(text)
        return _tokenize_simple(text)

    _tok.__name__ = f"{name}_tokenizer"
    return _tok, name

TOKENIZE = _tokenize_simple
TOKENIZER_NAME = "simple"

# =============== JSON 로더 ===============
def load_chunks_from_json(json_path: str | Path, encoding: str = "utf-8") -> List[str]:
    p = Path(json_path)
    obj = json.loads(p.read_text(encoding=encoding, errors="ignore"))

    if isinstance(obj, list):
        out: List[str] = []
        for it in obj:
            if isinstance(it, str) and it.strip():
                out.append(it.strip())
            elif isinstance(it, dict) and isinstance(it.get("text"), str) and it["text"].strip():
                out.append(it["text"].strip())
        if out:
            print(f"[JSON] 청크 {len(out)}개 로드")
            return out

    texts: List[str] = []
    def rec(o):
        if isinstance(o, dict):
            for k, v in o.items():
                if k.lower() == "text" and isinstance(v, str) and v.strip():
                    texts.append(v.strip())
                else:
                    rec(v)
        elif isinstance(o, list):
            for x in o:
                rec(x)
    rec(obj)
    if not texts:
        raise ValueError(f"❌ JSON 구조에서 text를 찾지 못함: {p}")
    print(f"[JSON] 청크 {len(texts)}개 로드(재귀)")
    return texts

# =============== prepare-delimited ===============
def prepare_delimited(json_path, out_dir, delim="---", doc_prefix="report", use_title_as_id=False):
    lines = load_chunks_from_json(json_path)
    lines = [s.strip() for s in lines if isinstance(s, str) and s.strip()]

    docs: List[List[str]] = []
    cur: List[str] = []
    for s in lines:
        if s == delim:
            if cur:
                docs.append(cur); cur = []
        else:
            cur.append(s)
    if cur:
        docs.append(cur)

    out_rows: List[str] = []
    chunk2doc: List[Dict] = []
    row_idx = 0
    for i, d in enumerate(docs, 1):
        doc_id = (d[0].strip() if use_title_as_id and d else f"{doc_prefix}{i}")
        for seg in d:
            out_rows.append(seg)
            chunk2doc.append({"row": row_idx, "doc_id": doc_id})
            row_idx += 1

    out_dir = Path(out_dir); out_dir.mkdir(parents=True, exist_ok=True)
    write_pretty_json(out_dir / "corpus_chunks.json", out_rows)
    write_pretty_json(out_dir / "bm25_chunk2doc.json", chunk2doc)
    print(f"[Prep] 문서 {len(docs)}개, 총 라인 {len(out_rows)}개")

# =============== BM25 인덱스 빌드 ===============
def build_index(docs: List[str], save_dir: str | Path, k1: float = 1.5, b: float = 0.75):
    save_dir = Path(save_dir)
    save_dir.mkdir(parents=True, exist_ok=True)
    print(f"[BM25] Build from {len(docs)} chunks ...")
    print(f"[Tokenizer] using: {TOKENIZER_NAME}")

    tokenized = [TOKENIZE(d) for d in docs]
    bm25 = BM25Okapi(tokenized, k1=k1, b=b)

    vocab: Dict[str, int] = {}
    for d in tokenized:
        for t in d:
            if t not in vocab:
                vocab[t] = len(vocab)

    rows, cols, data = [], [], []
    avgdl = bm25.avgdl
    doc_lens = bm25.doc_len
    idf = bm25.idf

    for i, toks in enumerate(tokenized):
        tf = Counter(toks)
        dl = doc_lens[i]
        for term, freq in tf.items():
            if term not in idf:
                continue
            t_id = vocab[term]
            w = idf[term] * ((bm25.k1 + 1) * freq) / (freq + bm25.k1 * (1 - bm25.b + bm25.b * dl / avgdl))
            if w > 0:
                rows.append(i); cols.append(t_id); data.append(w)

    W = csr_matrix((data, (rows, cols)), shape=(len(docs), len(vocab)), dtype=np.float32)

    save_npz(save_dir / "bm25_matrix.npz", W)
    write_pretty_json(save_dir / "bm25_vocab.json", vocab)
    write_pretty_json(save_dir / "bm25_docs.json", docs)
    write_pretty_json(save_dir / "bm25_meta.json", {"k1": k1, "b": b, "tokenizer": TOKENIZER_NAME})

    print(f"[BM25] Saved npz | chunks={len(docs)}, vocab={len(vocab)}, nnz={W.nnz}")

    return W, vocab, docs

def load_index(save_dir: str | Path):
    save_dir = Path(save_dir)
    W = load_npz(save_dir / "bm25_matrix.npz").astype(np.float32)
    vocab = json.loads((save_dir / "bm25_vocab.json").read_text(encoding="utf-8"))
    docs  = json.loads((save_dir / "bm25_docs.json").read_text(encoding="utf-8"))
    print(f"[BM25] Loaded matrix (chunks={W.shape[0]}, terms={W.shape[1]})")
    return W, vocab, docs

# =============== 문서 단위 집계 ===============
def _load_chunk2doc(path: str | Path):
    obj = json.loads(Path(path).read_text(encoding="utf-8"))
    return obj

def _aggregate_doc_scores(scores: np.ndarray, chunk2doc: List[Dict], mode: str = "topMsum"):
    buckets = defaultdict(list)
    for m in chunk2doc:
        buckets[m["doc_id"]].append(float(scores[m["row"]]))

    doc_scores = {}
    for doc, arr in buckets.items():
        arr.sort(reverse=True)

        if mode == "sum":
            s = sum(arr)
        elif mode == "max":
            s = arr[0] if arr else 0.0
        else:
            s = sum(arr[:4])

        doc_scores[doc] = s

    return sorted(doc_scores.items(), key=lambda x: x[1], reverse=True)

# =============== CLI ===============
def main():
    ap = argparse.ArgumentParser(description="BM25 minimal CLI")
    sub = ap.add_subparsers(dest="cmd", required=True)

    # prepare-delimited
    p_pd = sub.add_parser("prepare-delimited")
    p_pd.add_argument("--json", required=True)
    p_pd.add_argument("--out_dir", required=True)
    p_pd.add_argument("--delim", default="---")
    p_pd.add_argument("--doc_prefix", default="report")
    p_pd.add_argument("--use_title_as_id", action="store_true")

    common = argparse.ArgumentParser(add_help=False)
    common.add_argument("--tokenizer", choices=["simple", "okt", "mecab"], default="simple")

    # build-json
    p_j = sub.add_parser("build-json", parents=[common])
    p_j.add_argument("--json", required=True)
    p_j.add_argument("--save_dir", required=True)
    p_j.add_argument("--k1", type=float, default=1.5)
    p_j.add_argument("--b", type=float, default=0.75)

    # search-doc
    p_sd = sub.add_parser("search-doc", parents=[common])
    p_sd.add_argument("--save_dir", required=True)
    p_sd.add_argument("--query_json", required=True)
    p_sd.add_argument("--chunk2doc", required=True)
    p_sd.add_argument("--mode", choices=["sum", "max", "topMsum"], default="topMsum")
    p_sd.add_argument("--topk", type=int, default=3)

    args = ap.parse_args()

    global TOKENIZE, TOKENIZER_NAME
    TOKENIZE, TOKENIZER_NAME = make_tokenizer(getattr(args, "tokenizer", "simple"))
    print(f"[Tokenizer] using: {TOKENIZER_NAME}")

    if args.cmd == "prepare-delimited":
        prepare_delimited(
            json_path=args.json,
            out_dir=args.out_dir,
            delim=args.delim,
            doc_prefix=args.doc_prefix,
            use_title_as_id=bool(args.use_title_as_id),
        )

    elif args.cmd == "build-json":
        chunks = load_chunks_from_json(args.json)
        build_index(chunks, args.save_dir, k1=args.k1, b=args.b)

    elif args.cmd == "search-doc":
        W, vocab, _ = load_index(args.save_dir)

        q_chunks = load_chunks_from_json(args.query_json)
        q_tokens = []
        for s in q_chunks:
            q_tokens.extend(TOKENIZE(s))

        q_tokens = [t for t in q_tokens if t in vocab]
        if not q_tokens:
            raise SystemExit("[BM25] ❌ 쿼리 문서 토큰이 코퍼스 vocab에 없음")

        tf = Counter(q_tokens)
        ids = [vocab[t] for t in tf.keys()]
        vals = [float(tf[t]) for t in tf.keys()]

        q = csr_matrix((vals, ([0]*len(ids), ids)), shape=(1, W.shape[1]), dtype=np.float32)
        scores = (W @ q.T).toarray().ravel()

        c2d = _load_chunk2doc(args.chunk2doc)
        ranked = _aggregate_doc_scores(scores, c2d, mode=args.mode)

        top = ranked[:args.topk]
        print(f"\n[문서 Top-{args.topk}] (mode={args.mode}, M=4)")
        for doc_id, s in top:
            print(f"{s:12.6f}\t{doc_id}")

# =============== 외부 호출용 Wrapper ===============
def bm25_doc_similarity_from_chunks(
    chunks: List[str],
    save_dir: str | Path,
    chunk2doc_path: str | Path,
    tokenizer: str = "simple",
    mode: str = "topMsum",
    topk: int = 10,
):
    global TOKENIZE, TOKENIZER_NAME
    TOKENIZE, TOKENIZER_NAME = make_tokenizer(tokenizer)

    W, vocab, _ = load_index(save_dir)

    q_tokens = []
    for s in chunks:
        q_tokens.extend(TOKENIZE(s))

    q_tokens = [t for t in q_tokens if t in vocab]
    if not q_tokens:
        return []

    tf = Counter(q_tokens)
    ids = [vocab[t] for t in tf.keys()]
    vals = [float(tf[t]) for t in tf.keys()]

    q = csr_matrix((vals, ([0]*len(ids), ids)), shape=(1, W.shape[1]), dtype=np.float32)
    scores = (W @ q.T).toarray().ravel()

    c2d = _load_chunk2doc(chunk2doc_path)
    ranked = _aggregate_doc_scores(scores, c2d, mode=mode)

    return ranked[:topk]


if __name__ == "__main__":
    main()
