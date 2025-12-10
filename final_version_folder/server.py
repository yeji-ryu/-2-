# server.py — HS Capital RAG Pipeline API (UPLOAD → RUN_QUERY)

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pathlib import Path
from datetime import datetime
import shutil
import threading
import json
from fastapi.responses import StreamingResponse, FileResponse
import time


# === 파이프라인 코드 import ===
from rag_pipeline_v7 import (
    run_query,
    build_llm_payload,
    summarize_with_llm_local,
    save_markdown_summary
)

# =====================================================
# 기본 FastAPI 설정
# =====================================================

app = FastAPI(title="HS Capital RAG Analyzer API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["POST", "GET"],
    allow_headers=["*"],
)

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

BM25_DIR = BASE_DIR / "bm25_index"
BM25_DIR.mkdir(exist_ok=True)

ALLOWED_EXT = {"hwpx", "docx", "pdf"}


def validate_ext(filename: str) -> bool:
    return filename.split(".")[-1].lower() in ALLOWED_EXT


# =====================================================
# /analyze  — run_query만 즉시 수행 후 응답
# =====================================================
@app.post("/analyze")
async def analyze(file: UploadFile = File(...)):
    filename = file.filename
    ext = filename.split(".")[-1].lower()

    if ext not in ALLOWED_EXT:
        raise HTTPException(status_code=400, detail="지원하지 않는 파일 형식입니다.")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    save_path = UPLOAD_DIR / f"{timestamp}_{filename}"

    try:
        with save_path.open("wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

    # run_query 실행
    result = run_query(save_path, BM25_DIR)
    if not result:
        raise HTTPException(status_code=500, detail="run_query 결과 없음")

    # query 문서명
    query_doc = result["query_document"]

    # 백그라운드 작업 식별자
    task_id = f"{timestamp}_{filename}"

    # 작업 저장폴더 생성
    task_dir = BM25_DIR / task_id
    task_dir.mkdir(exist_ok=True)

    # run_query 결과 저장
    query_json_path = task_dir / "run_query.json"
    with open(query_json_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    # 🔥 백그라운드에서 LLM 실행
    def run_llm_async():
        try:
            print("[LLM] 스레드 시작")
            payload = build_llm_payload(result)
            print("[LLM] payload 생성 완료")

            md = summarize_with_llm_local(payload)
            print("[LLM] 요약 완료")

            save_markdown_summary(md, query_doc, out_dir=task_dir)
            print("[LLM] 마크다운 저장 완료")

            # 완료 표시 파일
            (task_dir / "LLM_DONE.flag").write_text("done")
            print("[LLM] LLM_DONE.flag 생성 완료")

        except Exception as e:
            print("[LLM] 오류 발생:", e)
            (task_dir / "LLM_DONE.flag").write_text(f"error: {e}")

    # 백그라운드Thread 시작
    threading.Thread(target=run_llm_async).start()

    # 클라이언트에 즉시 응답
    return {
        "status": "processing",
        "task_id": task_id,
        "run_query": result
    }


# =====================================================
# /llm_status — LLM 요약 완료 여부 확인
# =====================================================
@app.get("/llm_status")
def llm_status(task_id: str):
    flag_file = BM25_DIR / task_id / "LLM_DONE.flag"

    if not flag_file.exists():
        return {"done": False}

    content = flag_file.read_text().strip()
    if content.startswith("error"):
        return {"done": True, "error": content}

    return {"done": True}


# =====================================================
# /llm_summary — LLM 마크다운 반환
# =====================================================
@app.get("/llm_stream")
def llm_stream(task_id: str):
    def event_stream():
        task_dir = BM25_DIR / task_id
        flag_path = task_dir / "LLM_DONE.flag"

        while True:
            if flag_path.exists():
                yield "data: DONE\n\n"
                break
            time.sleep(0.3)  # 0.3초마다 체크 (지연 최소)
    return StreamingResponse(event_stream(), media_type="text/event-stream")


# =====================================================
# 서버 체크용
# =====================================================
@app.get("/")
def check():
    return {"status": "running"}

# =====================================================
# /llm_summary — LLM 마크다운 반환
# =====================================================
@app.get("/llm_summary")
def llm_summary(task_id: str):
    task_dir = BM25_DIR / task_id

    if not task_dir.exists():
        raise HTTPException(status_code=404, detail="해당 task_id 폴더가 없습니다.")

    # rag_pipeline_v7.save_markdown_summary 에서 저장한 md 파일 찾기
    md_files = list(task_dir.glob("*.md"))  # 또는 "rag_summary_*.md" 로 더 타이트하게
    if not md_files:
        raise HTTPException(status_code=404, detail="요약 Markdown 파일이 없습니다.")

    return FileResponse(md_files[0], media_type="text/markdown; charset=utf-8")
