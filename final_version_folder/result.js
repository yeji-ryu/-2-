// loading.html에서 저장해 둔 task_id 가져오기
const taskId = localStorage.getItem("task_id");

if (!taskId) {
    document.getElementById("markdown-output").textContent =
        "분석 작업 ID를 찾을 수 없습니다. 업로드 페이지에서 다시 시도해주세요.";
} else {
    // 서버의 /llm_summary 엔드포인트 호출
    fetch(`http://127.0.0.1:8000/llm_summary?task_id=${encodeURIComponent(taskId)}`)
        .then(res => {
            if (!res.ok) {
                throw new Error("요약 파일을 불러오지 못했습니다.");
            }
            return res.text();
        })
        .then(md => {
            document.getElementById("markdown-output").innerHTML =
                marked.parse(md);
        })
        .catch(err => {
            console.error("Markdown 불러오기 오류:", err);
            document.getElementById("markdown-output").textContent =
                "LLM 요약 파일을 불러올 수 없습니다.";
        });
}
