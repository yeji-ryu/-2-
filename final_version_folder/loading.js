/* 원그래프 생성 함수 */
function createCircleChart(percent) {
    const radius = 60;
    const circumference = 2 * Math.PI * radius;
    const offset = circumference * (1 - percent / 100);

    return `
        <div class="circle-chart">
            <svg width="150" height="150" viewBox="0 0 150 150">
                <g transform="rotate(-90 75 75)">
                    <circle class="circle-bg" cx="75" cy="75" r="${radius}"></circle>
                    <circle class="circle-progress"
                        cx="75" cy="75" r="${radius}"
                        stroke-dasharray="${circumference}"
                        stroke-dashoffset="${offset}">
                    </circle>
                </g>
                <text x="75" y="75" class="circle-text">${percent}%</text>
            </svg>
        </div>`;
}

async function startAnalysis() {

    const fileName = localStorage.getItem("uploaded_file_name");
    const base64data = localStorage.getItem("uploaded_file_data");

    const binary = atob(base64data.split(",")[1]);
    let bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);

    let blob = new Blob([bytes]);
    let form = new FormData();
    form.append("file", blob, fileName);

    // ① /analyze 호출 (run_query 즉시 반환)
    const res = await fetch("http://127.0.0.1:8000/analyze", {
        method: "POST",
        body: form
    });
    const data = await res.json();

    // 태스크 ID 저장
    const taskId = data.task_id;
    localStorage.setItem("task_id", taskId);

    // ② run_query 결과 표시 UI 업데이트
    const top3 = data.run_query.results.slice(0, 3);

    // 로딩 메시지 숨김
    document.querySelector(".spinner").style.display = "none";
    document.querySelector(".loading-title").style.display = "none";
    document.querySelectorAll(".loading-desc").forEach(el => el.style.display = "none");

    // run_query 결과 보이기
    document.getElementById("runquery-section").style.display = "block";
    document.getElementById("llm-loading").style.display = "block";

    document.getElementById("runquery-result").innerHTML =
        top3.map(doc => `
            <div class="doc-card">
                ${createCircleChart(doc.hybrid_pct.toFixed(0))}
                <div style="margin-top:10px;">${doc.doc_id}</div>
            </div>
        `).join("");

    // ③ LLM 상태 체크
    const evtSource = new EventSource(`http://127.0.0.1:8000/llm_stream?task_id=${taskId}`);

    evtSource.onmessage = (event) => {
        if (event.data === "DONE") {
            document.getElementById("llm-loading").style.display = "none";
            const btn = document.getElementById("go-result-btn");
            btn.style.display = "inline-block";
            btn.onclick = () => window.location.href = "result.html";

            evtSource.close(); // 이벤트 스트림 종료
        }
    };
}

startAnalysis();
