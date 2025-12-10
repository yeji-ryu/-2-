// upload.js (수정된 전체 코드)

// HTML에서 필요한 요소들을 가져옵니다.
const dropZone = document.querySelector(".drop-zone");
const fileInput = document.getElementById("file-input");
const browseBtn = document.querySelector(".browse-btn");
const fileList = document.querySelector(".file-list");
const cancelBtn = document.querySelector(".cancel-btn");

// 'Browse files' 버튼 클릭 시 숨겨진 파일 입력창을 클릭합니다.
browseBtn.addEventListener("click", () => {
    // 이미 파일이 있다면 추가로 열지 않음
    if (fileList.children.length > 0) {
        alert("하나의 파일만 업로드할 수 있습니다. 기존 파일을 삭제하고 시도해주세요.");
        return;
    }
    fileInput.click();
});

// 파일 입력창에 변화가 생기면(파일 선택 시) 파일 처리 함수를 호출합니다.
fileInput.addEventListener("change", (e) => {
    // 파일이 선택되었을 때만 처리
    if (e.target.files.length > 0) {
        handleFiles(e.target.files);
    }
});

// 드래그 앤 드롭 관련 이벤트 처리
// 1. 드래그해서 영역 위에 올렸을 때
dropZone.addEventListener("dragover", (e) => {
    e.preventDefault();
    if (fileList.children.length > 0) return; // 파일이 있으면 드래그 효과 없음
    dropZone.classList.add("dragover");
});

// 2. 드래그해서 영역 밖으로 나갔을 때
dropZone.addEventListener("dragleave", (e) => {
    e.preventDefault();
    dropZone.classList.remove("dragover");
});

// 3. 드롭했을 때
dropZone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropZone.classList.remove("dragover");
    
    // ✨ 핵심: 이미 파일이 있다면 추가하지 않고 경고창 표시
    if (fileList.children.length > 0) {
        alert("하나의 파일만 업로드할 수 있습니다.");
        return;
    }

    const files = e.dataTransfer.files;
    if (files.length) {
        fileInput.files = files;
        handleFiles(files);
    }
});

// 'Cancel' 버튼 클릭 시 파일 목록을 초기화합니다.
cancelBtn.addEventListener("click", () => {
    removeFile();
});

// 파일 처리 함수
function handleFiles(files) {
    // ✨ 핵심: 함수 시작 전에도 파일이 이미 있는지 한번 더 확인
    if (fileList.children.length > 0) {
        alert("하나의 파일만 업로드할 수 있습니다.");
        // 파일 선택창을 다시 열었을 때, 기존 파일이 있는데 새 파일을 선택한 경우
        // 파일 입력창을 비워줘야 다음에 같은 파일을 또 선택할 수 있음
        fileInput.value = "";
        return;
    }

    const file = files[0];
    if (file) {
        const fileItemHTML = `
            <div class="file-item">
                <img src="images/icon-file.svg" alt="file icon">
                <div class="file-info">
                    <strong>${file.name}</strong>
                    <span>${(file.size / 1024 / 1024).toFixed(2)} MB</span>
                </div>
                <button class="remove-btn" onclick="removeFile()">&times;</button>
            </div>
        `;
        fileList.innerHTML = fileItemHTML;
    }
}

// 파일 제거 함수
function removeFile() {
    fileList.innerHTML = "";
    fileInput.value = "";
}