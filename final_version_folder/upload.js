document.addEventListener("DOMContentLoaded", () => {
    const browseBtn = document.querySelector(".browse-btn");
    const fileInput = document.getElementById("file-input");
    const fileList = document.querySelector(".file-list");
    const dropZone = document.querySelector(".drop-zone");
    const cancelBtn = document.querySelector(".cancel-btn");
    const nextBtn = document.querySelector(".next-btn");

    const allowedExtensions = ["hwpx", "docx", "pdf"];

    // Browse files 클릭 → 숨겨진 file input 클릭
    browseBtn.addEventListener("click", () => fileInput.click());

    // 파일 선택
    fileInput.addEventListener("change", (e) => {
        handleFiles(e.target.files);
    });

    // Drag & Drop
    dropZone.addEventListener("dragover", (e) => {
        e.preventDefault();
        dropZone.classList.add("dragover");
    });

    dropZone.addEventListener("dragleave", () => {
        dropZone.classList.remove("dragover");
    });

    dropZone.addEventListener("drop", (e) => {
        e.preventDefault();
        dropZone.classList.remove("dragover");
        handleFiles(e.dataTransfer.files);
    });

    // 파일 처리 함수 (확장자 검사 포함)
    function handleFiles(files) {
        if (files.length === 0) return;

        const file = files[0];
        const ext = file.name.split('.').pop().toLowerCase();

        if (!allowedExtensions.includes(ext)) {
            alert("허용되지 않은 파일 형식입니다.\n(.hwpx, .docx, .pdf만 업로드 가능합니다.)");
            fileInput.value = "";
            return;
        }

        displayFile(file);
    }

    // 파일 목록 표시
    function displayFile(file) {
        fileList.innerHTML = ""; // 초기화

        const item = document.createElement("div");
        item.classList.add("file-item");
        item.innerHTML = `
            <div class="file-icon">📄</div>
            <div class="file-info">
                <p class="file-name">${file.name}</p>
                <p class="file-size">${(file.size / 1024).toFixed(1)} KB</p>
            </div>
        `;
        fileList.appendChild(item);
    }

    // Cancel → 초기화
    cancelBtn.addEventListener("click", () => {
        fileInput.value = "";
        fileList.innerHTML = "";
    });

    // -------------------------------
    // ⭐ Next 버튼 → 서버로 파일 업로드
    // -------------------------------
    nextBtn.addEventListener("click", () => {

        if (nextBtn.disabled) return;
        nextBtn.disabled = true;

        if (!fileInput.files.length) {
            alert("업로드할 파일을 선택해주세요.");
            nextBtn.disabled = false;
            return;
        }

        const file = fileInput.files[0];
        const reader = new FileReader();

        reader.onload = function () {
            localStorage.setItem("uploaded_file_name", file.name);
            localStorage.setItem("uploaded_file_data", reader.result);
            window.location.href = "loading.html";
        };

        reader.readAsDataURL(file);
    });
});
