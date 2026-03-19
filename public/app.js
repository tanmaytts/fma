// ─── DOM Elements ────────────────────────────────────────────
const dropzone = document.getElementById("dropzone");
const fileInput = document.getElementById("file-input");
const uploadCard = document.getElementById("upload-card");
const previewCard = document.getElementById("preview-card");
const imagePreview = document.getElementById("image-preview");
const clearBtn = document.getElementById("clear-btn");
const previewBtn = document.getElementById("preview-btn");
const downloadBtn = document.getElementById("download-btn");
const loadingCard = document.getElementById("loading-card");
const tableCard = document.getElementById("table-card");
const tableWrapper = document.getElementById("table-wrapper");
const rowCount = document.getElementById("row-count");
const toast = document.getElementById("toast");
const toastMsg = document.getElementById("toast-msg");

let selectedFile = null;

// ─── Drag & Drop ─────────────────────────────────────────────
dropzone.addEventListener("click", () => fileInput.click());

dropzone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropzone.classList.add("drag-over");
});

dropzone.addEventListener("dragleave", () => {
    dropzone.classList.remove("drag-over");
});

dropzone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropzone.classList.remove("drag-over");
    const file = e.dataTransfer.files[0];
    if (file && file.type.startsWith("image/")) {
        handleFile(file);
    } else {
        showToast("Please drop an image file.");
    }
});

fileInput.addEventListener("change", () => {
    if (fileInput.files[0]) {
        handleFile(fileInput.files[0]);
    }
});

// ─── File Handling ───────────────────────────────────────────
function handleFile(file) {
    selectedFile = file;

    // Show image preview
    const reader = new FileReader();
    reader.onload = (e) => {
        imagePreview.src = e.target.result;
        uploadCard.classList.add("hidden");
        previewCard.classList.remove("hidden");
        tableCard.classList.add("hidden");
    };
    reader.readAsDataURL(file);
}

// ─── Clear ───────────────────────────────────────────────────
clearBtn.addEventListener("click", () => {
    selectedFile = null;
    fileInput.value = "";
    imagePreview.src = "";
    previewCard.classList.add("hidden");
    tableCard.classList.add("hidden");
    uploadCard.classList.remove("hidden");
});

// ─── Preview Data ────────────────────────────────────────────
previewBtn.addEventListener("click", async () => {
    if (!selectedFile) return;

    setLoading(true);
    disableButtons(true);

    try {
        const formData = new FormData();
        formData.append("image", selectedFile);

        const res = await fetch("/preview", { method: "POST", body: formData });
        const json = await res.json();

        if (!res.ok) {
            throw new Error(json.error || "Server error");
        }

        renderTable(json.data);
    } catch (err) {
        showToast(err.message || "Failed to extract data.");
    } finally {
        setLoading(false);
        disableButtons(false);
    }
});

// ─── Download Excel ──────────────────────────────────────────
downloadBtn.addEventListener("click", async () => {
    if (!selectedFile) return;

    setLoading(true);
    disableButtons(true);

    try {
        const formData = new FormData();
        formData.append("image", selectedFile);

        const res = await fetch("/convert", { method: "POST", body: formData });

        if (!res.ok) {
            const json = await res.json();
            throw new Error(json.error || "Server error");
        }

        // Download the blob
        const blob = await res.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "extracted_data.xlsx";
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);

        showToast("✅ Excel file downloaded!", false);
    } catch (err) {
        showToast(err.message || "Failed to download Excel.");
    } finally {
        setLoading(false);
        disableButtons(false);
    }
});

// ─── Render Table ────────────────────────────────────────────
function renderTable(data) {
    if (!data || data.length === 0) {
        showToast("No table data found in the image.");
        return;
    }

    const headers = Object.keys(data[0]);

    let html = "<table><thead><tr>";
    headers.forEach((h) => {
        html += `<th>${escapeHtml(h)}</th>`;
    });
    html += "</tr></thead><tbody>";

    data.forEach((row) => {
        html += "<tr>";
        headers.forEach((h) => {
            html += `<td>${escapeHtml(String(row[h] ?? ""))}</td>`;
        });
        html += "</tr>";
    });

    html += "</tbody></table>";

    tableWrapper.innerHTML = html;
    rowCount.textContent = `${data.length} row${data.length !== 1 ? "s" : ""}`;
    tableCard.classList.remove("hidden");

    // Smooth scroll to table
    tableCard.scrollIntoView({ behavior: "smooth", block: "start" });
}

// ─── Helpers ─────────────────────────────────────────────────
function setLoading(show) {
    loadingCard.classList.toggle("hidden", !show);
}

function disableButtons(disabled) {
    previewBtn.disabled = disabled;
    downloadBtn.disabled = disabled;
}

function escapeHtml(str) {
    const div = document.createElement("div");
    div.textContent = str;
    return div.innerHTML;
}

let toastTimer;
function showToast(message, isError = true) {
    toastMsg.textContent = message;
    toast.querySelector(".toast-icon").textContent = isError ? "⚠️" : "✅";
    toast.classList.remove("hidden");

    // Trigger reflow for animation
    void toast.offsetWidth;
    toast.classList.add("visible");

    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => {
        toast.classList.remove("visible");
        setTimeout(() => toast.classList.add("hidden"), 350);
    }, 4000);
}
