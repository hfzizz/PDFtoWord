/**
 * PDF-to-Word Web UI — Frontend Application
 *
 * Handles: drag-and-drop upload, file list management, conversion
 * progress via SSE, result display, preview pagination, and history.
 */

// ── DOM References ──────────────────────────────────────────────────
const dropZone       = document.getElementById("dropZone");
const fileInput      = document.getElementById("fileInput");
const fileList       = document.getElementById("fileList");
const convertBtn     = document.getElementById("convertBtn");
const emptyState     = document.getElementById("emptyState");
const activeJobs     = document.getElementById("activeJobs");
const previewArea    = document.getElementById("previewArea");
const previewTitle   = document.getElementById("previewTitle");
const previewImage   = document.getElementById("previewImage");
const prevPageBtn    = document.getElementById("prevPage");
const nextPageBtn    = document.getElementById("nextPage");
const pageIndicator  = document.getElementById("pageIndicator");
const historyBody    = document.getElementById("historyBody");
const refreshHistory = document.getElementById("refreshHistory");

// ── State ───────────────────────────────────────────────────────────
let pendingFiles = [];          // Files selected but not yet uploaded
let activeJobMap = new Map();   // jobId -> {sse, data}
let previewJob   = null;        // jobId currently being previewed
let previewPage  = 0;
let previewTotal = 1;

// ── Drag & Drop ─────────────────────────────────────────────────────
dropZone.addEventListener("click", () => fileInput.click());

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
  const files = [...e.dataTransfer.files].filter(f => f.name.toLowerCase().endsWith(".pdf"));
  addFiles(files);
});

fileInput.addEventListener("change", () => {
  addFiles([...fileInput.files]);
  fileInput.value = "";
});

// ── File Management ─────────────────────────────────────────────────
function addFiles(files) {
  for (const f of files) {
    if (f.size > 50 * 1024 * 1024) {
      alert(`"${f.name}" exceeds the 50 MB limit.`);
      continue;
    }
    // Avoid duplicates.
    if (pendingFiles.some(p => p.name === f.name && p.size === f.size)) continue;
    pendingFiles.push(f);
  }
  renderFileList();
}

function removeFile(index) {
  pendingFiles.splice(index, 1);
  renderFileList();
}

function renderFileList() {
  if (pendingFiles.length === 0) {
    fileList.hidden = true;
    convertBtn.disabled = true;
    return;
  }
  fileList.hidden = false;
  convertBtn.disabled = false;
  fileList.innerHTML = pendingFiles.map((f, i) => `
    <div class="file-item">
      <span class="name" title="${esc(f.name)}">${esc(f.name)}</span>
      <span class="size">${formatSize(f.size)}</span>
      <button class="remove" onclick="removeFile(${i})" title="Remove">&times;</button>
    </div>
  `).join("");
}

// ── Convert ─────────────────────────────────────────────────────────
convertBtn.addEventListener("click", startConversion);

async function startConversion() {
  if (pendingFiles.length === 0) return;
  convertBtn.disabled = true;

  const settings = {
    ocr: document.getElementById("settingOcr").checked,
    skip_watermarks: document.getElementById("settingWatermarks").checked,
    password: document.getElementById("settingPassword").value || undefined,
  };

  // Hide empty state.
  emptyState.hidden = true;
  activeJobs.hidden = false;

  for (const file of pendingFiles) {
    const formData = new FormData();
    formData.append("file", file);
    formData.append("settings", JSON.stringify(settings));

    try {
      const resp = await fetch("/api/upload", { method: "POST", body: formData });
      if (!resp.ok) {
        const err = await resp.json();
        alert(`Upload failed for "${file.name}": ${err.error || resp.statusText}`);
        continue;
      }
      const { job_id, filename } = await resp.json();
      trackJob(job_id, filename);
    } catch (err) {
      alert(`Network error uploading "${file.name}": ${err.message}`);
    }
  }

  pendingFiles = [];
  renderFileList();
}

// ── Job Tracking (SSE) ──────────────────────────────────────────────
function trackJob(jobId, filename) {
  const sse = new EventSource(`/api/status/${jobId}/stream`);
  activeJobMap.set(jobId, { sse, data: null });

  sse.onmessage = (event) => {
    const data = JSON.parse(event.data);
    activeJobMap.get(jobId).data = data;
    renderJobs();

    if (data.status === "complete" || data.status === "failed") {
      sse.close();
      loadHistory();
    }
  };

  sse.onerror = () => {
    sse.close();
    // Fall back to polling.
    pollJob(jobId);
  };
}

async function pollJob(jobId) {
  const entry = activeJobMap.get(jobId);
  if (!entry) return;
  try {
    const resp = await fetch(`/api/status/${jobId}`);
    const data = await resp.json();
    entry.data = data;
    renderJobs();
    if (data.status !== "complete" && data.status !== "failed") {
      setTimeout(() => pollJob(jobId), 1000);
    } else {
      loadHistory();
    }
  } catch { /* retry */ setTimeout(() => pollJob(jobId), 2000); }
}

// ── Render Active Jobs ──────────────────────────────────────────────
function renderJobs() {
  const cards = [];
  for (const [jobId, { data }] of activeJobMap) {
    if (!data) continue;
    const statusCls = `status-${data.status}`;
    const progressCls = data.status === "complete" ? "complete" : "";

    let actionsHtml = "";
    if (data.status === "complete") {
      actionsHtml += `<button class="btn btn-success btn-small" onclick="downloadJob('${jobId}')">&#11015; Download .docx</button>`;
      actionsHtml += `<button class="btn btn-small" onclick="showPreview('${jobId}', ${data.page_count})">&#128196; Preview PDF</button>`;
    }
    if (data.status === "complete" || data.status === "failed") {
      actionsHtml += `<button class="btn btn-small" onclick="removeJob('${jobId}')">&times; Remove</button>`;
    }

    let qualityHtml = "";
    if (data.quality_report) {
      const qr = data.quality_report;
      const score = qr.quality_score ?? "?";
      const level = (qr.quality_level ?? "").toLowerCase();
      qualityHtml = `<span class="quality-badge quality-${level}">${score}/100 ${qr.quality_level || ""}</span>`;
      if (qr.metrics) {
        const m = qr.metrics;
        qualityHtml += `<div class="metrics">
          <span class="metric"><strong>${m.paragraphs ?? 0}</strong> paragraphs</span>
          <span class="metric"><strong>${m.tables ?? 0}</strong> tables</span>
          <span class="metric"><strong>${m.images ?? 0}</strong> images</span>
          <span class="metric"><strong>${m.headings ?? 0}</strong> headings</span>
          <span class="metric"><strong>${m.formatted_runs ?? 0}</strong> styled runs</span>
        </div>`;
      }
    }

    let errorHtml = "";
    if (data.error) {
      errorHtml = `<div class="error-text">${esc(data.error)}</div>`;
    }

    cards.push(`
      <div class="job-card" id="job-${jobId}">
        <div class="job-header">
          <span class="job-name" title="${esc(data.original_filename)}">${esc(data.original_filename)}</span>
          <span class="job-status ${statusCls}">${data.status}</span>
        </div>
        <div class="progress-bar-track">
          <div class="progress-bar-fill ${progressCls}" style="width:${data.progress}%"></div>
        </div>
        <div class="job-message">${esc(data.message || "")}</div>
        ${errorHtml}
        ${qualityHtml}
        <div class="job-actions">${actionsHtml}</div>
      </div>
    `);
  }
  activeJobs.innerHTML = cards.join("");
}

// ── Download ────────────────────────────────────────────────────────
function downloadJob(jobId) {
  window.open(`/api/download/${jobId}`, "_blank");
}

// ── Remove Job ──────────────────────────────────────────────────────
async function removeJob(jobId) {
  activeJobMap.delete(jobId);
  renderJobs();
  if (activeJobMap.size === 0) {
    emptyState.hidden = false;
    activeJobs.hidden = true;
  }
  try { await fetch(`/api/job/${jobId}`, { method: "DELETE" }); } catch {}
  loadHistory();
}

// ── Preview ─────────────────────────────────────────────────────────
function showPreview(jobId, pageCount) {
  previewJob = jobId;
  previewPage = 0;
  previewTotal = pageCount || 1;
  previewArea.hidden = false;
  renderPreview();
  loadDocxPreview(jobId);
}

function renderPreview() {
  pageIndicator.textContent = `Page ${previewPage + 1} / ${previewTotal}`;
  prevPageBtn.disabled = previewPage <= 0;
  nextPageBtn.disabled = previewPage >= previewTotal - 1;
  previewImage.innerHTML = `<img src="/api/preview/${previewJob}/${previewPage}" alt="Page ${previewPage + 1}">`;
}

async function loadDocxPreview(jobId) {
  const docxDiv = document.getElementById("docxPreview");
  docxDiv.innerHTML = `<p class="muted">Loading DOCX preview...</p>`;
  try {
    const resp = await fetch(`/api/docx-preview/${jobId}`);
    if (!resp.ok) { docxDiv.innerHTML = `<p class="muted">Preview unavailable.</p>`; return; }
    const data = await resp.json();

    let html = "";

    // Render paragraphs and tables interleaved (paragraphs first, then tables)
    for (const p of data.paragraphs) {
      const style = p.style.toLowerCase();
      if (style.startsWith("heading")) {
        const level = style.replace("heading ", "").replace("heading", "1");
        html += `<div class="docx-heading h${level}">${esc(p.text)}</div>`;
      } else {
        html += `<div class="docx-para">${esc(p.text)}</div>`;
      }
    }

    for (const t of data.tables) {
      html += `<table class="docx-table">`;
      for (let r = 0; r < t.rows.length; r++) {
        html += `<tr>`;
        const tag = r === 0 ? "th" : "td";
        for (const cell of t.rows[r]) {
          html += `<${tag}>${esc(cell)}</${tag}>`;
        }
        html += `</tr>`;
      }
      html += `</table>`;
    }

    if (!html) html = `<p class="muted">No content extracted.</p>`;
    docxDiv.innerHTML = html;
  } catch {
    docxDiv.innerHTML = `<p class="muted">Preview failed.</p>`;
  }
}

prevPageBtn.addEventListener("click", () => {
  if (previewPage > 0) { previewPage--; renderPreview(); }
});
nextPageBtn.addEventListener("click", () => {
  if (previewPage < previewTotal - 1) { previewPage++; renderPreview(); }
});

// ── History ─────────────────────────────────────────────────────────
async function loadHistory() {
  try {
    const resp = await fetch("/api/history");
    const jobs = await resp.json();
    if (jobs.length === 0) {
      historyBody.innerHTML = `<p class="muted">No conversions yet.</p>`;
      return;
    }
    historyBody.innerHTML = jobs.map(j => {
      const statusCls = `status-${j.status}`;
      const time = j.completed_at
        ? new Date(j.completed_at).toLocaleString()
        : new Date(j.created_at).toLocaleString();
      const score = j.quality_report?.quality_score;
      const scoreHtml = score != null ? `<span class="quality-badge quality-${(j.quality_report?.quality_level || "").toLowerCase()}">${score}/100</span>` : "";
      let actions = "";
      if (j.status === "complete") {
        actions = `<button class="btn btn-small" onclick="downloadJob('${j.id}')">Download</button>`;
      }
      return `<div class="history-row">
        <span class="job-status ${statusCls}" style="font-size:.72rem">${j.status}</span>
        <span class="name" title="${esc(j.original_filename)}">${esc(j.original_filename)}</span>
        ${scoreHtml}
        <span class="time">${time}</span>
        ${actions}
      </div>`;
    }).join("");
  } catch {}
}

refreshHistory.addEventListener("click", loadHistory);

// ── Utilities ───────────────────────────────────────────────────────
function esc(str) {
  const d = document.createElement("div");
  d.textContent = str || "";
  return d.innerHTML;
}

function formatSize(bytes) {
  if (bytes < 1024) return bytes + " B";
  if (bytes < 1048576) return (bytes / 1024).toFixed(1) + " KB";
  return (bytes / 1048576).toFixed(1) + " MB";
}

// ── Init ────────────────────────────────────────────────────────────
loadHistory();
