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
const convertHint    = document.getElementById("convertHint");
const feedbackArea   = document.getElementById("feedbackArea");
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
const styleEditor = document.getElementById("styleEditor");
const styleEditorJob = document.getElementById("styleEditorJob");
const stylePrompt = document.getElementById("stylePrompt");
const styleGeminiKey = document.getElementById("styleGeminiKey");
const styleModel = document.getElementById("styleModel");
const applyStyleBtn = document.getElementById("applyStyleBtn");
const settingAiEnhance = document.getElementById("settingAiEnhance");
const settingGeminiKey = document.getElementById("settingGeminiKey");
const aiKeySection = document.getElementById("aiKeySection");
const aiKeyHint = document.getElementById("aiKeyHint");

// ── State ───────────────────────────────────────────────────────────
let pendingFiles = [];          // Files selected but not yet uploaded
let activeJobMap = new Map();   // jobId -> {sse, data}
let previewJob   = null;        // jobId currently being previewed
let previewPage  = 0;
let previewTotal = 1;
let feedbackTimer = null;
let styleTargetJobId = null;

// ── Drag & Drop ─────────────────────────────────────────────────────
dropZone.addEventListener("click", () => fileInput.click());
dropZone.addEventListener("keydown", (e) => {
  if (e.key === "Enter" || e.key === " ") {
    e.preventDefault();
    fileInput.click();
  }
});

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
  let rejectedLarge = 0;
  let skippedDup = 0;
  let added = 0;
  for (const f of files) {
    if (f.size > 50 * 1024 * 1024) {
      rejectedLarge += 1;
      continue;
    }
    // Avoid duplicates.
    if (pendingFiles.some(p => p.name === f.name && p.size === f.size)) {
      skippedDup += 1;
      continue;
    }
    pendingFiles.push(f);
    added += 1;
  }
  if (rejectedLarge > 0) {
    showFeedback(`${rejectedLarge} file(s) were skipped (over 50 MB).`, "warn");
  } else if (skippedDup > 0 && added === 0) {
    showFeedback("All selected files were already in the list.", "info");
  } else if (added > 0) {
    showFeedback(`${added} file(s) added for conversion.`, "success", 2500);
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
    updateConvertState();
    return;
  }
  fileList.hidden = false;
  fileList.innerHTML = pendingFiles.map((f, i) => `
    <div class="file-item">
      <span class="name" title="${esc(f.name)}">${esc(f.name)}</span>
      <span class="size">${formatSize(f.size)}</span>
      <button class="remove" data-remove-index="${i}" title="Remove">&times;</button>
    </div>
  `).join("");
  updateConvertState();
}

fileList.addEventListener("click", (event) => {
  const btn = event.target.closest("button[data-remove-index]");
  if (!btn) return;
  const idx = Number(btn.getAttribute("data-remove-index"));
  if (!Number.isNaN(idx)) {
    removeFile(idx);
  }
});

// ── Convert ─────────────────────────────────────────────────────────
convertBtn.addEventListener("click", startConversion);

async function startConversion() {
  if (pendingFiles.length === 0) {
    updateConvertState();
    return;
  }
  
  const aiEnhanceEnabled = settingAiEnhance.checked;
  const geminiKey = settingGeminiKey.value.trim();
  
  // Validate API key if AI enhancement is enabled
  if (aiEnhanceEnabled && !geminiKey) {
    updateConvertState();
    showFeedback("Gemini API Key is required when AI Enhancement is enabled.", "error", 6000);
    return;
  }
  
  convertBtn.disabled = true;
  convertHint.textContent = "Uploading and starting conversion…";
  showFeedback("Starting conversion job(s)…", "info", 3000);

  const settings = {
    ocr: document.getElementById("settingOcr").checked,
    skip_watermarks: document.getElementById("settingWatermarks").checked,
    use_pdf2docx_lib: document.getElementById("settingUsePdf2docxLib").checked,
    password: document.getElementById("settingPassword").value || undefined,
    quality_mode: document.getElementById("settingQualityMode").value,
    quality_gate: document.getElementById("settingQualityGate").value,
    quality_engine_fallback: document.getElementById("settingQualityFallback").checked,
    ai_enabled: aiEnhanceEnabled,
    ai_compare: aiEnhanceEnabled, // compatibility with older backend code
    gemini_api_key: aiEnhanceEnabled ? geminiKey : undefined,
    ai_strategy: aiEnhanceEnabled ? document.getElementById("settingAiStrategy").value : undefined,
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
        let errText = resp.statusText;
        try {
          const err = await resp.json();
          errText = err.error || errText;
        } catch {
          // keep fallback status text
        }
        showFeedback(`Upload failed for "${file.name}": ${errText}`, "error", 7000);
        continue;
      }
      const { job_id, filename } = await resp.json();
      trackJob(job_id, filename, settings);
    } catch (err) {
      showFeedback(`Network error uploading "${file.name}": ${err.message}`, "error", 7000);
    }
  }

  pendingFiles = [];
  renderFileList();
}

// ── Job Tracking (SSE) ──────────────────────────────────────────────
function trackJob(jobId, filename, settings) {
  const sse = new EventSource(`/api/status/${jobId}/stream`);
  activeJobMap.set(jobId, { sse, data: null, settings });

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
  for (const [jobId, { data, settings }] of activeJobMap) {
    if (!data) continue;
    const statusCls = `status-${data.status}`;
    const progressCls = data.status === "complete" ? "complete" : "";
    const meta = data.settings_summary || {
      engine: settings?.use_pdf2docx_lib ? "pdf2docx_lib" : "custom",
      ai_enabled: !!settings?.ai_enabled,
      ai_strategy: settings?.ai_strategy || "B",
      quality_mode: settings?.quality_mode || "basic",
      quality_gate: settings?.quality_gate || "warn",
    };

    const metaHtml = `
      <div class="job-meta">
        <span class="meta-chip">Engine: ${meta.engine === "pdf2docx_lib" ? "pdf2docx library" : "custom"}</span>
        <span class="meta-chip">AI: ${meta.ai_enabled ? "on" : "off"}</span>
        <span class="meta-chip">Quality: ${meta.quality_mode || "basic"}/${meta.quality_gate || "warn"}</span>
        <span class="meta-chip">Strategy: ${meta.ai_strategy || "B"}</span>
      </div>
    `;

    let actionsHtml = "";
    if (data.status === "complete") {
      actionsHtml += `<button class="btn btn-success btn-small" onclick="downloadJob('${jobId}')">&#11015; Download .docx</button>`;
      actionsHtml += `<button class="btn btn-small" onclick="showPreview('${jobId}', ${data.page_count})">&#128196; Preview PDF</button>`;
      actionsHtml += `<button class="btn btn-small" onclick="openStyleEditor('${jobId}')">🎨 Style</button>`;
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
        ${metaHtml}
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

function openStyleEditor(jobId) {
  styleTargetJobId = jobId;
  styleEditor.hidden = false;
  styleEditorJob.textContent = `Target job: ${jobId}`;
  if (!stylePrompt.value.trim()) {
    stylePrompt.value = "Make the document look modern and professional with improved heading hierarchy and clean table headers.";
  }
  stylePrompt.focus();
}

async function applyStylePrompt() {
  if (!styleTargetJobId) {
    showFeedback("Select a completed job first using the Style button.", "warn", 5000);
    return;
  }
  const prompt = stylePrompt.value.trim();
  if (!prompt) {
    showFeedback("Enter a style prompt first.", "warn", 4000);
    return;
  }

  applyStyleBtn.disabled = true;
  showFeedback("Applying style changes…", "info", 3000);
  try {
    const resp = await fetch(`/api/style/${styleTargetJobId}/apply`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        prompt,
        gemini_api_key: styleGeminiKey.value.trim() || undefined,
        model: styleModel.value.trim() || undefined,
      }),
    });
    const data = await resp.json();
    if (!resp.ok) {
      showFeedback(data.error || "Style update failed.", "error", 7000);
      return;
    }

    showFeedback(data.summary || "Style updated.", "success", 5000);

    if (activeJobMap.has(styleTargetJobId)) {
      const entry = activeJobMap.get(styleTargetJobId);
      if (entry?.data) {
        entry.data.quality_report = data.quality_report || entry.data.quality_report;
        entry.data.message = "Style updated";
      }
      renderJobs();
    }
    if (previewJob === styleTargetJobId) {
      await loadDocxPreview(styleTargetJobId);
    }
    loadHistory();
  } catch (err) {
    showFeedback(`Style request failed: ${err.message}`, "error", 7000);
  } finally {
    applyStyleBtn.disabled = false;
  }
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
    // Fetch the actual .docx binary from the download endpoint
    const resp = await fetch(`/api/download/${jobId}`);
    if (!resp.ok) {
      docxDiv.innerHTML = `<p class="muted">Preview unavailable.</p>`;
      return;
    }
    const arrayBuffer = await resp.arrayBuffer();

    // Clear container and render with docx-preview library
    docxDiv.innerHTML = "";
    await docx.renderAsync(arrayBuffer, docxDiv, null, {
      className: "docx",
      inWrapper: true,
      ignoreWidth: false,
      ignoreHeight: false,
      ignoreFonts: false,
      breakPages: true,
      ignoreLastRenderedPageBreak: false,
      renderHeaders: true,
      renderFooters: true,
      renderFootnotes: true,
      renderEndnotes: true,
      useBase64URL: true,
    });

    // Auto-scale rendered pages to fit the container width
    const wrapper = docxDiv.querySelector(".docx-wrapper");
    if (wrapper) {
      const sections = wrapper.querySelectorAll("section.docx");
      const containerW = docxDiv.clientWidth - 20;
      sections.forEach((sec) => {
        const pageW = sec.offsetWidth;
        if (pageW > containerW && pageW > 0) {
          const scale = containerW / pageW;
          sec.style.transform = `scale(${scale})`;
          sec.style.transformOrigin = "top left";
          sec.style.marginBottom = `-${sec.offsetHeight * (1 - scale)}px`;
        }
      });
    }
  } catch (err) {
    console.error("DOCX preview error:", err);
    docxDiv.innerHTML = `<p class="muted">Preview failed: ${esc(err.message || "unknown error")}</p>`;
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
      const engine = j.settings_summary?.engine === "pdf2docx_lib" ? "pdf2docx" : "custom";
      const engineHtml = `<span class="meta-chip">${engine}</span>`;
      let actions = "";
      if (j.status === "complete") {
        actions = `<button class="btn btn-small" onclick="downloadJob('${j.id}')">Download</button>`;
      }
      return `<div class="history-row">
        <span class="job-status ${statusCls}" style="font-size:.72rem">${j.status}</span>
        <span class="name" title="${esc(j.original_filename)}">${esc(j.original_filename)}</span>
        ${engineHtml}
        ${scoreHtml}
        <span class="time">${time}</span>
        ${actions}
      </div>`;
    }).join("");
  } catch {
    showFeedback("Unable to refresh conversion history right now.", "warn", 5000);
  }
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

function showFeedback(message, type = "info", durationMs = 4500) {
  if (feedbackTimer) {
    clearTimeout(feedbackTimer);
    feedbackTimer = null;
  }
  feedbackArea.hidden = false;
  feedbackArea.className = `feedback ${type}`;
  feedbackArea.textContent = message;
  if (durationMs > 0) {
    feedbackTimer = setTimeout(() => {
      feedbackArea.hidden = true;
      feedbackArea.textContent = "";
    }, durationMs);
  }
}

function updateConvertState() {
  const hasFiles = pendingFiles.length > 0;
  const aiEnabled = settingAiEnhance.checked;
  const hasGeminiKey = settingGeminiKey.value.trim().length > 0;
  const aiValid = !aiEnabled || hasGeminiKey;

  convertBtn.disabled = !hasFiles || !aiValid;
  settingGeminiKey.classList.toggle("invalid", aiEnabled && !hasGeminiKey);
  aiKeyHint.hidden = !(aiEnabled && !hasGeminiKey);

  if (!hasFiles) {
    convertHint.textContent = "Add at least one PDF file to start conversion.";
  } else if (!aiValid) {
    convertHint.textContent = "Gemini API Key is required when AI Enhancement is enabled.";
  } else {
    convertHint.textContent = "Ready to convert.";
  }
}

// ── Init ────────────────────────────────────────────────────────────
loadHistory();

// Toggle API key section visibility when AI enhancement checkbox changes
settingAiEnhance.addEventListener("change", (e) => {
  aiKeySection.style.display = e.target.checked ? "block" : "none";
  updateConvertState();
});

settingGeminiKey.addEventListener("input", updateConvertState);
aiKeySection.style.display = settingAiEnhance.checked ? "block" : "none";
applyStyleBtn.addEventListener("click", applyStylePrompt);
updateConvertState();
