(() => {
  const dropzone = document.getElementById("dropzone");
  const fileInput = document.getElementById("fileInput");
  const browseBtn = document.getElementById("browseBtn");
  const filenameDisplay = document.getElementById("filenameDisplay");
  const errorBox = document.getElementById("errorBox");
  const loadingBox = document.getElementById("loadingBox");
  const uploadSection = document.getElementById("uploadSection");
  const resultsSection = document.getElementById("resultsSection");
  const resetBtn = document.getElementById("resetBtn");

  const tabs = document.querySelectorAll(".tab-btn");
  const panels = document.querySelectorAll(".tab-panel");

  tabs.forEach((btn) => {
    btn.addEventListener("click", () => {
      tabs.forEach((b) => b.classList.remove("active"));
      panels.forEach((p) => p.classList.remove("active"));
      btn.classList.add("active");
      document.getElementById(`panel-${btn.dataset.tab}`).classList.add("active");
    });
  });

  browseBtn.addEventListener("click", () => fileInput.click());

  fileInput.addEventListener("change", () => {
    if (fileInput.files.length > 0) {
      handleFile(fileInput.files[0]);
    }
  });

  ["dragenter", "dragover"].forEach((evt) => {
    dropzone.addEventListener(evt, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.add("dragover");
    });
  });

  ["dragleave", "drop"].forEach((evt) => {
    dropzone.addEventListener(evt, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.remove("dragover");
    });
  });

  dropzone.addEventListener("drop", (e) => {
    const files = e.dataTransfer.files;
    if (files.length > 0) {
      handleFile(files[0]);
    }
  });

  resetBtn.addEventListener("click", () => {
    resultsSection.hidden = true;
    uploadSection.hidden = false;
    fileInput.value = "";
    filenameDisplay.textContent = "";
    hideError();
  });

  function showError(message) {
    errorBox.textContent = message;
    errorBox.hidden = false;
  }

  function hideError() {
    errorBox.hidden = true;
    errorBox.textContent = "";
  }

  async function handleFile(file) {
    hideError();
    filenameDisplay.textContent = file.name;

    const ext = file.name.split(".").pop().toLowerCase();
    if (!["csv", "xlsx"].includes(ext)) {
      showError("지원하지 않는 파일 형식입니다. CSV(.csv) 또는 Excel(.xlsx) 파일만 업로드해주세요.");
      return;
    }
    if (file.size === 0) {
      showError("빈 파일입니다. 데이터가 포함된 파일을 업로드해주세요.");
      return;
    }

    const formData = new FormData();
    formData.append("file", file);

    loadingBox.hidden = false;
    try {
      const res = await fetch("/api/analyze", { method: "POST", body: formData });
      const data = await res.json();
      if (!res.ok) {
        showError(data.detail || "알 수 없는 오류가 발생했습니다.");
        return;
      }
      renderResults(data);
      uploadSection.hidden = true;
      resultsSection.hidden = false;
    } catch (err) {
      showError("서버와 통신하는 중 오류가 발생했습니다. 잠시 후 다시 시도해주세요.");
    } finally {
      loadingBox.hidden = true;
    }
  }

  function escapeHtml(value) {
    if (value === null || value === undefined) return "";
    return String(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  function fmtNum(v, decimals = 2) {
    if (v === null || v === undefined) return "-";
    return Number(v).toLocaleString("ko-KR", {
      minimumFractionDigits: 0,
      maximumFractionDigits: decimals,
    });
  }

  function renderResults(data) {
    renderSummary(data);
    renderPreview(data);
    renderInfo(data);
    renderNumeric(data);
    renderCategorical(data);

    tabs.forEach((b) => b.classList.remove("active"));
    panels.forEach((p) => p.classList.remove("active"));
    tabs[0].classList.add("active");
    document.getElementById("panel-preview").classList.add("active");
  }

  function renderSummary(data) {
    const el = document.getElementById("summaryCards");
    el.innerHTML = `
      <div class="summary-card"><div class="label">전체 행 수</div><div class="value">${fmtNum(data.shape.rows, 0)}</div></div>
      <div class="summary-card"><div class="label">전체 열 수</div><div class="value">${fmtNum(data.shape.cols, 0)}</div></div>
      <div class="summary-card"><div class="label">메모리 사용량</div><div class="value">${escapeHtml(data.memory_readable)}</div></div>
    `;
  }

  function renderPreview(data) {
    const table = document.getElementById("previewTable");
    const cols = data.preview_columns;
    let html = "<thead><tr>" + cols.map((c) => `<th>${escapeHtml(c)}</th>`).join("") + "</tr></thead><tbody>";
    if (data.preview_rows.length === 0) {
      html += `<tr><td colspan="${cols.length}">표시할 데이터가 없습니다.</td></tr>`;
    } else {
      for (const row of data.preview_rows) {
        html += "<tr>" + cols.map((c) => `<td>${escapeHtml(row[c] === null ? "" : row[c])}</td>`).join("") + "</tr>";
      }
    }
    html += "</tbody>";
    table.innerHTML = html;
  }

  function renderInfo(data) {
    const tbody = document.querySelector("#infoTable tbody");
    tbody.innerHTML = data.columns_info
      .map(
        (c) => `
      <tr>
        <td class="col-name">${escapeHtml(c.name)}</td>
        <td>${escapeHtml(c.dtype)}</td>
        <td class="num">${fmtNum(c.non_null, 0)}</td>
        <td class="num">${fmtNum(c.null_count, 0)}</td>
        <td class="num">${fmtNum(c.null_pct)}%</td>
      </tr>`
      )
      .join("");
  }

  function renderNumeric(data) {
    const tbody = document.querySelector("#numericTable tbody");
    const emptyNote = document.getElementById("numericEmptyNote");
    const table = document.getElementById("numericTable");
    if (data.numeric_stats.length === 0) {
      table.hidden = true;
      emptyNote.hidden = false;
      tbody.innerHTML = "";
      return;
    }
    table.hidden = false;
    emptyNote.hidden = true;
    tbody.innerHTML = data.numeric_stats
      .map(
        (c) => `
      <tr>
        <td class="col-name">${escapeHtml(c.name)}</td>
        <td class="num">${fmtNum(c.count, 0)}</td>
        <td class="num">${fmtNum(c.mean)}</td>
        <td class="num">${fmtNum(c.std)}</td>
        <td class="num">${fmtNum(c.min)}</td>
        <td class="num">${fmtNum(c.p25)}</td>
        <td class="num">${fmtNum(c.p50)}</td>
        <td class="num">${fmtNum(c.p75)}</td>
        <td class="num">${fmtNum(c.max)}</td>
      </tr>`
      )
      .join("");
  }

  function renderCategorical(data) {
    const tbody = document.querySelector("#categoricalTable tbody");
    const emptyNote = document.getElementById("categoricalEmptyNote");
    const table = document.getElementById("categoricalTable");
    if (data.categorical_stats.length === 0) {
      table.hidden = true;
      emptyNote.hidden = false;
      tbody.innerHTML = "";
      return;
    }
    table.hidden = false;
    emptyNote.hidden = true;
    tbody.innerHTML = data.categorical_stats
      .map(
        (c) => `
      <tr>
        <td class="col-name">${escapeHtml(c.name)}</td>
        <td class="num">${fmtNum(c.count, 0)}</td>
        <td class="num">${fmtNum(c.unique, 0)}</td>
        <td>${escapeHtml(c.top)}</td>
        <td class="num">${fmtNum(c.freq, 0)}</td>
      </tr>`
      )
      .join("");
  }
})();
