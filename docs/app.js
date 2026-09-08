(() => {
  const MAX_FILE_SIZE = 20 * 1024 * 1024; // 20MB

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

  // ---------- File parsing (runs entirely in the browser) ----------

  function decodeText(buffer) {
    // Try UTF-8 first (strict), then fall back to Korean encodings so
    // CSV files saved from Excel on Windows (EUC-KR/CP949) still work.
    const attempts = ["utf-8", "euc-kr"];
    for (const enc of attempts) {
      try {
        const decoder = new TextDecoder(enc, { fatal: true });
        return decoder.decode(buffer);
      } catch (e) {
        continue;
      }
    }
    return new TextDecoder("utf-8").decode(buffer);
  }

  function normalizeValue(v) {
    if (v === undefined || v === null) return null;
    if (typeof v === "string" && v.trim() === "") return null;
    return v;
  }

  async function parseCsv(file) {
    const buffer = await file.arrayBuffer();
    const text = decodeText(buffer);
    const parsed = Papa.parse(text, {
      header: true,
      dynamicTyping: true,
      skipEmptyLines: true,
    });
    if (parsed.errors && parsed.errors.length > 0) {
      const fatal = parsed.errors.some((e) => e.type !== "FieldMismatch");
      if (fatal && (!parsed.data || parsed.data.length === 0)) {
        throw new Error("parse");
      }
    }
    const columns = parsed.meta.fields || [];
    const rows = parsed.data.map((row) => {
      const out = {};
      for (const col of columns) out[col] = normalizeValue(row[col]);
      return out;
    });
    return { columns, rows };
  }

  async function parseExcel(file) {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const records = XLSX.utils.sheet_to_json(sheet, { defval: null, raw: true });
    const columns = records.length > 0 ? Object.keys(records[0]) : (XLSX.utils.sheet_to_json(sheet, { header: 1 })[0] || []);
    const rows = records.map((row) => {
      const out = {};
      for (const col of columns) {
        let v = row[col];
        if (v instanceof Date) v = v.toISOString().slice(0, 10);
        out[col] = normalizeValue(v);
      }
      return out;
    });
    return { columns, rows };
  }

  // ---------- Stats (pandas .info() / .describe() equivalent) ----------

  function round2(n) {
    if (n === null || n === undefined || Number.isNaN(n)) return null;
    return Math.round((n + Number.EPSILON) * 100) / 100;
  }

  function quantile(sortedArr, q) {
    const n = sortedArr.length;
    if (n === 0) return null;
    if (n === 1) return sortedArr[0];
    const idx = q * (n - 1);
    const lo = Math.floor(idx);
    const hi = Math.ceil(idx);
    if (lo === hi) return sortedArr[lo];
    return sortedArr[lo] + (sortedArr[hi] - sortedArr[lo]) * (idx - lo);
  }

  function inferDtype(values) {
    const nonNull = values.filter((v) => v !== null);
    if (nonNull.length === 0) return "object";
    const hasNull = nonNull.length < values.length;
    const allNumber = nonNull.every((v) => typeof v === "number" && Number.isFinite(v));
    if (allNumber) {
      const allInt = nonNull.every((v) => Number.isInteger(v));
      // pandas promotes an integer column to float64 when it contains
      // missing values, since NaN cannot be represented in an int array.
      return allInt && !hasNull ? "int64" : "float64";
    }
    const allBool = nonNull.every((v) => typeof v === "boolean");
    if (allBool && !hasNull) return "bool";
    return "object";
  }

  function formatBytes(n) {
    let val = n;
    const units = ["B", "KB", "MB", "GB"];
    for (let i = 0; i < units.length; i++) {
      if (val < 1024 || units[i] === "GB") {
        return i === 0 ? `${Math.round(val)} B` : `${val.toFixed(2)} ${units[i]}`;
      }
      val /= 1024;
    }
    return `${val.toFixed(2)} TB`;
  }

  function estimateMemoryBytes(columns, rows) {
    let total = 0;
    for (const row of rows) {
      for (const col of columns) {
        const v = row[col];
        if (v === null) total += 8;
        else if (typeof v === "number") total += 8;
        else if (typeof v === "boolean") total += 1;
        else total += new TextEncoder().encode(String(v)).length + 8;
      }
    }
    return total;
  }

  function buildAnalysis(columns, rows) {
    const nRows = rows.length;
    const nCols = columns.length;

    const columnsInfo = [];
    const numericStats = [];
    const categoricalStats = [];

    for (const col of columns) {
      const values = rows.map((r) => r[col]);
      const nonNullValues = values.filter((v) => v !== null);
      const nonNull = nonNullValues.length;
      const nullCount = nRows - nonNull;
      const nullPct = nRows ? round2((nullCount / nRows) * 100) : 0;
      const dtype = inferDtype(values);

      columnsInfo.push({
        name: col,
        dtype,
        non_null: nonNull,
        null_count: nullCount,
        null_pct: nullPct,
      });

      if (dtype === "int64" || dtype === "float64") {
        const nums = nonNullValues.map(Number).sort((a, b) => a - b);
        const count = nums.length;
        const mean = count ? nums.reduce((a, b) => a + b, 0) / count : null;
        let std = null;
        if (count > 1) {
          const variance = nums.reduce((acc, v) => acc + (v - mean) ** 2, 0) / (count - 1);
          std = Math.sqrt(variance);
        }
        numericStats.push({
          name: col,
          count,
          mean: round2(mean),
          std: round2(std),
          min: count ? round2(nums[0]) : null,
          p25: round2(quantile(nums, 0.25)),
          p50: round2(quantile(nums, 0.5)),
          p75: round2(quantile(nums, 0.75)),
          max: count ? round2(nums[nums.length - 1]) : null,
        });
      } else {
        const count = nonNullValues.length;
        if (count === 0) {
          categoricalStats.push({ name: col, count: 0, unique: 0, top: null, freq: 0 });
        } else {
          const freqMap = new Map();
          for (const v of nonNullValues) {
            const key = String(v);
            freqMap.set(key, (freqMap.get(key) || 0) + 1);
          }
          let top = null;
          let freq = 0;
          for (const [k, c] of freqMap.entries()) {
            if (c > freq) {
              freq = c;
              top = k;
            }
          }
          categoricalStats.push({
            name: col,
            count,
            unique: freqMap.size,
            top,
            freq,
          });
        }
      }
    }

    const memBytes = estimateMemoryBytes(columns, rows);
    const previewRows = rows.slice(0, 10);

    return {
      shape: { rows: nRows, cols: nCols },
      memory_bytes: memBytes,
      memory_readable: formatBytes(memBytes),
      columns_info: columnsInfo,
      numeric_stats: numericStats,
      categorical_stats: categoricalStats,
      preview_columns: columns,
      preview_rows: previewRows,
    };
  }

  // ---------- Main flow ----------

  async function handleFile(file) {
    hideError();
    filenameDisplay.textContent = file.name;

    const ext = file.name.includes(".") ? file.name.split(".").pop().toLowerCase() : "";
    if (!["csv", "xlsx"].includes(ext)) {
      showError("지원하지 않는 파일 형식입니다. CSV(.csv) 또는 Excel(.xlsx) 파일만 업로드해주세요.");
      return;
    }
    if (file.size === 0) {
      showError("빈 파일입니다. 데이터가 포함된 파일을 업로드해주세요.");
      return;
    }
    if (file.size > MAX_FILE_SIZE) {
      showError(`파일 크기가 너무 큽니다. 최대 ${MAX_FILE_SIZE / (1024 * 1024)}MB까지 업로드할 수 있습니다.`);
      return;
    }

    loadingBox.hidden = false;
    try {
      const { columns, rows } = ext === "csv" ? await parseCsv(file) : await parseExcel(file);

      if (!columns || columns.length === 0) {
        showError("파일에서 열(컬럼)을 찾을 수 없습니다.");
        return;
      }
      if (!rows || rows.length === 0) {
        showError("파일에 데이터 행이 없습니다. 헤더 외에 실제 데이터가 필요합니다.");
        return;
      }

      const data = buildAnalysis(columns, rows);
      renderResults(data);
      uploadSection.hidden = true;
      resultsSection.hidden = false;
    } catch (err) {
      showError("파일을 읽는 중 오류가 발생했습니다. 파일 형식이나 내용을 확인해주세요.");
    } finally {
      loadingBox.hidden = true;
    }
  }

  // ---------- Rendering ----------

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
      <div class="summary-card"><div class="label">메모리 사용량(추정)</div><div class="value">${escapeHtml(data.memory_readable)}</div></div>
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
