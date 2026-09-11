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

  const prepStatusBox = document.getElementById("prepStatus");
  const missingSummaryEl = document.getElementById("missingSummary");
  const missingFormEl = document.getElementById("missingForm");
  const missingColumnSelect = document.getElementById("missingColumnSelect");
  const missingMethodSelect = document.getElementById("missingMethodSelect");
  const applyMissingBtn = document.getElementById("applyMissingBtn");
  const outlierColumnSelect = document.getElementById("outlierColumnSelect");
  const boxplotWrap = document.getElementById("boxplotWrap");
  const outlierSummaryEl = document.getElementById("outlierSummary");
  const applyOutlierBtn = document.getElementById("applyOutlierBtn");
  const resetDataBtn = document.getElementById("resetDataBtn");

  const previewRowCountInput = document.getElementById("previewRowCount");
  const previewHintEl = document.getElementById("previewHint");
  const downloadCsvBtn = document.getElementById("downloadCsvBtn");

  // originalData never changes once a file is loaded; workingData is the
  // mutable copy that missing-value/outlier preprocessing operates on.
  let originalData = null;
  let workingData = null;
  let currentBaseName = "data";

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
    hidePrepStatus();
    originalData = null;
    workingData = null;
    previewRowCountInput.value = 10;
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

  function cloneData(data) {
    return { columns: [...data.columns], rows: data.rows.map((r) => ({ ...r })) };
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

    return {
      shape: { rows: nRows, cols: nCols },
      memory_bytes: memBytes,
      memory_readable: formatBytes(memBytes),
      columns_info: columnsInfo,
      numeric_stats: numericStats,
      categorical_stats: categoricalStats,
    };
  }

  // ---------- Main flow ----------

  async function handleFile(file) {
    hideError();
    filenameDisplay.textContent = file.name;
    const dotIdx = file.name.lastIndexOf(".");
    currentBaseName = dotIdx > 0 ? file.name.slice(0, dotIdx) : file.name;

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

      originalData = { columns, rows };
      workingData = cloneData(originalData);
      previewRowCountInput.value = 10;
      renderResults(buildAnalysis(workingData.columns, workingData.rows));
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

  function fmtValue(v) {
    if (v === null || v === undefined) return "-";
    return typeof v === "number" ? fmtNum(v) : escapeHtml(String(v));
  }

  function renderAllTabs(data) {
    renderSummary(data);
    renderPreviewFromWorkingData();
    renderInfo(data);
    renderNumeric(data);
    renderCategorical(data);
    renderPrepPanel(data);
  }

  function renderResults(data) {
    renderAllTabs(data);

    tabs.forEach((b) => b.classList.remove("active"));
    panels.forEach((p) => p.classList.remove("active"));
    tabs[0].classList.add("active");
    document.getElementById("panel-preview").classList.add("active");
  }

  // Re-analyze workingData after a preprocessing action and refresh every
  // tab in place, without changing which tab is currently active.
  function refreshWorkingData() {
    renderAllTabs(buildAnalysis(workingData.columns, workingData.rows));
  }

  function renderSummary(data) {
    const el = document.getElementById("summaryCards");
    el.innerHTML = `
      <div class="summary-card"><div class="label">전체 행 수</div><div class="value">${fmtNum(data.shape.rows, 0)}</div></div>
      <div class="summary-card"><div class="label">전체 열 수</div><div class="value">${fmtNum(data.shape.cols, 0)}</div></div>
      <div class="summary-card"><div class="label">메모리 사용량(추정)</div><div class="value">${escapeHtml(data.memory_readable)}</div></div>
    `;
  }

  // The preview table's row count is user-controlled and independent of the
  // analysis stats, so it reads straight from workingData instead of going
  // through buildAnalysis (which would be wasted work on every keystroke).
  function renderPreviewFromWorkingData() {
    if (!workingData) return;
    const totalRows = workingData.rows.length;
    let n = parseInt(previewRowCountInput.value, 10);
    if (!Number.isFinite(n) || n < 1) n = 1;
    if (n > totalRows) n = totalRows;
    renderPreview({
      preview_columns: workingData.columns,
      preview_rows: workingData.rows.slice(0, n),
      total_rows: totalRows,
    });
  }

  previewRowCountInput.addEventListener("input", renderPreviewFromWorkingData);
  previewRowCountInput.addEventListener("change", () => {
    if (!workingData) return;
    let n = parseInt(previewRowCountInput.value, 10);
    const totalRows = workingData.rows.length;
    if (!Number.isFinite(n) || n < 1) n = 1;
    if (n > totalRows) n = totalRows || 1;
    previewRowCountInput.value = n;
    renderPreviewFromWorkingData();
  });

  function renderPreview(data) {
    const table = document.getElementById("previewTable");
    const cols = data.preview_columns;
    if (typeof data.total_rows === "number") {
      previewHintEl.textContent = `전체 ${fmtNum(data.total_rows, 0)}행 중 ${fmtNum(data.preview_rows.length, 0)}행 표시`;
    }
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

  // ---------- Preprocessing: missing values + outliers ----------

  function populateSelect(selectEl, items) {
    const prevValue = selectEl.value;
    selectEl.innerHTML = "";
    for (const item of items) {
      const opt = document.createElement("option");
      opt.value = item.value;
      opt.textContent = item.label;
      selectEl.appendChild(opt);
    }
    if (items.some((i) => i.value === prevValue)) selectEl.value = prevValue;
  }

  function showPrepStatus(message) {
    prepStatusBox.textContent = message;
    prepStatusBox.hidden = false;
  }

  function hidePrepStatus() {
    prepStatusBox.hidden = true;
    prepStatusBox.textContent = "";
  }

  function renderPrepPanel(data) {
    const missingCols = data.columns_info.filter((c) => c.null_count > 0);
    populateSelect(
      missingColumnSelect,
      missingCols.map((c) => ({ value: c.name, label: `${c.name} (결측 ${c.null_count}개)` }))
    );
    if (missingCols.length === 0) {
      missingSummaryEl.textContent = "현재 결측치가 있는 컬럼이 없습니다.";
      missingFormEl.hidden = true;
    } else {
      missingSummaryEl.textContent = `결측치가 있는 컬럼: ${missingCols.map((c) => c.name).join(", ")}`;
      missingFormEl.hidden = false;
      updateMissingMethodOptions();
    }

    const numericCols = data.numeric_stats.map((s) => s.name);
    populateSelect(
      outlierColumnSelect,
      numericCols.map((c) => ({ value: c, label: c }))
    );
    if (numericCols.length === 0) {
      boxplotWrap.innerHTML = "";
      outlierSummaryEl.textContent = "숫자형 컬럼이 없어 이상치를 확인할 수 없습니다.";
      applyOutlierBtn.disabled = true;
      applyOutlierBtn.dataset.column = "";
    } else {
      updateOutlierView();
    }
  }

  function updateMissingMethodOptions() {
    const col = missingColumnSelect.value;
    if (!col) return;
    const values = workingData.rows.map((r) => r[col]);
    const nonNullCount = values.filter((v) => v !== null).length;
    const dtype = inferDtype(values);
    const isNumeric = dtype === "int64" || dtype === "float64";
    for (const opt of missingMethodSelect.options) {
      if (opt.value === "mean" || opt.value === "median") {
        opt.disabled = !isNumeric || nonNullCount === 0;
      } else if (opt.value === "mode") {
        opt.disabled = nonNullCount === 0;
      }
    }
    if (missingMethodSelect.options[missingMethodSelect.selectedIndex]?.disabled) {
      const firstEnabled = [...missingMethodSelect.options].find((o) => !o.disabled);
      if (firstEnabled) missingMethodSelect.value = firstEnabled.value;
    }
  }

  missingColumnSelect.addEventListener("change", updateMissingMethodOptions);

  applyMissingBtn.addEventListener("click", () => {
    if (!workingData) return;
    const col = missingColumnSelect.value;
    const method = missingMethodSelect.value;
    if (!col) return;

    const beforeRows = workingData.rows.length;
    const beforeMissing = workingData.rows.filter((r) => r[col] === null).length;
    if (beforeMissing === 0) return;

    if (method === "drop") {
      workingData.rows = workingData.rows.filter((r) => r[col] !== null);
      showPrepStatus(
        `'${col}' 컬럼의 결측치가 있는 행 ${beforeMissing}개를 삭제했습니다. (${beforeRows}행 → ${workingData.rows.length}행)`
      );
    } else {
      const nonNull = workingData.rows.map((r) => r[col]).filter((v) => v !== null);
      if (nonNull.length === 0) {
        showPrepStatus(`'${col}' 컬럼에는 대체에 사용할 값이 없습니다. 행 삭제를 이용해주세요.`);
        return;
      }
      let fillValue;
      let methodLabel;
      if (method === "mean") {
        const nums = nonNull.map(Number);
        fillValue = round2(nums.reduce((a, b) => a + b, 0) / nums.length);
        methodLabel = "평균값";
      } else if (method === "median") {
        const nums = nonNull.map(Number).sort((a, b) => a - b);
        fillValue = round2(quantile(nums, 0.5));
        methodLabel = "중앙값";
      } else {
        const freqMap = new Map();
        for (const v of nonNull) {
          const key = String(v);
          freqMap.set(key, (freqMap.get(key) || 0) + 1);
        }
        let bestKey = null;
        let bestCount = 0;
        for (const [k, c] of freqMap.entries()) {
          if (c > bestCount) {
            bestCount = c;
            bestKey = k;
          }
        }
        fillValue = nonNull.find((v) => String(v) === bestKey);
        methodLabel = "최빈값";
      }
      workingData.rows = workingData.rows.map((r) => (r[col] === null ? { ...r, [col]: fillValue } : r));
      showPrepStatus(`'${col}' 컬럼의 결측치 ${beforeMissing}개를 ${methodLabel}(${fmtValue(fillValue)})로 대체했습니다.`);
    }
    refreshWorkingData();
  });

  function computeOutlierStats(column) {
    const values = workingData.rows
      .map((r) => r[column])
      .filter((v) => v !== null)
      .map(Number)
      .sort((a, b) => a - b);
    const q1 = quantile(values, 0.25);
    const q2 = quantile(values, 0.5);
    const q3 = quantile(values, 0.75);
    const iqr = q3 - q1;
    const lower = q1 - 1.5 * iqr;
    const upper = q3 + 1.5 * iqr;
    const outliers = values.filter((v) => v < lower || v > upper);
    const inliers = values.filter((v) => v >= lower && v <= upper);
    const whiskerMin = inliers.length ? inliers[0] : values[0];
    const whiskerMax = inliers.length ? inliers[inliers.length - 1] : values[values.length - 1];
    return { values, q1, q2, q3, iqr, lower, upper, outliers, whiskerMin, whiskerMax };
  }

  function buildBoxPlotSvg(stats) {
    const { q1, q2, q3, lower, upper, outliers, whiskerMin, whiskerMax } = stats;

    // Scale by the whisker/fence range (not raw outlier values) so a single
    // extreme outlier can't crush the box down to a sliver. Outlier dots are
    // clamped to the plot edge instead of stretching the axis to fit them.
    const width = 680;
    const height = 180;
    const marginX = 50;
    const plotW = width - marginX * 2;
    const rawMin = Math.min(whiskerMin, lower);
    const rawMax = Math.max(whiskerMax, upper);
    const span = rawMax - rawMin || Math.abs(rawMax) || 1;
    const pad = span * 0.18;
    const domainMin = rawMin - pad;
    const domainMax = rawMax + pad;
    const scale = (v) => marginX + ((v - domainMin) / (domainMax - domainMin || 1)) * plotW;
    const clampX = (v) => Math.min(width - 16, Math.max(16, scale(v)));

    const cy = 92;
    const boxH = 44;
    const capH = 20;
    const x1 = scale(q1);
    const x2 = scale(q3);
    const xMed = scale(q2);
    const xWMin = scale(whiskerMin);
    const xWMax = scale(whiskerMax);
    const boxLeft = Math.min(x1, x2);
    const boxWidth = Math.max(Math.abs(x2 - x1), 1);

    const fenceTop = 20;
    const fenceBottom = height - 34;
    const xLower = scale(lower);
    const xUpper = scale(upper);
    const fenceLines = `
      <line x1="${xLower.toFixed(1)}" y1="${fenceTop}" x2="${xLower.toFixed(1)}" y2="${fenceBottom}" stroke="var(--danger)" stroke-width="1" stroke-dasharray="4 3" opacity="0.7" />
      <text x="${xLower.toFixed(1)}" y="${fenceTop - 6}" font-size="10" fill="var(--danger)" text-anchor="middle">하한 ${round2(lower)}</text>
      <line x1="${xUpper.toFixed(1)}" y1="${fenceTop}" x2="${xUpper.toFixed(1)}" y2="${fenceBottom}" stroke="var(--danger)" stroke-width="1" stroke-dasharray="4 3" opacity="0.7" />
      <text x="${xUpper.toFixed(1)}" y="${fenceTop - 6}" font-size="10" fill="var(--danger)" text-anchor="middle">상한 ${round2(upper)}</text>
    `;

    const outlierDots = outliers
      .map((v) => {
        const x = clampX(v);
        const beyond = Math.abs(x - scale(v)) > 0.5;
        return `<circle cx="${x.toFixed(1)}" cy="${cy}" r="4" fill="var(--danger)" opacity="0.9"><title>${round2(v)}${beyond ? " (범위 밖)" : ""}</title></circle>`;
      })
      .join("");

    const whiskerLabelY = cy - boxH / 2 - 10;
    const whiskerLabels = `
      <text x="${xWMin.toFixed(1)}" y="${whiskerLabelY}" font-size="10" fill="var(--text-muted)" text-anchor="middle">min ${round2(whiskerMin)}</text>
      <text x="${xWMax.toFixed(1)}" y="${whiskerLabelY}" font-size="10" fill="var(--text-muted)" text-anchor="middle">max ${round2(whiskerMax)}</text>
    `;

    const tickY = cy + boxH / 2 + capH / 2 + 8;
    const quartileTicks = [
      { v: q1, label: "Q1" },
      { v: q2, label: "중앙값" },
      { v: q3, label: "Q3" },
    ]
      .map(({ v, label }) => {
        const x = scale(v).toFixed(1);
        return `<line x1="${x}" y1="${tickY}" x2="${x}" y2="${tickY + 5}" stroke="var(--navy)" stroke-width="1" />
          <text x="${x}" y="${tickY + 18}" font-size="10" fill="var(--navy)" text-anchor="middle" font-weight="600">${round2(v)}</text>
          <text x="${x}" y="${tickY + 30}" font-size="9" fill="var(--text-muted)" text-anchor="middle">${label}</text>`;
      })
      .join("");

    return `<svg viewBox="0 0 ${width} ${height}" xmlns="http://www.w3.org/2000/svg" style="width:100%;height:auto;display:block;">
      ${fenceLines}
      <line x1="${xWMin.toFixed(1)}" y1="${cy}" x2="${boxLeft.toFixed(1)}" y2="${cy}" stroke="var(--navy)" stroke-width="2" />
      <line x1="${(boxLeft + boxWidth).toFixed(1)}" y1="${cy}" x2="${xWMax.toFixed(1)}" y2="${cy}" stroke="var(--navy)" stroke-width="2" />
      <line x1="${xWMin.toFixed(1)}" y1="${cy - capH / 2}" x2="${xWMin.toFixed(1)}" y2="${cy + capH / 2}" stroke="var(--navy)" stroke-width="2" />
      <line x1="${xWMax.toFixed(1)}" y1="${cy - capH / 2}" x2="${xWMax.toFixed(1)}" y2="${cy + capH / 2}" stroke="var(--navy)" stroke-width="2" />
      <rect x="${boxLeft.toFixed(1)}" y="${cy - boxH / 2}" width="${boxWidth.toFixed(1)}" height="${boxH}" fill="var(--navy-pale)" stroke="var(--navy)" stroke-width="2" />
      <line x1="${xMed.toFixed(1)}" y1="${cy - boxH / 2}" x2="${xMed.toFixed(1)}" y2="${cy + boxH / 2}" stroke="var(--navy)" stroke-width="2" />
      ${whiskerLabels}
      ${outlierDots}
      ${quartileTicks}
    </svg>`;
  }

  function updateOutlierView() {
    const col = outlierColumnSelect.value;
    if (!col || !workingData) return;
    const stats = computeOutlierStats(col);
    if (stats.values.length === 0) {
      boxplotWrap.innerHTML = "";
      outlierSummaryEl.textContent = "이 컬럼에는 값이 없습니다.";
      applyOutlierBtn.disabled = true;
      applyOutlierBtn.dataset.column = "";
      return;
    }
    boxplotWrap.innerHTML = buildBoxPlotSvg(stats);
    outlierSummaryEl.innerHTML = `
      Q1: ${fmtNum(round2(stats.q1))} · 중앙값: ${fmtNum(round2(stats.q2))} · Q3: ${fmtNum(round2(stats.q3))} · IQR: ${fmtNum(round2(stats.iqr))}<br />
      하한(Q1−1.5×IQR): ${fmtNum(round2(stats.lower))} · 상한(Q3+1.5×IQR): ${fmtNum(round2(stats.upper))}<br />
      이상치 개수: ${stats.outliers.length}개${
      stats.outliers.length ? " (" + stats.outliers.map((v) => fmtNum(v)).join(", ") + ")" : ""
    }
    `;
    applyOutlierBtn.disabled = stats.outliers.length === 0;
    applyOutlierBtn.dataset.column = col;
  }

  outlierColumnSelect.addEventListener("change", updateOutlierView);

  applyOutlierBtn.addEventListener("click", () => {
    if (!workingData) return;
    const col = applyOutlierBtn.dataset.column;
    if (!col) return;
    const stats = computeOutlierStats(col);
    if (stats.outliers.length === 0) return;
    const beforeRows = workingData.rows.length;
    workingData.rows = workingData.rows.filter((r) => {
      const v = r[col];
      if (v === null) return true;
      return v >= stats.lower && v <= stats.upper;
    });
    showPrepStatus(
      `'${col}' 컬럼의 이상치 ${stats.outliers.length}개 행을 삭제했습니다. (${beforeRows}행 → ${workingData.rows.length}행)`
    );
    refreshWorkingData();
  });

  resetDataBtn.addEventListener("click", () => {
    if (!originalData) return;
    workingData = cloneData(originalData);
    showPrepStatus("원본 데이터로 되돌렸습니다.");
    refreshWorkingData();
  });

  // ---------- CSV export (current working data, after any preprocessing) ----------

  function escapeCsvField(v) {
    if (v === null || v === undefined) return "";
    const s = String(v);
    if (/[",\n\r]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
    return s;
  }

  function rowsToCsv(columns, rows) {
    const lines = [columns.map(escapeCsvField).join(",")];
    for (const row of rows) {
      lines.push(columns.map((c) => escapeCsvField(row[c])).join(","));
    }
    return lines.join("\r\n");
  }

  downloadCsvBtn.addEventListener("click", () => {
    if (!workingData) return;
    // Leading BOM so Excel on Windows opens the Korean text as UTF-8 instead of guessing.
    const csv = "﻿" + rowsToCsv(workingData.columns, workingData.rows);
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${currentBaseName}_processed.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  });
})();
