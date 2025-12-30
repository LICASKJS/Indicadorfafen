(() => {
  const DEFAULT_WORKBOOKS = [
    {
      id: "fafen-se",
      label: "FAFEN-SE",
      path: "dados/FORM.1341.REV00 - CONTROLE DE VALIDADE DE TREINAMENTOS FAFEN-SE (2).xlsx",
    },
    {
      id: "bahia",
      label: "Bahia",
      path: "dados/FORM.1341.REV00 - CONTROLE DE VALIDADE DE TREINAMENTOS - Bahia (4).xlsx",
    },
  ];

  const MAX_RENDER_ROWS = 800;
  const EXPIRE_WARNING_DAYS = 30;

  const validityStatusMeta = {
    prazo: { label: "Em dia", hint: "Dentro do prazo" },
    atrasado: { label: "A vencer", hint: `Vence em até ${EXPIRE_WARNING_DAYS} dias` },
    vencido: { label: "Vencido", hint: "Prazo expirado" },
  };

  const monthLabels = [
    "Janeiro",
    "Fevereiro",
    "Março",
    "Abril",
    "Maio",
    "Junho",
    "Julho",
    "Agosto",
    "Setembro",
    "Outubro",
    "Novembro",
    "Dezembro",
  ];

  const numberFormatter = new Intl.NumberFormat("pt-BR");

  const state = {
    workbook: null,
    workbookLabel: "",
    workbookSource: "",
    sheetName: "",
    table: null,
    chart: null,
  };

  const el = {
    navPrimary: document.querySelector("[data-nav-primary]"),
    workbookSelect: document.getElementById("workbookSelect"),
    sheetSelect: document.getElementById("sheetSelect"),
    workbookUpload: document.getElementById("workbookUpload"),
    uploadTrigger: document.querySelector("[data-upload-trigger]"),
    workbookMeta: document.getElementById("workbookMeta"),
    loadAlert: document.getElementById("loadAlert"),

    searchInput: document.getElementById("searchInput"),
    filterStartDate: document.getElementById("filterStartDate"),
    filterEndDate: document.getElementById("filterEndDate"),
    filterMonth: document.getElementById("filterMonth"),
    filterYear: document.getElementById("filterYear"),
    clearFilters: document.getElementById("clearFilters"),

    tableTitle: document.querySelector("[data-table-title]"),
    tableHead: document.querySelector("#dataTable thead"),
    tableBody: document.querySelector("#dataTable tbody"),
    tableFootnote: document.getElementById("tableFootnote"),

    chartCanvas: document.getElementById("dataChart"),
    chartEmpty: document.getElementById("chartEmpty"),
    chartMetrics: document.getElementById("chartMetrics"),
  };

  function initNav() {
    if (!el.navPrimary) return;
    el.navPrimary.innerHTML = `<a class="nav-link active" href="index.html">INDICADOR FAFEN</a>`;
  }

  function setAlert(message, { kind = "info" } = {}) {
    if (!el.loadAlert) return;
    el.loadAlert.textContent = message || "";
    el.loadAlert.classList.toggle("is-error", kind === "error");
  }

  function setLoading(isLoading) {
    if (!el.uploadTrigger) return;
    el.uploadTrigger.classList.toggle("is-loading", Boolean(isLoading));
  }

  function clearTable() {
    el.tableHead.innerHTML = "";
    el.tableBody.innerHTML = "";
  }

  function fillMonthSelect() {
    el.filterMonth.innerHTML = "";
    el.filterMonth.append(new Option("Todos", ""));
    monthLabels.forEach((label, index) => el.filterMonth.append(new Option(label, String(index + 1))));
  }

  function fillYearSelect(years) {
    const current = el.filterYear.value;
    el.filterYear.innerHTML = "";
    el.filterYear.append(new Option("Todos", ""));
    years.forEach((year) => el.filterYear.append(new Option(String(year), String(year))));
    if (years.includes(Number(current))) el.filterYear.value = current;
  }

  function fillWorkbookSelect() {
    el.workbookSelect.innerHTML = "";
    DEFAULT_WORKBOOKS.forEach((wb) => el.workbookSelect.append(new Option(wb.label, wb.id)));
    el.workbookSelect.append(new Option("Arquivo carregado.", "uploaded"));
    el.workbookSelect.value = DEFAULT_WORKBOOKS[0]?.id ?? "uploaded";
  }

  function fillSheetSelect(names) {
    el.sheetSelect.innerHTML = "";
    names.forEach((name) => el.sheetSelect.append(new Option(name, name)));
  }

  function normalizeHeader(value) {
    if (value == null) return "";
    return String(value).replace(/\s+/g, " ").trim();
  }

  function normalizeForMatch(value) {
    return String(value ?? "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase()
      .trim();
  }

  function isHiddenColumn(columnName) {
    const normalized = normalizeForMatch(columnName);
    if (!normalized) return false;
    if (/^coluna(s)?(\s+\d+)?(\s*\(|$)/.test(normalized)) return true;
    if (normalized === "empresa" || normalized.includes("empresa")) return true;
    if (/\bcpf\b/.test(normalized)) return true;
    if (/\bchave\b/.test(normalized)) return true;
    if (/\bid\s*petrobras\b/.test(normalized)) return true;
    return false;
  }

  function getVisibleColumnIndices(columns) {
    const indices = [];
    for (let colIndex = 0; colIndex < columns.length; colIndex += 1) {
      if (isHiddenColumn(columns[colIndex])) continue;
      indices.push(colIndex);
    }
    return indices;
  }

  function isEmptyCell(value) {
    return value == null || (typeof value === "string" && value.trim() === "");
  }

  function decodeSheetRange(sheet) {
    const ref = sheet?.["!ref"];
    if (!ref) return null;
    try {
      return XLSX.utils.decode_range(ref);
    } catch {
      return null;
    }
  }

  function buildColumnLetters(maxCol) {
    const letters = [];
    for (let c = 0; c <= maxCol; c += 1) letters.push(XLSX.utils.encode_col(c));
    return letters;
  }

  function getCellValue(sheet, columnLetters, rowIndex, colIndex) {
    const addr = `${columnLetters[colIndex]}${rowIndex + 1}`;
    const cell = sheet?.[addr];
    return cell ? cell.v : null;
  }

  function readRow(sheet, columnLetters, rowIndex, maxCol) {
    const row = Array(maxCol + 1).fill(null);
    for (let c = 0; c <= maxCol; c += 1) row[c] = getCellValue(sheet, columnLetters, rowIndex, c);
    return row;
  }

  function readMatrix(sheet, columnLetters, startRow, endRow, maxCol) {
    const matrix = [];
    for (let r = startRow; r <= endRow; r += 1) matrix.push(readRow(sheet, columnLetters, r, maxCol));
    return matrix;
  }

  function inferKeyColumns(columns) {
    const normalized = columns.map((name) => normalizeForMatch(name));
    const keyColumns = [];

    const nameIndex = normalized.findIndex((name) => name === "nome completo" || name.includes("nome completo"));
    if (nameIndex !== -1) keyColumns.push(nameIndex);

    const cpfIndex = normalized.findIndex((name) => name === "cpf" || name.includes("cpf"));
    if (cpfIndex !== -1) keyColumns.push(cpfIndex);

    const empresaIndex = normalized.findIndex((name) => name === "empresa" || name.includes("empresa"));
    if (empresaIndex !== -1) keyColumns.push(empresaIndex);

    if (keyColumns.length) return Array.from(new Set(keyColumns));
    return [0, 1, 2].filter((idx) => idx < columns.length);
  }

  function findHeaderRowIndex(matrix) {
    const limit = Math.min(matrix.length, 50);
    let bestIndex = 0;
    let bestScore = -Infinity;

    for (let i = 0; i < limit; i += 1) {
      const row = matrix[i];
      const cells = row.map((v) => normalizeHeader(v)).filter(Boolean);
      const normalized = new Set(cells.map((v) => v.toUpperCase()));

      const stringCount = row.filter((v) => typeof v === "string" && v.trim()).length;
      const bonus =
        (normalized.has("EMPRESA") ? 60 : 0) +
        (normalized.has("NOME COMPLETO") || normalized.has("NOME") ? 60 : 0) +
        (normalized.has("CPF") ? 20 : 0) +
        (normalized.has("FUNÇÃO") || normalized.has("FUNCAO") ? 10 : 0);

      const score = stringCount + bonus;
      if (score > bestScore) {
        bestScore = score;
        bestIndex = i;
      }
    }

    return bestIndex;
  }

  function detectHeaderRowCount(matrix, headerRowIndex) {
    const next = matrix[headerRowIndex + 1];
    if (!next) return 1;
    const normalized = next
      .filter((v) => typeof v === "string" && v.trim())
      .map((v) => normalizeForMatch(v));
    const dataTokens = normalized.filter((v) => v.includes("data") || v.includes("valid") || v.includes("realiza"));
    if (dataTokens.length >= 6) return 2;
    return 1;
  }

  function buildColumnNames(matrix, headerRowIndex, headerRowCount) {
    const header1 = matrix[headerRowIndex] ?? [];
    const header2 = headerRowCount === 2 ? matrix[headerRowIndex + 1] ?? [] : [];
    const maxLen = Math.max(header1.length, header2.length);

    const seen = new Map();
    const columns = [];

    for (let colIndex = 0; colIndex < maxLen; colIndex += 1) {
      const h1 = normalizeHeader(header1[colIndex]);
      const h2 = normalizeHeader(header2[colIndex]);

      let name = "";
      if (h1 && h2 && h2 !== h1) name = `${h1} - ${h2}`;
      else name = h1 || h2 || `Coluna ${colIndex + 1}`;

      const count = seen.get(name) ?? 0;
      seen.set(name, count + 1);
      if (count > 0) name = `${name} (${count + 1})`;

      columns.push(name);
    }

    return columns;
  }

  function coerceToDate(value) {
    if (value == null) return null;
    if (value instanceof Date && !Number.isNaN(value.getTime())) return value;

    if (typeof value === "number" && Number.isFinite(value)) {
      if (value < 20000 || value > 90000) return null;
      const parsed = XLSX.SSF?.parse_date_code?.(value);
      if (parsed && parsed.y && parsed.m && parsed.d) {
        const date = new Date(parsed.y, parsed.m - 1, parsed.d);
        if (!Number.isNaN(date.getTime())) return date;
      }
      return null;
    }

    if (typeof value === "string") {
      const raw = value.trim();
      if (!raw) return null;
      const iso = new Date(raw);
      if (!Number.isNaN(iso.getTime())) return iso;

      const brMatch = raw.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
      if (brMatch) {
        const [, dd, mm, yyyy] = brMatch;
        const date = new Date(Number(yyyy), Number(mm) - 1, Number(dd));
        if (!Number.isNaN(date.getTime())) return date;
      }
    }

    return null;
  }

  function inferColumnMeta(columns, rows, { treatAllDatesAsValidity = false } = {}) {
    const dateColumns = [];
    const validityDateColumns = [];
    const numericColumns = [];

    const sampleSize = Math.min(rows.length, 250);

    for (let colIndex = 0; colIndex < columns.length; colIndex += 1) {
      let nonEmpty = 0;
      let dateHits = 0;
      let numericHits = 0;

      for (let i = 0; i < sampleSize; i += 1) {
        const value = rows[i]?.[colIndex];
        if (isEmptyCell(value)) continue;
        nonEmpty += 1;
        if (typeof value === "number" && Number.isFinite(value)) numericHits += 1;
        if (coerceToDate(value)) dateHits += 1;
      }

      if (nonEmpty === 0) continue;
      const dateRatio = dateHits / nonEmpty;
      const numericRatio = numericHits / nonEmpty;

      const header = normalizeForMatch(columns[colIndex]);
      const isLikelyDate = dateRatio >= 0.65 || (dateHits >= 2 && dateRatio >= 0.4);
      const isLikelyNumeric = numericRatio >= 0.85 && !header.includes("cpf");

      if (isLikelyNumeric) numericColumns.push(colIndex);
      if (!isLikelyDate) continue;

      dateColumns.push(colIndex);
      const looksLikeRealizacao = header.includes("realiza");
      const looksLikeValidade = header.includes("valid") || header.includes("venc") || header.includes("aso");
      if (looksLikeValidade || (!looksLikeRealizacao && treatAllDatesAsValidity)) validityDateColumns.push(colIndex);
    }

    return { dateColumns, validityDateColumns, numericColumns };
  }

  function parseSheetAsTable(sheetName, sheet) {
    const range = decodeSheetRange(sheet);
    if (!range) return null;

    const maxCol = range.e.c;
    const maxRow = range.e.r;
    const columnLetters = buildColumnLetters(maxCol);

    const headerScanEnd = Math.min(maxRow, range.s.r + 200);
    const headerMatrix = readMatrix(sheet, columnLetters, range.s.r, headerScanEnd, maxCol);
    if (headerMatrix.length === 0) return null;

    const headerRowIndexInScan = findHeaderRowIndex(headerMatrix);
    const headerRowCount =
      sheetName === "Controle de Validade" ? detectHeaderRowCount(headerMatrix, headerRowIndexInScan) : 1;
    const columns = buildColumnNames(headerMatrix, headerRowIndexInScan, headerRowCount);

    const dataStartRow = range.s.r + headerRowIndexInScan + headerRowCount;
    const keyColumns = sheetName === "Controle de Validade" ? inferKeyColumns(columns) : [];

    const rows = [];
    const emptyStreakLimit = sheetName === "Controle de Validade" ? 80 : 40;
    let emptyStreak = 0;

    const maxRowsToRead = Math.min(maxRow, dataStartRow + (sheetName === "Controle de Validade" ? 5000 : 1500));
    for (let rowIndex = dataStartRow; rowIndex <= maxRowsToRead; rowIndex += 1) {
      let isCandidate = true;
      if (keyColumns.length) {
        isCandidate = keyColumns.some((colIndex) => !isEmptyCell(getCellValue(sheet, columnLetters, rowIndex, colIndex)));
      }

      if (!isCandidate) {
        emptyStreak += 1;
        if (emptyStreak >= emptyStreakLimit) break;
        continue;
      }

      const row = readRow(sheet, columnLetters, rowIndex, maxCol);
      if (!row.some((cell) => !isEmptyCell(cell))) {
        emptyStreak += 1;
        if (emptyStreak >= emptyStreakLimit) break;
        continue;
      }

      emptyStreak = 0;
      rows.push(row);
    }

    const meta = inferColumnMeta(columns, rows, { treatAllDatesAsValidity: sheetName === "Controle de Validade" });
    return { columns, rows, meta };
  }

  function formatDate(date) {
    try {
      return new Intl.DateTimeFormat("pt-BR").format(date);
    } catch {
      return "";
    }
  }

  function startOfDay(date) {
    const d = new Date(date);
    d.setHours(0, 0, 0, 0);
    return d;
  }

  function getValidityStatus(date) {
    const today = startOfDay(new Date());
    const target = startOfDay(date);
    const diffDays = Math.floor((target - today) / 86400000);
    if (diffDays < 0) return "vencido";
    if (diffDays <= EXPIRE_WARNING_DAYS) return "atrasado";
    return "prazo";
  }

  function isValidityColumn(colIndex, meta) {
    if (meta.validityDateColumns.length > 0) return meta.validityDateColumns.includes(colIndex);
    return meta.dateColumns.includes(colIndex);
  }

  function rowMatchesSearch(row, search) {
    if (!search) return true;
    const haystack = row.map((v) => normalizeForMatch(v)).join(" ");
    return haystack.includes(search);
  }

  function rowMatchesDateFilters(row, meta, filters) {
    const { startDate, endDate, month, year } = filters;
    const hasAnyDateFilter = Boolean(startDate || endDate || month || year);
    if (!hasAnyDateFilter) return true;

    const start = startDate ? startOfDay(startDate) : null;
    const end = endDate ? startOfDay(endDate) : null;

    const monthNum = month ? Number(month) : null;
    const yearNum = year ? Number(year) : null;

    const candidates = [];
    for (const colIndex of meta.validityDateColumns.length ? meta.validityDateColumns : meta.dateColumns) {
      const date = coerceToDate(row[colIndex]);
      if (date) candidates.push(date);
    }

    if (candidates.length === 0) return false;

    return candidates.some((date) => {
      const d = startOfDay(date);
      if (start && d < start) return false;
      if (end && d > end) return false;
      if (monthNum && d.getMonth() + 1 !== monthNum) return false;
      if (yearNum && d.getFullYear() !== yearNum) return false;
      return true;
    });
  }

  function buildYearsFromData(rows, meta) {
    const years = new Set();
    for (const row of rows) {
      for (const colIndex of meta.validityDateColumns.length ? meta.validityDateColumns : meta.dateColumns) {
        const date = coerceToDate(row[colIndex]);
        if (!date) continue;
        years.add(date.getFullYear());
      }
    }
    return Array.from(years).sort((a, b) => a - b);
  }

  function renderTable(columns, rows, meta) {
    el.tableHead.innerHTML = "";
    el.tableBody.innerHTML = "";

    if (!columns.length) {
      clearTable();
      return;
    }

    const visibleColumnIndices = getVisibleColumnIndices(columns);
    if (visibleColumnIndices.length === 0) {
      clearTable();
      return;
    }

    const headRow = document.createElement("tr");
    visibleColumnIndices.forEach((colIndex) => {
      const name = columns[colIndex];
      const th = document.createElement("th");
      th.textContent = name;
      if (meta.numericColumns.includes(colIndex)) th.classList.add("is-numeric");
      headRow.appendChild(th);
    });
    const thead = document.createElement("thead");
    thead.appendChild(headRow);
    el.tableHead.replaceWith(thead);
    el.tableHead = thead;

    const tbody = document.createElement("tbody");
    const rowsToRender = rows.slice(0, MAX_RENDER_ROWS);

    for (const row of rowsToRender) {
      const tr = document.createElement("tr");
      visibleColumnIndices.forEach((colIndex) => {
        const td = document.createElement("td");
        const value = row[colIndex];

        if (meta.numericColumns.includes(colIndex)) td.classList.add("is-numeric");

        const date = coerceToDate(value);
        if (date && meta.dateColumns.includes(colIndex)) {
          const wrapper = document.createElement("div");
          wrapper.className = "cell-date";
          const valueEl = document.createElement("span");
          valueEl.className = "cell-date-value";
          valueEl.textContent = formatDate(date);
          wrapper.appendChild(valueEl);

          if (isValidityColumn(colIndex, meta)) {
            const status = getValidityStatus(date);
            const statusMeta = validityStatusMeta[status];
            const badge = document.createElement("span");
            badge.className = `status-badge status-${status}`;
            badge.textContent = statusMeta?.label ?? status;
            if (statusMeta?.hint) badge.title = statusMeta.hint;
            wrapper.appendChild(badge);
          }

          td.appendChild(wrapper);
        } else if (value == null) {
          td.textContent = "";
        } else {
          td.textContent = String(value);
        }

        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    }

    el.tableBody.replaceWith(tbody);
    el.tableBody = tbody;
  }

  function renderFootnote({ totalRows, filteredRows, renderedRows, meta, counts }) {
    const parts = [];
    parts.push(`<strong>Linhas:</strong> ${renderedRows}/${filteredRows} (total ${totalRows})`);

    if (counts && (counts.prazo + counts.atrasado + counts.vencido) > 0) {
      parts.push(
        `<strong>Validades:</strong> ${validityStatusMeta.prazo.label} ${counts.prazo}, ${validityStatusMeta.atrasado.label} ${counts.atrasado}, ${validityStatusMeta.vencido.label} ${counts.vencido}`
      );
    } else if (meta && meta.dateColumns.length) {
      parts.push(`<strong>Datas:</strong> ${meta.dateColumns.length} coluna(s) detectada(s)`);
    }

    el.tableFootnote.innerHTML = parts.join(" · ");
  }

  function computeValidityCounts(rows, meta) {
    const counts = { prazo: 0, atrasado: 0, vencido: 0 };
    const cols = meta.validityDateColumns.length ? meta.validityDateColumns : meta.dateColumns;
    if (!cols.length) return counts;

    for (const row of rows) {
      for (const colIndex of cols) {
        const date = coerceToDate(row[colIndex]);
        if (!date) continue;
        const status = getValidityStatus(date);
        counts[status] += 1;
      }
    }
    return counts;
  }

  function destroyChart() {
    if (!state.chart) return;
    state.chart.destroy();
    state.chart = null;
  }

  function renderChartMetrics(counts) {
    if (!el.chartMetrics) return;

    const safeCounts = {
      prazo: Number(counts?.prazo ?? 0),
      atrasado: Number(counts?.atrasado ?? 0),
      vencido: Number(counts?.vencido ?? 0),
    };

    const items = [
      {
        key: "prazo",
        label: validityStatusMeta.prazo.label,
        hint: validityStatusMeta.prazo.hint,
        className: "chart-metric--prazo",
      },
      {
        key: "atrasado",
        label: validityStatusMeta.atrasado.label,
        hint: validityStatusMeta.atrasado.hint,
        className: "chart-metric--atrasado",
      },
      {
        key: "vencido",
        label: validityStatusMeta.vencido.label,
        hint: validityStatusMeta.vencido.hint,
        className: "chart-metric--vencido",
      },
    ];

    el.chartMetrics.innerHTML = items
      .map((item) => {
        const value = numberFormatter.format(safeCounts[item.key]);
        return `
          <div class="chart-metric ${item.className}" title="${item.hint}">
            <span class="metric-dot" aria-hidden="true"></span>
            <span class="metric-label">${item.label}</span>
            <span class="metric-value">${value}</span>
          </div>
        `;
      })
      .join("");
  }

  function renderChart(counts) {
    renderChartMetrics(counts);

    const total = counts.prazo + counts.atrasado + counts.vencido;
    if (!total) {
      destroyChart();
      el.chartEmpty.classList.add("is-visible");
      return;
    }

    el.chartEmpty.classList.remove("is-visible");

    if (!window.Chart) {
      destroyChart();
      el.chartEmpty.classList.add("is-visible");
      setAlert("Biblioteca de gráficos não carregou (CDN).", { kind: "error" });
      return;
    }

    const ctx = el.chartCanvas.getContext("2d");

    const computed = window.getComputedStyle(document.body);
    const theme = {
      textPrimary: computed.getPropertyValue("--text-primary").trim() || "#1e293b",
      textSecondary: computed.getPropertyValue("--text-secondary").trim() || "#64748b",
      border: computed.getPropertyValue("--border").trim() || "#e2e8f0",
    };

    const makeGradient = (from, to) => {
      const height = el.chartCanvas.clientHeight || el.chartCanvas.height || 320;
      const gradient = ctx.createLinearGradient(0, 0, 0, height);
      gradient.addColorStop(0, from);
      gradient.addColorStop(1, to);
      return gradient;
    };

    const palette = [
      makeGradient("rgba(34, 197, 94, 0.95)", "rgba(34, 197, 94, 0.35)"),
      makeGradient("rgba(245, 158, 11, 0.95)", "rgba(245, 158, 11, 0.35)"),
      makeGradient("rgba(239, 68, 68, 0.95)", "rgba(239, 68, 68, 0.35)"),
    ];

    const percentOnTime = Math.round((counts.prazo / total) * 100);

    const arcShadowPlugin = {
      id: "arcShadow",
      beforeDatasetDraw(chart, args, pluginOptions) {
        if (args.index !== 0) return;
        const shadowColor = pluginOptions?.color ?? "rgba(15, 23, 42, 0.18)";
        chart.ctx.save();
        chart.ctx.shadowColor = shadowColor;
        chart.ctx.shadowBlur = pluginOptions?.blur ?? 18;
        chart.ctx.shadowOffsetY = pluginOptions?.offsetY ?? 10;
        chart.ctx.shadowOffsetX = pluginOptions?.offsetX ?? 0;
      },
      afterDatasetDraw(chart, args) {
        if (args.index !== 0) return;
        chart.ctx.restore();
      },
    };

    const centerLabelPlugin = {
      id: "centerLabel",
      afterDraw(chart, _args, pluginOptions) {
        const meta = chart.getDatasetMeta(0);
        const firstArc = meta?.data?.[0];
        if (!firstArc) return;

        const x = firstArc.x;
        const y = firstArc.y;
        const innerRadius = firstArc.innerRadius ?? Math.min(chart.chartArea.width, chart.chartArea.height) / 3;

        const valueText = pluginOptions?.valueText ?? "";
        const labelText = pluginOptions?.labelText ?? "";
        const detailText = pluginOptions?.detailText ?? "";

        chart.ctx.save();
        chart.ctx.textAlign = "center";
        chart.ctx.textBaseline = "middle";

        const valueSize = Math.max(18, Math.min(42, innerRadius * 0.42));
        const labelSize = Math.max(10, Math.min(14, innerRadius * 0.16));
        const detailSize = Math.max(10, Math.min(12, innerRadius * 0.14));

        chart.ctx.fillStyle = pluginOptions?.valueColor ?? "#1e293b";
        chart.ctx.font = `800 ${valueSize}px Inter, -apple-system, BlinkMacSystemFont, \"Segoe UI\", Roboto, sans-serif`;
        chart.ctx.fillText(valueText, x, y - innerRadius * 0.12);

        chart.ctx.fillStyle = pluginOptions?.labelColor ?? "#64748b";
        chart.ctx.font = `700 ${labelSize}px Inter, -apple-system, BlinkMacSystemFont, \"Segoe UI\", Roboto, sans-serif`;
        chart.ctx.fillText(labelText, x, y + innerRadius * 0.18);

        if (detailText) {
          chart.ctx.fillStyle = pluginOptions?.detailColor ?? "#94a3b8";
          chart.ctx.font = `600 ${detailSize}px Inter, -apple-system, BlinkMacSystemFont, \"Segoe UI\", Roboto, sans-serif`;
          chart.ctx.fillText(detailText, x, y + innerRadius * 0.36);
        }

        chart.ctx.restore();
      },
    };

    const data = {
      labels: [validityStatusMeta.prazo.label, validityStatusMeta.atrasado.label, validityStatusMeta.vencido.label],
      datasets: [
        {
          label: "Validades",
          data: [counts.prazo, counts.atrasado, counts.vencido],
          backgroundColor: palette,
          borderColor: "rgba(255, 255, 255, 0.85)",
          borderWidth: 3,
          hoverBorderColor: "rgba(255, 255, 255, 1)",
          hoverOffset: 10,
          spacing: 4,
          borderRadius: 10,
        },
      ],
    };

    const options = {
      responsive: true,
      maintainAspectRatio: false,
      cutout: "72%",
      layout: { padding: 6 },
      animation: { duration: 900, easing: "easeOutQuart" },
      plugins: {
        legend: {
          display: false,
          position: "bottom",
          labels: {
            usePointStyle: true,
            pointStyle: "circle",
            boxWidth: 10,
            boxHeight: 10,
            padding: 18,
            color: theme.textSecondary,
            font: { size: 12, weight: "600" },
          },
        },
        tooltip: {
          backgroundColor: "rgba(15, 23, 42, 0.92)",
          padding: 12,
          titleColor: "#f8fafc",
          bodyColor: "#e2e8f0",
          displayColors: true,
          callbacks: {
            label(context) {
              const value = Number(context.parsed ?? 0);
              const pct = total ? Math.round((value / total) * 100) : 0;
              return `${context.label}: ${numberFormatter.format(value)} (${pct}%)`;
            },
          },
        },
        arcShadow: { color: "rgba(15, 23, 42, 0.18)", blur: 18, offsetY: 10 },
        centerLabel: {
          valueText: `${percentOnTime}%`,
          labelText: "Em dia",
          detailText: `${numberFormatter.format(counts.prazo)} de ${numberFormatter.format(total)}`,
          valueColor: theme.textPrimary,
          labelColor: theme.textSecondary,
          detailColor: theme.textSecondary,
        },
      },
    };

    if (state.chart) {
      state.chart.data = data;
      state.chart.options = options;
      state.chart.update();
      return;
    }

    state.chart = new Chart(ctx, { type: "doughnut", data, options, plugins: [arcShadowPlugin, centerLabelPlugin] });
  }

  function getFilters() {
    const startDate = el.filterStartDate.value ? new Date(el.filterStartDate.value) : null;
    const endDate = el.filterEndDate.value ? new Date(el.filterEndDate.value) : null;
    return {
      search: normalizeForMatch(el.searchInput.value),
      startDate,
      endDate,
      month: el.filterMonth.value,
      year: el.filterYear.value,
    };
  }

  function applyFiltersAndRender() {
    if (!state.table) return;
    const { columns, rows, meta } = state.table;
    const filters = getFilters();

    const filtered = rows.filter((row) => rowMatchesSearch(row, filters.search) && rowMatchesDateFilters(row, meta, filters));
    const counts = computeValidityCounts(filtered, meta);

    renderTable(columns, filtered, meta);
    renderChart(counts);
    renderFootnote({
      totalRows: rows.length,
      filteredRows: filtered.length,
      renderedRows: Math.min(filtered.length, MAX_RENDER_ROWS),
      meta,
      counts,
    });

    const years = buildYearsFromData(rows, meta);
    fillYearSelect(years);
  }

  async function loadWorkbookFromArrayBuffer(buffer, { label, source }) {
    if (!window.XLSX) throw new Error("Biblioteca XLSX não carregou (CDN).");
    if (!buffer) throw new Error("Arquivo vazio.");

    const workbook = XLSX.read(buffer, { type: "array", cellDates: true, cellNF: false });
    state.workbook = workbook;
    state.workbookLabel = label;
    state.workbookSource = source;

    fillSheetSelect(workbook.SheetNames);
    const defaultSheet =
      workbook.SheetNames.includes("Controle de Validade") ? "Controle de Validade" : workbook.SheetNames[0] ?? "";
    el.sheetSelect.value = defaultSheet;
    await loadSelectedSheet();
  }

  function readFileAsArrayBuffer(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = () => reject(new Error("Falha ao ler o arquivo."));
      reader.readAsArrayBuffer(file);
    });
  }

  async function loadWorkbookFromFile(file) {
    if (!file) return;
    setLoading(true);
    setAlert("Carregando planilha do arquivo…");
    try {
      const buffer = await readFileAsArrayBuffer(file);
      await loadWorkbookFromArrayBuffer(buffer, { label: file.name, source: "upload" });
      el.workbookSelect.value = "uploaded";
      setAlert("");
    } catch (err) {
      setAlert(err?.message ?? "Não foi possível ler o arquivo.", { kind: "error" });
    } finally {
      setLoading(false);
    }
  }

  async function loadWorkbookFromPath(workbookDef) {
    setLoading(true);
    setAlert(`Carregando ${workbookDef.label}…`);
    try {
      const res = await fetch(encodeURI(workbookDef.path), { cache: "no-store" });
      if (!res.ok) throw new Error(`Não foi possível carregar: ${workbookDef.path}`);
      const buffer = await res.arrayBuffer();
      await loadWorkbookFromArrayBuffer(buffer, { label: workbookDef.label, source: workbookDef.path });
      setAlert("");
    } catch (err) {
      setAlert(
        "Não consegui carregar automaticamente a planilha. Se você abriu o arquivo via 'file://', rode um servidor local (ex.: `python -m http.server`) ou use o botão (+) para carregar o Excel.",
        { kind: "error" }
      );
      clearTable();
      destroyChart();
      el.chartEmpty.classList.add("is-visible");
    } finally {
      setLoading(false);
    }
  }

  async function loadSelectedSheet() {
    if (!state.workbook) return;

    const sheetName = el.sheetSelect.value;
    const sheet = state.workbook.Sheets[sheetName];
    if (!sheet) return;

    state.sheetName = sheetName;
    setAlert("");

    const parsed = parseSheetAsTable(sheetName, sheet);
    if (!parsed) {
      state.table = null;
      clearTable();
      destroyChart();
      el.chartEmpty.classList.add("is-visible");
      return;
    }

    state.table = parsed;
    el.tableTitle.textContent = sheetName;
    el.workbookMeta.textContent = `${state.workbookLabel} · ${parsed.rows.length} linha(s)`;
    applyFiltersAndRender();
  }

  function clearFilters() {
    el.searchInput.value = "";
    el.filterStartDate.value = "";
    el.filterEndDate.value = "";
    el.filterMonth.value = "";
    el.filterYear.value = "";
    applyFiltersAndRender();
  }

  function bindEvents() {
    el.workbookSelect.addEventListener("change", async () => {
      const id = el.workbookSelect.value;
      if (id === "uploaded") {
        setAlert("Use o botão (+) para carregar um Excel.", { kind: "info" });
        return;
      }
      const wb = DEFAULT_WORKBOOKS.find((w) => w.id === id);
      if (!wb) return;
      await loadWorkbookFromPath(wb);
    });

    el.workbookUpload.addEventListener("change", async (e) => {
      const file = e.target.files?.[0];
      await loadWorkbookFromFile(file);
      e.target.value = "";
    });

    el.sheetSelect.addEventListener("change", loadSelectedSheet);

    [el.searchInput, el.filterStartDate, el.filterEndDate].forEach((input) => input.addEventListener("input", applyFiltersAndRender));
    [el.filterMonth, el.filterYear].forEach((select) => select.addEventListener("change", applyFiltersAndRender));

    el.clearFilters.addEventListener("click", clearFilters);
  }

  async function start() {
    initNav();
    fillMonthSelect();
    fillWorkbookSelect();
    bindEvents();

    const first = DEFAULT_WORKBOOKS[0];
    if (first) await loadWorkbookFromPath(first);
  }

  start();
})();
