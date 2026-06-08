"use strict";

const $ = (id) => document.getElementById(id);

const ui = {
  screenHome: $("screenHome"),
  screenUpload: $("screenUpload"),
  screenSearch: $("screenSearch"),
  btnStartHome: $("btnStartHome"),
  btnLoadFiles: $("btnLoadFiles"),
  btnResetUpload: $("btnResetUpload"),
  btnChangeFiles: $("btnChangeFiles"),
  btnPrintReport: $("btnPrintReport"),
  btnToTop: $("btnToTop"),
  themeToggle: $("themeToggle"),
  fileInputEX: $("fileInputEX"),
  fileInputNVE: $("fileInputNVE"),
  fileMetaEX: $("fileMetaEX"),
  fileMetaNVE: $("fileMetaNVE"),
  fileStateEX: $("fileStateEX"),
  fileStateNVE: $("fileStateNVE"),
  uploadAlert: $("uploadAlert"),
  ncmInput: $("ncmInput"),
  hint: $("hint"),
  pillLoadedEX: $("pillLoadedEX"),
  pillLoadedNVE: $("pillLoadedNVE"),
  pillResultEX: $("pillResultEX"),
  pillResultNVE: $("pillResultNVE"),
  pillExpiredEX: $("pillExpiredEX"),
  countNVE: $("countNVE"),
  countEX: $("countEX"),
  tableNVE: $("tableNVE"),
  tableEX: $("tableEX"),
  wrapNVE: $("wrapNVE"),
  wrapEX: $("wrapEX"),
  emptyNVE: $("emptyNVE"),
  emptyEX: $("emptyEX"),
};

const dataStore = {
  ex: null,
  nve: null,
  currentNcm: "",
};

let debounceTimer = null;

const DISPLAY_COLS = {
  EX: ["EX", "DESCRIÇÃO", "INÍCIO DA VIGÊNCIA", "FIM DA VIGÊNCIA", "STATUS"],
  NVE: [
    "NIVEL",
    "ATRIBUTO",
    "DESCRIÇÃO ATRIBUTO",
    "ESPECIFICACAO",
    "DESCRIÇÃO ESPECIFICAÇÃO",
  ],
};

const HEADER_SYNONYMS = {
  NCM: ["NCM"],
  EX: ["EX"],
  DESCRIÇÃO: ["DESCRIÇÃO", "DESCRICAO"],
  "INÍCIO DA VIGÊNCIA": ["INÍCIO DA VIGÊNCIA", "INICIO DA VIGENCIA"],
  "FIM DA VIGÊNCIA": ["FIM DA VIGÊNCIA", "FIM DA VIGENCIA"],
  NIVEL: ["NIVEL", "NÍVEL"],
  ATRIBUTO: ["ATRIBUTO"],
  "DESCRIÇÃO ATRIBUTO": [
    "DESCRIÇÃO ATRIBUTO",
    "DESCRICAO ATRIBUTO",
    "DESCRIÇÃOATRIBUTO",
    "DESCRICAOATRIBUTO",
  ],
  ESPECIFICACAO: ["ESPECIFICACAO", "ESPECIFICAÇÃO"],
  "DESCRIÇÃO ESPECIFICAÇÃO": [
    "DESCRIÇÃO ESPECIFICAÇÃO",
    "DESCRICAO ESPECIFICACAO",
    "DESCRIÇÃOESPECIFICAÇÃO",
    "DESCRICAOESPECIFICACAO",
  ],
};

const REQUIRED_HEADERS = {
  EX: ["NCM", "EX", "FIM DA VIGÊNCIA"],
  NVE: ["NCM", "ATRIBUTO"],
};

function setTheme(theme) {
  document.documentElement.setAttribute("data-theme", theme);
  document.body.setAttribute("data-theme", theme);
}

function initTheme() {
  const dark =
    window.matchMedia &&
    window.matchMedia("(prefers-color-scheme: dark)").matches;
  setTheme(dark ? "dark" : "light");
}

function normalizeHeader(value) {
  return String(value ?? "")
    .replace(/\u00A0/g, " ")
    .replace(/[\r\n\t]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

function cleanText(value) {
  return String(value ?? "")
    .replace(/x000D/gi, " ")
    .replace(/_x000D_/gi, " ")
    .replace(/\u000d/gi, " ")
    .replace(/Sensitivity Label.*$/gi, "")
    .replace(/\u00A0/g, " ")
    .replace(/[\r\n\t]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function cleanNcmInput(raw) {
  return String(raw ?? "").replace(/\D/g, "");
}

function formatNcm(ncm) {
  const digits = cleanNcmInput(ncm);
  if (digits.length !== 8) return digits;
  return `${digits.slice(0, 4)}.${digits.slice(4, 6)}.${digits.slice(6)}`;
}

function setPill(el, label, value, tone = "") {
  el.className = "pill" + (tone ? ` ${tone}` : "");
  el.textContent = `${label}: ${value}`;
}

function showUploadError(message) {
  ui.uploadAlert.textContent = message;
  ui.uploadAlert.classList.remove("hidden");
}

function hideUploadError() {
  ui.uploadAlert.textContent = "";
  ui.uploadAlert.classList.add("hidden");
}

function setFileMeta(target, file) {
  if (!file) {
    target.textContent = "";
    return;
  }
  const mb = (file.size / (1024 * 1024)).toFixed(2);
  target.textContent = `${file.name} • ${mb} MB`;
}

function sheetToRows(sheet) {
  return XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
}

function sanitizeRowKeys(rows) {
  return rows.map((row) => {
    const out = {};
    for (const [key, value] of Object.entries(row)) {
      out[cleanText(key)] = cleanText(value);
    }
    return out;
  });
}

function buildHeaderMap(headers) {
  const normalizedToOriginal = new Map();
  headers.forEach((header) => {
    const normalized = normalizeHeader(header);
    if (normalized && !normalizedToOriginal.has(normalized))
      normalizedToOriginal.set(normalized, header);
  });

  const resolved = {};
  for (const [logical, synonyms] of Object.entries(HEADER_SYNONYMS)) {
    let found = null;
    for (const synonym of synonyms) {
      const normalized = normalizeHeader(synonym);
      if (normalizedToOriginal.has(normalized)) {
        found = normalizedToOriginal.get(normalized);
        break;
      }
    }
    resolved[logical] = found;
  }
  return resolved;
}

function validateHeaders(kind, headerMap) {
  const missing = REQUIRED_HEADERS[kind].filter((key) => !headerMap[key]);
  return {
    valid: missing.length === 0,
    missing,
  };
}

function resolveDisplayColumns(kind, headerMap, headersInSheet) {
  const list = DISPLAY_COLS[kind].map((logical) => {
    if (logical === "STATUS")
      return { key: "STATUS", header: "STATUS", title: "STATUS" };
    let actual = headerMap[logical] || null;
    if (!actual) {
      const found = headersInSheet.find(
        (h) => normalizeHeader(h) === normalizeHeader(logical),
      );
      actual = found || null;
    }
    return { key: logical, header: actual, title: logical };
  });
  return list;
}

function buildPrefixIndexes(rows, headerMap) {
  const ncmHeader = headerMap.NCM || "NCM";
  const prefix4Map = new Map();
  const allRowsByNcm = new Map();

  for (const row of rows) {
    const ncm = cleanNcmInput(row[ncmHeader] ?? row.NCM ?? "");
    if (!ncm) continue;
    row.__NCM_NORMALIZED = ncm;

    if (!allRowsByNcm.has(ncm)) allRowsByNcm.set(ncm, []);
    allRowsByNcm.get(ncm).push(row);

    const prefix4 = ncm.slice(0, 4);
    if (prefix4.length < 4) continue;
    if (!prefix4Map.has(prefix4)) prefix4Map.set(prefix4, []);
    prefix4Map.get(prefix4).push({ ncm, row });
  }

  return { prefix4Map, allRowsByNcm };
}

function parseDateFlexible(value) {
  const raw = cleanText(value);
  if (!raw) return null;

  let m = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (m) {
    let month = Number(m[1]);
    let day = Number(m[2]);
    let year = Number(m[3]);
    if (year < 100) year += 2000;
    const dt = new Date(year, month - 1, day);
    return Number.isNaN(dt.getTime()) ? null : dt;
  }

  m = raw.match(/^(\d{1,2})-(\d{1,2})-(\d{2,4})$/);
  if (m) {
    const dt = new Date(
      Number(m[3]) < 100 ? Number(m[3]) + 2000 : Number(m[3]),
      Number(m[2]) - 1,
      Number(m[1]),
    );
    return Number.isNaN(dt.getTime()) ? null : dt;
  }

  const parsed = new Date(raw);
  if (!Number.isNaN(parsed.getTime())) return parsed;
  return null;
}

function getExStatus(row, headerMap) {
  const endHeader = headerMap["FIM DA VIGÊNCIA"] || "FIM DA VIGÊNCIA";
  const value = row[endHeader] ?? "";
  const date = parseDateFlexible(value);
  if (!date) return { text: "SEM DATA", tone: "unknown", expired: false };

  const today = new Date();
  const current = new Date(
    today.getFullYear(),
    today.getMonth(),
    today.getDate(),
  );
  const end = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  if (end < current) return { text: "VENCIDO", tone: "expired", expired: true };
  return { text: "VÁLIDO", tone: "valid", expired: false };
}

async function readExcelBase(file, kind) {
  if (!file) throw new Error("Nenhum arquivo selecionado.");
  const lower = file.name.toLowerCase();
  if (!(lower.endsWith(".xlsx") || lower.endsWith(".xls"))) {
    throw new Error("Selecione um arquivo Excel (.xlsx ou .xls).");
  }
  const data = await file.arrayBuffer();
  const wb = XLSX.read(data, { type: "array" });
  const firstSheetName = wb.SheetNames[0];
  if (!firstSheetName) throw new Error("Arquivo sem abas legíveis.");
  const sheet = wb.Sheets[firstSheetName];
  const rows = sanitizeRowKeys(sheetToRows(sheet));
  const headers = rows.length ? Object.keys(rows[0]) : [];
  const headerMap = buildHeaderMap(headers);
  const validation = validateHeaders(kind, headerMap);
  if (!validation.valid) {
    throw new Error(
      `Arquivo ${kind} inválido. Colunas mínimas ausentes: ${validation.missing.join(", ")}.`,
    );
  }

  const displayCols = resolveDisplayColumns(kind, headerMap, headers);
  const indexes = buildPrefixIndexes(rows, headerMap);

  return {
    kind,
    sourceName: file.name,
    headerMap,
    displayCols,
    rows,
    prefix4Map: indexes.prefix4Map,
    allRowsByNcm: indexes.allRowsByNcm,
  };
}

function updateUploadButtonState() {
  ui.btnLoadFiles.disabled = !dataStore.ex && !dataStore.nve;
}

function resetRenderedResults() {
  renderTables([], [], [], []);
  setPill(ui.pillResultEX, "EX resultado", "—");
  setPill(ui.pillResultNVE, "NVE resultado", "—");
  setPill(ui.pillExpiredEX, "EX vencidos", "—");
  ui.hint.textContent = "Digite pelo menos 4 dígitos para iniciar a busca.";
}

function resetAll() {
  hideUploadError();
  dataStore.ex = null;
  dataStore.nve = null;
  dataStore.currentNcm = "";
  ui.fileInputEX.value = "";
  ui.fileInputNVE.value = "";
  setFileMeta(ui.fileMetaEX, null);
  setFileMeta(ui.fileMetaNVE, null);
  ui.fileStateEX.textContent = "Nenhum arquivo EX carregado.";
  ui.fileStateNVE.textContent = "Nenhum arquivo NVE carregado.";
  ui.ncmInput.value = "";
  setPill(ui.pillLoadedEX, "EX carregado", "—");
  setPill(ui.pillLoadedNVE, "NVE carregado", "—");
  resetRenderedResults();
  updateUploadButtonState();
}

function openUploadScreen() {
  ui.screenHome.classList.add("hidden");
  ui.screenSearch.classList.add("hidden");
  ui.screenUpload.classList.remove("hidden");
}

function openSearchScreen() {
  ui.screenHome.classList.add("hidden");
  ui.screenUpload.classList.add("hidden");
  ui.screenSearch.classList.remove("hidden");
  ui.ncmInput.focus();
}

async function handleBaseSelection(
  kind,
  input,
  metaEl,
  stateEl,
  loadedPill,
  label,
) {
  hideUploadError();
  const file = input.files && input.files[0] ? input.files[0] : null;
  setFileMeta(metaEl, file);
  if (!file) {
    dataStore[kind.toLowerCase()] = null;
    stateEl.textContent = `Nenhum arquivo ${label} carregado.`;
    setPill(loadedPill, `${label} carregado`, "NÃO", "bad");
    updateUploadButtonState();
    return;
  }

  try {
    const base = await readExcelBase(file, kind);
    dataStore[kind.toLowerCase()] = base;
    stateEl.textContent = `${label} carregado com sucesso. ${base.rows.length} linha(s) lida(s).`;
    setPill(loadedPill, `${label} carregado`, "SIM", "ok");
  } catch (error) {
    dataStore[kind.toLowerCase()] = null;
    stateEl.textContent = `${label} não carregado.`;
    setPill(loadedPill, `${label} carregado`, "NÃO", "bad");
    showUploadError(error.message || String(error));
  }
  updateUploadButtonState();
}

function searchByPrefix(base, ncmPrefix) {
  if (!base) return [];
  const prefix4 = ncmPrefix.slice(0, 4);
  const bucket = base.prefix4Map.get(prefix4) || [];
  if (ncmPrefix.length === 4) return bucket.map((item) => item.row);
  return bucket
    .filter((item) => item.ncm.startsWith(ncmPrefix))
    .map((item) => item.row);
}

function renderTable(tableEl, rows, cols, kind, headerMap) {
  const thead = tableEl.querySelector("thead");
  const tbody = tableEl.querySelector("tbody");
  thead.textContent = "";
  tbody.textContent = "";

  if (!cols || !cols.length) return;

  const trHead = document.createElement("tr");
  cols.forEach((col) => {
    const th = document.createElement("th");
    th.textContent = col.title;
    trHead.appendChild(th);
  });
  thead.appendChild(trHead);

  rows.forEach((row) => {
    const tr = document.createElement("tr");
    if (kind === "EX") {
      const status = getExStatus(row, headerMap);
      if (status.expired) tr.classList.add("expired");
    }

    cols.forEach((col) => {
      const td = document.createElement("td");
      if (kind === "EX" && col.key === "STATUS") {
        const status = getExStatus(row, headerMap);
        const badge = document.createElement("span");
        badge.className = `badge-status ${status.tone}`;
        badge.textContent = status.text;
        td.appendChild(badge);
      } else {
        const value = col.header ? (row[col.header] ?? "") : "";
        td.textContent = cleanText(value);
      }
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
}

function toggleTableVisibility(rows, wrapEl, emptyEl) {
  if (rows.length) {
    wrapEl.classList.remove("hidden");
    emptyEl.classList.add("hidden");
  } else {
    wrapEl.classList.add("hidden");
    emptyEl.classList.remove("hidden");
  }
}

function renderTables(nveRows, exRows, nveCols, exCols) {
  renderTable(
    ui.tableNVE,
    nveRows,
    nveCols,
    "NVE",
    dataStore.nve?.headerMap || {},
  );
  renderTable(ui.tableEX, exRows, exCols, "EX", dataStore.ex?.headerMap || {});
  toggleTableVisibility(nveRows, ui.wrapNVE, ui.emptyNVE);
  toggleTableVisibility(exRows, ui.wrapEX, ui.emptyEX);
  ui.countNVE.textContent = `${nveRows.length} linha(s)`;
  ui.countEX.textContent = `${exRows.length} linha(s)`;
}

function runSearch() {
  const ncm = cleanNcmInput(ui.ncmInput.value);
  dataStore.currentNcm = ncm;

  if (ncm.length < 4) {
    ui.hint.textContent = "Digite pelo menos 4 dígitos para iniciar a busca.";
    setPill(ui.pillResultNVE, "NVE resultado", "—");
    setPill(ui.pillResultEX, "EX resultado", "—");
    setPill(ui.pillExpiredEX, "EX vencidos", "—");
    renderTables(
      [],
      [],
      dataStore.nve?.displayCols || [],
      dataStore.ex?.displayCols || [],
    );
    return;
  }

  const nveRows = searchByPrefix(dataStore.nve, ncm);
  const exRows = searchByPrefix(dataStore.ex, ncm);
  const expiredCount = exRows.filter(
    (row) => getExStatus(row, dataStore.ex.headerMap).expired,
  ).length;

  ui.hint.textContent = `NCM normalizado: ${ncm}${ncm.length === 8 ? ` • formatado: ${formatNcm(ncm)}` : ""}`;
  setPill(
    ui.pillResultNVE,
    "NVE resultado",
    nveRows.length ? "SIM" : "NÃO",
    nveRows.length ? "ok" : "bad",
  );
  setPill(
    ui.pillResultEX,
    "EX resultado",
    exRows.length ? "SIM" : "NÃO",
    exRows.length ? "ok" : "bad",
  );
  setPill(
    ui.pillExpiredEX,
    "EX vencidos",
    String(expiredCount),
    expiredCount ? "warn" : "ok",
  );

  renderTables(
    nveRows,
    exRows,
    dataStore.nve?.displayCols || [],
    dataStore.ex?.displayCols || [],
  );
}

function escapeHtml(text) {
  return String(text ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function buildTableHtml(rows, cols, kind, headerMap, title) {
  if (!rows.length || !cols.length) return "";
  const headers = cols.map((c) => `<th>${escapeHtml(c.title)}</th>`).join("");
  const body = rows
    .map((row) => {
      const status = kind === "EX" ? getExStatus(row, headerMap) : null;
      const cells = cols
        .map((col) => {
          if (kind === "EX" && col.key === "STATUS") {
            return `<td>${escapeHtml(status.text)}</td>`;
          }
          const value = col.header ? cleanText(row[col.header] ?? "") : "";
          return `<td>${escapeHtml(value)}</td>`;
        })
        .join("");
      return `<tr${status && status.expired ? ' class="expired"' : ""}>${cells}</tr>`;
    })
    .join("");

  return `
        <section class="section">
          <h2>${escapeHtml(title)}</h2>
          <div class="meta">${rows.length} linha(s)</div>
          <table>
            <thead><tr>${headers}</tr></thead>
            <tbody>${body}</tbody>
          </table>
        </section>
      `;
}

function openPrintWindow() {
  const ncm = cleanNcmInput(ui.ncmInput.value);
  if (ncm.length < 4) {
    alert("Digite um NCM com pelo menos 4 dígitos para gerar o relatório.");
    return;
  }

  const nveBase = dataStore.nve;
  const exBase = dataStore.ex;

  const nveRows = searchByPrefix(nveBase, ncm);
  const exRows = searchByPrefix(exBase, ncm);

  if (!nveRows.length && !exRows.length) {
    alert("Não há linhas para exportar no NCM informado.");
    return;
  }

  const htmlNVE = buildTableHtml(
    nveRows,
    nveBase ? nveBase.displayCols : [],
    "NVE",
    nveBase ? nveBase.headerMap : {},
    "NVE",
  );

  const htmlEX = buildTableHtml(
    exRows,
    exBase ? exBase.displayCols : [],
    "EX",
    exBase ? exBase.headerMap : {},
    "EX",
  );

  const when = new Date().toLocaleString("pt-BR");
  const formatted = formatNcm(ncm);

  const doc = `<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<title>Relatório NCM ${escapeHtml(formatted || ncm)}</title>
<style>
body{font-family:Arial,sans-serif;margin:24px;color:#111}
.meta{margin:0 0 14px;color:#444;font-size:12px}
h1{margin:0 0 8px;font-size:22px}
h2{margin:22px 0 8px;font-size:18px}
.section{margin-bottom:20px;page-break-inside:avoid}
table{width:100%;border-collapse:collapse;table-layout:fixed}
th,td{border:1px solid #ccc;padding:8px;vertical-align:top;white-space:pre-wrap;word-break:break-word;overflow-wrap:anywhere;font-size:12px}
th{background:#f2f2f2;text-align:left}
tr.expired td{background:#fff2f2}
</style>
</head>
<body>
<h1>Lista filtrada NVE/Ex-tarifário para o NCM ${escapeHtml(formatted || ncm)}</h1>
<p class="meta">Gerado em ${escapeHtml(when)}</p>
${htmlNVE}
${htmlEX}
</body>
</html>`;

  const w = window.open("", "_blank");
  if (!w) {
    return;
  }

  w.document.open();
  w.document.write(doc);
  w.document.close();
  setTimeout(() => {
    try {
      w.focus();
      w.print();
    } catch (e) {}
  }, 80);
}

ui.btnStartHome.addEventListener("click", openUploadScreen);

ui.fileInputEX.addEventListener("change", () =>
  handleBaseSelection(
    "EX",
    ui.fileInputEX,
    ui.fileMetaEX,
    ui.fileStateEX,
    ui.pillLoadedEX,
    "EX",
  ),
);
ui.fileInputNVE.addEventListener("change", () =>
  handleBaseSelection(
    "NVE",
    ui.fileInputNVE,
    ui.fileMetaNVE,
    ui.fileStateNVE,
    ui.pillLoadedNVE,
    "NVE",
  ),
);

ui.btnLoadFiles.addEventListener("click", () => {
  if (!dataStore.ex && !dataStore.nve) {
    showUploadError("Carregue pelo menos uma planilha válida para continuing.");
    return;
  }
  hideUploadError();
  openSearchScreen();
  resetRenderedResults();
});

ui.btnResetUpload.addEventListener("click", resetAll);

ui.btnChangeFiles.addEventListener("click", () => {
  openUploadScreen();
});

ui.ncmInput.addEventListener("input", () => {
  clearTimeout(debounceTimer);
  debounceTimer = setTimeout(runSearch, 80);
});

ui.btnPrintReport.addEventListener("click", openPrintWindow);
ui.btnToTop.addEventListener("click", () =>
  window.scrollTo({ top: 0, behavior: "smooth" }),
);
ui.themeToggle.addEventListener("click", () => {
  const current =
    document.documentElement.getAttribute("data-theme") || "light";
  setTheme(current === "dark" ? "light" : "dark");
});

initTheme();
resetAll();
setPill(ui.pillLoadedEX, "EX carregado", "—");
setPill(ui.pillLoadedNVE, "NVE carregado", "—");
