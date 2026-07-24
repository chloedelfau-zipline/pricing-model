const state = {
  lines: [],
  report: null,
  activeView: "bom",
};

const isFilePreview = window.location.protocol === "file:";

const elements = {
  apiMode: document.querySelector("#apiMode"),
  fileInput: document.querySelector("#fileInput"),
  fileLabel: document.querySelector("#fileLabel"),
  dropZone: document.querySelector("#dropZone"),
  exportButton: document.querySelector("#exportButton"),
  addLineButton: document.querySelector("#addLineButton"),
  clearButton: document.querySelector("#clearButton"),
  quoteButton: document.querySelector("#quoteButton"),
  statusText: document.querySelector("#statusText"),
  lineCount: document.querySelector("#lineCount"),
  bomBody: document.querySelector("#bomBody"),
  mpnFilter: document.querySelector("#mpnFilter"),
  partTypeFilter: document.querySelector("#partTypeFilter"),
  procurementFilter: document.querySelector("#procurementFilter"),
  clearFiltersButton: document.querySelector("#clearFiltersButton"),
  defaultDemand: document.querySelector("#defaultDemand"),
  bufferPct: document.querySelector("#bufferPct"),
  country: document.querySelector("#country"),
  currency: document.querySelector("#currency"),
  supplierFilter: document.querySelector("#supplierFilter"),
  authorizedOnly: document.querySelector("#authorizedOnly"),
  matchLimit: document.querySelector("#matchLimit"),
  summaryGrid: document.querySelector("#summaryGrid"),
  quotedMetric: document.querySelector("#quotedMetric"),
  openMetric: document.querySelector("#openMetric"),
  riskMetric: document.querySelector("#riskMetric"),
  costMetric: document.querySelector("#costMetric"),
  bomView: document.querySelector("#bomView"),
  reportPanel: document.querySelector("#reportView"),
  bomViewButton: document.querySelector("#bomViewButton"),
  reportViewButton: document.querySelector("#reportViewButton"),
  reportMeta: document.querySelector("#reportMeta"),
  reportBody: document.querySelector("#reportBody"),
};

const columnAliases = {
  ziplinePn: [
    "id",
    "itemid",
    "itemnumber",
    "partid",
    "ziplinepn",
    "ziplinepartnumber",
    "zippn",
    "zlpn",
    "zlpartnumber",
  ],
  lineNumber: [
    "line",
    "lineno",
    "item",
    "itemno",
    "reference",
    "ref",
    "findnumber",
  ],
  description: [
    "description",
    "desc",
    "partdescription",
    "comment",
    "value",
    "revisionname",
    "itemname",
    "objectname",
    "name",
  ],
  mpn: [
    "mpn",
    "manufacturerpartnumber",
    "mfgpartnumber",
    "mfrpartnumber",
    "manufacturerpn",
    "mfrpn",
    "partnumber",
    "part",
  ],
  alternatives: [
    "alternatives",
    "alternate",
    "alternates",
    "alt",
    "alternatempn",
    "approvedalternatives",
    "approvedalternates",
    "substitutes",
  ],
  bomQty: [
    "qty",
    "quantity",
    "bomqty",
    "quantityper",
    "qtyper",
    "perassembly",
    "qtyassy",
    "qtyperassembly",
  ],
  partType: ["parttype", "type", "itemtype"],
  procurementIntent: [
    "procurementintent",
    "procurement",
    "sourcingintent",
    "buybuild",
    "makebuy",
  ],
};

function normalizeHeader(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "");
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let field = "";
  let inQuotes = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const next = text[index + 1];

    if (char === '"' && inQuotes && next === '"') {
      field += '"';
      index += 1;
      continue;
    }

    if (char === '"') {
      inQuotes = !inQuotes;
      continue;
    }

    if ((char === "," || char === "\t") && !inQuotes) {
      row.push(field);
      field = "";
      continue;
    }

    if ((char === "\n" || char === "\r") && !inQuotes) {
      if (char === "\r" && next === "\n") index += 1;
      row.push(field);
      if (row.some((cell) => cell.trim() !== "")) rows.push(row);
      row = [];
      field = "";
      continue;
    }

    field += char;
  }

  row.push(field);
  if (row.some((cell) => cell.trim() !== "")) rows.push(row);
  return rows;
}

function readUInt16(view, offset) {
  return view.getUint16(offset, true);
}

function readUInt32(view, offset) {
  return view.getUint32(offset, true);
}

function findZipEndOfCentralDirectory(view) {
  for (let offset = view.byteLength - 22; offset >= 0; offset -= 1) {
    if (readUInt32(view, offset) === 0x06054b50) return offset;
  }
  throw new Error("This does not look like a readable Excel workbook.");
}

async function inflateRaw(bytes) {
  if (!("DecompressionStream" in window)) {
    throw new Error(
      "This browser cannot read Excel BOMs directly. Open the app at http://127.0.0.1:4180 and upload again.",
    );
  }

  try {
    const stream = new Blob([bytes]).stream().pipeThrough(
      new DecompressionStream("deflate-raw"),
    );
    return new Uint8Array(await new Response(stream).arrayBuffer());
  } catch {
    throw new Error(
      "This browser cannot unzip Excel BOMs directly. Open the app at http://127.0.0.1:4180 and upload again.",
    );
  }
}

async function readZipEntries(arrayBuffer) {
  const bytes = new Uint8Array(arrayBuffer);
  const view = new DataView(arrayBuffer);
  const decoder = new TextDecoder("utf-8");
  const entries = new Map();
  const eocd = findZipEndOfCentralDirectory(view);
  const entryCount = readUInt16(view, eocd + 10);
  let offset = readUInt32(view, eocd + 16);

  for (let entryIndex = 0; entryIndex < entryCount; entryIndex += 1) {
    if (readUInt32(view, offset) !== 0x02014b50) {
      throw new Error("The Excel workbook directory is invalid.");
    }

    const compressionMethod = readUInt16(view, offset + 10);
    const compressedSize = readUInt32(view, offset + 20);
    const fileNameLength = readUInt16(view, offset + 28);
    const extraLength = readUInt16(view, offset + 30);
    const commentLength = readUInt16(view, offset + 32);
    const localHeaderOffset = readUInt32(view, offset + 42);
    const fileName = decoder.decode(
      bytes.subarray(offset + 46, offset + 46 + fileNameLength),
    );

    if (readUInt32(view, localHeaderOffset) !== 0x04034b50) {
      throw new Error(`The Excel workbook entry is invalid for ${fileName}.`);
    }

    const localNameLength = readUInt16(view, localHeaderOffset + 26);
    const localExtraLength = readUInt16(view, localHeaderOffset + 28);
    const dataStart = localHeaderOffset + 30 + localNameLength + localExtraLength;
    const compressed = bytes.subarray(dataStart, dataStart + compressedSize);

    let data = null;
    if (compressionMethod === 0) {
      data = new Uint8Array(compressed);
    } else if (compressionMethod === 8) {
      data = await inflateRaw(compressed);
    }

    if (data) entries.set(fileName, data);
    offset += 46 + fileNameLength + extraLength + commentLength;
  }

  return entries;
}

function decodeXml(value) {
  return String(value || "")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

function getXmlAttr(attrs, name) {
  const match = attrs.match(new RegExp(`(?:^|\\s)${name}="([^"]*)"`, "i"));
  return match ? decodeXml(match[1]) : "";
}

function extractTagText(xml, tagName) {
  const match = xml.match(new RegExp(`<${tagName}\\b[^>]*>([\\s\\S]*?)<\\/${tagName}>`, "i"));
  return match ? decodeXml(match[1]) : "";
}

function parseSharedStrings(xml) {
  if (!xml) return [];
  const strings = [];
  const stringRegex = /<si\b[\s\S]*?<\/si>/gi;

  for (const stringMatch of xml.matchAll(stringRegex)) {
    const textParts = [];
    for (const textMatch of stringMatch[0].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/gi)) {
      textParts.push(decodeXml(textMatch[1]));
    }
    strings.push(textParts.join(""));
  }

  return strings;
}

function columnIndexFromCellRef(ref) {
  const column = String(ref || "").match(/[A-Z]+/i)?.[0] || "";
  let index = 0;
  for (const char of column.toUpperCase()) {
    index = index * 26 + char.charCodeAt(0) - 64;
  }
  return Math.max(0, index - 1);
}

function parseWorksheetRows(xml, sharedStrings) {
  const rows = [];
  const rowRegex = /<row\b[^>]*>[\s\S]*?<\/row>/gi;

  for (const rowMatch of xml.matchAll(rowRegex)) {
    const cells = [];
    const cellRegex = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/gi;
    let fallbackColumn = 0;

    for (const cellMatch of rowMatch[0].matchAll(cellRegex)) {
      const attrs = cellMatch[1];
      const body = cellMatch[2] || "";
      const type = getXmlAttr(attrs, "t");
      const ref = getXmlAttr(attrs, "r");
      const columnIndex = ref ? columnIndexFromCellRef(ref) : fallbackColumn;
      fallbackColumn = columnIndex + 1;

      let value = "";
      if (type === "s") {
        value = sharedStrings[Number(extractTagText(body, "v"))] || "";
      } else if (type === "inlineStr") {
        value = extractTagText(body, "t");
      } else {
        value = extractTagText(body, "v");
      }

      cells[columnIndex] = value;
    }

    if (cells.some((cell) => String(cell || "").trim())) {
      rows.push(cells.map((cell) => cell ?? ""));
    }
  }

  return rows;
}

async function parseWorkbookRows(arrayBuffer) {
  const entries = await readZipEntries(arrayBuffer);
  const decoder = new TextDecoder("utf-8");
  const sharedStrings = parseSharedStrings(
    entries.has("xl/sharedStrings.xml")
      ? decoder.decode(entries.get("xl/sharedStrings.xml"))
      : "",
  );

  const worksheets = [...entries.entries()]
    .filter(([name]) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(name))
    .map(([name, contents]) => ({
      name,
      rows: parseWorksheetRows(decoder.decode(contents), sharedStrings),
    }))
    .sort((a, b) => b.rows.length - a.rows.length);

  const bestSheet = worksheets.find((sheet) => sheet.rows.length > 1) || worksheets[0];
  if (!bestSheet?.rows.length) {
    throw new Error("No BOM rows were found in the workbook.");
  }

  return bestSheet.rows;
}

function splitAlternatives(value) {
  return String(value || "")
    .split(/[;|\n]/)
    .map((item) => item.trim())
    .filter(Boolean);
}

function findColumnIndex(headers, field) {
  const normalizedHeaders = headers.map(normalizeHeader);
  return normalizedHeaders.findIndex((header) =>
    columnAliases[field].includes(header),
  );
}

function rowsToLines(rows) {
  if (!rows.length) return [];

  const headers = rows[0];
  const indexes = Object.fromEntries(
    Object.keys(columnAliases).map((field) => [field, findColumnIndex(headers, field)]),
  );

  const hasRecognizedHeader =
    indexes.mpn >= 0 ||
    indexes.bomQty >= 0 ||
    indexes.ziplinePn >= 0 ||
    indexes.description >= 0;
  const dataRows = hasRecognizedHeader ? rows.slice(1) : rows;
  const fallback = {
    ziplinePn: 0,
    lineNumber: 1,
    description: 2,
    mpn: 3,
    alternatives: 4,
    bomQty: 5,
    partType: 6,
    procurementIntent: 7,
  };

  return dataRows
    .map((row, index) => {
      const get = (field) => {
        const preferred = indexes[field];
        const fallbackIndex = fallback[field];
        const cell =
          preferred >= 0
            ? row[preferred]
            : !hasRecognizedHeader
              ? row[fallbackIndex]
              : "";
        return String(cell || "").trim();
      };

      return {
        id: makeId(),
        ziplinePn: get("ziplinePn"),
        lineNumber: get("lineNumber") || String(index + 1),
        description: get("description"),
        mpn: get("mpn"),
        alternatives: splitAlternatives(get("alternatives")),
        bomQty: get("bomQty") || "1",
        partType: get("partType"),
        procurementIntent: get("procurementIntent"),
      };
    })
    .filter((line) => line.mpn || line.description || line.ziplinePn);
}

function setStatus(message, tone = "") {
  elements.statusText.textContent = message;
  elements.statusText.dataset.tone = tone;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function formatCurrency(value, currency = "USD", maxDigits = 2) {
  const number = Number(value || 0);
  try {
    return new Intl.NumberFormat(undefined, {
      style: "currency",
      currency,
      maximumFractionDigits: maxDigits,
    }).format(number);
  } catch {
    return `${currency} ${number.toFixed(maxDigits)}`;
  }
}

function formatNumber(value) {
  return new Intl.NumberFormat().format(Number(value || 0));
}

function assetUrl(path) {
  return isFilePreview ? path.replace(/^\//, "") : path;
}

function numberFrom(value, fallback = 0) {
  const parsed = Number(String(value ?? "").replace(/,/g, "").trim());
  return Number.isFinite(parsed) ? parsed : fallback;
}

function calculateRequiredQty(line) {
  const bomQty = Math.max(0, numberFrom(line.bomQty, 0));
  const buildDemand = Math.max(0, numberFrom(elements.defaultDemand.value, 0));
  const bufferPct = Math.max(0, numberFrom(elements.bufferPct.value, 0));
  const requiredQty = (bomQty * buildDemand * (100 + bufferPct)) / 100;
  return Math.ceil(requiredQty - 1e-9);
}

function lineCountText(count) {
  return `${count} ${count === 1 ? "line" : "lines"}`;
}

function normalizedText(value) {
  return String(value || "").trim().toLowerCase();
}

function uniqueFieldValues(field) {
  const seen = new Set();
  const values = [];

  for (const line of state.lines) {
    const value = String(line[field] || "").trim();
    const key = value.toLowerCase();
    if (value && !seen.has(key)) {
      seen.add(key);
      values.push(value);
    }
  }

  return values.sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
}

function renderSelectOptions(select, values, allLabel) {
  const selectedValue = select.value;
  select.innerHTML = [
    `<option value="">${escapeHtml(allLabel)}</option>`,
    ...values.map((value) => {
      const selected = value === selectedValue ? " selected" : "";
      return `<option value="${escapeHtml(value)}"${selected}>${escapeHtml(value)}</option>`;
    }),
  ].join("");

  if (selectedValue && !values.includes(selectedValue)) {
    select.value = "";
  }
}

function renderFilterOptions() {
  renderSelectOptions(
    elements.partTypeFilter,
    uniqueFieldValues("partType"),
    "All part types",
  );
  renderSelectOptions(
    elements.procurementFilter,
    uniqueFieldValues("procurementIntent"),
    "All procurement intents",
  );
}

function hasActiveFilters() {
  return Boolean(
    elements.mpnFilter.value.trim() ||
      elements.partTypeFilter.value ||
      elements.procurementFilter.value,
  );
}

function lineMatchesFilters(line) {
  const mpnFilter = normalizedText(elements.mpnFilter.value);
  const partTypeFilter = normalizedText(elements.partTypeFilter.value);
  const procurementFilter = normalizedText(elements.procurementFilter.value);
  const candidateMpns = [line.mpn, ...(line.alternatives || [])]
    .map(normalizedText)
    .join(" ");

  return (
    (!mpnFilter || candidateMpns.includes(mpnFilter)) &&
    (!partTypeFilter || normalizedText(line.partType) === partTypeFilter) &&
    (!procurementFilter ||
      normalizedText(line.procurementIntent) === procurementFilter)
  );
}

function getFilteredLines() {
  return state.lines.filter(lineMatchesFilters);
}

function clearFilterInputs() {
  elements.mpnFilter.value = "";
  elements.partTypeFilter.value = "";
  elements.procurementFilter.value = "";
}

function makeId() {
  if (window.crypto?.randomUUID) {
    return window.crypto.randomUUID();
  }

  return `id-${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`;
}

function getSettings() {
  return {
    defaultDemand: elements.defaultDemand.value,
    bufferPct: elements.bufferPct.value,
    country: elements.country.value,
    currency: elements.currency.value,
    allowedSuppliers: elements.supplierFilter.value
      .split(",")
      .map((supplier) => supplier.trim())
      .filter(Boolean),
    authorizedOnly: elements.authorizedOnly.checked,
    matchLimit: elements.matchLimit.value,
  };
}

function renderSummary(summary = null) {
  const currency = summary?.currency || elements.currency.value || "USD";
  elements.quotedMetric.textContent = summary
    ? `${summary.quotedLines}/${summary.lineCount}`
    : "0";
  elements.openMetric.textContent = formatNumber(summary?.unresolvedLines || 0);
  elements.riskMetric.textContent = formatNumber(summary?.stockRisks || 0);
  elements.costMetric.textContent = formatCurrency(summary?.totalCost || 0, currency);
}

function showView(view) {
  const nextView = view === "report" && state.report ? "report" : "bom";
  const isReport = nextView === "report";
  state.activeView = nextView;

  elements.bomView.hidden = isReport;
  elements.reportPanel.hidden = !isReport;
  elements.bomViewButton.classList.toggle("is-active", !isReport);
  elements.reportViewButton.classList.toggle("is-active", isReport);
  elements.bomViewButton.setAttribute("aria-selected", String(!isReport));
  elements.reportViewButton.setAttribute("aria-selected", String(isReport));
  elements.lineCount.hidden = isReport;
  elements.reportMeta.hidden = !isReport || !state.report;
}

function resetReportView() {
  state.report = null;
  elements.exportButton.disabled = true;
  elements.reportViewButton.disabled = true;
  elements.reportBody.innerHTML = "";
  elements.reportMeta.textContent = "";
  renderSummary();
  showView("bom");
}

function addEmptyLine() {
  resetReportView();
  state.lines.push({
    id: makeId(),
    ziplinePn: "",
    lineNumber: String(state.lines.length + 1),
    description: "",
    mpn: "",
    alternatives: [],
    bomQty: "1",
    partType: "",
    procurementIntent: "",
  });
  renderBom();
}

function updateLine(id, field, value) {
  const line = state.lines.find((item) => item.id === id);
  if (!line) return;
  if (field === "alternatives") {
    line.alternatives = splitAlternatives(value);
  } else {
    line[field] = value;
  }
}

function removeLine(id) {
  resetReportView();
  state.lines = state.lines.filter((line) => line.id !== id);
  renderBom();
}

function refreshComputedQuantities() {
  for (const row of elements.bomBody.querySelectorAll("tr[data-id]")) {
    const line = state.lines.find((item) => item.id === row.dataset.id);
    const target = row.querySelector("[data-computed-required]");
    if (line && target) {
      target.textContent = formatNumber(calculateRequiredQty(line));
    }
  }
}

function renderBom() {
  renderFilterOptions();

  const visibleLines = getFilteredLines();
  elements.lineCount.textContent = hasActiveFilters()
    ? `${lineCountText(visibleLines.length)} shown of ${lineCountText(state.lines.length)}`
    : lineCountText(state.lines.length);

  if (!state.lines.length) {
    elements.bomBody.innerHTML = `
      <tr class="empty-row">
        <td colspan="8">No BOM lines.</td>
      </tr>
    `;
    return;
  }

  if (!visibleLines.length) {
    elements.bomBody.innerHTML = `
      <tr class="empty-row">
        <td colspan="8">No BOM lines match the current filters.</td>
      </tr>
    `;
    return;
  }

  elements.bomBody.innerHTML = visibleLines
    .map(
      (line) => `
        <tr data-id="${escapeHtml(line.id)}">
          <td>
            <input data-field="ziplinePn" value="${escapeHtml(line.ziplinePn)}" />
          </td>
          <td>
            <input data-field="description" value="${escapeHtml(line.description)}" />
          </td>
          <td>
            <input data-field="mpn" value="${escapeHtml(line.mpn)}" />
          </td>
          <td>
            <input data-field="bomQty" type="number" min="0" step="any" value="${escapeHtml(
              line.bomQty,
            )}" />
          </td>
          <td>
            <input data-field="partType" value="${escapeHtml(line.partType)}" />
          </td>
          <td>
            <input data-field="procurementIntent" value="${escapeHtml(
              line.procurementIntent,
            )}" />
          </td>
          <td>
            <span class="computed-qty" data-computed-required>${formatNumber(
              calculateRequiredQty(line),
            )}</span>
          </td>
          <td>
            <button class="icon-button" data-action="remove" title="Remove line" type="button" aria-label="Remove line">
              <svg viewBox="0 0 24 24" aria-hidden="true">
                <path d="M3 6h18M8 6V4h8v2m-9 0 1 15h8l1-15" />
              </svg>
            </button>
          </td>
        </tr>
      `,
    )
    .join("");
}

function stockBadgeClass(status) {
  if (status === "Quoted") return "good";
  if (status === "Stock short" || status === "No match" || status === "No priced offer") {
    return "bad";
  }
  return "warn";
}

function renderReport(report) {
  state.report = report;
  elements.exportButton.disabled = false;
  elements.reportViewButton.disabled = false;

  renderSummary(report.summary);

  elements.reportMeta.textContent = "Nexar/Octopart live data";

  elements.reportBody.innerHTML = report.lines
    .map((line) => {
      const rec = line.recommendation;
      const runner = line.runnerUp;
      const unit = rec
        ? `${formatCurrency(rec.unitPrice, rec.currency, 4)} @ ${formatNumber(
            rec.priceBreakQty,
          )}+`
        : "";
      const stock =
        rec?.inventoryLevel === null || rec?.inventoryLevel === undefined
          ? "Unknown"
          : formatNumber(rec.inventoryLevel);
      const marketAvailability =
        rec?.marketAvailability === null || rec?.marketAvailability === undefined
          ? ""
          : formatNumber(rec.marketAvailability);
      const lead = rec?.leadDays ? `${rec.leadDays} days` : "";
      const total = rec ? formatCurrency(rec.totalCost, rec.currency) : "";
      const runnerDelta =
        rec && runner ? Number(runner.totalCost || 0) - Number(rec.totalCost || 0) : 0;
      const runnerText =
        runner && rec
          ? `Next: ${runner.supplier} ${formatCurrency(
              runner.totalCost,
              runner.currency || rec.currency || report.summary.currency,
            )} (${runnerDelta >= 0 ? "+" : ""}${formatCurrency(
              runnerDelta,
              runner.currency || rec.currency || report.summary.currency,
            )})`
          : "";
      const inputText =
        rec?.inputMpn &&
        rec?.quotedMpn &&
        normalizedText(rec.inputMpn) !== normalizedText(rec.quotedMpn)
          ? `Input: ${rec.inputMpn}`
          : "";
      const matchedManufacturer = rec?.manufacturer
        ? `Matched manufacturer: ${rec.manufacturer}`
        : "";
      const orderText = rec?.minimumOrderQty > 1 ? `MOQ ${rec.minimumOrderQty}` : "";
      const skuText = rec?.supplierSku ? `Seller SKU: ${rec.supplierSku}` : "";
      const partSubtext = [
        line.description,
        line.partType,
        line.procurementIntent,
        rec?.sourceRole,
        inputText,
        matchedManufacturer,
        orderText,
        skuText,
        (line.similarParts || []).length ? `Similar: ${line.similarParts.join(", ")}` : "",
      ]
        .filter(Boolean)
        .join(" | ");

      return `
        <tr>
          <td>${escapeHtml(line.ziplinePn || line.lineNumber)}</td>
          <td>${formatNumber(line.requiredQty)}</td>
          <td>
            <div class="part-cell">
              <strong>${escapeHtml(
                rec?.quotedMpn || line.primaryMpn || "No manufacturer MPN",
              )}</strong>
              <span class="subtle">${escapeHtml(partSubtext)}</span>
            </div>
          </td>
          <td>
            <div class="supplier-cell">
              <strong>${escapeHtml(rec?.supplier || "")}</strong>
              ${runnerText ? `<span class="subtle">${escapeHtml(runnerText)}</span>` : ""}
            </div>
          </td>
          <td>${escapeHtml(unit)}</td>
          <td>${escapeHtml(total)}</td>
          <td>
            <div class="stock-cell">
              <strong>${escapeHtml(stock)}</strong>
              ${
                marketAvailability
                  ? `<span class="subtle">Market ${escapeHtml(marketAvailability)}</span>`
                  : ""
              }
            </div>
          </td>
          <td>${escapeHtml(lead)}</td>
          <td><span class="badge ${stockBadgeClass(line.status)}">${escapeHtml(
            line.status,
          )}</span></td>
          <td>${
            rec?.url
              ? `<a class="report-link" href="${escapeHtml(rec.url)}" target="_blank" rel="noreferrer">Open</a>`
              : ""
          }</td>
        </tr>
      `;
    })
    .join("");

  showView("report");
}

function quoteCsv(report) {
  const rows = [
    [
      "Line",
      "Zipline PN",
      "Description",
      "Primary MPN",
      "Candidates",
      "Part Type",
      "Procurement Intent",
      "Required Qty",
      "Recommended MPN",
      "Matched Manufacturer",
      "Supplier",
      "Supplier SKU",
      "Next Best Supplier",
      "Next Best Total",
      "Next Best Delta",
      "Unit Price",
      "Currency",
      "Price Break",
      "MOQ",
      "Quoted Qty",
      "Extended Cost",
      "Inventory",
      "Market Availability",
      "Lead Days",
      "Status",
      "URL",
    ],
    ...report.lines.map((line) => {
      const rec = line.recommendation || {};
      const runner = line.runnerUp || {};
      const runnerDelta =
        line.recommendation && line.runnerUp
          ? Number(line.runnerUp.totalCost || 0) -
            Number(line.recommendation.totalCost || 0)
          : "";
      return [
        line.lineNumber,
        line.ziplinePn,
        line.description,
        line.primaryMpn,
        line.candidateMpns.join("; "),
        line.partType || "",
        line.procurementIntent || "",
        line.requiredQty,
        rec.quotedMpn || "",
        rec.manufacturer || "",
        rec.supplier || "",
        rec.supplierSku || "",
        runner.supplier || "",
        runner.totalCost ?? "",
        runnerDelta,
        rec.unitPrice ?? "",
        rec.currency || report.summary.currency,
        rec.priceBreakQty ?? "",
        rec.minimumOrderQty ?? "",
        rec.quotedQty ?? "",
        rec.totalCost ?? "",
        rec.inventoryLevel ?? "",
        rec.marketAvailability ?? "",
        rec.leadDays ?? "",
        line.status,
        rec.url || "",
      ];
    }),
  ];

  return rows
    .map((row) =>
      row
        .map((cell) => {
          const value = String(cell ?? "");
          return /[",\n]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value;
        })
        .join(","),
    )
    .join("\n");
}

async function loadFile(file) {
  const extension = file.name.split(".").pop().toLowerCase();
  let rows;

  elements.dropZone.classList.remove("uploaded");
  elements.fileLabel.textContent = "Importing BOM...";
  setStatus("");

  if (["xlsx", "xlsm"].includes(extension)) {
    const workbookBuffer = await file.arrayBuffer();

    if (!isFilePreview) {
      try {
        const response = await fetch(`/api/import-bom?filename=${encodeURIComponent(file.name)}`, {
          method: "POST",
          headers: { "Content-Type": "application/octet-stream" },
          body: workbookBuffer.slice(0),
        });
        const payload = await response.json();
        if (!response.ok) throw new Error(payload.error || "Workbook import failed.");
        rows = payload.rows;
      } catch {
        rows = await parseWorkbookRows(workbookBuffer);
      }
    } else {
      rows = await parseWorkbookRows(workbookBuffer);
    }
  } else {
    const text = await file.text();
    rows = parseCsv(text);
  }

  const lines = rowsToLines(rows);
  if (!lines.length) {
    throw new Error("No BOM lines were found in that file.");
  }

  state.lines = lines;
  clearFilterInputs();
  resetReportView();
  elements.dropZone.classList.add("uploaded");
  elements.fileLabel.textContent = `BOM uploaded: ${file.name}`;
  renderBom();
  setStatus("");
}

async function handleSelectedFile(file) {
  if (!file) return;

  try {
    await loadFile(file);
  } catch (error) {
    elements.dropZone.classList.remove("uploaded");
    elements.fileLabel.textContent = "Upload BOM";
    setStatus(error.message || "Unable to upload BOM.", "error");
  } finally {
    elements.fileInput.value = "";
  }
}

async function generateReport() {
  const filteredLines = getFilteredLines();
  const validLines = filteredLines.filter(
    (line) =>
      Number(line.bomQty) > 0 && (line.mpn || line.description || line.ziplinePn),
  );
  if (!validLines.length) {
    setStatus(
      hasActiveFilters()
        ? "No matching BOM lines with a quantity."
        : "Add at least one BOM line with a quantity.",
      "warning",
    );
    return;
  }

  elements.quoteButton.disabled = true;
  elements.quoteButton.textContent = "Generating...";
  setStatus("Generating supplier report.");

  try {
    if (isFilePreview) {
      setStatus(
        "Open the local server version of the app to generate live Nexar/Octopart data.",
        "error",
      );
      return;
    }

    const response = await fetch("/api/quote", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        lines: validLines,
        settings: getSettings(),
      }),
    });

    const payload = await response.json();
    if (!response.ok) throw new Error(payload.error || "Quote request failed.");

    renderReport(payload);
    setStatus("Report generated with live Octopart data.");
  } catch (error) {
    setStatus(error.message || "Unable to generate report.", "error");
  } finally {
    elements.quoteButton.disabled = false;
    elements.quoteButton.textContent = "Generate Report";
  }
}

async function checkHealth() {
  if (!elements.apiMode) {
    return;
  }

  if (isFilePreview) {
    elements.apiMode.textContent = "Local file preview";
    return;
  }

  try {
    const response = await fetch("/api/health");
    const health = await response.json();
    elements.apiMode.textContent = health.nexarConfigured
      ? "Live Nexar/Octopart API"
      : "Nexar API not configured";
  } catch {
    elements.apiMode.textContent = "Server unavailable";
  }
}

elements.fileInput.addEventListener("change", (event) => {
  const [file] = event.target.files;
  handleSelectedFile(file);
});

elements.dropZone.addEventListener("dragover", (event) => {
  event.preventDefault();
  elements.dropZone.classList.add("dragging");
});

elements.dropZone.addEventListener("dragleave", () => {
  elements.dropZone.classList.remove("dragging");
});

elements.dropZone.addEventListener("drop", (event) => {
  event.preventDefault();
  elements.dropZone.classList.remove("dragging");
  const [file] = event.dataTransfer.files;
  handleSelectedFile(file);
});

elements.bomBody.addEventListener("input", (event) => {
  const input = event.target.closest("input");
  if (!input) return;
  const row = input.closest("tr");
  updateLine(row.dataset.id, input.dataset.field, input.value);
  resetReportView();
  if (input.dataset.field === "bomQty") {
    refreshComputedQuantities();
  }
  if (input.dataset.field === "partType" || input.dataset.field === "procurementIntent") {
    renderFilterOptions();
  }
});

elements.bomBody.addEventListener("click", (event) => {
  const button = event.target.closest("button[data-action='remove']");
  if (!button) return;
  removeLine(button.closest("tr").dataset.id);
});

elements.addLineButton.addEventListener("click", addEmptyLine);
elements.bomViewButton.addEventListener("click", () => {
  showView("bom");
});
elements.reportViewButton.addEventListener("click", () => {
  showView("report");
});
elements.mpnFilter.addEventListener("input", () => {
  resetReportView();
  renderBom();
});
elements.partTypeFilter.addEventListener("change", () => {
  resetReportView();
  renderBom();
});
elements.procurementFilter.addEventListener("change", () => {
  resetReportView();
  renderBom();
});
elements.clearFiltersButton.addEventListener("click", () => {
  clearFilterInputs();
  resetReportView();
  renderBom();
});
elements.defaultDemand.addEventListener("input", () => {
  resetReportView();
  refreshComputedQuantities();
});
elements.bufferPct.addEventListener("input", () => {
  resetReportView();
  refreshComputedQuantities();
});
[
  elements.country,
  elements.currency,
  elements.supplierFilter,
  elements.authorizedOnly,
  elements.matchLimit,
].forEach((input) => {
  input.addEventListener("input", resetReportView);
  input.addEventListener("change", resetReportView);
});
elements.clearButton.addEventListener("click", () => {
  state.lines = [];
  clearFilterInputs();
  resetReportView();
  elements.dropZone.classList.remove("uploaded");
  elements.fileLabel.textContent = "Upload BOM";
  renderBom();
  setStatus("");
});
elements.quoteButton.addEventListener("click", generateReport);
elements.exportButton.addEventListener("click", () => {
  if (!state.report) return;
  const blob = new Blob([quoteCsv(state.report)], { type: "text/csv" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = `supplier-report-${new Date().toISOString().slice(0, 10)}.csv`;
  anchor.click();
  URL.revokeObjectURL(url);
});

resetReportView();
renderBom();
checkHealth();
