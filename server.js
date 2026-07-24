import http from "node:http";
import { readFile } from "node:fs/promises";
import { existsSync, readFileSync } from "node:fs";
import path from "node:path";
import { inflateRawSync } from "node:zlib";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const publicDir = path.join(__dirname, "public");

loadEnvFile();

const port = Number(process.env.PORT || 4173);
const host = process.env.HOST || "127.0.0.1";

const NEXAR_TOKEN_URL =
  process.env.NEXAR_TOKEN_URL || "https://identity.nexar.com/connect/token";
const NEXAR_GRAPHQL_URL =
  process.env.NEXAR_GRAPHQL_URL || "https://api.nexar.com/graphql";

const mimeTypes = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".csv": "text/csv; charset=utf-8",
  ".svg": "image/svg+xml",
};

let tokenCache = {
  accessToken: "",
  expiresAt: 0,
};

function loadEnvFile() {
  const envPath = path.join(__dirname, ".env");
  if (!existsSync(envPath)) return;

  const contents = readFileSync(envPath, "utf8");
  for (const line of contents.split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) continue;

    const separatorIndex = trimmed.indexOf("=");
    if (separatorIndex === -1) continue;

    const key = trimmed.slice(0, separatorIndex).trim();
    let value = trimmed.slice(separatorIndex + 1).trim();
    if (
      (value.startsWith('"') && value.endsWith('"')) ||
      (value.startsWith("'") && value.endsWith("'"))
    ) {
      value = value.slice(1, -1);
    }

    if (key && process.env[key] === undefined) {
      process.env[key] = value;
    }
  }
}

function hasNexarCredentials() {
  return Boolean(process.env.NEXAR_CLIENT_ID && process.env.NEXAR_CLIENT_SECRET);
}

function sendJson(res, statusCode, payload) {
  res.writeHead(statusCode, {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store",
  });
  res.end(JSON.stringify(payload));
}

function sendText(res, statusCode, text) {
  res.writeHead(statusCode, { "Content-Type": "text/plain; charset=utf-8" });
  res.end(text);
}

function normalizeHeader(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "");
}

function normalizeMpn(value) {
  return String(value || "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, " ");
}

function numberFrom(value, fallback = 0) {
  const parsed = Number(String(value ?? "").replace(/,/g, "").trim());
  return Number.isFinite(parsed) ? parsed : fallback;
}

function optionalNumber(value) {
  if (value === null || value === undefined || String(value).trim() === "") {
    return null;
  }
  const parsed = Number(String(value).replace(/,/g, "").trim());
  return Number.isFinite(parsed) ? parsed : null;
}

function gqlString(value) {
  return JSON.stringify(String(value ?? ""));
}

function uniqueStrings(values) {
  const seen = new Set();
  const output = [];
  for (const value of values) {
    const clean = String(value || "").trim();
    const key = clean.toLowerCase();
    if (clean && !seen.has(key)) {
      seen.add(key);
      output.push(clean);
    }
  }
  return output;
}

function searchKindForPartNumber(value) {
  const text = String(value || "").trim();
  if (/-ND$/i.test(text)) return "sku";
  return "mpn";
}

function buildPartQuery(value, settings) {
  const searchKind = searchKindForPartNumber(value);
  const limit = searchKind === "sku" ? 1 : Number(settings.matchLimit || 1);
  return `{ ${searchKind}: ${gqlString(value)}, limit: ${limit} }`;
}

async function readJsonBody(req) {
  let body = "";
  for await (const chunk of req) {
    body += chunk;
    if (body.length > 2_000_000) {
      throw new Error("Request body is too large.");
    }
  }
  return body ? JSON.parse(body) : {};
}

async function readBinaryBody(req, maxBytes = 25_000_000) {
  const chunks = [];
  let totalLength = 0;

  for await (const chunk of req) {
    totalLength += chunk.length;
    if (totalLength > maxBytes) {
      throw new Error("Uploaded BOM is too large.");
    }
    chunks.push(chunk);
  }

  return Buffer.concat(chunks, totalLength);
}

async function getNexarToken() {
  if (tokenCache.accessToken && tokenCache.expiresAt > Date.now() + 60_000) {
    return tokenCache.accessToken;
  }

  const form = new URLSearchParams();
  form.set("grant_type", "client_credentials");
  form.set("client_id", process.env.NEXAR_CLIENT_ID);
  form.set("client_secret", process.env.NEXAR_CLIENT_SECRET);
  form.set("scope", "supply.domain");

  const response = await fetch(NEXAR_TOKEN_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body: form,
  });

  if (!response.ok) {
    const detail = await response.text();
    throw new Error(`Nexar token request failed (${response.status}): ${detail}`);
  }

  const data = await response.json();
  if (!data.access_token) {
    throw new Error("Nexar token response did not include an access token.");
  }

  const expiresInSeconds = Number(data.expires_in || 3600);
  tokenCache = {
    accessToken: data.access_token,
    expiresAt: Date.now() + Math.max(expiresInSeconds - 120, 60) * 1000,
  };

  return data.access_token;
}

function buildNexarQuery(mpns, settings) {
  const queryList = mpns
    .map((mpn) => buildPartQuery(mpn, settings))
    .join(", ");

  const args = [`queries: [${queryList}]`];
  if (settings.country) args.push(`country: ${gqlString(settings.country)}`);
  if (settings.currency) args.push(`currency: ${gqlString(settings.currency)}`);
  if (settings.allowedSuppliers?.length) {
    const distributors = settings.allowedSuppliers.map(gqlString).join(", ");
    args.push(`options: { filters: { distributor_id: [${distributors}] } }`);
  }

  const sellerArgs = settings.authorizedOnly ? "(authorizedOnly: true)" : "";

  return `
    query BomMatch {
      supMultiMatch(${args.join(", ")}) {
        hits
        reference
        parts {
          id
          mpn
          name
          totalAvail
          octopartUrl
          estimatedFactoryLeadDays
          manufacturer {
            name
          }
          similarParts {
            id
            mpn
            manufacturer {
              name
            }
          }
          sellers${sellerArgs} {
            company {
              name
            }
            offers {
              id
              sku
              clickUrl
              inventoryLevel
              factoryLeadDays
              moq
              packaging
              prices {
                quantity
                price
                currency
                convertedPrice
                convertedCurrency
              }
            }
          }
        }
      }
    }
  `;
}

function readUInt16(buffer, offset) {
  return buffer.readUInt16LE(offset);
}

function readUInt32(buffer, offset) {
  return buffer.readUInt32LE(offset);
}

function findZipEndOfCentralDirectory(buffer) {
  const minimumLength = 22;
  for (let offset = buffer.length - minimumLength; offset >= 0; offset -= 1) {
    if (readUInt32(buffer, offset) === 0x06054b50) {
      return offset;
    }
  }
  throw new Error("Uploaded file is not a readable Excel workbook.");
}

function readZipEntries(buffer) {
  const entries = new Map();
  const eocd = findZipEndOfCentralDirectory(buffer);
  const entryCount = readUInt16(buffer, eocd + 10);
  let offset = readUInt32(buffer, eocd + 16);

  for (let entryIndex = 0; entryIndex < entryCount; entryIndex += 1) {
    if (readUInt32(buffer, offset) !== 0x02014b50) {
      throw new Error("Excel workbook central directory is invalid.");
    }

    const compressionMethod = readUInt16(buffer, offset + 10);
    const compressedSize = readUInt32(buffer, offset + 20);
    const fileNameLength = readUInt16(buffer, offset + 28);
    const extraLength = readUInt16(buffer, offset + 30);
    const commentLength = readUInt16(buffer, offset + 32);
    const localHeaderOffset = readUInt32(buffer, offset + 42);
    const fileName = buffer
      .subarray(offset + 46, offset + 46 + fileNameLength)
      .toString("utf8");

    if (readUInt32(buffer, localHeaderOffset) !== 0x04034b50) {
      throw new Error(`Excel workbook local header is invalid for ${fileName}.`);
    }

    const localNameLength = readUInt16(buffer, localHeaderOffset + 26);
    const localExtraLength = readUInt16(buffer, localHeaderOffset + 28);
    const dataStart = localHeaderOffset + 30 + localNameLength + localExtraLength;
    const compressed = buffer.subarray(dataStart, dataStart + compressedSize);

    let data;
    if (compressionMethod === 0) {
      data = Buffer.from(compressed);
    } else if (compressionMethod === 8) {
      data = inflateRawSync(compressed);
    } else {
      data = null;
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

function parseWorkbookRows(buffer) {
  const entries = readZipEntries(buffer);
  const sharedStrings = parseSharedStrings(
    entries.get("xl/sharedStrings.xml")?.toString("utf8") || "",
  );

  const worksheets = [...entries.entries()]
    .filter(([name]) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(name))
    .map(([name, contents]) => ({
      name,
      rows: parseWorksheetRows(contents.toString("utf8"), sharedStrings),
    }))
    .sort((a, b) => b.rows.length - a.rows.length);

  const bestSheet = worksheets.find((sheet) => sheet.rows.length > 1) || worksheets[0];
  if (!bestSheet?.rows.length) {
    throw new Error("No BOM rows were found in the workbook.");
  }

  return bestSheet;
}

async function fetchNexarParts(mpns, settings) {
  const token = await getNexarToken();
  const matches = [];
  const chunkSize = 20;

  for (let index = 0; index < mpns.length; index += chunkSize) {
    const chunk = mpns.slice(index, index + chunkSize);
    const chunkKinds = chunk.map(searchKindForPartNumber);
    const response = await fetch(NEXAR_GRAPHQL_URL, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ query: buildNexarQuery(chunk, settings) }),
    });

    if (!response.ok) {
      const detail = await response.text();
      throw new Error(`Nexar GraphQL request failed (${response.status}): ${detail}`);
    }

    const data = await response.json();
    if (data.errors?.length) {
      throw new Error(data.errors.map((error) => error.message).join("; "));
    }

    const batch = Array.isArray(data.data?.supMultiMatch)
      ? data.data.supMultiMatch
      : [];

    batch.forEach((match, offset) => {
      matches.push({
        requestedMpn: match.reference || chunk[offset],
        searchKind: chunkKinds[offset],
        hits: match.hits || 0,
        parts: match.parts || [],
      });
    });
  }

  return matches;
}

function normalizeInventory(value) {
  if (value === null || value === undefined || value === "") return null;
  const parsed = Number(String(value).replace(/[^0-9.-]+/g, ""));
  return Number.isFinite(parsed) ? parsed : null;
}

function choosePriceBreak(prices, requiredQty) {
  const cleanPrices = (prices || [])
    .map((price) => {
      const convertedPrice = optionalNumber(price.convertedPrice);
      const nativePrice = optionalNumber(price.price);
      return {
        quantity: Math.max(1, numberFrom(price.quantity, 1)),
        price: convertedPrice ?? nativePrice,
        currency:
          convertedPrice !== null
            ? price.convertedCurrency || price.currency || ""
            : price.currency || "",
      };
    })
    .filter((price) => price.price !== null && price.price >= 0)
    .sort((a, b) => a.quantity - b.quantity);

  if (!cleanPrices.length) return null;

  const applicable = cleanPrices.filter((price) => price.quantity <= requiredQty);
  return applicable.length ? applicable.at(-1) : cleanPrices[0];
}

function evaluateBom(lines, matches, settings, source) {
  const partsByMpn = new Map();
  for (const match of matches) {
    const key = normalizeMpn(match.requestedMpn);
    partsByMpn.set(key, {
      hits: match.hits || 0,
      searchKind: match.searchKind || "mpn",
      parts: match.parts || [],
    });
  }

  const supplierFilter = settings.allowedSuppliers || [];
  const supplierFilterSet = new Set(supplierFilter.map((name) => name.toLowerCase()));

  const reportLines = lines.map((line, index) => {
    const bomQty = Math.max(0, numberFrom(line.bomQty, 0));
    const demand = Math.max(0, numberFrom(settings.defaultDemand, 0));
    const bufferPct = Math.max(0, numberFrom(settings.bufferPct, 10));
    const rawRequiredQty = (bomQty * demand * (100 + bufferPct)) / 100;
    const requiredQty = Math.ceil(rawRequiredQty - 1e-9);

    const candidateMpns = uniqueStrings([
      line.mpn,
      ...(Array.isArray(line.alternatives) ? line.alternatives : []),
    ]);

    const offers = [];
    const similarParts = [];
    let matchedPartsCount = 0;

    for (const candidateMpn of candidateMpns) {
      const match = partsByMpn.get(normalizeMpn(candidateMpn));
      if (!match) continue;

      for (const part of match.parts) {
        if (
          match.searchKind === "mpn" &&
          normalizeMpn(part.mpn) !== normalizeMpn(candidateMpn)
        ) {
          continue;
        }

        matchedPartsCount += 1;

        for (const similarPart of part.similarParts || []) {
          similarParts.push({
            mpn: similarPart.mpn,
            manufacturer: similarPart.manufacturer?.name || "",
          });
        }

        const marketAvailability = normalizeInventory(part.totalAvail);

        for (const seller of part.sellers || []) {
          const supplier = seller.company?.name || "Unknown supplier";
          if (
            supplierFilterSet.size &&
            !supplierFilterSet.has(supplier.toLowerCase())
          ) {
            continue;
          }

          for (const offer of seller.offers || []) {
            const minimumOrderQty = Math.max(1, numberFrom(offer.moq, 1));
            const priceBreak = choosePriceBreak(
              offer.prices,
              Math.max(requiredQty, minimumOrderQty, 1),
            );
            if (!priceBreak) continue;

            const quotedQty = Math.max(requiredQty, minimumOrderQty, priceBreak.quantity);
            const inventoryLevel = normalizeInventory(offer.inventoryLevel);
            const stockStatus =
              inventoryLevel === null
                ? "unknown"
                : inventoryLevel >= requiredQty
                  ? "sufficient"
                  : "short";
            const totalCost = priceBreak.price * quotedQty;

            offers.push({
              inputMpn: candidateMpn,
              quotedMpn: part.mpn || candidateMpn,
              manufacturer: part.manufacturer?.name || "",
              supplierSku: offer.sku || "",
              supplier,
              sourceRole:
                normalizeMpn(candidateMpn) === normalizeMpn(line.mpn)
                  ? "Primary"
                  : "Approved alternative",
              unitPrice: priceBreak.price,
              currency: priceBreak.currency || settings.currency || "USD",
              priceBreakQty: priceBreak.quantity,
              minimumOrderQty,
              packaging: offer.packaging || "",
              quotedQty,
              totalCost,
              inventoryLevel,
              marketAvailability,
              stockStatus,
              leadDays: offer.factoryLeadDays || part.estimatedFactoryLeadDays || null,
              url: offer.clickUrl || part.octopartUrl || "",
            });
          }
        }
      }
    }

    offers.sort((a, b) => {
      const stockRank = { sufficient: 0, unknown: 1, short: 2 };
      return (
        stockRank[a.stockStatus] - stockRank[b.stockStatus] ||
        a.totalCost - b.totalCost ||
        a.unitPrice - b.unitPrice ||
        a.supplier.localeCompare(b.supplier)
      );
    });

    const recommendation = offers[0] || null;
    let status = "Quoted";
    if (!requiredQty) status = "Check quantity";
    else if (!candidateMpns.length) status = "Missing MPN";
    else if (!matchedPartsCount) status = "No match";
    else if (!offers.length) status = "No priced offer";
    else if (recommendation.stockStatus === "short") status = "Stock short";
    else if (recommendation.stockStatus === "unknown") status = "Stock unknown";

    const runnerUp = offers.find((offer, offerIndex) => {
      if (offerIndex === 0) return false;
      return (
        offer.supplier !== recommendation?.supplier ||
        offer.quotedMpn !== recommendation?.quotedMpn
      );
    });

    return {
      id: line.id || `line-${index + 1}`,
      ziplinePn: line.ziplinePn || "",
      lineNumber: line.lineNumber || String(index + 1),
      description: line.description || "",
      primaryMpn: line.mpn || "",
      candidateMpns,
      partType: line.partType || "",
      procurementIntent: line.procurementIntent || "",
      bomQty,
      demand,
      bufferPct,
      requiredQty,
      status,
      recommendation,
      runnerUp,
      offerCount: offers.length,
      matchedPartsCount,
      similarParts: uniqueStrings(
        similarParts.map((part) =>
          part.manufacturer ? `${part.mpn} (${part.manufacturer})` : part.mpn,
        ),
      ).slice(0, 5),
    };
  });

  const quotedLines = reportLines.filter((line) => line.recommendation);
  const totalCost = quotedLines.reduce(
    (sum, line) => sum + Number(line.recommendation?.totalCost || 0),
    0,
  );
  const unresolvedLines = reportLines.filter((line) => !line.recommendation).length;
  const stockRisks = reportLines.filter((line) =>
    ["Stock short", "Stock unknown"].includes(line.status),
  ).length;

  return {
    generatedAt: new Date().toISOString(),
    source,
    settings,
    summary: {
      lineCount: reportLines.length,
      quotedLines: quotedLines.length,
      unresolvedLines,
      stockRisks,
      totalCost,
      currency: settings.currency || "USD",
    },
    lines: reportLines,
  };
}

function sanitizeSettings(rawSettings = {}) {
  return {
    defaultDemand: Math.max(0, numberFrom(rawSettings.defaultDemand, 1)),
    bufferPct: Math.max(0, numberFrom(rawSettings.bufferPct, 10)),
    country: String(rawSettings.country || "US").trim().toUpperCase().slice(0, 2),
    currency: String(rawSettings.currency || "USD").trim().toUpperCase().slice(0, 3),
    authorizedOnly: rawSettings.authorizedOnly !== false,
    matchLimit: Math.min(5, Math.max(1, numberFrom(rawSettings.matchLimit, 1))),
    allowedSuppliers: uniqueStrings(
      Array.isArray(rawSettings.allowedSuppliers)
        ? rawSettings.allowedSuppliers
        : String(rawSettings.allowedSuppliers || "")
            .split(",")
            .map((value) => value.trim()),
    ),
  };
}

function sanitizeLines(rawLines) {
  if (!Array.isArray(rawLines)) return [];
  return rawLines.slice(0, 250).map((line, index) => ({
    id: String(line.id || `line-${index + 1}`),
    ziplinePn: String(line.ziplinePn || "").trim(),
    lineNumber: String(line.lineNumber || index + 1),
    description: String(line.description || "").trim(),
    mpn: String(line.mpn || "").trim(),
    alternatives: uniqueStrings(
      Array.isArray(line.alternatives)
        ? line.alternatives
        : String(line.alternatives || "")
            .split(/[;|\n]/)
            .map((value) => value.trim()),
    ),
    bomQty: numberFrom(line.bomQty, 0),
    partType: String(line.partType || "").trim(),
    procurementIntent: String(line.procurementIntent || "").trim(),
    demand: line.demand === "" || line.demand === undefined ? "" : numberFrom(line.demand, 0),
  }));
}

async function handleBomImport(req, res) {
  try {
    const buffer = await readBinaryBody(req);
    const workbook = parseWorkbookRows(buffer);

    return sendJson(res, 200, {
      sheet: workbook.name,
      rows: workbook.rows,
      rowCount: workbook.rows.length,
    });
  } catch (error) {
    return sendJson(res, 400, {
      error: error.message || "Unable to import workbook.",
    });
  }
}

async function handleQuote(req, res) {
  try {
    const body = await readJsonBody(req);
    const lines = sanitizeLines(body.lines);
    const settings = sanitizeSettings(body.settings);

    if (!lines.length) {
      return sendJson(res, 400, { error: "Upload or enter at least one BOM line." });
    }

    const mpns = uniqueStrings(
      lines.flatMap((line) => [line.mpn, ...(line.alternatives || [])]),
    );

    if (!hasNexarCredentials()) {
      return sendJson(res, 503, {
        error:
          "Nexar credentials are not configured. Add NEXAR_CLIENT_ID and NEXAR_CLIENT_SECRET before generating supplier recommendations.",
      });
    }

    const matches = mpns.length ? await fetchNexarParts(mpns, settings) : [];

    return sendJson(res, 200, evaluateBom(lines, matches, settings, "nexar"));
  } catch (error) {
    return sendJson(res, 500, {
      error: error.message || "Unable to generate the quote report.",
    });
  }
}

async function serveStatic(req, res) {
  const url = new URL(req.url, `http://${req.headers.host || "localhost"}`);
  const rawPathname = decodeURIComponent(url.pathname);
  const pathname = rawPathname === "/" ? "/index.html" : rawPathname;
  const filePath = path.normalize(path.join(publicDir, pathname));

  if (!filePath.startsWith(publicDir) || !existsSync(filePath)) {
    return sendText(res, 404, "Not found");
  }

  const extension = path.extname(filePath);
  const content = await readFile(filePath);
  res.writeHead(200, {
    "Content-Type": mimeTypes[extension] || "application/octet-stream",
    "Cache-Control": "no-store",
  });
  res.end(content);
}

const server = http.createServer(async (req, res) => {
  if (req.method === "GET" && req.url?.startsWith("/api/health")) {
    return sendJson(res, 200, {
      ok: true,
      apiSource: hasNexarCredentials() ? "nexar" : "unconfigured",
      nexarConfigured: hasNexarCredentials(),
    });
  }

  if (req.method === "POST" && req.url?.startsWith("/api/quote")) {
    return handleQuote(req, res);
  }

  if (req.method === "POST" && req.url?.startsWith("/api/import-bom")) {
    return handleBomImport(req, res);
  }

  if (req.method === "GET") {
    try {
      return await serveStatic(req, res);
    } catch (error) {
      return sendText(res, 500, error.message || "Unable to serve file");
    }
  }

  return sendText(res, 405, "Method not allowed");
});

server.listen(port, host, () => {
  console.log(`BOM Supplier Optimizer running at http://${host}:${port}`);
  console.log(
    hasNexarCredentials()
      ? "Using Nexar/Octopart API credentials from the environment."
      : "No Nexar credentials found. Supplier recommendations are disabled.",
  );
});
