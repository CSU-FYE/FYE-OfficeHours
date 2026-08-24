/*
 * A small .xlsx reader. No dependencies.
 *
 * An .xlsx file is a zip of XML documents. Everything needed to read one is
 * already in the browser: DecompressionStream inflates the entries and
 * DOMParser reads the XML. This exposes just enough to turn a workbook into
 * plain rows of strings — no styles, no formulas, no formatting.
 *
 *   const sheets = await readWorkbook(arrayBuffer);
 *   sheets.shifts  // -> [["name","day",...], ["Dr. Harvey","Monday",...], ...]
 */

const EOCD_SIG = 0x06054b50;
const CD_SIG = 0x02014b50;

/** Locate the end-of-central-directory record, which sits at the end of the file. */
function findEocd(view) {
  // It is followed only by an optional comment, so scan back over the max comment size.
  const max = Math.min(view.byteLength, 0xffff + 22);
  for (let i = 22; i <= max; i++) {
    const at = view.byteLength - i;
    if (view.getUint32(at, true) === EOCD_SIG) return at;
  }
  throw new Error("Not a valid .xlsx file (no zip end-of-directory record).");
}

/** Map every entry in the zip to {offset, compression, size} by filename. */
function readCentralDirectory(buffer) {
  const view = new DataView(buffer);
  const eocd = findEocd(view);
  const count = view.getUint16(eocd + 10, true);
  let at = view.getUint32(eocd + 16, true);

  if (at === 0xffffffff) throw new Error("Zip64 .xlsx files are not supported.");

  const entries = new Map();
  const decoder = new TextDecoder();
  for (let i = 0; i < count; i++) {
    if (view.getUint32(at, true) !== CD_SIG) break;
    const compression = view.getUint16(at + 10, true);
    const compressedSize = view.getUint32(at + 20, true);
    const nameLen = view.getUint16(at + 28, true);
    const extraLen = view.getUint16(at + 30, true);
    const commentLen = view.getUint16(at + 32, true);
    const localOffset = view.getUint32(at + 42, true);
    const name = decoder.decode(new Uint8Array(buffer, at + 46, nameLen));
    entries.set(name, { compression, compressedSize, localOffset });
    at += 46 + nameLen + extraLen + commentLen;
  }
  return entries;
}

/** Inflate one entry to text. */
async function readEntry(buffer, entry) {
  const view = new DataView(buffer);
  // The local header repeats the name and extra fields with their own lengths,
  // and only they give the true start of the data.
  const nameLen = view.getUint16(entry.localOffset + 26, true);
  const extraLen = view.getUint16(entry.localOffset + 28, true);
  const start = entry.localOffset + 30 + nameLen + extraLen;
  const bytes = new Uint8Array(buffer, start, entry.compressedSize);

  if (entry.compression === 0) return new TextDecoder().decode(bytes);
  if (entry.compression !== 8) {
    throw new Error(`Unsupported zip compression method ${entry.compression}.`);
  }
  if (typeof DecompressionStream === "undefined") {
    throw new Error("This browser is too old to read .xlsx files (no DecompressionStream).");
  }
  const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  return await new Response(stream).text();
}

const parseXml = (text) => new DOMParser().parseFromString(text, "application/xml");

/** Concatenate every <t> under a node — shared strings may be split into runs. */
const textOf = (node) =>
  Array.from(node.getElementsByTagName("t"))
    .map((t) => t.textContent)
    .join("");

/** "BC12" -> 54 (zero-based column index). */
function columnIndex(ref) {
  let n = 0;
  for (const ch of ref) {
    const code = ch.charCodeAt(0);
    if (code < 65 || code > 90) break;
    n = n * 26 + (code - 64);
  }
  return n - 1;
}

function parseSheet(xml, sharedStrings) {
  const rows = [];
  for (const row of xml.getElementsByTagName("row")) {
    const cells = [];
    for (const c of row.getElementsByTagName("c")) {
      const at = columnIndex(c.getAttribute("r") || "");
      const type = c.getAttribute("t");
      let value = "";

      if (type === "inlineStr") {
        value = textOf(c);
      } else {
        const v = c.getElementsByTagName("v")[0];
        const raw = v ? v.textContent : "";
        if (raw === "") value = "";
        else if (type === "s") value = sharedStrings[Number(raw)] ?? "";
        else if (type === "e") value = "";
        else if (type === "b") value = raw === "1" ? "TRUE" : "FALSE";
        else if (type === "str") value = raw;
        else value = Number(raw); // numeric, including dates and times as serials
      }

      if (at >= 0) {
        while (cells.length < at) cells.push("");
        cells[at] = value;
      } else {
        cells.push(value);
      }
    }
    // Rows carry their own index; a skipped row must not shift the ones after it.
    const index = Number(row.getAttribute("r") || rows.length + 1) - 1;
    while (rows.length < index) rows.push([]);
    rows[index] = cells;
  }
  return rows;
}

/** Read a whole workbook into { sheetName: rows }. */
export async function readWorkbook(buffer) {
  const entries = readCentralDirectory(buffer);
  const get = (name) => {
    const entry = entries.get(name);
    if (!entry) throw new Error(`Malformed .xlsx: missing ${name}`);
    return readEntry(buffer, entry);
  };

  let sharedStrings = [];
  if (entries.has("xl/sharedStrings.xml")) {
    const xml = parseXml(await get("xl/sharedStrings.xml"));
    sharedStrings = Array.from(xml.getElementsByTagName("si")).map(textOf);
  }

  // Sheet names live in workbook.xml but point at files only via relationship ids.
  const relsXml = parseXml(await get("xl/_rels/workbook.xml.rels"));
  const targets = new Map();
  for (const rel of relsXml.getElementsByTagName("Relationship")) {
    let target = rel.getAttribute("Target") || "";
    if (target.startsWith("/")) target = target.slice(1);
    else if (!target.startsWith("xl/")) target = "xl/" + target.replace(/^\.\//, "");
    targets.set(rel.getAttribute("Id"), target);
  }

  const workbookXml = parseXml(await get("xl/workbook.xml"));
  const sheets = {};
  for (const sheet of workbookXml.getElementsByTagName("sheet")) {
    const name = sheet.getAttribute("name");
    const rid =
      sheet.getAttribute("r:id") ||
      sheet.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
    const path = targets.get(rid);
    if (!path || !entries.has(path)) continue;
    sheets[name] = parseSheet(parseXml(await readEntry(buffer, entries.get(path))), sharedStrings);
  }
  return sheets;
}

/**
 * Rows -> objects keyed by the header row, lowercased and trimmed.
 * Blank rows are dropped. Each object carries the source row number for error
 * messages that point back at the spreadsheet.
 */
export function toRecords(rows) {
  if (!rows || !rows.length) return [];
  const headerIndex = rows.findIndex((r) => r.some((c) => String(c ?? "").trim()));
  if (headerIndex < 0) return [];
  const headers = rows[headerIndex].map((h) => String(h ?? "").trim().toLowerCase());

  const records = [];
  for (let i = headerIndex + 1; i < rows.length; i++) {
    const row = rows[i] || [];
    if (!row.some((c) => String(c ?? "").trim())) continue;
    const record = { __row: i + 1 };
    headers.forEach((h, j) => {
      if (h) record[h] = row[j] ?? "";
    });
    records.push(record);
  }
  return records;
}
