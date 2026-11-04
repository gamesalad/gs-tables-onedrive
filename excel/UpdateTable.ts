/**
 * GameSalad table API backed by an Excel workbook in OneDrive: the POST side.
 *
 * Port of doPost()/updateTable() from the Google Apps Script version. Power Automate
 * calls this with "Run script" and passes the `params` (table JSON) and `sig`
 * (SHA-1 of params) form fields GameSalad posted. The script writes the table to
 * the sheet named after its table ID, creating the sheet if needed, and returns
 * { statusCode, body } for the flow's Response action.
 *
 * Sheet layout is documented in GetTable.ts.
 *
 * There is no LockService in Office Scripts. Turn on concurrency control
 * (degree of parallelism 1) on the flow's trigger instead; see the README.
 */

// Reject requests whose `sig` isn't the SHA-1 of `params`. This only catches
// tampering in transit: anyone can compute a SHA-1, so it is not authentication.
const REQUIRE_SIGNATURE = true;

// Clear the sheet before writing, so a table with fewer rows than last time
// doesn't leave old rows behind. The original script left this commented out.
const CLEAR_SHEET_BEFORE_WRITE = true;

// Store everything as text. Keeps Excel from turning "007" into 7 or "1/2"
// into a date, and stops values starting with "=" from becoming formulas.
// GameSalad gets the values back as strings either way.
const WRITE_AS_TEXT = true;

const INDEX_SHEET = "_tables";

interface GSProperty {
  Name: string;
  Value: string | number | boolean;
}

interface GSNode {
  Children: GSNode[];
  Name: string;
  Properties: GSProperty[];
}

interface ScriptResult {
  statusCode: number;
  body: string;
}

function main(workbook: ExcelScript.Workbook, params: string, sig: string): ScriptResult {
  if (!params) {
    return failure("Missing params");
  }
  if (REQUIRE_SIGNATURE && sha1Hex(params) !== (sig || "").trim().toLowerCase()) {
    return failure("Signature mismatch");
  }

  try {
    const table = JSON.parse(params) as GSNode;
    if (!table.Children || table.Children.length < 2) {
      return failure("Table JSON must have a headers child and a data child");
    }

    const tableId = table.Children[1].Name;
    const header = parseHeader(table.Children[0].Properties || []);
    const rows = (table.Children[1].Properties || []).map((p) => splitRow(String(p.Value)));

    const sheet = getOrCreateSheet(workbook, tableId);
    if (CLEAR_SHEET_BEFORE_WRITE) {
      const used = sheet.getUsedRange();
      if (used) {
        used.clear(ExcelScript.ClearApplyTo.contents);
      }
    }

    // Rows 1-2 from column B: column types, then column names.
    if (header.types.length > 0) {
      writeBlock(sheet, 0, 1, [header.types, header.names]);
    }

    // Row 3 onward from column A: row label, then cells.
    if (rows.length > 0) {
      writeBlock(sheet, 2, 0, rows);
    }

    return { statusCode: 200, body: JSON.stringify({ Status: "Success" }) };
  } catch (e) {
    return failure(String(e));
  }
}

function parseHeader(props: GSProperty[]): { names: string[]; types: string[] } {
  // Columns come in as "0-<col>-name" / "0-<col>-type"; use <col> rather than
  // relying on property order.
  const names: string[] = [];
  const types: string[] = [];
  let count = 0;
  props.forEach((p) => {
    const m = /^0-(\d+)-(name|type)$/.exec(p.Name);
    if (!m) {
      return;
    }
    const col = parseInt(m[1], 10) - 1;
    if (col < 0) {
      return;
    }
    (m[2] === "name" ? names : types)[col] = String(p.Value);
    count = Math.max(count, col + 1);
  });
  for (let i = 0; i < count; i++) {
    names[i] = names[i] === undefined ? "" : names[i];
    types[i] = types[i] === undefined ? "" : types[i];
  }
  return { names, types };
}

// Split a GameSalad row string on "|", treating "\|" as a literal pipe.
function splitRow(rowString: string): string[] {
  const cells: string[] = [];
  let current = "";
  for (let i = 0; i < rowString.length; i++) {
    const ch = rowString[i];
    if (ch === "\\" && rowString[i + 1] === "|") {
      current += "|";
      i++;
    } else if (ch === "|") {
      cells.push(current);
      current = "";
    } else {
      current += ch;
    }
  }
  cells.push(current);
  return cells;
}

// Write a 2D block in one call, padding short rows so the block is rectangular.
function writeBlock(sheet: ExcelScript.Worksheet, row: number, col: number, data: string[][]): void {
  let width = 0;
  data.forEach((r) => (width = Math.max(width, r.length)));
  const padded: string[][] = data.map((r) => {
    const out = r.slice();
    while (out.length < width) {
      out.push("");
    }
    return out;
  });
  const range = sheet.getRangeByIndexes(row, col, padded.length, width);
  if (WRITE_AS_TEXT) {
    range.setNumberFormat("@");
  }
  range.setValues(padded);
}

function getOrCreateSheet(workbook: ExcelScript.Workbook, tableId: string): ExcelScript.Worksheet {
  const existing = findSheet(workbook, tableId);
  if (existing) {
    return existing;
  }
  if (isValidSheetName(tableId)) {
    return workbook.addWorksheet(tableId);
  }

  // The ID can't be a sheet name (too long or has reserved characters), so make
  // up a name and record the mapping in the index sheet.
  const base = tableId.replace(/[:\\\/?*\[\]]/g, "_").replace(/^'+|'+$/g, "").slice(0, 27) || "table";
  let name = base;
  for (let n = 1; workbook.getWorksheet(name) || name.toLowerCase() === "history"; n++) {
    name = `${base}_${n}`;
  }
  const sheet = workbook.addWorksheet(name);

  let index = workbook.getWorksheet(INDEX_SHEET);
  if (!index) {
    index = workbook.addWorksheet(INDEX_SHEET);
    writeBlock(index, 0, 0, [["tableId", "sheet"]]);
  }
  const used = index.getUsedRange(true);
  const nextRow = used ? used.getRowIndex() + used.getRowCount() : 1;
  writeBlock(index, nextRow, 0, [[tableId, name]]);

  return sheet;
}

function findSheet(workbook: ExcelScript.Workbook, tableId: string): ExcelScript.Worksheet | undefined {
  const index = workbook.getWorksheet(INDEX_SHEET);
  if (index) {
    const used = index.getUsedRange(true);
    if (used) {
      const rows = used.getValues();
      for (let i = 0; i < rows.length; i++) {
        if (String(rows[i][0]) === tableId) {
          return workbook.getWorksheet(String(rows[i][1]));
        }
      }
    }
  }
  if (isValidSheetName(tableId)) {
    return workbook.getWorksheet(tableId);
  }
  return undefined;
}

function isValidSheetName(name: string): boolean {
  return (
    name.length > 0 &&
    name.length <= 31 &&
    !/[:\\\/?*\[\]]/.test(name) &&
    name[0] !== "'" &&
    name[name.length - 1] !== "'" &&
    name.toLowerCase() !== "history"
  );
}

function failure(message: string): ScriptResult {
  // HTTP 200 with Status "Failure", matching the original script.
  return { statusCode: 200, body: JSON.stringify({ Status: "Failure", Error: message }) };
}

// SHA-1 of the UTF-8 bytes of `message`, as lowercase hex. Matches
// Utilities.computeDigest(SHA_1, message) in the Apps Script version.
function sha1Hex(message: string): string {
  const bytes = utf8Bytes(message);
  const bitLenHi = Math.floor(bytes.length / 0x20000000);
  const bitLenLo = (bytes.length * 8) >>> 0;
  bytes.push(0x80);
  while (bytes.length % 64 !== 56) {
    bytes.push(0);
  }
  [bitLenHi, bitLenLo].forEach((w) => bytes.push((w >>> 24) & 255, (w >>> 16) & 255, (w >>> 8) & 255, w & 255));

  const rotl = (x: number, n: number): number => (x << n) | (x >>> (32 - n));
  const h: number[] = [0x67452301, 0xefcdab89, 0x98badcfe, 0x10325476, 0xc3d2e1f0];
  const w: number[] = [];

  for (let off = 0; off < bytes.length; off += 64) {
    for (let i = 0; i < 16; i++) {
      const j = off + i * 4;
      w[i] = (bytes[j] << 24) | (bytes[j + 1] << 16) | (bytes[j + 2] << 8) | bytes[j + 3];
    }
    for (let i = 16; i < 80; i++) {
      w[i] = rotl(w[i - 3] ^ w[i - 8] ^ w[i - 14] ^ w[i - 16], 1);
    }
    let a = h[0];
    let b = h[1];
    let c = h[2];
    let d = h[3];
    let e = h[4];
    for (let i = 0; i < 80; i++) {
      let f: number;
      let k: number;
      if (i < 20) {
        f = (b & c) | (~b & d);
        k = 0x5a827999;
      } else if (i < 40) {
        f = b ^ c ^ d;
        k = 0x6ed9eba1;
      } else if (i < 60) {
        f = (b & c) | (b & d) | (c & d);
        k = 0x8f1bbcdc;
      } else {
        f = b ^ c ^ d;
        k = 0xca62c1d6;
      }
      const t = (rotl(a, 5) + f + e + k + w[i]) | 0;
      e = d;
      d = c;
      c = rotl(b, 30);
      b = a;
      a = t;
    }
    h[0] = (h[0] + a) | 0;
    h[1] = (h[1] + b) | 0;
    h[2] = (h[2] + c) | 0;
    h[3] = (h[3] + d) | 0;
    h[4] = (h[4] + e) | 0;
  }
  return h.map((x) => ("00000000" + (x >>> 0).toString(16)).slice(-8)).join("");
}

function utf8Bytes(s: string): number[] {
  const out: number[] = [];
  for (let i = 0; i < s.length; i++) {
    let cp = s.charCodeAt(i);
    if (cp >= 0xd800 && cp <= 0xdbff && i + 1 < s.length) {
      const lo = s.charCodeAt(i + 1);
      if (lo >= 0xdc00 && lo <= 0xdfff) {
        cp = 0x10000 + ((cp - 0xd800) << 10) + (lo - 0xdc00);
        i++;
      }
    }
    if (cp < 0x80) {
      out.push(cp);
    } else if (cp < 0x800) {
      out.push(0xc0 | (cp >> 6), 0x80 | (cp & 63));
    } else if (cp < 0x10000) {
      out.push(0xe0 | (cp >> 12), 0x80 | ((cp >> 6) & 63), 0x80 | (cp & 63));
    } else {
      out.push(0xf0 | (cp >> 18), 0x80 | ((cp >> 12) & 63), 0x80 | ((cp >> 6) & 63), 0x80 | (cp & 63));
    }
  }
  return out;
}
