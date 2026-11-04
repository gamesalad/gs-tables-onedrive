/**
 * GameSalad table API backed by CSV files in OneDrive: the POST side.
 *
 * The flow passes the `params` (table JSON) and `sig` (SHA-1 of params) form
 * fields GameSalad posted. This script checks the signature, converts the table
 * to CSV, and returns:
 *   { statusCode, body, fileName, csv }
 * The flow then writes `csv` to `<folder>/<fileName>` in OneDrive and sends
 * `body` back to the game.
 *
 * It never touches `workbook`; see CsvGetTable.ts for why a host .xlsx is needed.
 * Each POST replaces the whole CSV file.
 */

// Reject requests whose `sig` isn't the SHA-1 of `params`. This only catches
// tampering in transit: anyone can compute a SHA-1, so it is not authentication.
const REQUIRE_SIGNATURE = true;

// Start the file with a UTF-8 byte order mark. Excel desktop needs it to show
// non-ASCII text correctly when someone opens the CSV by hand. CsvGetTable
// ignores it.
const WRITE_BOM = true;

interface GSProperty {
  Name: string;
  Value: string | number | boolean;
}

interface GSNode {
  Children: GSNode[];
  Name: string;
  Properties: GSProperty[];
}

interface CsvScriptResult {
  statusCode: number;
  body: string;
  fileName: string;
  csv: string;
}

function main(workbook: ExcelScript.Workbook, params: string, sig: string): CsvScriptResult {
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
    // OneDrive would silently replace these characters with "_" on create, and
    // the GET flow would then not find the file by its table ID.
    if (!tableId || /["*:<>?\/\\|\x00-\x1f]/.test(tableId) || /^[ .]|[ .]$/.test(tableId)) {
      return failure(`Table ID '${tableId}' can't be used as a OneDrive file name`);
    }

    const header = parseHeader(table.Children[0].Properties || []);
    const rows = (table.Children[1].Properties || []).map((p) => splitRow(String(p.Value)));
    const grid: string[][] = [[""].concat(header.types), [""].concat(header.names)].concat(rows);

    return {
      statusCode: 200,
      body: JSON.stringify({ Status: "Success" }),
      fileName: `${tableId}.csv`,
      csv: (WRITE_BOM ? "\uFEFF" : "") + toCsv(grid),
    };
  } catch (e) {
    return failure(String(e));
  }
}

function parseHeader(props: GSProperty[]): { names: string[]; types: string[] } {
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

// RFC 4180 output: CRLF line endings, fields quoted only when they need it.
// Rows are padded to the same width so Excel and other tools see a clean grid.
function toCsv(grid: string[][]): string {
  let width = 0;
  grid.forEach((r) => (width = Math.max(width, r.length)));
  return grid
    .map((r) => {
      const cells: string[] = [];
      for (let i = 0; i < width; i++) {
        const v = i < r.length ? r[i] : "";
        cells.push(/[",\r\n]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v);
      }
      return cells.join(",");
    })
    .join("\r\n") + "\r\n";
}

function failure(message: string): CsvScriptResult {
  // HTTP 200 with Status "Failure", matching the original script.
  return { statusCode: 200, body: JSON.stringify({ Status: "Failure", Error: message }), fileName: "", csv: "" };
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
