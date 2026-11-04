/**
 * GameSalad table API backed by CSV files in OneDrive: the GET side.
 *
 * The flow reads `<tableId>.csv` with OneDrive's "Get file content using path"
 * and passes the text here with the table ID. This script turns the CSV into a
 * GameSalad table and returns { statusCode, body } for the Response action.
 *
 * It never touches `workbook`. Office Scripts can only run against an Excel
 * workbook, so the flow points "Run script" at an empty host .xlsx; the table
 * data lives only in the CSV files.
 *
 * CSV layout is the same grid as the sheet version:
 *   Row 1, from column 2: column types   (cell 1 is left blank)
 *   Row 2, from column 2: column names   (cell 1 is left blank)
 *   Row 3 onward:         row label, then cells
 */

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

interface FileContentWrapper {
  "$content"?: string;
}

function main(workbook: ExcelScript.Workbook, tableId: string, csv: string): ScriptResult {
  if (!tableId) {
    return failure(400, "Missing tableId");
  }

  const rows = parseCsv(unwrapFileContent(csv || ""));
  if (rows.length === 0) {
    return failure(404, `CSV for table '${tableId}' is empty`);
  }

  // Pad to a rectangle, like Apps Script's getDataRange().
  let width = 0;
  rows.forEach((r) => (width = Math.max(width, r.length)));
  rows.forEach((r) => {
    while (r.length < width) {
      r.push("");
    }
  });

  return { statusCode: 200, body: JSON.stringify(buildTable(tableId, rows)) };
}

function buildTable(tableId: string, values: string[][]): GSNode {
  const colTypes = values.length > 0 ? values[0] : [];
  const colHeaders = values.length > 1 ? values[1] : [];

  const headerProps: GSProperty[] = [
    { Name: "rowCount", Value: Math.max(values.length - 2, 0) },
    { Name: "columnCount", Value: Math.max(colTypes.length - 1, 0) },
  ];
  for (let i = 1; i < colHeaders.length; i++) {
    headerProps.push({ Name: `0-${i}-name`, Value: colHeaders[i] });
    headerProps.push({ Name: `0-${i}-type`, Value: colTypes[i] });
  }

  const rowProps: GSProperty[] = [];
  for (let i = 2; i < values.length; i++) {
    const rowString = values[i].map((cell) => cell.replace(/\|/g, "\\|")).join("|");
    rowProps.push({ Name: (i - 1).toString(), Value: rowString });
  }

  return {
    Children: [
      { Children: [], Name: `${tableId}_headers`, Properties: headerProps },
      { Children: [], Name: tableId, Properties: rowProps },
    ],
    Name: "",
    Properties: [],
  };
}

// RFC 4180 CSV: quoted fields, "" escapes, CRLF or LF line endings, optional BOM.
// Trailing all-blank rows are dropped.
function parseCsv(text: string): string[][] {
  if (text.charCodeAt(0) === 0xfeff) {
    text = text.slice(1);
  }
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let inQuotes = false;
  let fieldStarted = false;

  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQuotes) {
      if (c === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        field += c;
      }
    } else if (c === '"' && !fieldStarted) {
      inQuotes = true;
      fieldStarted = true;
    } else if (c === ",") {
      row.push(field);
      field = "";
      fieldStarted = false;
    } else if (c === "\r" || c === "\n") {
      if (c === "\r" && text[i + 1] === "\n") {
        i++;
      }
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
      fieldStarted = false;
    } else {
      field += c;
      fieldStarted = true;
    }
  }
  if (fieldStarted || row.length > 0) {
    row.push(field);
    rows.push(row);
  }

  while (rows.length > 0 && rows[rows.length - 1].every((cell) => cell === "")) {
    rows.pop();
  }
  return rows;
}

// If Power Automate hands over the file as {"$content-type": ..., "$content": <base64>}
// instead of plain text, decode it. Otherwise return the input unchanged.
function unwrapFileContent(input: string): string {
  const trimmed = input.trim();
  if (trimmed[0] !== "{" || trimmed.indexOf('"$content"') < 0) {
    return input;
  }
  try {
    const wrapper = JSON.parse(trimmed) as FileContentWrapper;
    if (typeof wrapper["$content"] === "string") {
      return utf8Decode(base64Decode(wrapper["$content"]));
    }
  } catch (e) {
    // Not JSON after all; treat it as CSV text.
  }
  return input;
}

function base64Decode(b64: string): number[] {
  const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
  const clean = b64.replace(/[^A-Za-z0-9+/]/g, "");
  const out: number[] = [];
  let buffer = 0;
  let bits = 0;
  for (let i = 0; i < clean.length; i++) {
    buffer = (buffer << 6) | alphabet.indexOf(clean[i]);
    bits += 6;
    if (bits >= 8) {
      bits -= 8;
      out.push((buffer >> bits) & 255);
    }
  }
  return out;
}

function utf8Decode(bytes: number[]): string {
  const parts: string[] = [];
  for (let i = 0; i < bytes.length; ) {
    const b = bytes[i];
    let cp: number;
    let extra: number;
    if (b < 0x80) {
      cp = b;
      extra = 0;
    } else if (b >= 0xf0) {
      cp = b & 7;
      extra = 3;
    } else if (b >= 0xe0) {
      cp = b & 15;
      extra = 2;
    } else {
      cp = b & 31;
      extra = 1;
    }
    for (let j = 1; j <= extra; j++) {
      cp = (cp << 6) | (bytes[i + j] & 63);
    }
    i += extra + 1;
    if (cp > 0xffff) {
      cp -= 0x10000;
      parts.push(String.fromCharCode(0xd800 + (cp >> 10), 0xdc00 + (cp & 1023)));
    } else {
      parts.push(String.fromCharCode(cp));
    }
  }
  return parts.join("");
}

function failure(statusCode: number, message: string): ScriptResult {
  return { statusCode, body: JSON.stringify({ Status: "Failure", Error: message }) };
}
