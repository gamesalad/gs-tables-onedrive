/**
 * GameSalad table API backed by an Excel workbook in OneDrive: the GET side.
 *
 * Port of doGet()/getTable() from the Google Apps Script version. Power Automate
 * calls this with "Run script" and passes the `tableId` query parameter it received.
 * The script returns { statusCode, body }; the flow's Response action sends `body`
 * (a GameSalad table in JSON) back to the game.
 *
 * Sheet layout (same as the Google Sheets version):
 *   Row 1, from column B: column types
 *   Row 2, from column B: column names
 *   Column A, from row 3: row labels
 *   B3 onward:            cell data
 *
 * The sheet is named after the table ID. Excel limits sheet names to 31 characters,
 * so tables with longer IDs are mapped to a sheet through the `_tables` sheet,
 * which UpdateTable creates when it needs it.
 */

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

function main(workbook: ExcelScript.Workbook, tableId: string): ScriptResult {
  if (!tableId) {
    return failure(400, "Missing tableId");
  }

  const sheet = findSheet(workbook, tableId);
  if (!sheet) {
    return failure(404, `No sheet found for table '${tableId}'`);
  }

  const used = sheet.getUsedRange(true);
  if (!used) {
    return failure(404, `Sheet for table '${tableId}' is empty`);
  }

  // Start the range at A1 even when A1 is blank (it usually is), so
  // row and column indexes match the layout above.
  const values = sheet.getRange("A1").getBoundingRect(used).getValues();

  return { statusCode: 200, body: JSON.stringify(buildTable(tableId, values)) };
}

function buildTable(tableId: string, values: (string | number | boolean)[][]): GSNode {
  const colTypes = values.length > 0 ? values[0] : [];
  const colHeaders = values.length > 1 ? values[1] : [];

  const headerProps: GSProperty[] = [
    { Name: "rowCount", Value: Math.max(values.length - 2, 0) }, // skip type and header rows
    { Name: "columnCount", Value: Math.max(colTypes.length - 1, 0) }, // skip row label column
  ];
  for (let i = 1; i < colHeaders.length; i++) {
    headerProps.push({ Name: `0-${i}-name`, Value: colHeaders[i] });
    headerProps.push({ Name: `0-${i}-type`, Value: colTypes[i] });
  }

  // GameSalad wants each row as one pipe-delimited string, row label first.
  const rowProps: GSProperty[] = [];
  for (let i = 2; i < values.length; i++) {
    const rowString = values[i].map((cell) => String(cell).replace(/\|/g, "\\|")).join("|");
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

function failure(statusCode: number, message: string): ScriptResult {
  return { statusCode, body: JSON.stringify({ Status: "Failure", Error: message }) };
}
