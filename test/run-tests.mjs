// Local tests for the Office Scripts. Run with: node test/run-tests.mjs  (Node 22.13+ / 23.2+)
//
// Each script is type-stripped and run in a sandbox against a small in-memory
// mock of the ExcelScript API (only the calls these scripts make). This checks
// the scripts' logic. It does not prove they compile in the Office Scripts editor.

import { readFileSync } from "node:fs";
import { stripTypeScriptTypes } from "node:module";
import { createHash } from "node:crypto";
import vm from "node:vm";
import assert from "node:assert/strict";

const root = new URL("..", import.meta.url);

function load(path) {
  const src = stripTypeScriptTypes(readFileSync(new URL(path, root), "utf8"));
  const ctx = { ExcelScript: { ClearApplyTo: { contents: "Contents", all: "All" } } };
  vm.createContext(ctx);
  vm.runInContext(src, ctx);
  return ctx;
}

// ---- Minimal ExcelScript mock -------------------------------------------------

class MockRange {
  constructor(sheet, r, c, rows, cols) {
    Object.assign(this, { sheet, r, c, rows, cols });
  }
  getValues() {
    const out = [];
    for (let i = 0; i < this.rows; i++) {
      const row = [];
      for (let j = 0; j < this.cols; j++) row.push(this.sheet.cells.get(`${this.r + i},${this.c + j}`) ?? "");
      out.push(row);
    }
    return out;
  }
  setValues(values) {
    assert.equal(values.length, this.rows, "setValues row count must match range");
    values.forEach((row, i) => {
      assert.equal(row.length, this.cols, "setValues column count must match range");
      row.forEach((v, j) => {
        const key = `${this.r + i},${this.c + j}`;
        if (v === "") this.sheet.cells.delete(key);
        else this.sheet.cells.set(key, v);
      });
    });
  }
  setNumberFormat(fmt) {
    this.sheet.formats.push({ r: this.r, c: this.c, rows: this.rows, cols: this.cols, fmt });
  }
  clear() {
    for (let i = 0; i < this.rows; i++) for (let j = 0; j < this.cols; j++) this.sheet.cells.delete(`${this.r + i},${this.c + j}`);
  }
  getBoundingRect(other) {
    const r = Math.min(this.r, other.r);
    const c = Math.min(this.c, other.c);
    const r2 = Math.max(this.r + this.rows, other.r + other.rows);
    const c2 = Math.max(this.c + this.cols, other.c + other.cols);
    return new MockRange(this.sheet, r, c, r2 - r, c2 - c);
  }
  getRowIndex() { return this.r; }
  getRowCount() { return this.rows; }
}

class MockSheet {
  constructor(name) {
    this.name = name;
    this.cells = new Map();
    this.formats = [];
  }
  getUsedRange() {
    if (this.cells.size === 0) return undefined;
    let r1 = Infinity, c1 = Infinity, r2 = -1, c2 = -1;
    for (const key of this.cells.keys()) {
      const [r, c] = key.split(",").map(Number);
      r1 = Math.min(r1, r); c1 = Math.min(c1, c); r2 = Math.max(r2, r); c2 = Math.max(c2, c);
    }
    return new MockRange(this, r1, c1, r2 - r1 + 1, c2 - c1 + 1);
  }
  getRange(addr) {
    assert.equal(addr, "A1");
    return new MockRange(this, 0, 0, 1, 1);
  }
  getRangeByIndexes(r, c, rows, cols) {
    assert.ok(rows > 0 && cols > 0, "range must be non-empty");
    return new MockRange(this, r, c, rows, cols);
  }
}

class MockWorkbook {
  constructor() { this.sheets = []; }
  getWorksheet(name) { return this.sheets.find((s) => s.name.toLowerCase() === name.toLowerCase()); }
  addWorksheet(name) {
    assert.ok(name.length <= 31, `sheet name too long: ${name}`);
    assert.ok(!this.getWorksheet(name), `duplicate sheet: ${name}`);
    const s = new MockSheet(name);
    this.sheets.push(s);
    return s;
  }
}

// ---- Helpers ------------------------------------------------------------------

const sha1 = (s) => createHash("sha1").update(s, "utf8").digest("hex");

function makeTable(tableId, columns, rows) {
  const headerProps = [
    { Name: "rowCount", Value: rows.length },
    { Name: "columnCount", Value: columns.length },
  ];
  columns.forEach(([name, type], i) => {
    headerProps.push({ Name: `0-${i + 1}-name`, Value: name });
    headerProps.push({ Name: `0-${i + 1}-type`, Value: type });
  });
  return JSON.stringify({
    Children: [
      { Children: [], Name: `${tableId}_headers`, Properties: headerProps },
      { Children: [], Name: tableId, Properties: rows.map((r, i) => ({ Name: String(i + 1), Value: r })) },
    ],
    Name: "",
    Properties: [],
  });
}

let passed = 0;
function test(name, fn) {
  fn();
  passed++;
  console.log(`ok - ${name}`);
}

const getX = load("excel/GetTable.ts");
const updX = load("excel/UpdateTable.ts");
const getC = load("csv/CsvGetTable.ts");
const updC = load("csv/CsvUpdateTable.ts");

const COLS = [["Player", "text"], ["Score", "integer"]];
const ROWS = ["Row 1|Alice|100", "Row 2|B\\|ob, \"the\" builder|85", "Row 3|Zoë 🎮|007"];

// ---- SHA-1 --------------------------------------------------------------------

test("sha1 matches node crypto (ascii, unicode, block boundaries)", () => {
  const inputs = ["", "abc", "a".repeat(55), "a".repeat(56), "a".repeat(64), "a".repeat(1000), "Zoë 🎮 |\\|", makeTable("t", COLS, ROWS)];
  for (const s of inputs) {
    assert.equal(updX.sha1Hex(s), sha1(s));
    assert.equal(updC.sha1Hex(s), sha1(s));
  }
});

// ---- Excel version ------------------------------------------------------------

test("excel: POST then GET round-trips the table", () => {
  const wb = new MockWorkbook();
  const params = makeTable("Scores", COLS, ROWS);
  const post = updX.main(wb, params, sha1(params));
  assert.equal(post.statusCode, 200);
  assert.equal(JSON.parse(post.body).Status, "Success");

  const sheet = wb.getWorksheet("Scores");
  assert.deepEqual(sheet.getRange("A1").getBoundingRect(sheet.getUsedRange()).getValues(), [
    ["", "text", "integer"],
    ["", "Player", "Score"],
    ["Row 1", "Alice", "100"],
    ["Row 2", 'B|ob, "the" builder', "85"],
    ["Row 3", "Zoë 🎮", "007"],
  ]);
  assert.ok(sheet.formats.every((f) => f.fmt === "@"));

  const get = getX.main(wb, "Scores");
  assert.equal(get.statusCode, 200);
  const table = JSON.parse(get.body);
  assert.equal(table.Children[0].Name, "Scores_headers");
  assert.deepEqual(table.Children[0].Properties.slice(0, 2), [
    { Name: "rowCount", Value: 3 },
    { Name: "columnCount", Value: 2 },
  ]);
  assert.deepEqual(table.Children[1].Properties.map((p) => p.Value), ROWS);
});

test("excel: shorter POST clears old rows", () => {
  const wb = new MockWorkbook();
  const p1 = makeTable("Scores", COLS, ROWS);
  updX.main(wb, p1, sha1(p1));
  const p2 = makeTable("Scores", COLS, ["Row 1|Solo|1"]);
  updX.main(wb, p2, sha1(p2));
  const table = JSON.parse(getX.main(wb, "Scores").body);
  assert.deepEqual(table.Children[1].Properties.map((p) => p.Value), ["Row 1|Solo|1"]);
});

test("excel: bad signature is rejected and nothing is written", () => {
  const wb = new MockWorkbook();
  const params = makeTable("Scores", COLS, ROWS);
  const post = updX.main(wb, params, "deadbeef");
  assert.equal(JSON.parse(post.body).Status, "Failure");
  assert.equal(wb.sheets.length, 0);
});

test("excel: long table IDs go through the _tables index", () => {
  const wb = new MockWorkbook();
  const id = "id_0123456789abcdef0123456789abcdef_long";
  const params = makeTable(id, COLS, ROWS);
  assert.equal(JSON.parse(updX.main(wb, params, sha1(params)).body).Status, "Success");
  assert.ok(wb.getWorksheet("_tables"));
  assert.equal(JSON.parse(updX.main(wb, params, sha1(params)).body).Status, "Success"); // reuses sheet
  assert.equal(wb.sheets.length, 2);
  const table = JSON.parse(getX.main(wb, id).body);
  assert.equal(table.Children[1].Name, id);
  assert.deepEqual(table.Children[1].Properties.map((p) => p.Value), ROWS);
});

test("excel: GET for a missing table returns 404", () => {
  assert.equal(getX.main(new MockWorkbook(), "Nope").statusCode, 404);
});

// ---- CSV version --------------------------------------------------------------

test("csv: POST then GET round-trips the table", () => {
  const params = makeTable("Scores", COLS, ROWS);
  const post = updC.main(null, params, sha1(params));
  assert.equal(post.statusCode, 200);
  assert.equal(post.fileName, "Scores.csv");
  assert.equal(
    post.csv,
    '\uFEFF,text,integer\r\n,Player,Score\r\nRow 1,Alice,100\r\nRow 2,"B|ob, ""the"" builder",85\r\nRow 3,Zoë 🎮,007\r\n',
  );
  const get = getC.main(null, "Scores", post.csv);
  assert.equal(get.statusCode, 200);
  assert.deepEqual(JSON.parse(get.body).Children[1].Properties.map((p) => p.Value), ROWS);
});

test("csv: reads Excel-style CSV (BOM, LF, quoted newline, blank trailing rows)", () => {
  const csv = '﻿,text,text\n,Name,Note\nr1,"multi\nline",x\nr2,short\n,,\n';
  const table = JSON.parse(getC.main(null, "T", csv).body);
  assert.deepEqual(table.Children[0].Properties.slice(0, 2), [
    { Name: "rowCount", Value: 2 },
    { Name: "columnCount", Value: 2 },
  ]);
  assert.deepEqual(table.Children[1].Properties.map((p) => p.Value), ["r1|multi\nline|x", "r2|short|"]);
});

test("csv: decodes a {$content-type,$content} wrapper", () => {
  const csv = ",text\n,Name\nr1,Zoë 🎮\n";
  const wrapped = JSON.stringify({ "$content-type": "text/csv", "$content": Buffer.from(csv, "utf8").toString("base64") });
  const table = JSON.parse(getC.main(null, "T", wrapped).body);
  assert.deepEqual(table.Children[1].Properties.map((p) => p.Value), ["r1|Zoë 🎮"]);
});

test("csv: table IDs that aren't valid file names are rejected", () => {
  const params = makeTable("a/b", COLS, ROWS);
  const post = updC.main(null, params, sha1(params));
  assert.equal(JSON.parse(post.body).Status, "Failure");
  assert.equal(post.csv, "");
});

test("csv: empty file returns 404", () => {
  assert.equal(getC.main(null, "T", "").statusCode, 404);
});

console.log(`\n${passed} tests passed`);
