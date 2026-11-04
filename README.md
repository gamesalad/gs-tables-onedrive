# GameSalad tables on OneDrive (Power Automate + Office Scripts)

This is a port of the [Google Sheets / Apps Script version](https://gist.github.com/tant42/add5a37bd3143974362c981ff3d3fb4b). It lets a GameSalad game load and save tables through its Network behaviors, with the data kept in OneDrive.

There are two variants. Pick one:

| | **A. Excel workbook** (`excel/`) | **B. CSV files** (`csv/`) |
|---|---|---|
| Where table data lives | One `.xlsx` workbook, one sheet per table | One `<tableId>.csv` file per table, in a OneDrive folder |
| Office Scripts | `GetTable.ts`, `UpdateTable.ts` | `CsvGetTable.ts`, `CsvUpdateTable.ts` |
| Excel file needed? | Yes, it holds the data | **Yes, but only as an empty "script host".** See below |
| Closest to the original | ✔ (sheet = tab, same layout) | Same layout, one file per table |

> **Why the CSV variant still needs an Excel file:** Power Automate can only run an Office Script through the Excel Online (Business) **Run script** action, and that action has to point at an `.xlsx` workbook. It can't run a script against a CSV. In variant B the flow reads and writes the CSV with the **OneDrive for Business** connector, then passes the text to and from the script. The script ignores the workbook it runs against. Make one empty workbook (e.g. `GameSalad Script Host.xlsx`) and never put data in it.

## How it works

```
GameSalad ──GET  ?tableId=Scores──▶  Flow "GS Table – Get"   ──Run script──▶ GetTable / CsvGetTable    ──▶ JSON table
GameSalad ──POST params=…&sig=…──▶  Flow "GS Table – Update" ──Run script──▶ UpdateTable / CsvUpdateTable ──▶ {"Status": …}
```

Apps Script published `doGet` and `doPost` together as a single web app. Here each one is its own flow, with its own URL.

## Requirements

- **A Microsoft 365 work or school account.** Office Scripts and the Excel Online (Business) connector don't work with personal Microsoft accounts or personal OneDrive.
- **Office Scripts enabled** for your tenant. Check that Excel on the web has an **Automate** tab. If it doesn't, ask your M365 admin.
- **Power Automate Premium** (or pay-as-you-go). The **When an HTTP request is received** trigger is a premium connector.
- Usage limits on **Run script** apply, including a daily cap on calls per user. Check Microsoft's current numbers in [Platform limits with Office Scripts](https://learn.microsoft.com/office/dev/scripts/testing/platform-limits) before you ship a game that calls this heavily.

## Table layout (both variants)

This is the same layout the Apps Script version used:

|   | A | B | C |
|---|---|---|---|
| **1** | *(blank)* | `text` | `integer` | ← column types |
| **2** | *(blank)* | `Player` | `Score` | ← column names |
| **3** | `Row 1` | `Alice` | `100` | ← row label, then cells |
| **4** | `Row 2` | `Bob` | `85` | |

In variant A this is a sheet named after the table ID. In variant B it's `<tableId>.csv`, written like this:

```csv
,text,integer
,Player,Score
Row 1,Alice,100
Row 2,Bob,85
```

The type values above are placeholders. Use whatever GameSalad sends. The easiest way to seed a table is to POST it from the game once.

---

## Variant A: Excel workbook

### A1. Create the workbook

1. In OneDrive, create a workbook, e.g. `GameSalad Tables.xlsx`.
2. Optional: add a sheet per table using the layout above. The update flow creates sheets that don't exist yet.

### A2. Add the scripts

1. Open the workbook in **Excel on the web** → **Automate** tab → **New Script**.
2. Replace the editor contents with [`excel/GetTable.ts`](excel/GetTable.ts). Rename the script to **GetTable** and save it.
3. Do the same with [`excel/UpdateTable.ts`](excel/UpdateTable.ts) and save it as **UpdateTable**.

The scripts are saved to your OneDrive under `Documents/Office Scripts/`. The flow must belong to the same account.

### A3. Flow "GS Table – Get"

At [make.powerautomate.com](https://make.powerautomate.com) → **Create** → **Instant cloud flow** → skip → add these steps:

1. **Trigger: When an HTTP request is received** (Request connector)
   - *Who can trigger the flow?* **Anyone**
   - *Method:* **GET** (under advanced parameters)
2. **Action: Run script** (Excel Online (Business))
   - *Location:* OneDrive for Business · *Document Library:* OneDrive · *File:* `GameSalad Tables.xlsx`
   - *Script:* **GetTable**
   - *tableId:* expression `triggerOutputs()?['queries']?['tableId']`
3. **Action: Response** (Request connector)
   - *Status Code:* expression `body('Run_script')?['result']?['statusCode']`
   - *Headers:* `Content-Type` = `application/json`
   - *Body:* expression `json(body('Run_script')?['result']?['body'])`

Save the flow. The trigger then shows its **HTTP URL**. Copy it.

> The expressions refer to steps by name (`Run_script` = the step called "Run script"). If you rename a step, update the expressions that use it.

### A4. Flow "GS Table – Update"

1. **Trigger: When an HTTP request is received**
   - *Who can trigger the flow?* **Anyone** · *Method:* **POST**
   - Trigger **⋯ → Settings → Concurrency control: On, Degree of parallelism: 1**. This replaces Apps Script's `LockService`: only one write runs at a time, and the rest wait in a queue.
2. **Action: Run script**: same file, *Script:* **UpdateTable**
   - *params:* expression `triggerFormDataValue('params')`
   - *sig:* expression `triggerFormDataValue('sig')`
3. **Action: Response**: same three fields as A3.

Save the flow and copy its URL.

### A5. Options in `UpdateTable.ts`

| Constant | Default | Meaning |
|---|---|---|
| `REQUIRE_SIGNATURE` | `true` | Reject posts where `sig` ≠ SHA-1(`params`). This catches tampering in transit, but anyone can compute a SHA-1, so it is not authentication. |
| `CLEAR_SHEET_BEFORE_WRITE` | `true` | Clear the sheet's contents before writing (formatting is kept), so fewer rows don't leave stale rows behind. The original left this commented out. |
| `WRITE_AS_TEXT` | `true` | Format written cells as Text so Excel doesn't turn `007` into `7`, `1/2` into a date, or `=…` into a formula. |

**Long table IDs:** Excel limits sheet names to 31 characters and doesn't allow `: \ / ? * [ ]`. If a table ID breaks those rules, `UpdateTable` creates a sheet with a shortened name and records the mapping in a `_tables` sheet. `GetTable` reads that mapping. Don't delete or rename `_tables` or the mapped sheets.

---

## Variant B: CSV files

### B1. Create the folder and the host workbook

1. In OneDrive, create a folder: **`GameSalad Tables`** (the steps below assume `/GameSalad Tables`).
2. Optional: put table CSVs in it, named **`<tableId>.csv`**, using the layout above. The update flow creates missing files.
3. Create an **empty** workbook anywhere in OneDrive, e.g. `GameSalad Script Host.xlsx`. It's only there so the scripts have something to run against (see the note at the top).

### B2. Add the scripts

Open the host workbook in Excel on the web → **Automate** → **New Script**. Save [`csv/CsvGetTable.ts`](csv/CsvGetTable.ts) as **CsvGetTable** and [`csv/CsvUpdateTable.ts`](csv/CsvUpdateTable.ts) as **CsvUpdateTable**.

### B3. Flow "GS Table CSV – Get"

1. **Trigger: When an HTTP request is received**: *Anyone*, *Method* **GET**
2. **Action: Get file content using path** (OneDrive for Business)
   - *File Path:* expression `concat('/GameSalad Tables/', triggerOutputs()?['queries']?['tableId'], '.csv')`
   - *Infer Content Type:* **Yes**
3. **Action: Run script** (Excel Online (Business))
   - *File:* `GameSalad Script Host.xlsx` · *Script:* **CsvGetTable**
   - *tableId:* expression `triggerOutputs()?['queries']?['tableId']`
   - *csv:* expression `string(body('Get_file_content_using_path'))`
4. **Action: Response**
   - *Status Code:* `body('Run_script')?['result']?['statusCode']`
   - *Headers:* `Content-Type` = `application/json`
   - *Body:* `json(body('Run_script')?['result']?['body'])`
5. **Parallel branch for a missing file.** On step 2, use **+ → Add a parallel branch** and add a second **Response**:
   - **Settings → Run after:** *Get file content using path* **has failed** (untick *is successful*)
   - *Status Code:* `404` · *Headers:* `Content-Type` = `application/json`
   - *Body:* `{"Status":"Failure","Error":"Table not found"}`

> The connector sometimes returns file content as `{"$content-type": …, "$content": "<base64>"}` instead of plain text. `string(...)` turns either form into a string, and `CsvGetTable` detects and decodes the wrapped form.

### B4. Flow "GS Table CSV – Update"

1. **Trigger: When an HTTP request is received**: *Anyone*, *Method* **POST**, **Concurrency control: On, degree 1** (as in A4).
2. **Action: Run script**: host workbook, *Script:* **CsvUpdateTable**
   - *params:* `triggerFormDataValue('params')` · *sig:* `triggerFormDataValue('sig')`
3. **Condition "Has file to write"**: `body('Run_script')?['result']?['fileName']` *is not equal to* (leave blank).
   - **If yes:**
     1. **Get file metadata using path**: *File Path:* `concat('/GameSalad Tables/', body('Run_script')?['result']?['fileName'])`
     2. **Condition "File exists"**: **Settings → Run after:** *Get file metadata using path* **is successful** *and* **has failed**.
        Expression: `equals(actions('Get_file_metadata_using_path')?['status'], 'Succeeded')` *is equal to* `true`
        - **If yes → Update file** (OneDrive for Business): *File:* `body('Get_file_metadata_using_path')?['Id']` · *File Content:* `body('Run_script')?['result']?['csv']`
        - **If no → Create file**: *Folder Path:* `/GameSalad Tables` · *File Name:* `body('Run_script')?['result']?['fileName']` · *File Content:* `body('Run_script')?['result']?['csv']`
   - **If no:** leave empty. The script rejected the request, and its reason is in `body`.
4. **Action: Response** (after the condition): same three fields as B3 step 4.
5. **Parallel branch for write errors.** On step 3, add a parallel **Response** with **Run after:** *Has file to write* **has failed**. Use Status `500` and Body `{"Status":"Failure","Error":"Could not write CSV"}`.

The update flow checks for the file and then either updates or creates it. It doesn't rely on **Create file** to overwrite, because the connector docs don't say what Create file does when the file already exists.

### B5. Options and CSV caveats

- `CsvUpdateTable.ts` has `REQUIRE_SIGNATURE` (as in variant A) and `WRITE_BOM` (default `true`). The BOM marks the file as UTF-8, so Excel desktop shows accents and emoji correctly when someone opens the CSV by hand.
- Each POST **replaces the whole file**.
- The table ID becomes the file name, so IDs containing `" * : < > ? / \ |`, or starting or ending with a space or `.`, are rejected with a Failure. If GameSalad's IDs can contain those characters, use variant A.
- If you edit a CSV in Excel desktop and save it, Excel may reformat values (`007` → `7`, long numbers → `1E+15`). Edit in a text editor, or import as Text.

---

## Pointing GameSalad at it

- In the network behavior that **loads** a table, use the **Get** flow's URL.
- In the network behavior that **sends** a table, use the **Update** flow's URL.

The flow URLs contain a signature (`sig=…`) that works as the access key. Anyone who has the URL can call the flow. Don't commit the URLs to public repos.

### ⚠️ Check before relying on it

These points come from reading the Apps Script code, not from testing a real game:

1. **Query string on GET.** The Apps Script version read `tableId` from `e.parameter`, so these flows expect `?tableId=…` in the query string. A flow URL already has a query string (`?api-version=…&sig=…`). If GameSalad blindly appends `?tableId=…`, the URL breaks. If it appends `&tableId=…`, or merges parameters properly, it works. Check the first GET in the flow's **run history**.
2. **POST encoding.** These flows expect `params` and `sig` as `application/x-www-form-urlencoded` body fields, which is what `e.parameter` reads in Apps Script for a POST. If GameSalad sends them some other way, `triggerFormDataValue` fails and the run history will show it.
3. **Latency.** A Run script call usually takes a few seconds. That's slower than Apps Script, so allow for it in your network timeout and loading UI.

## Testing with curl

```bash
GET_URL='<Get flow URL>'
POST_URL='<Update flow URL>'

PARAMS='{"Children":[{"Children":[],"Name":"Scores_headers","Properties":[{"Name":"rowCount","Value":2},{"Name":"columnCount","Value":2},{"Name":"0-1-name","Value":"Player"},{"Name":"0-1-type","Value":"text"},{"Name":"0-2-name","Value":"Score"},{"Name":"0-2-type","Value":"integer"}]},{"Children":[],"Name":"Scores","Properties":[{"Name":"1","Value":"Row 1|Alice|100"},{"Name":"2","Value":"Row 2|Bob|85"}]}],"Name":"","Properties":[]}'
SIG=$(printf %s "$PARAMS" | shasum -a 1 | cut -d' ' -f1)

# Save (should print {"Status":"Success"})
curl -s -X POST "$POST_URL" --data-urlencode "params=$PARAMS" --data-urlencode "sig=$SIG"

# Load (should print the table back)
curl -s "$GET_URL&tableId=Scores"
```

## Differences from the Apps Script version

- **Locking:** Apps Script used `LockService`. Here, trigger concurrency is set to 1 instead.
- **Setup:** no `setup()` step or stored document ID. The flow's Run script step points at the file directly.
- **Pipe escaping:** the original's `replace("|", "\\|")` only escaped the **first** pipe in each cell. These scripts escape every `|` as `\|` and split rows on unescaped pipes.
- **Header parsing:** column names and types are matched by the column index in `0-<n>-name` / `0-<n>-type`, not by property order.
- **Writes:** cells are written as one block instead of one at a time, and the sheet is cleared first by default (`CLEAR_SHEET_BEFORE_WRITE`).
- **Missing table on GET:** returns HTTP 404 with `{"Status":"Failure"}`. The original threw an error instead.
- **No polyfills:** the `Array.prototype.map`/`filter` polyfills are gone, since Office Scripts runs modern JavaScript.

## Local tests

```bash
node test/run-tests.mjs   # Node 22.13+ (uses node:module stripTypeScriptTypes)
```

The tests strip the TypeScript types and run each script's `main` against an in-memory mock of the few ExcelScript calls the scripts make. They check SHA-1 against Node's `crypto` and round-trip tables through both variants, covering pipes, quotes, commas, newlines, emoji, a BOM, the base64 wrapper, long table IDs, and bad signatures. **They don't cover the real Office Scripts runtime or Power Automate.** Do a real end-to-end run with the curl commands above.
