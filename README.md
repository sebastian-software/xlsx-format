# xlsx-format

[![npm version](https://img.shields.io/npm/v/xlsx-format)](https://www.npmjs.com/package/xlsx-format)
[![CI](https://github.com/sebastian-software/xlsx-format/actions/workflows/ci.yml/badge.svg)](https://github.com/sebastian-software/xlsx-format/actions/workflows/ci.yml)
[![license](https://img.shields.io/npm/l/xlsx-format)](LICENSE)
[![node](https://img.shields.io/node/v/xlsx-format)](https://nodejs.org/)
[![bun](https://img.shields.io/badge/Bun-tested-f9f1e1?logo=bun)](https://bun.sh/)
[![TypeScript](https://img.shields.io/badge/TypeScript-strict-blue)](https://www.typescriptlang.org/)

The XLSX library your bundler will thank you for. Zero dependencies. Fully async. Works in Node.js and the browser.

```bash
npm install xlsx-format
```

```typescript
import { readFile, writeFile, sheetToJson, jsonToSheet, createWorkbook } from "xlsx-format";

// Read an Excel file into JSON
const workbook = await readFile("report.xlsx");
const rows = sheetToJson(workbook.Sheets[workbook.SheetNames[0]]);

// Write JSON back to Excel
const sheet = jsonToSheet([
	{ Name: "Alice", Revenue: 48000 },
	{ Name: "Bob", Revenue: 52000 },
]);
await writeFile(createWorkbook(sheet, "Q4 Sales"), "output.xlsx");
```

## Why xlsx-format?

Most projects just need XLSX -- but the popular libraries ship with support for dozens of legacy formats, pull in 7-9 runtime dependencies, and lock you into synchronous APIs that block the event loop.

xlsx-format does one thing well: read and write modern Excel files. The result is a library you can actually tree-shake, `await`, and ship to the browser without a separate bundle.

|                     | **xlsx-format**                | **SheetJS (xlsx)**      | **ExcelJS** |
| ------------------- | ------------------------------ | ----------------------- | ----------- |
| **Written in**      | TypeScript (strict)            | JavaScript (with .d.ts) | TypeScript  |
| **Async**           | Yes (streaming ZIP)            | No                      | Partial     |
| **Module format**   | ESM + CJS                      | CJS only                | CJS only    |
| **Tree-shakeable**  | Yes                            | No                      | Partial     |
| **Runtime deps**    | 0                              | 7                       | 9           |
| **Browser support** | Yes (`read` / `write`)         | Yes (separate bundle)   | No          |
| **Formats**         | XLSX / XLSM / CSV / TSV / HTML | 30+ formats             | XLSX / CSV  |
| **API style**       | Named exports, async           | Namespace object        | Class-based |
| **License**         | Apache 2.0                     | Apache 2.0              | MIT         |

## What it handles

- **Cell data** -- strings, numbers, booleans, dates, formulas, comments, hyperlinks
- **Number formatting** -- full SSF engine with the same format codes Excel uses (`#,##0.00`, `yyyy-mm-dd`, custom patterns)
- **Sheet structure** -- multiple sheets, merge regions, column widths, row heights, frozen panes, auto-filters
- **Metadata** -- defined names, document properties, sheet visibility
- **Format conversion** -- JSON, arrays, CSV, TSV, HTML tables (read and write)

## Runs everywhere

**Node.js >= 22** -- full support including `readFile` / `writeFile` for filesystem access.

**Browsers** -- `read()` and `write()` work in any modern browser with `Uint8Array` or `ArrayBuffer`. No Node.js APIs needed. Only `readFile()` / `writeFile()` require Node.

```typescript
// Browser: read from a File input
const buffer = await file.arrayBuffer();
const workbook = await read(buffer);

// Browser: trigger a download
const data = await write(workbook, { type: "array" });
const blob = new Blob([data], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
const url = URL.createObjectURL(blob);
```

## API

### Reading

```typescript
// From a Uint8Array or ArrayBuffer
const workbook = await read(buffer);

// From a file (Node.js) -- format detected from extension
const workbook = await readFile("spreadsheet.xlsx");
const workbook = await readFile("data.csv");
const workbook = await readFile("report.html");

// From a plain string (CSV or HTML auto-detected)
const workbook = await read(csvString, { type: "string" });
```

### Writing

```typescript
// To a Uint8Array (XLSX)
const bytes = await write(workbook);

// To CSV / TSV / HTML string
const csv = await write(workbook, { bookType: "csv", type: "string" });
const tsv = await write(workbook, { bookType: "tsv", type: "string" });
const html = await write(workbook, { bookType: "html", type: "string" });

// To a file (Node.js) -- format detected from extension
await writeFile(workbook, "output.xlsx");
await writeFile(workbook, "output.csv");
await writeFile(workbook, "output.html");
```

### Converting data

```typescript
// Sheet -> JSON objects (first row = headers)
const rows = sheetToJson(sheet);
// [{ Name: "Alice", Age: 30 }, { Name: "Bob", Age: 25 }]

// Sheet -> array of arrays
const arrays = sheetToJson(sheet, { header: 1 });
// [["Name", "Age"], ["Alice", 30], ["Bob", 25]]

// Sheet -> CSV / HTML
const csv = sheetToCsv(sheet);
const html = sheetToHtml(sheet);

// JSON / arrays / CSV / HTML -> Sheet
const sheet = jsonToSheet([{ Name: "Alice", Age: 30 }]);
const sheet = arrayToSheet([
	["Name", "Age"],
	["Alice", 30],
]);
const sheet = csvToSheet("Name,Age\nAlice,30");
const sheet = htmlToSheet("<table><tr><td>Name</td></tr></table>");
```

### Workbook helpers

```typescript
const wb = createWorkbook(firstSheet, "Sheet1");
appendSheet(wb, secondSheet, "Sheet2");
setSheetVisibility(wb, 1, "hidden");
```

### Cell utilities

```typescript
setCellNumberFormat(sheet, "B2", "#,##0.00");
setCellHyperlink(sheet, "A1", "https://example.com");
addCellComment(sheet, "C3", "Check this value", "Alice");
setArrayFormula(sheet, "D1:D10", "=A1:A10*B1:B10");
```

### Cell addresses

```typescript
decodeCell("B3"); // { r: 2, c: 1 }
encodeCell({ r: 2, c: 1 }); // "B3"
decodeRange("A1:C5"); // { s: { r: 0, c: 0 }, e: { r: 4, c: 2 } }
encodeRange(range); // "A1:C5"
```

## Switching from SheetJS

The API is intentionally close to SheetJS. Three things change:

1. `read()` and `write()` are `async` (ZIP uses streaming)
2. Named imports replace the namespace: `import { read } from "xlsx-format"`
3. Utility names are camelCase: `sheetToJson` instead of `XLSX.utils.sheet_to_json`

```diff
- import XLSX from "xlsx";
+ import { read, write, sheetToJson, sheetToCsv } from "xlsx-format";

- const wb = XLSX.read(buffer);
+ const wb = await read(buffer);

- const rows = XLSX.utils.sheet_to_json(ws);
+ const rows = sheetToJson(ws);

- const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
+ const buf = await write(wb, { type: "buffer" });

- const csv = XLSX.utils.sheet_to_csv(ws);
+ const csv = sheetToCsv(ws);
```

Cell objects keep the same shape: `{ t: "n", v: 42, w: "42" }` works exactly as before.

## Acknowledgments

Based on the work of [SheetJS](https://github.com/SheetJS/sheetjs), originally created by SheetJS LLC. Thank you to the SheetJS team and its contributors for building the foundation this library stands on.

## License

Apache 2.0 -- see [LICENSE](LICENSE) for details.
