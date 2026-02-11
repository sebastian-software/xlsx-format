# xlsx-format

[![npm version](https://img.shields.io/npm/v/xlsx-format)](https://www.npmjs.com/package/xlsx-format)
[![CI](https://github.com/nickelow/xlsx-format/actions/workflows/ci.yml/badge.svg)](https://github.com/nickelow/xlsx-format/actions/workflows/ci.yml)
[![license](https://img.shields.io/npm/l/xlsx-format)](LICENSE)
[![bundle size](https://img.shields.io/bundlephobia/minzip/xlsx-format)](https://bundlephobia.com/package/xlsx-format)
[![TypeScript](https://img.shields.io/badge/TypeScript-strict-blue)](https://www.typescriptlang.org/)

Read and write XLSX spreadsheets in Node.js. No runtime dependencies. 184 KB unminified, 42 KB gzipped.

```bash
npm install xlsx-format
```

```typescript
import { read, write, sheetToJson, jsonToSheet, createWorkbook, appendSheet } from "xlsx-format";

// Read an Excel file into JSON
const workbook = await read(fs.readFileSync("report.xlsx"));
const rows = sheetToJson(workbook.Sheets[workbook.SheetNames[0]]);

// Write JSON back to Excel
const sheet = jsonToSheet([
	{ Name: "Alice", Revenue: 48000 },
	{ Name: "Bob", Revenue: 52000 },
]);
const wb = createWorkbook(sheet, "Q4 Sales");
fs.writeFileSync("output.xlsx", await write(wb, { type: "buffer" }));
```

## Why this exists

SheetJS (the `xlsx` npm package) supports every spreadsheet format ever made -- XLS, XLSB, ODS, CSV, DBF, and more. That coverage comes at a cost: 7 runtime dependencies, 7.5 MB unpacked, and source code that's difficult to read or contribute to.

Most projects only need XLSX. xlsx-format strips away everything else and rewrites the core in modern TypeScript. It also reads and writes CSV, TSV, and HTML through the same `read()`/`write()` API.

## How it compares

|                      | **xlsx-format**                | **SheetJS (xlsx)**        | **ExcelJS**                   |
| -------------------- | ------------------------------ | ------------------------- | ----------------------------- |
| **Formats**          | XLSX / XLSM / CSV / TSV / HTML | 30+ formats               | XLSX / CSV                    |
| **Bundle (ESM)**     | 184 KB                         | ~1 MB (full)              | ~1 MB                         |
| **Gzipped**          | 42 KB                          | ~330 KB                   | ~250 KB                       |
| **Runtime deps**     | 0                              | 7 (cfb, ssf, codepage...) | 9 (jszip, archiver, saxes...) |
| **TypeScript**       | Written in TS                  | JS with .d.ts             | Written in TS                 |
| **Tree-shakeable**   | Yes (ESM)                      | No                        | Partial                       |
| **Module format**    | ESM + CJS                      | CJS (+ browser bundle)    | CJS (+ browser bundle)        |
| **ZIP handling**     | Built-in (DecompressionStream) | cfb + custom              | jszip + archiver + unzipper   |
| **Node requirement** | >= 18                          | >= 0.8                    | >= 16                         |
| **API compatible**   | ~90% (read/write/utils)        | --                        | Different API                 |
| **License**          | Apache 2.0                     | Apache 2.0                | MIT                           |

## What it can do

**Read and write:** Cell values (strings, numbers, booleans, dates), formulas, number formats, multiple sheets, defined names, comments, hyperlinks, merge regions, column widths, row heights, sheet visibility, frozen panes, auto-filters, document properties.

**Convert to/from:** JSON objects, arrays of arrays, CSV, HTML tables.

**Number formatting:** Full SSF (SpreadSheet Format) engine -- the same format codes Excel uses (`#,##0.00`, `yyyy-mm-dd`, custom patterns).

## API

### Reading

```typescript
// From a Buffer or Uint8Array
const workbook = await read(buffer);

// From a file path (Node.js) — format detected from extension
const workbook = await readFile("spreadsheet.xlsx");
const workbook = await readFile("data.csv");
const workbook = await readFile("report.html");

// From a plain string (CSV or HTML auto-detected)
const workbook = await read(csvString, { type: "string" });
```

### Writing

```typescript
// To a Buffer (XLSX)
const buffer = await write(workbook, { type: "buffer" });

// To CSV / TSV / HTML string
const csv = await write(workbook, { bookType: "csv", type: "string" });
const tsv = await write(workbook, { bookType: "tsv", type: "string" });
const html = await write(workbook, { bookType: "html", type: "string" });

// Directly to a file (Node.js) — format detected from extension
await writeFile(workbook, "output.xlsx");
await writeFile(workbook, "output.csv");
await writeFile(workbook, "output.html");
```

### Sheet to data

```typescript
// Sheet -> array of objects (first row = headers)
const rows = sheetToJson(sheet);
// [{ Name: "Alice", Age: 30 }, { Name: "Bob", Age: 25 }]

// Sheet -> array of arrays (no headers)
const arrays = sheetToJson(sheet, { header: 1 });
// [["Name", "Age"], ["Alice", 30], ["Bob", 25]]

// Sheet -> CSV string
const csv = sheetToCsv(sheet);

// Sheet -> HTML table
const html = sheetToHtml(sheet);
```

### Data to sheet

```typescript
// Array of objects -> sheet
const sheet = jsonToSheet([{ Name: "Alice", Age: 30 }]);

// Array of arrays -> sheet
const sheet = arrayToSheet([
	["Name", "Age"],
	["Alice", 30],
]);

// CSV string -> sheet
const sheet = csvToSheet("Name,Age\nAlice,30");

// HTML table -> sheet
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

### Cell address encoding

```typescript
decodeCell("B3"); // { r: 2, c: 1 }
encodeCell({ r: 2, c: 1 }); // "B3"
decodeRange("A1:C5"); // { s: { r: 0, c: 0 }, e: { r: 4, c: 2 } }
encodeRange(range); // "A1:C5"
```

## Migration from SheetJS

The API mirrors SheetJS where it matters. If you're already using SheetJS for XLSX files, the main changes are:

1. `XLSX.read()` and `XLSX.write()` are now `async` (ZIP decompression uses streams)
2. Named imports instead of a namespace: `import { read, write } from "xlsx-format"`
3. Utility functions are top-level exports: `sheetToJson` instead of `XLSX.utils.sheet_to_json`

```diff
- import XLSX from "xlsx";
+ import { read, write, sheetToJson } from "xlsx-format";

- const wb = XLSX.read(buffer);
+ const wb = await read(buffer);

- const rows = XLSX.utils.sheet_to_json(ws);
+ const rows = sheetToJson(ws);

- const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
+ const buf = await write(wb, { type: "buffer" });

- const csv = XLSX.utils.sheet_to_csv(ws);
+ const csv = sheetToCsv(ws);
+ // or write the whole workbook as CSV:
+ const csv = await write(wb, { bookType: "csv", type: "string" });
```

The cell object shape is unchanged: `{ t: "n", v: 42, w: "42" }` works exactly the same way.

## License

Apache 2.0 -- see [LICENSE](LICENSE) for details.
