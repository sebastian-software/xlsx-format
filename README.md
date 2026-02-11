# xlsx-format

[![npm version](https://img.shields.io/npm/v/xlsx-format)](https://www.npmjs.com/package/xlsx-format)
[![CI](https://github.com/sebastian-software/xlsx-format/actions/workflows/ci.yml/badge.svg)](https://github.com/sebastian-software/xlsx-format/actions/workflows/ci.yml)
[![codecov](https://codecov.io/gh/sebastian-software/xlsx-format/graph/badge.svg)](https://codecov.io/gh/sebastian-software/xlsx-format)
[![license](https://img.shields.io/npm/l/xlsx-format)](LICENSE)
[![node](https://img.shields.io/node/v/xlsx-format)](https://nodejs.org/)
[![bun](https://img.shields.io/badge/Bun-tested-f9f1e1?logo=bun)](https://bun.sh/)
[![TypeScript](https://img.shields.io/badge/TypeScript-strict-blue)](https://www.typescriptlang.org/)

The XLSX library your bundler will thank you for. Zero dependencies. Fully async. Works in Node.js and the browser.

**[Documentation](https://sebastian-software.github.io/xlsx-format/)** | **[API Reference](https://sebastian-software.github.io/xlsx-format/api-reference)**

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

|                     | **xlsx-format**                | **SheetJS (xlsx)**      | **ExcelJS**  |
| ------------------- | ------------------------------ | ----------------------- | ------------ |
| **Written in**      | TypeScript (strict)            | JavaScript (with .d.ts) | TypeScript   |
| **Async**           | Yes (streaming ZIP)            | No                      | Partial      |
| **Module format**   | ESM + CJS                      | CJS only                | CJS only     |
| **Tree-shakeable**  | Yes                            | No                      | Partial      |
| **Runtime deps**    | 0                              | 7                       | 9            |
| **Browser support** | Yes (`read` / `write`)         | Yes (separate bundle)   | No           |
| **Formats**         | XLSX / XLSM / CSV / TSV / HTML | 30+ formats             | XLSX / CSV   |
| **API style**       | Named exports, async           | Namespace object        | Class-based  |
| **Test coverage**   | 91% ([Codecov][codecov])       | Not measured            | Not measured |
| **License**         | Apache 2.0                     | Apache 2.0              | MIT          |

[codecov]: https://codecov.io/gh/sebastian-software/xlsx-format

For a detailed feature matrix (cell data, formulas, styles, comments, hyperlinks, and more), see [Why xlsx-format?](https://sebastian-software.github.io/xlsx-format/guide/why-xlsx-format) in the docs.

## Runs everywhere

**Node.js >= 22** -- full support including `readFile` / `writeFile` for filesystem access.

**Browsers** -- `read()` and `write()` work in any modern browser with `Uint8Array` or `ArrayBuffer`. No Node.js APIs needed. Only `readFile()` / `writeFile()` require Node.

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
```

Cell objects keep the same shape: `{ t: "n", v: 42, w: "42" }` works exactly as before. For a full function mapping table, see the [Migration Guide](https://sebastian-software.github.io/xlsx-format/guide/migration).

## Acknowledgments

Based on the work of [SheetJS](https://github.com/SheetJS/sheetjs), originally created by SheetJS LLC. Thank you to the SheetJS team and its contributors for building the foundation this library stands on.

## License

Apache 2.0 -- see [LICENSE](LICENSE) for details.

Copyright (C) 2012-present SheetJS LLC (original work)\
Copyright (C) 2025-present [Sebastian Software GmbH](https://sebastian-software.de) (modifications)
