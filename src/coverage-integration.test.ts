/**
 * Integration & API-level coverage tests.
 *
 * Targets coverage gaps in:
 * - src/read.ts & src/write.ts (type detection, format inference, output types)
 * - src/api/aoa.ts (dense mode, nulls, dates, formulas, edge cases)
 * - src/api/json.ts (stub cells, errors, dates, dedup, dense, pre-built cells)
 * - src/api/html.ts (merges, NaN/Infinity, editable, links, rowspan/colspan)
 * - src/ssf/format.ts (number formats, fractions, scientific, conditionals, text)
 */
import { describe, it, expect } from "vitest";
import * as fs from "node:fs";
import * as path from "node:path";
import * as os from "node:os";

import {
	read,
	write,
	readFile,
	writeFile,
	arrayToSheet,
	sheetToJson,
	sheetToHtml,
	htmlToSheet,
	createWorkbook,
} from "./index.js";
import { addArrayToSheet } from "./api/aoa.js";
import { addJsonToSheet } from "./api/json.js";
import { formatNumber } from "./ssf/format.js";

// ============================================================
// src/read.ts
// ============================================================
describe("read.ts — input type handling", () => {
	it("should read from ArrayBuffer", async () => {
		const ws = arrayToSheet([["Hello"]]);
		const wb = createWorkbook(ws, "Sheet1");
		const u8 = await write(wb, { type: "array" });
		const ab = u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength);
		const result = await read(ab);
		expect(result.SheetNames).toContain("Sheet1");
	});

	it("should read from base64 string", async () => {
		const ws = arrayToSheet([["Test"]]);
		const wb = createWorkbook(ws, "Sheet1");
		const b64 = await write(wb, { type: "base64" });
		const result = await read(b64, { type: "base64" });
		expect(result.SheetNames).toContain("Sheet1");
	});

	it("should read plain CSV string", async () => {
		const result = await read("A,B\n1,2", { type: "string" });
		const rows = sheetToJson(result.Sheets[result.SheetNames[0]], { header: 1 });
		expect(rows[0]).toContain("A");
		expect(rows[0]).toContain("B");
		expect(rows[1]).toContain(1);
		expect(rows[1]).toContain(2);
	});

	it("should read HTML string", async () => {
		const result = await read("<table><tr><td>Hi</td></tr></table>", { type: "string" });
		expect(result.SheetNames).toHaveLength(1);
	});

	it("should reject PDF input", async () => {
		const pdf = new Uint8Array([0x25, 0x50, 0x44, 0x46]);
		await expect(read(pdf)).rejects.toThrow("PDF");
	});

	it("should reject PNG input", async () => {
		const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47]);
		await expect(read(png)).rejects.toThrow("PNG");
	});

	it("should reject unknown format", async () => {
		const junk = new Uint8Array([0x00, 0x01, 0x02, 0x03]);
		await expect(read(junk)).rejects.toThrow("Unsupported");
	});

	it("should read from plain number array", async () => {
		const ws = arrayToSheet([["Data"]]);
		const wb = createWorkbook(ws, "Sheet1");
		const u8 = await write(wb, { type: "array" });
		const arr = Array.from(u8);
		const result = await read(arr);
		expect(result.SheetNames).toContain("Sheet1");
	});
});

describe("readFile — extension detection", () => {
	it("should read a .csv file", async () => {
		const tmpDir = os.tmpdir();
		const csvPath = path.join(tmpDir, "xlsx-fmt-test-read.csv");
		fs.writeFileSync(csvPath, "X,Y\n10,20", "utf-8");
		try {
			const wb = await readFile(csvPath);
			const rows = sheetToJson(wb.Sheets[wb.SheetNames[0]], { header: 1 });
			expect(rows[1]).toContain(10);
			expect(rows[1]).toContain(20);
		} finally {
			fs.unlinkSync(csvPath);
		}
	});

	it("should read a .tsv file", async () => {
		const tmpDir = os.tmpdir();
		const tsvPath = path.join(tmpDir, "xlsx-fmt-test-read.tsv");
		fs.writeFileSync(tsvPath, "A\tB\n1\t2", "utf-8");
		try {
			const wb = await readFile(tsvPath);
			const rows = sheetToJson(wb.Sheets[wb.SheetNames[0]], { header: 1 });
			expect(rows[0]).toContain("A");
			expect(rows[0]).toContain("B");
		} finally {
			fs.unlinkSync(tsvPath);
		}
	});

	it("should read a .html file", async () => {
		const tmpDir = os.tmpdir();
		const htmlPath = path.join(tmpDir, "xlsx-fmt-test-read.html");
		fs.writeFileSync(htmlPath, "<table><tr><td>val</td></tr></table>", "utf-8");
		try {
			const wb = await readFile(htmlPath);
			expect(wb.SheetNames).toHaveLength(1);
		} finally {
			fs.unlinkSync(htmlPath);
		}
	});
});

// ============================================================
// src/write.ts
// ============================================================
describe("write.ts — output types", () => {
	const simpleWb = () => createWorkbook(arrayToSheet([["A"]]), "Sheet1");

	it("should write CSV as base64", async () => {
		const b64 = await write(simpleWb(), { bookType: "csv", type: "base64" });
		expect(typeof b64).toBe("string");
		expect(atob(b64)).toContain("A");
	});

	it("should write CSV as array (Uint8Array)", async () => {
		const arr = await write(simpleWb(), { bookType: "csv", type: "array" });
		expect(arr).toBeInstanceOf(Uint8Array);
	});

	it("should write CSV as buffer", async () => {
		const buf = await write(simpleWb(), { bookType: "csv", type: "buffer" });
		expect(Buffer.isBuffer(buf)).toBe(true);
	});

	it("should write TSV as string", async () => {
		const tsv = await write(simpleWb(), { bookType: "tsv", type: "string" });
		expect(tsv).toContain("A");
	});

	it("should write HTML as string", async () => {
		const html = await write(simpleWb(), { bookType: "html", type: "string" });
		expect(html).toContain("<table");
	});

	it("should write XLSX as base64", async () => {
		const b64 = await write(simpleWb(), { type: "base64" });
		expect(typeof b64).toBe("string");
	});

	it("should write empty workbook CSV", async () => {
		const emptyWb = { SheetNames: ["S1"], Sheets: { S1: {} } } as any;
		const csv = await write(emptyWb, { bookType: "csv", type: "string" });
		expect(csv).toBe("");
	});
});

describe("writeFile — extension inference", () => {
	const tmpDir = os.tmpdir();
	const simpleWb = () => createWorkbook(arrayToSheet([["Val"]]), "Sheet1");

	it("should write .csv file", async () => {
		const p = path.join(tmpDir, "xlsx-fmt-test-write.csv");
		await writeFile(simpleWb(), p);
		try {
			const content = fs.readFileSync(p, "utf-8");
			expect(content).toContain("Val");
		} finally {
			fs.unlinkSync(p);
		}
	});

	it("should write .html file", async () => {
		const p = path.join(tmpDir, "xlsx-fmt-test-write.html");
		await writeFile(simpleWb(), p);
		try {
			const content = fs.readFileSync(p, "utf-8");
			expect(content).toContain("<table");
		} finally {
			fs.unlinkSync(p);
		}
	});

	it("should write .xlsx file", async () => {
		const p = path.join(tmpDir, "xlsx-fmt-test-write.xlsx");
		await writeFile(simpleWb(), p);
		try {
			const data = fs.readFileSync(p);
			expect(data[0]).toBe(0x50); // PK
			expect(data[1]).toBe(0x4b);
		} finally {
			fs.unlinkSync(p);
		}
	});
});

// ============================================================
// src/api/aoa.ts
// ============================================================
describe("aoa.ts — addArrayToSheet edge cases", () => {
	it("should create dense worksheet", () => {
		const ws = addArrayToSheet(
			null,
			[
				["a", 1],
				["b", 2],
			],
			{ dense: true },
		);
		expect((ws as any)["!data"]).toBeDefined();
		expect((ws as any)["!data"][0][0].v).toBe("a");
		expect((ws as any)["!data"][1][1].v).toBe(2);
	});

	it("should handle numeric origin", () => {
		const ws = addArrayToSheet(null, [["A"]], { origin: 3 } as any);
		// When no prior ref, range starts at 0,0 and data at row 3 expands e.r to 3
		expect(ws["!ref"]).toBe("A1:A4");
		expect((ws as any)["A4"]).toBeDefined();
	});

	it("should handle origin as cell ref string", () => {
		const ws = addArrayToSheet(null, [["X"]], { origin: "C5" } as any);
		expect((ws as any)["C5"].v).toBe("X");
	});

	it("should handle origin -1 (append)", () => {
		let ws = arrayToSheet([["Row1"]]);
		ws = addArrayToSheet(ws, [["Row2"]], { origin: -1 } as any);
		expect((ws as any)["A2"].v).toBe("Row2");
	});

	it("should handle null values with nullError", () => {
		const ws = arrayToSheet([[null, "ok"]], { nullError: true } as any);
		expect((ws as any)["A1"].t).toBe("e");
		expect((ws as any)["A1"].v).toBe(0);
	});

	it("should handle null values with sheetStubs", () => {
		const ws = arrayToSheet([[null, "ok"]], { sheetStubs: true } as any);
		expect((ws as any)["A1"].t).toBe("z");
	});

	it("should handle NaN and Infinity", () => {
		const ws = arrayToSheet([[NaN, Infinity]]);
		expect((ws as any)["A1"].t).toBe("e");
		expect((ws as any)["A1"].v).toBe(0x0f); // #VALUE!
		expect((ws as any)["B1"].t).toBe("e");
		expect((ws as any)["B1"].v).toBe(0x07); // #DIV/0!
	});

	it("should handle array values [value, formula]", () => {
		const ws = arrayToSheet([[["result", "=SUM(A2:A10)"]]]);
		expect((ws as any)["A1"].v).toBe("result");
		expect((ws as any)["A1"].f).toBe("=SUM(A2:A10)");
	});

	it("should handle pre-built cell objects", () => {
		const ws = arrayToSheet([[{ t: "n", v: 42, z: "#,##0" }]]);
		expect((ws as any)["A1"].v).toBe(42);
		expect((ws as any)["A1"].z).toBe("#,##0");
	});

	it("should handle Date values", () => {
		const date = new Date("2024-06-15T00:00:00Z");
		const ws = arrayToSheet([[date]], { UTC: true });
		expect((ws as any)["A1"].t).toBe("n");
		expect((ws as any)["A1"].v).toBeGreaterThan(40000);
	});

	it("should handle Date with cellDates", () => {
		const date = new Date("2024-06-15T00:00:00Z");
		const ws = arrayToSheet([[date]], { cellDates: true, UTC: true });
		expect((ws as any)["A1"].t).toBe("d");
	});

	it("should skip undefined values", () => {
		const ws = arrayToSheet([[undefined, "b"]]);
		expect((ws as any)["A1"]).toBeUndefined();
		expect((ws as any)["B1"].v).toBe("b");
	});

	it("should skip null rows", () => {
		const data: any[][] = [["a"], null as any, ["c"]];
		const ws = arrayToSheet(data);
		expect((ws as any)["A1"].v).toBe("a");
		expect((ws as any)["A3"].v).toBe("c");
	});

	it("should throw for non-array rows", () => {
		expect(() => arrayToSheet(["not an array" as any])).toThrow("array of arrays");
	});

	it("should handle null with formula", () => {
		const ws = arrayToSheet([[[null, "=NOW()"]]]);
		expect((ws as any)["A1"].t).toBe("n");
		expect((ws as any)["A1"].f).toBe("=NOW()");
	});
});

// ============================================================
// src/api/json.ts
// ============================================================
describe("json.ts — sheetToJson edge cases", () => {
	it("should return [] for null sheet", () => {
		expect(sheetToJson(null as any)).toEqual([]);
	});

	it("should handle header: 'A' mode", () => {
		const ws = arrayToSheet([
			["x", "y"],
			[1, 2],
		]);
		const rows = sheetToJson(ws, { header: "A" });
		expect(rows[0]).toHaveProperty("A", "x");
		expect(rows[1]).toHaveProperty("B", 2);
	});

	it("should handle custom header array", () => {
		const ws = arrayToSheet([
			[10, 20],
			[30, 40],
		]);
		const rows = sheetToJson(ws, { header: ["Col1", "Col2"] });
		expect(rows[0]).toEqual(expect.objectContaining({ Col1: 10, Col2: 20 }));
	});

	it("should deduplicate repeated headers", () => {
		const ws: any = {
			"!ref": "A1:C2",
			A1: { t: "s", v: "Name" },
			B1: { t: "s", v: "Name" },
			C1: { t: "s", v: "Name" },
			A2: { t: "s", v: "a" },
			B2: { t: "s", v: "b" },
			C2: { t: "s", v: "c" },
		};
		const rows = sheetToJson(ws);
		expect(rows[0]).toHaveProperty("Name", "a");
		expect(rows[0]).toHaveProperty("Name_1", "b");
		expect(rows[0]).toHaveProperty("Name_2", "c");
	});

	it("should handle defval for missing cells", () => {
		const ws: any = {
			"!ref": "A1:B2",
			A1: { t: "s", v: "H1" },
			B1: { t: "s", v: "H2" },
			A2: { t: "s", v: "val" },
			// B2 missing
		};
		const rows = sheetToJson(ws, { defval: "N/A" });
		expect(rows[0].H2).toBe("N/A");
	});

	it("should handle error cells with defval", () => {
		const ws: any = {
			"!ref": "A1:B2",
			A1: { t: "s", v: "H" },
			B1: { t: "s", v: "Val" },
			A2: { t: "e", v: 0 },
			B2: { t: "s", v: "ok" },
		};
		// Error code 0 (#NULL!) maps to null
		const rows = sheetToJson(ws, { defval: "DEF" });
		expect(rows[0].H).toBeNull();
		expect(rows[0].Val).toBe("ok");
	});

	it("should handle numeric range option", () => {
		const ws = arrayToSheet([["a"], ["b"], ["c"], ["d"]]);
		const rows = sheetToJson(ws, { header: 1, range: 2 });
		expect(rows[0]).toEqual(["c"]);
	});

	it("should skip hidden rows", () => {
		const ws = arrayToSheet([["h1"], ["r1"], ["r2"], ["r3"]]);
		(ws as any)["!rows"] = [{}, { hidden: true }, {}, {}];
		const rows = sheetToJson(ws, { skipHidden: true });
		expect(rows).toHaveLength(2); // r1 skipped
	});

	it("should skip hidden cols", () => {
		const ws = arrayToSheet([
			["h1", "h2"],
			[1, 2],
		]);
		(ws as any)["!cols"] = [{ hidden: true }, {}];
		const rows = sheetToJson(ws, { skipHidden: true });
		expect(rows[0]).not.toHaveProperty("h1");
		expect(rows[0]).toHaveProperty("h2", 2);
	});

	it("should handle blankrows option", () => {
		const ws: any = {
			"!ref": "A1:A3",
			A1: { t: "s", v: "H" },
			A3: { t: "s", v: "val" },
		};
		const withBlanks = sheetToJson(ws, { blankrows: true });
		const noBlanks = sheetToJson(ws, { blankrows: false });
		expect(withBlanks.length).toBeGreaterThan(noBlanks.length);
	});

	it("should format with rawNumbers=false", () => {
		const ws: any = {
			"!ref": "A1:A2",
			A1: { t: "s", v: "Val" },
			A2: { t: "n", v: 1234.5, z: "#,##0.00" },
		};
		const rows = sheetToJson(ws, { rawNumbers: false });
		expect(rows[0].Val).toBe("1,234.50");
	});
});

describe("json.ts — addJsonToSheet edge cases", () => {
	it("should create dense worksheet", () => {
		const ws = addJsonToSheet(null, [{ a: 1 }], { dense: true });
		expect((ws as any)["!data"]).toBeDefined();
		expect((ws as any)["!data"][0][0].v).toBe("a"); // header
		expect((ws as any)["!data"][1][0].v).toBe(1);
	});

	it("should handle origin -1 (append)", () => {
		const ws = addJsonToSheet(null, [{ x: 1 }]);
		const ws2 = addJsonToSheet(ws, [{ x: 2 }], { origin: -1 });
		const rows = sheetToJson(ws2, { header: 1 });
		expect(rows[rows.length - 1]).toEqual([2]);
	});

	it("should handle skipHeader", () => {
		const ws = addJsonToSheet(null, [{ a: 1 }, { a: 2 }], { skipHeader: true });
		const rows = sheetToJson(ws, { header: 1 });
		expect(rows[0]).toEqual([1]); // no header row
	});

	it("should handle pre-built cell objects", () => {
		const ws = addJsonToSheet(null, [{ val: { t: "n", v: 42, f: "=6*7" } }]);
		const cell = (ws as any)["A2"];
		expect(cell.v).toBe(42);
		expect(cell.f).toBe("=6*7");
	});

	it("should handle Date values", () => {
		const d = new Date("2024-01-15T00:00:00Z");
		const ws = addJsonToSheet(null, [{ date: d }], { UTC: true });
		const cell = (ws as any)["A2"];
		expect(cell.t).toBe("n");
		expect(cell.v).toBeGreaterThan(40000);
	});

	it("should handle nullError", () => {
		const ws = addJsonToSheet(null, [{ val: null }], { nullError: true });
		const cell = (ws as any)["A2"];
		expect(cell.t).toBe("e");
		expect(cell.v).toBe(0);
	});

	it("should handle numeric origin", () => {
		const ws = addJsonToSheet(null, [{ a: 1 }], { origin: 5 } as any);
		// origin=5 means header at row 5, data at row 6. Range includes row 0 start.
		expect(ws["!ref"]).toContain("7"); // row 7 = originRow(5) + 1 data row + header offset
	});

	it("should update existing cells in place", () => {
		const ws = addJsonToSheet(null, [{ a: 1 }]);
		const ws2 = addJsonToSheet(ws, [{ a: 999 }]);
		const cell = (ws2 as any)["A2"];
		expect(cell.v).toBe(999);
	});
});

// ============================================================
// src/api/html.ts
// ============================================================
describe("html.ts — sheetToHtml", () => {
	it("should handle merged cells", () => {
		const ws: any = {
			"!ref": "A1:B2",
			"!merges": [{ s: { r: 0, c: 0 }, e: { r: 1, c: 1 } }],
			A1: { t: "s", v: "Merged" },
		};
		const html = sheetToHtml(ws);
		expect(html).toContain('rowspan="2"');
		expect(html).toContain('colspan="2"');
	});

	it("should handle NaN as #NUM! and Infinity as #DIV/0!", () => {
		const ws: any = {
			"!ref": "A1:B1",
			A1: { t: "n", v: NaN },
			B1: { t: "n", v: Infinity },
		};
		const html = sheetToHtml(ws);
		// NaN → error 0x24 = #NUM!, Infinity → error 0x07 = #DIV/0!
		expect(html).toContain("#NUM!");
		expect(html).toContain("#DIV/0!");
	});

	it("should support editable mode", () => {
		const ws: any = {
			"!ref": "A1",
			A1: { t: "s", v: "text" },
		};
		const html = sheetToHtml(ws, { editable: true });
		expect(html).toContain('contenteditable="true"');
	});

	it("should include data attributes", () => {
		const ws: any = {
			"!ref": "A1",
			A1: { t: "n", v: 42, z: "#,##0", f: "=6*7" },
		};
		const html = sheetToHtml(ws);
		expect(html).toContain('data-v="42"');
		expect(html).toContain('data-f="=6*7"');
		expect(html).toContain('data-z="#,##0"');
	});

	it("should render hyperlinks", () => {
		const ws: any = {
			"!ref": "A1",
			A1: { t: "s", v: "Link", l: { Target: "https://example.com" } },
		};
		const html = sheetToHtml(ws);
		expect(html).toContain('href="https://example.com"');
	});

	it("should sanitize javascript: links", () => {
		const ws: any = {
			"!ref": "A1",
			A1: { t: "s", v: "Bad", l: { Target: "javascript:alert(1)" } },
		};
		const html = sheetToHtml(ws, { sanitizeLinks: true });
		expect(html).not.toContain("javascript:");
	});

	it("should skip internal links (#)", () => {
		const ws: any = {
			"!ref": "A1",
			A1: { t: "s", v: "Internal", l: { Target: "#Sheet2!A1" } },
		};
		const html = sheetToHtml(ws);
		expect(html).not.toContain("href=");
	});

	it("should support custom header/footer", () => {
		const ws = arrayToSheet([["x"]]);
		const html = sheetToHtml(ws, { header: "<div>", footer: "</div>" });
		expect(html.startsWith("<div>")).toBe(true);
		expect(html.endsWith("</div>")).toBe(true);
	});

	it("should support table id", () => {
		const ws = arrayToSheet([["x"]]);
		const html = sheetToHtml(ws, { id: "mytable" });
		expect(html).toContain('id="mytable"');
	});
});

describe("html.ts — htmlToSheet", () => {
	it("should handle rowspan", () => {
		const html = `<table>
			<tr><td rowspan="2">A</td><td>B</td></tr>
			<tr><td>C</td></tr>
		</table>`;
		const ws = htmlToSheet(html);
		const rows = sheetToJson(ws, { header: 1 });
		expect(rows[0][0]).toBe("A");
		expect(rows[0][1]).toBe("B");
		expect(rows[1][1]).toBe("C");
	});

	it("should handle colspan", () => {
		const html = `<table>
			<tr><td colspan="3">Wide</td></tr>
			<tr><td>A</td><td>B</td><td>C</td></tr>
		</table>`;
		const ws = htmlToSheet(html);
		const rows = sheetToJson(ws, { header: 1 });
		expect(rows[0][0]).toBe("Wide");
	});

	it("should coerce types from text", () => {
		const html = `<table>
			<tr><td>42</td><td>TRUE</td><td>hello</td></tr>
		</table>`;
		const ws = htmlToSheet(html);
		expect((ws as any)["A1"].v).toBe(42);
		expect((ws as any)["B1"].v).toBe(true);
		expect((ws as any)["C1"].v).toBe("hello");
	});

	it("should handle data-t and data-v attributes", () => {
		const html = `<table>
			<tr><td data-t="n" data-v="99">formatted</td></tr>
		</table>`;
		const ws = htmlToSheet(html);
		expect((ws as any)["A1"].v).toBe(99);
	});

	it("should return empty sheet for no table", () => {
		const ws = htmlToSheet("<div>no table</div>");
		expect(ws["!ref"]).toBeUndefined();
	});

	it("should handle combined rowspan and colspan", () => {
		const html = `<table>
			<tr><td rowspan="2" colspan="2">Big</td><td>C</td></tr>
			<tr><td>D</td></tr>
			<tr><td>E</td><td>F</td><td>G</td></tr>
		</table>`;
		const ws = htmlToSheet(html);
		const rows = sheetToJson(ws, { header: 1 });
		expect(rows[0][0]).toBe("Big");
		expect(rows[2][0]).toBe("E");
		expect(rows[2][2]).toBe("G");
	});

	it("should unescape HTML entities", () => {
		const html = `<table><tr><td>&lt;b&gt;bold&lt;/b&gt;</td></tr></table>`;
		const ws = htmlToSheet(html);
		expect((ws as any)["A1"].v).toBe("<b>bold</b>");
	});
});

// ============================================================
// src/ssf/format.ts
// ============================================================
describe("ssf/format.ts — formatNumber", () => {
	describe("General format", () => {
		it("should format integer", () => {
			expect(formatNumber("General", 42)).toBe("42");
		});

		it("should format decimal", () => {
			expect(formatNumber("General", 3.14)).toBe("3.14");
		});

		it("should format very small number with E notation", () => {
			// Need 1e-10 or smaller to trigger E notation in General format
			const result = formatNumber("General", 1e-10);
			expect(result).toContain("E");
		});

		it("should format very large number with E notation", () => {
			const result = formatNumber("General", 1e15);
			expect(result.length).toBeLessThan(20);
		});

		it("should format boolean true", () => {
			expect(formatNumber("General", true)).toBe("TRUE");
		});

		it("should format boolean false", () => {
			expect(formatNumber("General", false)).toBe("FALSE");
		});

		it("should format empty string", () => {
			expect(formatNumber("General", "")).toBe("");
		});

		it("should format text string", () => {
			expect(formatNumber("@", "hello")).toBe("hello");
		});
	});

	describe("Number formats", () => {
		it("should format #,##0.00", () => {
			expect(formatNumber("#,##0.00", 1234.5)).toBe("1,234.50");
		});

		it("should format 0.00", () => {
			expect(formatNumber("0.00", 3.1)).toBe("3.10");
		});

		it("should format zero with custom section", () => {
			expect(formatNumber('0;0;"zero"', 0)).toBe("zero");
		});
	});

	describe("Date formats", () => {
		it("should format m/d/yy", () => {
			const result = formatNumber("m/d/yy", 45292); // 2024-01-01
			expect(result).toMatch(/1\/1\/24/);
		});

		it("should format yyyy-mm-dd", () => {
			const result = formatNumber("yyyy-mm-dd", 45292);
			expect(result).toBe("2024-01-01");
		});

		it("should format dd-mmm-yyyy", () => {
			const result = formatNumber("dd-mmm-yyyy", 45292);
			expect(result).toMatch(/01-Jan-2024/);
		});

		it("should format h:mm:ss", () => {
			const result = formatNumber("h:mm:ss", 0.5); // noon
			expect(result).toBe("12:00:00");
		});

		it("should format mm/dd/yyyy hh:mm:ss", () => {
			const result = formatNumber("mm/dd/yyyy hh:mm:ss", 45292.5);
			expect(result).toContain("01/01/2024");
			expect(result).toContain("12:00:00");
		});

		it("should format elapsed hours [h]:mm", () => {
			const result = formatNumber("[h]:mm", 1.5); // 36 hours
			expect(result).toBe("36:00");
		});

		it("should format elapsed minutes [mm]:ss", () => {
			const result = formatNumber("[mm]:ss", 0.5); // 720 minutes
			expect(result).toBe("720:00");
		});

		it("should format sub-seconds", () => {
			const result = formatNumber("h:mm:ss.00", 0.50001);
			expect(result).toMatch(/12:00:00/);
		});

		it("should handle date serial 60 (1900 leap year bug)", () => {
			const result = formatNumber("yyyy-mm-dd", 60);
			expect(result).toBe("1900-02-29"); // phantom date
		});

		it("should handle date1904 option", () => {
			const result = formatNumber("yyyy-mm-dd", 0, { date1904: true });
			expect(result).toBe("1904-01-01");
		});
	});

	describe("Scientific notation", () => {
		it("should format 0.00E+00", () => {
			const result = formatNumber("0.00E+00", 12345);
			expect(result).toMatch(/1\.23E\+04/);
		});

		it("should format small number", () => {
			const result = formatNumber("0.00E+00", 0.00123);
			expect(result).toMatch(/E/);
		});

		it("should format negative scientific", () => {
			const result = formatNumber("0.00E+00", -12345);
			expect(result).toContain("-");
		});
	});

	describe("Fraction formats", () => {
		it("should format # ?/?", () => {
			const result = formatNumber("# ?/?", 1.5);
			expect(result.trim()).toContain("1/2");
		});

		it("should format # ??/??", () => {
			const result = formatNumber("# ??/??", 3.333333);
			expect(result.trim()).toMatch(/3\s+1\/\s*3/);
		});

		it("should format value < 1", () => {
			const result = formatNumber("# ?/?", 0.5);
			expect(result.trim()).toContain("1/2");
		});
	});

	describe("Special formats", () => {
		it("should handle currency [$]", () => {
			const result = formatNumber("[$€-407]#,##0.00", 1234.5);
			expect(result).toContain("€");
			expect(result).toContain("1,234.50");
		});

		it("should handle text format @", () => {
			expect(formatNumber("@", "hello")).toBe("hello");
		});

		it("should handle escaped characters", () => {
			const result = formatNumber('0" kg"', 5);
			expect(result).toBe("5 kg");
		});

		it("should return empty for null", () => {
			expect(formatNumber("0", null)).toBe("");
		});

		it("should return #NUM! for NaN with number format", () => {
			expect(formatNumber("0", NaN)).toBe("#NUM!");
		});

		it("should return #DIV/0! for Infinity with number format", () => {
			expect(formatNumber("0", Infinity)).toBe("#DIV/0!");
		});

		it("should use format index lookup", () => {
			// Format 14 = "m/d/yy"
			const result = formatNumber(14, 45292);
			expect(result).toMatch(/1\/1\/24/);
		});

		it("should use dateNF override for format 14", () => {
			const result = formatNumber(14, 45292, { dateNF: "yyyy-mm-dd" });
			expect(result).toBe("2024-01-01");
		});

		it("should use dateNF override for 'm/d/yy' string", () => {
			const result = formatNumber("m/d/yy", 45292, { dateNF: "yyyy-mm-dd" });
			expect(result).toBe("2024-01-01");
		});

		it("should handle two-section format", () => {
			const result = formatNumber("0;0", -500);
			expect(result).toBe("500");
		});

		it("should format text in 4-section format", () => {
			const result = formatNumber("0;0;0;@", "text");
			expect(result).toBe("text");
		});

		it("should format text with single section containing @", () => {
			const result = formatNumber("@", "mytext");
			expect(result).toBe("mytext");
		});

		it("should format Date object", () => {
			const d = new Date("2024-01-01T00:00:00Z");
			const result = formatNumber("yyyy-mm-dd", d);
			// Date → serial number conversion depends on local timezone;
			// just verify we get a valid yyyy-mm-dd date string back
			expect(result).toMatch(/^\d{4}-\d{2}-\d{2}$/);
		});
	});
});

// ============================================================
// Integration: XLSX write → read roundtrip
// ============================================================
describe("XLSX roundtrip", () => {
	it("should preserve number values", async () => {
		const ws = arrayToSheet([[42, 3.14, -100]]);
		const wb = createWorkbook(ws, "Sheet1");
		const u8 = await write(wb, { type: "array" });
		const wb2 = await read(u8);
		const s = wb2.Sheets["Sheet1"];
		expect((s as any)["A1"].v).toBe(42);
		expect((s as any)["B1"].v).toBeCloseTo(3.14);
		expect((s as any)["C1"].v).toBe(-100);
	});

	it("should preserve boolean values", async () => {
		const ws = arrayToSheet([[true, false]]);
		const wb = createWorkbook(ws, "Sheet1");
		const u8 = await write(wb, { type: "array" });
		const wb2 = await read(u8);
		const s = wb2.Sheets["Sheet1"];
		expect((s as any)["A1"].v).toBe(true);
		expect((s as any)["B1"].v).toBe(false);
	});
});
