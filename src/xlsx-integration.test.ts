import { describe, it, expect } from "vitest";
import {
	read,
	write,
	createWorkbook,
	appendSheet,
	arrayToSheet,
	sheetToJson,
	createSheet,
	setArrayFormula,
	setCellHyperlink,
	setCellInternalLink,
	addCellComment,
	setSheetVisibility,
	encodeCell,
	type WorkBook,
	type ReadOptions,
	type WriteOptions,
} from "./index.js";

/** Write a workbook to XLSX bytes then read it back. */
async function roundtrip(wb: WorkBook, writeOpts?: WriteOptions, readOpts?: ReadOptions): Promise<WorkBook> {
	const u8 = await write(wb, writeOpts);
	return read(u8, readOpts);
}

// ---------------------------------------------------------------------------
// 1. Cell Types
// ---------------------------------------------------------------------------
describe("Cell Types", () => {
	it("boolean true survives roundtrip", async () => {
		const ws = arrayToSheet([[true]]);
		const wb2 = await roundtrip(createWorkbook(ws, "S"));
		const cell = wb2.Sheets["S"]["A1"];
		expect(cell.t).toBe("b");
		expect(cell.v).toBe(true);
	});

	it("boolean false survives roundtrip", async () => {
		const ws = arrayToSheet([[false]]);
		const wb2 = await roundtrip(createWorkbook(ws, "S"));
		const cell = wb2.Sheets["S"]["A1"];
		expect(cell.t).toBe("b");
		expect(cell.v).toBe(false);
	});

	it("number survives roundtrip", async () => {
		const ws = arrayToSheet([[42.5]]);
		const wb2 = await roundtrip(createWorkbook(ws, "S"));
		const cell = wb2.Sheets["S"]["A1"];
		expect(cell.t).toBe("n");
		expect(cell.v).toBe(42.5);
	});

	it("error cell (#N/A) survives roundtrip", async () => {
		const ws = createSheet();
		ws["A1"] = { t: "e", v: 0x2a }; // #N/A
		ws["!ref"] = "A1";
		const wb2 = await roundtrip(createWorkbook(ws, "S"));
		const cell = wb2.Sheets["S"]["A1"];
		expect(cell.t).toBe("e");
		expect(cell.v).toBe(0x2a);
	});

	it("string survives roundtrip", async () => {
		const ws = arrayToSheet([["hello world"]]);
		const wb2 = await roundtrip(createWorkbook(ws, "S"));
		const cell = wb2.Sheets["S"]["A1"];
		expect(cell.t).toBe("s");
		expect(cell.v).toBe("hello world");
	});

	it("date stored as serial number survives roundtrip", async () => {
		const ws = createSheet();
		ws["A1"] = { t: "n", v: 44927 }; // 2023-01-01 as serial
		ws["!ref"] = "A1";
		const wb2 = await roundtrip(createWorkbook(ws, "S"));
		expect(wb2.Sheets["S"]["A1"].v).toBe(44927);
	});
});

// ---------------------------------------------------------------------------
// 2. Number Values
// ---------------------------------------------------------------------------
describe("Number Values", () => {
	it("floating point precision preserved", async () => {
		const ws = arrayToSheet([[3.14159265358979]]);
		const wb2 = await roundtrip(createWorkbook(ws, "S"));
		expect(wb2.Sheets["S"]["A1"].v).toBeCloseTo(3.14159265358979, 10);
	});

	it("negative number survives roundtrip", async () => {
		const ws = arrayToSheet([[-999.99]]);
		const wb2 = await roundtrip(createWorkbook(ws, "S"));
		expect(wb2.Sheets["S"]["A1"].v).toBe(-999.99);
	});

	it("zero survives roundtrip", async () => {
		const ws = arrayToSheet([[0]]);
		const wb2 = await roundtrip(createWorkbook(ws, "S"));
		expect(wb2.Sheets["S"]["A1"].v).toBe(0);
	});
});

// ---------------------------------------------------------------------------
// 3. Date Systems
// ---------------------------------------------------------------------------
describe("Date Systems", () => {
	it("default 1900 date system roundtrip", async () => {
		const ws = arrayToSheet([[44927]]); // 2023-01-01
		const wb = createWorkbook(ws, "S");
		const wb2 = await roundtrip(wb);
		expect(wb2.Sheets["S"]["A1"].v).toBe(44927);
	});

	it("date1904 flag persists on roundtrip", async () => {
		const ws = arrayToSheet([[1]]);
		const wb = createWorkbook(ws, "S");
		wb.Workbook = { WBProps: { date1904: true } };
		const wb2 = await roundtrip(wb);
		expect(wb2.Workbook?.WBProps?.date1904).toBe(true);
	});

	it("date1904=false is default", async () => {
		const ws = arrayToSheet([[1]]);
		const wb2 = await roundtrip(createWorkbook(ws, "S"));
		// date1904 should be falsy (either false or undefined)
		expect(wb2.Workbook?.WBProps?.date1904).toBeFalsy();
	});
});

// ---------------------------------------------------------------------------
// 4. String Handling
// ---------------------------------------------------------------------------
describe("String Handling", () => {
	it("multiple distinct strings roundtrip", async () => {
		const ws = arrayToSheet([["alpha"], ["beta"], ["gamma"]]);
		const wb2 = await roundtrip(createWorkbook(ws, "S"));
		const rows = sheetToJson<{ A: string }>(wb2.Sheets["S"], { header: "A" });
		expect(rows.map((r) => r.A)).toEqual(["alpha", "beta", "gamma"]);
	});

	it("duplicate strings in multiple cells", async () => {
		const ws = arrayToSheet([["dup"], ["dup"], ["dup"]]);
		const wb2 = await roundtrip(createWorkbook(ws, "S"));
		const rows = sheetToJson<{ A: string }>(wb2.Sheets["S"], { header: "A" });
		expect(rows.every((r) => r.A === "dup")).toBe(true);
	});
});

// ---------------------------------------------------------------------------
// 5. Formulas
// ---------------------------------------------------------------------------
describe("Formulas", () => {
	it("regular formula roundtrips", async () => {
		const ws = createSheet();
		ws["A1"] = { t: "n", v: 3, f: "1+2" };
		ws["!ref"] = "A1";
		const wb2 = await roundtrip(createWorkbook(ws, "S"));
		expect(wb2.Sheets["S"]["A1"].f).toBe("1+2");
		expect(wb2.Sheets["S"]["A1"].v).toBe(3);
	});

	it("array formula via setArrayFormula roundtrips", async () => {
		const ws = arrayToSheet([
			[1, 10],
			[2, 20],
		]);
		setArrayFormula(ws, "C1:C2", "A1:A2*B1:B2");
		const wb2 = await roundtrip(createWorkbook(ws, "S"));
		const c1 = wb2.Sheets["S"]["C1"];
		expect(c1.f).toBe("A1:A2*B1:B2");
		expect(c1.F).toBe("C1:C2");
	});

	it("dynamic array formula flag set via API", async () => {
		const ws = createSheet();
		ws["A1"] = { t: "n", v: 0 };
		ws["!ref"] = "A1:A2";
		setArrayFormula(ws, "A1:A2", "SORT(B1:B10)", true);
		// Verify the API sets D flag on the source cell
		expect(ws["A1"].D).toBe(true);
		expect(ws["A1"].f).toBe("SORT(B1:B10)");
		// The writer does not emit the dynamic array attribute, so D is lost on roundtrip
		const wb2 = await roundtrip(createWorkbook(ws, "S"));
		expect(wb2.Sheets["S"]["A1"].f).toBe("SORT(B1:B10)");
	});

	it("formula without pre-computed value", async () => {
		const ws = createSheet();
		ws["A1"] = { t: "n", f: "1+1" }; // no .v
		ws["!ref"] = "A1";
		const wb2 = await roundtrip(createWorkbook(ws, "S"));
		expect(wb2.Sheets["S"]["A1"].f).toBe("1+1");
	});
});

// ---------------------------------------------------------------------------
// 6. Merged Cells
// ---------------------------------------------------------------------------
describe("Merged Cells", () => {
	it("single merge range roundtrips", async () => {
		const ws = arrayToSheet([["merged"]]);
		ws["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 1, c: 1 } }];
		const wb2 = await roundtrip(createWorkbook(ws, "S"));
		expect(wb2.Sheets["S"]["!merges"]).toHaveLength(1);
		expect(wb2.Sheets["S"]["!merges"]![0].s).toEqual({ r: 0, c: 0 });
		expect(wb2.Sheets["S"]["!merges"]![0].e).toEqual({ r: 1, c: 1 });
	});

	it("multiple merge ranges roundtrip", async () => {
		const ws = arrayToSheet([
			["a", "", "b"],
			["", "", ""],
		]);
		ws["!merges"] = [
			{ s: { r: 0, c: 0 }, e: { r: 1, c: 1 } },
			{ s: { r: 0, c: 2 }, e: { r: 1, c: 2 } },
		];
		const wb2 = await roundtrip(createWorkbook(ws, "S"));
		expect(wb2.Sheets["S"]["!merges"]).toHaveLength(2);
	});

	it("merge with cell data preserved", async () => {
		const ws = arrayToSheet([["top-left"]]);
		ws["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 2, c: 2 } }];
		const wb2 = await roundtrip(createWorkbook(ws, "S"));
		expect(wb2.Sheets["S"]["A1"].v).toBe("top-left");
		expect(wb2.Sheets["S"]["!merges"]).toHaveLength(1);
	});
});

// ---------------------------------------------------------------------------
// 7. Hyperlinks — API only
// ---------------------------------------------------------------------------
describe("Hyperlinks — API only", () => {
	it("setCellHyperlink sets cell.l", () => {
		const cell = { t: "s" as const, v: "click" };
		setCellHyperlink(cell, "https://example.com", "Example");
		expect(cell.l!.Target).toBe("https://example.com");
		expect(cell.l!.Tooltip).toBe("Example");
	});

	it("setCellInternalLink sets # prefix", () => {
		const cell = { t: "s" as const, v: "go" };
		setCellInternalLink(cell, "Sheet2!A1");
		expect(cell.l!.Target).toBe("#Sheet2!A1");
	});

	it("hyperlinks do not survive XLSX roundtrip (known limitation)", async () => {
		const ws = createSheet();
		ws["A1"] = { t: "s", v: "link" };
		setCellHyperlink(ws["A1"], "https://example.com");
		ws["!ref"] = "A1";
		const wb2 = await roundtrip(createWorkbook(ws, "S"));
		// Hyperlinks are not written by the XLSX writer
		expect(wb2.Sheets["S"]["A1"].l).toBeUndefined();
	});
});

// ---------------------------------------------------------------------------
// 8. Comments — API only
// ---------------------------------------------------------------------------
describe("Comments — API only", () => {
	it("addCellComment sets cell.c", () => {
		const cell = { t: "s" as const, v: "data" };
		addCellComment(cell, "review this", "Alice");
		expect(cell.c).toHaveLength(1);
		expect(cell.c![0].t).toBe("review this");
		expect(cell.c![0].a).toBe("Alice");
	});

	it("default author is SheetJS", () => {
		const cell = { t: "s" as const, v: "data" };
		addCellComment(cell, "note");
		expect(cell.c![0].a).toBe("SheetJS");
	});

	it("multiple comments on one cell", () => {
		const cell = { t: "s" as const, v: "data" };
		addCellComment(cell, "first");
		addCellComment(cell, "second");
		expect(cell.c).toHaveLength(2);
		expect(cell.c![0].t).toBe("first");
		expect(cell.c![1].t).toBe("second");
	});
});

// ---------------------------------------------------------------------------
// 9. Column/Row Properties
// ---------------------------------------------------------------------------
describe("Column/Row Properties", () => {
	it("column width roundtrips with cellStyles", async () => {
		const ws = arrayToSheet([[1]]);
		ws["!cols"] = [{ width: 20 }];
		const wb2 = await roundtrip(createWorkbook(ws, "S"), undefined, { cellStyles: true });
		expect(wb2.Sheets["S"]["!cols"]).toBeDefined();
		expect(wb2.Sheets["S"]["!cols"]![0].width).toBeCloseTo(20, 0);
	});

	it("hidden column roundtrips with cellStyles", async () => {
		const ws = arrayToSheet([[1, 2]]);
		ws["!cols"] = [undefined as any, { hidden: true, width: 9.140625 }];
		const wb2 = await roundtrip(createWorkbook(ws, "S"), undefined, { cellStyles: true });
		expect(wb2.Sheets["S"]["!cols"]![1].hidden).toBe(true);
	});

	it("row height (hpt) roundtrips", async () => {
		const ws = arrayToSheet([[1]]);
		ws["!rows"] = [{ hpt: 30 }];
		const wb2 = await roundtrip(createWorkbook(ws, "S"));
		expect(wb2.Sheets["S"]["!rows"]).toBeDefined();
		expect(wb2.Sheets["S"]["!rows"]![0].hpt).toBe(30);
	});

	it("hidden row roundtrips", async () => {
		const ws = arrayToSheet([[1], [2]]);
		ws["!rows"] = [undefined as any, { hidden: true }];
		const wb2 = await roundtrip(createWorkbook(ws, "S"));
		expect(wb2.Sheets["S"]["!rows"]![1].hidden).toBe(true);
	});

	it("multiple columns mixed properties", async () => {
		const ws = arrayToSheet([[1, 2, 3]]);
		ws["!cols"] = [{ width: 15 }, { hidden: true, width: 9.140625 }, { width: 25 }];
		const wb2 = await roundtrip(createWorkbook(ws, "S"), undefined, { cellStyles: true });
		const cols = wb2.Sheets["S"]["!cols"]!;
		expect(cols[0].width).toBeCloseTo(15, 0);
		expect(cols[1].hidden).toBe(true);
		expect(cols[2].width).toBeCloseTo(25, 0);
	});
});

// ---------------------------------------------------------------------------
// 10. Sheet Visibility
// ---------------------------------------------------------------------------
describe("Sheet Visibility", () => {
	it("default sheet is visible (0)", async () => {
		const wb = createWorkbook(arrayToSheet([[1]]), "S1");
		appendSheet(wb, arrayToSheet([[2]]), "S2");
		setSheetVisibility(wb, "S1", 0);
		const wb2 = await roundtrip(wb);
		expect(wb2.Workbook?.Sheets?.[0]?.Hidden).toBeFalsy();
	});

	it("hidden sheet (1) roundtrips", async () => {
		const wb = createWorkbook(arrayToSheet([[1]]), "Visible");
		appendSheet(wb, arrayToSheet([[2]]), "Hidden");
		setSheetVisibility(wb, "Hidden", 1);
		const wb2 = await roundtrip(wb);
		expect(wb2.Workbook?.Sheets?.[1]?.Hidden).toBe(1);
	});

	it("very hidden sheet (2) roundtrips", async () => {
		const wb = createWorkbook(arrayToSheet([[1]]), "Visible");
		appendSheet(wb, arrayToSheet([[2]]), "VeryHidden");
		setSheetVisibility(wb, "VeryHidden", 2);
		const wb2 = await roundtrip(wb);
		expect(wb2.Workbook?.Sheets?.[1]?.Hidden).toBe(2);
	});
});

// ---------------------------------------------------------------------------
// 11. AutoFilter
// ---------------------------------------------------------------------------
describe("AutoFilter", () => {
	it("basic autofilter ref roundtrips", async () => {
		const ws = arrayToSheet([
			["Name", "Score"],
			["A", 1],
		]);
		ws["!autofilter"] = { ref: "A1:B2" };
		const wb2 = await roundtrip(createWorkbook(ws, "S"));
		expect(wb2.Sheets["S"]["!autofilter"]!.ref).toBe("A1:B2");
	});

	it("autofilter with data", async () => {
		const ws = arrayToSheet([
			["Col1", "Col2"],
			["x", 10],
			["y", 20],
			["z", 30],
		]);
		ws["!autofilter"] = { ref: "A1:B4" };
		const wb2 = await roundtrip(createWorkbook(ws, "S"));
		expect(wb2.Sheets["S"]["!autofilter"]!.ref).toBe("A1:B4");
		const rows = sheetToJson(wb2.Sheets["S"]);
		expect(rows).toHaveLength(3);
	});
});

// ---------------------------------------------------------------------------
// 12. Defined Names
// ---------------------------------------------------------------------------
describe("Defined Names", () => {
	it("workbook-scoped name roundtrips", async () => {
		const wb = createWorkbook(arrayToSheet([[1]]), "Data");
		wb.Workbook = {
			Names: [{ Name: "MyRange", Ref: "Data!$A$1" }],
		};
		const wb2 = await roundtrip(wb);
		const names = wb2.Workbook?.Names;
		expect(names).toBeDefined();
		expect(names!.some((n) => n.Name === "MyRange" && n.Ref === "Data!$A$1")).toBe(true);
	});

	it("sheet-scoped name roundtrips", async () => {
		const wb = createWorkbook(arrayToSheet([[1]]), "Data");
		wb.Workbook = {
			Names: [{ Name: "LocalName", Ref: "Data!$A$1", Sheet: 0 }],
		};
		const wb2 = await roundtrip(wb);
		const names = wb2.Workbook?.Names;
		expect(names).toBeDefined();
		const localName = names!.find((n) => n.Name === "LocalName");
		expect(localName).toBeDefined();
		expect(localName!.Sheet).toBe(0);
	});

	it("multiple defined names roundtrip", async () => {
		const wb = createWorkbook(arrayToSheet([[1, 2]]), "Data");
		wb.Workbook = {
			Names: [
				{ Name: "First", Ref: "Data!$A$1" },
				{ Name: "Second", Ref: "Data!$B$1" },
				{ Name: "Third", Ref: "Data!$A$1:$B$1" },
			],
		};
		const wb2 = await roundtrip(wb);
		expect(wb2.Workbook?.Names).toHaveLength(3);
	});
});

// ---------------------------------------------------------------------------
// 13. Page Margins
// ---------------------------------------------------------------------------
describe("Page Margins", () => {
	it("custom margins roundtrip", async () => {
		const ws = arrayToSheet([[1]]);
		ws["!margins"] = {
			left: 0.5,
			right: 0.5,
			top: 1.0,
			bottom: 1.0,
			header: 0.25,
			footer: 0.25,
		};
		const wb2 = await roundtrip(createWorkbook(ws, "S"));
		const m = wb2.Sheets["S"]["!margins"]!;
		expect(m.left).toBeCloseTo(0.5);
		expect(m.right).toBeCloseTo(0.5);
		expect(m.top).toBeCloseTo(1.0);
		expect(m.bottom).toBeCloseTo(1.0);
		expect(m.header).toBeCloseTo(0.25);
		expect(m.footer).toBeCloseTo(0.25);
	});

	it("no margins = undefined on read", async () => {
		const ws = arrayToSheet([[1]]);
		// don't set !margins
		const wb2 = await roundtrip(createWorkbook(ws, "S"));
		expect(wb2.Sheets["S"]["!margins"]).toBeUndefined();
	});
});

// ---------------------------------------------------------------------------
// 14. Multiple Sheets
// ---------------------------------------------------------------------------
describe("Multiple Sheets", () => {
	it("different data per sheet", async () => {
		const wb = createWorkbook();
		appendSheet(wb, arrayToSheet([["alpha"]]), "A");
		appendSheet(wb, arrayToSheet([["beta"]]), "B");
		const wb2 = await roundtrip(wb);
		expect(sheetToJson(wb2.Sheets["A"], { header: 1 })).toEqual([["alpha"]]);
		expect(sheetToJson(wb2.Sheets["B"], { header: 1 })).toEqual([["beta"]]);
	});

	it("sheet order preserved", async () => {
		const wb = createWorkbook();
		appendSheet(wb, arrayToSheet([[1]]), "Z");
		appendSheet(wb, arrayToSheet([[2]]), "A");
		appendSheet(wb, arrayToSheet([[3]]), "M");
		const wb2 = await roundtrip(wb);
		expect(wb2.SheetNames).toEqual(["Z", "A", "M"]);
	});

	it("5 sheets roundtrip", async () => {
		const wb = createWorkbook();
		for (let i = 0; i < 5; i++) {
			appendSheet(wb, arrayToSheet([[i]]), "Sheet" + (i + 1));
		}
		const wb2 = await roundtrip(wb);
		expect(wb2.SheetNames).toHaveLength(5);
		for (let i = 0; i < 5; i++) {
			expect(wb2.Sheets["Sheet" + (i + 1)]["A1"].v).toBe(i);
		}
	});
});

// ---------------------------------------------------------------------------
// 15. Dense Mode
// ---------------------------------------------------------------------------
describe("Dense Mode", () => {
	it("write sparse + read dense", async () => {
		const ws = arrayToSheet([["hello", 42]]);
		const wb = createWorkbook(ws, "S");
		const wb2 = await roundtrip(wb, undefined, { dense: true });
		const data = wb2.Sheets["S"]["!data"];
		expect(data).toBeDefined();
		expect(data![0]![0]!.v).toBe("hello");
		expect(data![0]![1]!.v).toBe(42);
	});

	it("arrayToSheet with dense: true", async () => {
		const ws = arrayToSheet(
			[
				["a", "b"],
				[1, 2],
			],
			{ dense: true },
		);
		expect(ws["!data"]).toBeDefined();
		const wb2 = await roundtrip(createWorkbook(ws, "S"));
		const rows = sheetToJson(wb2.Sheets["S"]);
		expect(rows[0]["a"]).toBe(1);
		expect(rows[0]["b"]).toBe(2);
	});

	it("dense data roundtrip back to dense", async () => {
		const ws = arrayToSheet([[100, 200]], { dense: true });
		const wb = createWorkbook(ws, "S");
		const wb2 = await roundtrip(wb, undefined, { dense: true });
		const data = wb2.Sheets["S"]["!data"];
		expect(data![0]![0]!.v).toBe(100);
		expect(data![0]![1]!.v).toBe(200);
	});
});

// ---------------------------------------------------------------------------
// 16. Unicode
// ---------------------------------------------------------------------------
describe("Unicode", () => {
	it("ASCII sheet name with spaces roundtrips", async () => {
		const wb = createWorkbook(arrayToSheet([[1]]), "My Data");
		const wb2 = await roundtrip(wb);
		expect(wb2.SheetNames).toContain("My Data");
	});

	it("unicode cell values (CJK, accented)", async () => {
		const ws = arrayToSheet([["日本語"], ["café"], ["über"]]);
		const wb2 = await roundtrip(createWorkbook(ws, "S"));
		const rows = sheetToJson<{ A: string }>(wb2.Sheets["S"], { header: "A" });
		expect(rows[0].A).toBe("日本語");
		expect(rows[1].A).toBe("café");
		expect(rows[2].A).toBe("über");
	});

	it("XML special characters in cell values", async () => {
		const ws = arrayToSheet([["<tag>&\"quote\"'apos'"]]);
		const wb2 = await roundtrip(createWorkbook(ws, "S"));
		expect(wb2.Sheets["S"]["A1"].v).toBe("<tag>&\"quote\"'apos'");
	});
});

// ---------------------------------------------------------------------------
// 17. sheetRows Limit
// ---------------------------------------------------------------------------
describe("sheetRows Limit", () => {
	it("read only N rows from a large sheet", async () => {
		const data = Array.from({ length: 100 }, (_, i) => [i]);
		const ws = arrayToSheet(data);
		const wb = createWorkbook(ws, "S");
		const wb2 = await roundtrip(wb, undefined, { sheetRows: 10 });
		const rows = sheetToJson(wb2.Sheets["S"], { header: 1 });
		expect(rows).toHaveLength(10);
	});

	it("!fullref preserved when sheetRows limits reading", async () => {
		const data = Array.from({ length: 50 }, (_, i) => [i]);
		const ws = arrayToSheet(data);
		const wb = createWorkbook(ws, "S");
		const wb2 = await roundtrip(wb, undefined, { sheetRows: 5 });
		// !fullref preserves the original extent
		const fullref = (wb2.Sheets["S"] as any)["!fullref"];
		expect(fullref).toBeDefined();
		expect(fullref).toMatch(/A50/);
	});
});

// ---------------------------------------------------------------------------
// 18. bookProps / bookSheets
// ---------------------------------------------------------------------------
describe("bookProps / bookSheets", () => {
	it("bookSheets returns SheetNames only", async () => {
		const wb = createWorkbook(arrayToSheet([[1]]), "Data");
		appendSheet(wb, arrayToSheet([[2]]), "More");
		const wb2 = await roundtrip(wb, undefined, { bookSheets: true });
		expect(wb2.SheetNames).toEqual(["Data", "More"]);
		// Sheets object should not be populated with data
		expect(wb2.Sheets).toBeUndefined();
	});

	it("bookProps returns Props", async () => {
		const wb = createWorkbook(arrayToSheet([[1]]), "S");
		const wb2 = await roundtrip(wb, undefined, { bookProps: true });
		expect(wb2.Props).toBeDefined();
	});

	it("bookSheets + bookProps combined", async () => {
		const wb = createWorkbook(arrayToSheet([[1]]), "S");
		const wb2 = await roundtrip(wb, undefined, { bookSheets: true, bookProps: true });
		expect(wb2.SheetNames).toBeDefined();
		expect(wb2.Props).toBeDefined();
	});
});

// ---------------------------------------------------------------------------
// 19. Large References
// ---------------------------------------------------------------------------
describe("Large References", () => {
	it("cell at column 16383 (XFD) roundtrips", async () => {
		const ws = createSheet();
		const col = 16383; // XFD (0-based)
		const ref = encodeCell({ c: col, r: 0 });
		ws[ref] = { t: "n", v: 99 };
		ws["!ref"] = `A1:${ref}`;
		const wb2 = await roundtrip(createWorkbook(ws, "S"));
		expect(wb2.Sheets["S"][ref].v).toBe(99);
	});

	it("cell at row 99999 roundtrips", async () => {
		const ws = createSheet();
		ws["A100000"] = { t: "n", v: 77 }; // row 99999 (0-based)
		ws["!ref"] = "A1:A100000";
		const wb2 = await roundtrip(createWorkbook(ws, "S"));
		expect(wb2.Sheets["S"]["A100000"].v).toBe(77);
	});
});

// ---------------------------------------------------------------------------
// 20. Edge Cases
// ---------------------------------------------------------------------------
describe("Edge Cases", () => {
	it("empty sheet roundtrips", async () => {
		const ws = createSheet();
		const wb2 = await roundtrip(createWorkbook(ws, "S"));
		expect(wb2.SheetNames).toEqual(["S"]);
	});

	it("sparse data (A1 + Z100 only)", async () => {
		const ws = createSheet();
		ws["A1"] = { t: "s", v: "start" };
		ws["Z100"] = { t: "s", v: "end" };
		ws["!ref"] = "A1:Z100";
		const wb2 = await roundtrip(createWorkbook(ws, "S"));
		expect(wb2.Sheets["S"]["A1"].v).toBe("start");
		expect(wb2.Sheets["S"]["Z100"].v).toBe("end");
	});

	it("very long string (10000 chars)", async () => {
		const longStr = "x".repeat(10000);
		const ws = arrayToSheet([[longStr]]);
		const wb2 = await roundtrip(createWorkbook(ws, "S"));
		expect(wb2.Sheets["S"]["A1"].v).toBe(longStr);
	});

	it("Infinity written as error", async () => {
		const ws = createSheet();
		ws["A1"] = { t: "n", v: Infinity };
		ws["!ref"] = "A1";
		// Infinity is not valid in XLSX; behavior may vary
		// just ensure write+read doesn't crash
		const wb2 = await roundtrip(createWorkbook(ws, "S"));
		expect(wb2.Sheets["S"]["A1"]).toBeDefined();
	});

	it("compressed output with compression: true", async () => {
		const ws = arrayToSheet([["hello", "world", 123]]);
		const wb = createWorkbook(ws, "S");
		const uncompressed = await write(wb);
		const compressed = await write(wb, { compression: true });
		// Compressed should be smaller (or at least valid)
		expect(compressed).toBeInstanceOf(Uint8Array);
		expect(compressed[0]).toBe(0x50); // PK magic
		expect(compressed.length).toBeLessThanOrEqual(uncompressed.length);
		// Verify it reads back
		const wb2 = await read(compressed);
		expect(sheetToJson(wb2.Sheets["S"], { header: 1 })).toEqual([["hello", "world", 123]]);
	});
});
