import { describe, it, expect } from "vitest";
import {
	createWorkbook,
	appendSheet,
	createSheet,
	getSheetIndex,
	setSheetVisibility,
	setCellNumberFormat,
	setCellHyperlink,
	setCellInternalLink,
	addCellComment,
	setArrayFormula,
	sheetToFormulae,
} from "./book.js";
import { arrayToSheet } from "./aoa.js";
import type { CellObject } from "../types.js";

import { jsonToSheet } from "../index.js";

describe("createWorkbook", () => {
	it("should create an empty workbook", () => {
		const wb = createWorkbook();
		expect(wb.SheetNames).toEqual([]);
		expect(wb.Sheets).toEqual({});
	});

	it("should create a workbook with an initial sheet", () => {
		const ws = createSheet();
		const wb = createWorkbook(ws, "Data");
		expect(wb.SheetNames).toEqual(["Data"]);
		expect(wb.Sheets["Data"]).toBe(ws);
	});

	it("should default sheet name to Sheet1", () => {
		const ws = createSheet();
		const wb = createWorkbook(ws);
		expect(wb.SheetNames).toEqual(["Sheet1"]);
	});
});

describe("appendSheet", () => {
	it("should auto-generate sheet names", () => {
		const wb = createWorkbook();
		appendSheet(wb, createSheet());
		appendSheet(wb, createSheet());
		expect(wb.SheetNames).toEqual(["Sheet1", "Sheet2"]);
	});

	it("should throw on duplicate name", () => {
		const wb = createWorkbook();
		appendSheet(wb, createSheet(), "Test");
		expect(() => appendSheet(wb, createSheet(), "Test")).toThrow("already exists");
	});

	it("should roll names when roll=true", () => {
		const wb = createWorkbook();
		appendSheet(wb, createSheet(), "Sheet1");
		const name = appendSheet(wb, createSheet(), "Sheet1", true);
		expect(name).toBe("Sheet2");
	});

	it("should return the used sheet name", () => {
		const wb = createWorkbook();
		const name = appendSheet(wb, createSheet(), "MySheet");
		expect(name).toBe("MySheet");
	});
});

describe("getSheetIndex", () => {
	it("should find sheet by name", () => {
		const wb = createWorkbook();
		appendSheet(wb, createSheet(), "A");
		appendSheet(wb, createSheet(), "B");
		expect(getSheetIndex(wb, "A")).toBe(0);
		expect(getSheetIndex(wb, "B")).toBe(1);
	});

	it("should find sheet by index", () => {
		const wb = createWorkbook();
		appendSheet(wb, createSheet(), "A");
		expect(getSheetIndex(wb, 0)).toBe(0);
	});

	it("should throw for unknown name", () => {
		const wb = createWorkbook();
		expect(() => getSheetIndex(wb, "X")).toThrow("Cannot find sheet");
	});

	it("should throw for out-of-range index", () => {
		const wb = createWorkbook();
		expect(() => getSheetIndex(wb, 5)).toThrow("Cannot find sheet");
	});
});

describe("setSheetVisibility", () => {
	it("should set visibility on a sheet", () => {
		const wb = createWorkbook();
		appendSheet(wb, createSheet(), "A");
		setSheetVisibility(wb, "A", 1);
		expect(wb.Workbook!.Sheets![0].Hidden).toBe(1);
	});

	it("should throw for invalid visibility value", () => {
		const wb = createWorkbook();
		appendSheet(wb, createSheet(), "A");
		expect(() => {
			setSheetVisibility(wb, "A", 3 as any);
		}).toThrow("Bad sheet visibility");
	});
});

describe("setCellNumberFormat", () => {
	it("should set z property on cell", () => {
		const cell: CellObject = { t: "n", v: 42 };
		setCellNumberFormat(cell, "#,##0.00");
		expect(cell.z).toBe("#,##0.00");
	});
});

describe("setCellHyperlink", () => {
	it("should set hyperlink on cell", () => {
		const cell: CellObject = { t: "s", v: "click" };
		setCellHyperlink(cell, "https://example.com", "Example");
		expect(cell.l!.Target).toBe("https://example.com");
		expect(cell.l!.Tooltip).toBe("Example");
	});

	it("should remove hyperlink when target is falsy", () => {
		const cell: CellObject = { t: "s", v: "click", l: { Target: "https://example.com" } };
		setCellHyperlink(cell);
		expect(cell.l).toBeUndefined();
	});
});

describe("setCellInternalLink", () => {
	it("should prefix target with #", () => {
		const cell: CellObject = { t: "s", v: "link" };
		setCellInternalLink(cell, "Sheet2!A1");
		expect(cell.l!.Target).toBe("#Sheet2!A1");
	});
});

describe("addCellComment", () => {
	it("should add a comment to a cell", () => {
		const cell: CellObject = { t: "s", v: "test" };
		addCellComment(cell, "review this", "Alice");
		expect(cell.c).toHaveLength(1);
		expect(cell.c![0].t).toBe("review this");
		expect(cell.c![0].a).toBe("Alice");
	});

	it("should default author to SheetJS", () => {
		const cell: CellObject = { t: "s", v: "test" };
		addCellComment(cell, "note");
		expect(cell.c![0].a).toBe("SheetJS");
	});

	it("should append multiple comments", () => {
		const cell: CellObject = { t: "s", v: "test" };
		addCellComment(cell, "first");
		addCellComment(cell, "second");
		expect(cell.c).toHaveLength(2);
	});
});

describe("setArrayFormula", () => {
	it("should set formula on top-left and F on all cells", () => {
		const ws = arrayToSheet([
			[1, 2],
			[3, 4],
		]);
		setArrayFormula(ws, "C1:C2", "=A1:A2*B1:B2");
		expect((ws as any)["C1"].f).toBe("=A1:A2*B1:B2");
		expect((ws as any)["C1"].F).toBe("C1:C2");
		expect((ws as any)["C2"].F).toBe("C1:C2");
		expect((ws as any)["C2"].f).toBeUndefined();
	});

	it("should set formula cells in dense worksheets", () => {
		const ws = arrayToSheet(
			[
				[1, 2],
				[3, 4],
			],
			{ dense: true },
		);

		setArrayFormula(ws, "C1:C2", "=A1:A2*B1:B2");

		expect(ws["!data"]![0]![2]!.f).toBe("=A1:A2*B1:B2");
		expect(ws["!data"]![0]![2]!.F).toBe("C1:C2");
		expect(ws["!data"]![1]![2]!.F).toBe("C1:C2");
		expect(ws["!data"]![1]![2]!.f).toBeUndefined();
	});
});

describe("sheetToFormulae", () => {
	it("should return empty array for null sheet", () => {
		expect(sheetToFormulae(null as any)).toEqual([]);
	});

	it("should return cell values as ref=value strings", () => {
		const ws = arrayToSheet([
			["Name", "Age"],
			["Alice", 30],
		]);
		const formulae = sheetToFormulae(ws);
		expect(formulae).toContain("A1='Name");
		expect(formulae).toContain("B1='Age");
		expect(formulae).toContain("A2='Alice");
		expect(formulae).toContain("B2=30");
	});
});

describe("book.ts: API edge cases", () => {
	it("appendSheet auto-generates name", () => {
		const wb = createWorkbook();
		appendSheet(wb, arrayToSheet([["A"]]));
		expect(wb.SheetNames).toContain("Sheet1");
		appendSheet(wb, arrayToSheet([["B"]]));
		expect(wb.SheetNames).toContain("Sheet2");
	});

	it("appendSheet with roll option", () => {
		const wb = createWorkbook(arrayToSheet([["A"]]), "Sheet1");
		const name = appendSheet(wb, arrayToSheet([["B"]]), "Sheet1", true);
		expect(name).toBe("Sheet2");
		expect(wb.SheetNames).toContain("Sheet2");
	});

	it("appendSheet rejects duplicate without roll", () => {
		const wb = createWorkbook(arrayToSheet([["A"]]), "Sheet1");
		expect(() => appendSheet(wb, arrayToSheet([["B"]]), "Sheet1")).toThrow("already exists");
	});

	it("getSheetIndex finds by name", () => {
		const wb = createWorkbook(arrayToSheet([["A"]]), "MySheet");
		expect(getSheetIndex(wb, "MySheet")).toBe(0);
	});

	it("getSheetIndex finds by index", () => {
		const wb = createWorkbook(arrayToSheet([["A"]]), "Sheet1");
		expect(getSheetIndex(wb, 0)).toBe(0);
	});

	it("getSheetIndex throws for missing name", () => {
		const wb = createWorkbook(arrayToSheet([["A"]]), "Sheet1");
		expect(() => getSheetIndex(wb, "NoSuchSheet")).toThrow("Cannot find");
	});

	it("getSheetIndex throws for out-of-range index", () => {
		const wb = createWorkbook(arrayToSheet([["A"]]), "Sheet1");
		expect(() => getSheetIndex(wb, 99)).toThrow("Cannot find");
	});

	it("setSheetVisibility rejects invalid value", () => {
		const wb = createWorkbook(arrayToSheet([["A"]]), "Sheet1");
		expect(() => {
			setSheetVisibility(wb, 0, 99 as any);
		}).toThrow("Bad sheet visibility");
	});

	it("setSheetVisibility initializes Workbook metadata", () => {
		const wb = createWorkbook(arrayToSheet([["A"]]), "Sheet1");
		setSheetVisibility(wb, 0, 1);
		expect(wb.Workbook).toBeDefined();
		expect(wb.Workbook!.Sheets![0].Hidden).toBe(1);
	});

	it("createSheet creates empty sheet", () => {
		const ws = createSheet();
		expect(ws).toBeDefined();
	});

	it("createSheet dense mode", () => {
		const ws = createSheet({ dense: true });
		expect(ws["!data"]).toBeDefined();
	});

	it("setCellNumberFormat sets z property", () => {
		const cell: any = { t: "n", v: 42 };
		setCellNumberFormat(cell, "#,##0.00");
		expect(cell.z).toBe("#,##0.00");
	});

	it("setCellHyperlink sets link", () => {
		const cell: any = { t: "s", v: "Click" };
		setCellHyperlink(cell, "https://example.com", "Example");
		expect(cell.l.Target).toBe("https://example.com");
		expect(cell.l.Tooltip).toBe("Example");
	});

	it("setCellHyperlink removes link when falsy", () => {
		const cell: any = { t: "s", v: "Click", l: { Target: "https://example.com" } };
		setCellHyperlink(cell);
		expect(cell.l).toBeUndefined();
	});

	it("setCellInternalLink adds # prefix", () => {
		const cell: any = { t: "s", v: "Click" };
		setCellInternalLink(cell, "Sheet2!A1");
		expect(cell.l.Target).toBe("#Sheet2!A1");
	});

	it("addCellComment adds comment", () => {
		const cell: any = { t: "s", v: "Value" };
		addCellComment(cell, "Check this", "Alice");
		expect(cell.c).toHaveLength(1);
		expect(cell.c[0].t).toBe("Check this");
		expect(cell.c[0].a).toBe("Alice");
	});

	it("addCellComment uses default author", () => {
		const cell: any = { t: "s", v: "Value" };
		addCellComment(cell, "Note");
		expect(cell.c[0].a).toBe("SheetJS");
	});

	it("sheetToFormulae extracts formulas", () => {
		const ws: any = {
			A1: { t: "n", v: 42 },
			A2: { t: "s", v: "Hello" },
			A3: { t: "b", v: true },
			A4: { t: "n", v: 10, f: "SUM(A1:A1)" },
			"!ref": "A1:A4",
		};
		const formulae = sheetToFormulae(ws);
		expect(formulae.length).toBeGreaterThanOrEqual(4);
		expect(formulae).toContain("A1=42");
		expect(formulae).toContain("A2='Hello");
		expect(formulae).toContain("A3=TRUE");
		expect(formulae).toContain("A4=SUM(A1:A1)");
	});

	it("sheetToFormulae handles array formulas", () => {
		const ws = arrayToSheet([[1], [2], [3]]);
		setArrayFormula(ws, "B1:B3", "A1:A3*2");
		const formulae = sheetToFormulae(ws);
		const arrayFmla = formulae.find((f) => f.includes("A1:A3*2"));
		expect(arrayFmla).toBeDefined();
		expect(arrayFmla).toContain("B1:B3");
	});

	it("sheetToFormulae with dense mode", () => {
		const ws = arrayToSheet([["A", 1]], { dense: true });
		const formulae = sheetToFormulae(ws);
		expect(formulae.length).toBeGreaterThanOrEqual(2);
	});

	it("sheetToFormulae returns empty for null sheet", () => {
		expect(sheetToFormulae(null as any)).toEqual([]);
		expect(sheetToFormulae({} as any)).toEqual([]);
	});

	it("sheetToFormulae with w property", () => {
		const ws: any = {
			A1: { t: "n", v: 42, w: "42.00" },
			"!ref": "A1:A1",
		};
		// When v is present, it uses v for numeric
		const formulae = sheetToFormulae(ws);
		expect(formulae).toContain("A1=42");
	});

	it("setArrayFormula dense mode", () => {
		const ws = createSheet({ dense: true });
		ws["!ref"] = "A1:A1";
		(ws as any)["!data"] = [[{ t: "n", v: 1 }]];
		setArrayFormula(ws, "B1:B2", "A1*2", true);
		expect((ws as any)["!data"][0][1].f).toBe("A1*2");
		expect((ws as any)["!data"][0][1].D).toBe(true);
	});

	it("setArrayFormula expands ref", () => {
		const ws = arrayToSheet([[1]]);
		setArrayFormula(ws, "C5:D6", "A1*2");
		const ref = ws["!ref"];
		expect(ref).toContain("D6");
	});
});

describe("book.ts: additional API coverage", () => {
	it("sheetToFormulae returns formula strings", () => {
		const ws: any = {};
		ws["!ref"] = "A1:B1";
		ws["A1"] = { t: "n", v: 1 };
		ws["B1"] = { t: "n", v: 2, f: "A1+1" };
		const formulae = sheetToFormulae(ws);
		expect(formulae.some((f: string) => f.includes("A1+1"))).toBe(true);
	});

	it("setArrayFormula sets formula on range", () => {
		const ws: any = {};
		ws["!ref"] = "A1:B2";
		ws["A1"] = { t: "n", v: 1 };
		ws["A2"] = { t: "n", v: 2 };
		setArrayFormula(ws, "B1:B2", "A1:A2*2");
		expect(ws["B1"]?.f).toBe("A1:A2*2");
		expect(ws["B1"]?.F).toBe("B1:B2");
	});

	it("setCellHyperlink sets external link", () => {
		const cell: any = { t: "s", v: "click" };
		setCellHyperlink(cell, "https://example.com");
		expect(cell.l?.Target).toBe("https://example.com");
	});

	it("setCellInternalLink sets internal link", () => {
		const cell: any = { t: "s", v: "go" };
		setCellInternalLink(cell, "Sheet2!A1");
		expect(cell.l?.Target).toBe("#Sheet2!A1");
	});

	it("appendSheet with roll option handles duplicates", () => {
		const wb = createWorkbook(jsonToSheet([{ a: 1 }]), "Sheet1");
		appendSheet(wb, jsonToSheet([{ b: 2 }]), "Sheet1", true);
		expect(wb.SheetNames).toHaveLength(2);
		expect(wb.SheetNames[1]).not.toBe("Sheet1");
	});

	it("addCellComment adds comment to cell", () => {
		const cell: any = { t: "n", v: 1 };
		addCellComment(cell, "A comment", "Author");
		expect(cell.c).toBeDefined();
		expect(cell.c?.[0]?.t).toBe("A comment");
		expect(cell.c?.[0]?.a).toBe("Author");
	});

	it("setCellNumberFormat sets format", () => {
		const cell: any = { t: "n", v: 1.5 };
		setCellNumberFormat(cell, "#,##0.00");
		expect(cell.z).toBe("#,##0.00");
	});
});
