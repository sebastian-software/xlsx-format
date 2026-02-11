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
} from "../src/api/book.js";
import { arrayToSheet } from "../src/api/aoa.js";
import type { CellObject } from "../src/types.js";

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
