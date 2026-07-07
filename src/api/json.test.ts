import { describe, it, expect } from "vitest";
import { arrayToSheet, sheetToJson } from "../index.js";
import { addJsonToSheet } from "./json.js";

describe("json.ts — sheetToJson edge cases", () => {
	it("should return [] for null sheet", () => {
		expect(sheetToJson(null as any)).toStrictEqual([]);
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
		expect(rows[0]).toStrictEqual(expect.objectContaining({ Col1: 10, Col2: 20 }));
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
		expect(rows[0]).toStrictEqual(["c"]);
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
		expect(rows.at(-1)).toStrictEqual([2]);
	});

	it("should handle skipHeader", () => {
		const ws = addJsonToSheet(null, [{ a: 1 }, { a: 2 }], { skipHeader: true });
		const rows = sheetToJson(ws, { header: 1 });
		expect(rows[0]).toStrictEqual([1]); // no header row
	});

	it("should handle pre-built cell objects", () => {
		const ws = addJsonToSheet(null, [{ val: { t: "n", v: 42, f: "=6*7" } }]);
		const cell = (ws as any).A2;
		expect(cell.v).toBe(42);
		expect(cell.f).toBe("=6*7");
	});

	it("should handle Date values", () => {
		const d = new Date("2024-01-15T00:00:00Z");
		const ws = addJsonToSheet(null, [{ date: d }], { UTC: true });
		const cell = (ws as any).A2;
		expect(cell.t).toBe("n");
		expect(cell.v).toBeGreaterThan(40000);
	});

	it("should handle nullError", () => {
		const ws = addJsonToSheet(null, [{ val: null }], { nullError: true });
		const cell = (ws as any).A2;
		expect(cell.t).toBe("e");
		expect(cell.v).toBe(0);
	});

	it("should handle numeric origin", () => {
		const ws = addJsonToSheet(null, [{ a: 1 }], { origin: 5 });
		// origin=5 means header at row 5, data at row 6. Range includes row 0 start.
		expect(ws["!ref"]).toContain("7"); // row 7 = originRow(5) + 1 data row + header offset
	});

	it("should update existing cells in place", () => {
		const ws = addJsonToSheet(null, [{ a: 1 }]);
		const ws2 = addJsonToSheet(ws, [{ a: 999 }]);
		const cell = (ws2 as any).A2;
		expect(cell.v).toBe(999);
	});
});
