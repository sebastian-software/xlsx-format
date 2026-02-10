import { describe, it, expect } from "vitest";
import {
	read,
	write,
	book_new,
	book_append_sheet,
	aoa_to_sheet,
	json_to_sheet,
	sheet_to_json,
	sheet_to_csv,
	sheet_to_html,
	sheet_add_aoa,
	sheet_new,
} from "../src/index.js";

describe("roundtrip", () => {
	it("should write and re-read a simple workbook", () => {
		const data = [
			["Name", "Age", "Active"],
			["Alice", 30, true],
			["Bob", 25, false],
			["Charlie", 35, true],
		];
		const ws = aoa_to_sheet(data);
		const wb = book_new();
		book_append_sheet(wb, ws, "People");

		const u8 = write(wb);
		expect(u8).toBeInstanceOf(Uint8Array);
		expect(u8.length).toBeGreaterThan(0);

		// PK magic bytes
		expect(u8[0]).toBe(0x50);
		expect(u8[1]).toBe(0x4b);

		const wb2 = read(u8);
		expect(wb2.SheetNames).toEqual(["People"]);

		const rows = sheet_to_json(wb2.Sheets["People"]);
		expect(rows.length).toBe(3);
		expect(rows[0]["Name"]).toBe("Alice");
		expect(rows[0]["Age"]).toBe(30);
		expect(rows[0]["Active"]).toBe(true);
		expect(rows[1]["Name"]).toBe("Bob");
		expect(rows[2]["Name"]).toBe("Charlie");
	});

	it("should roundtrip multiple sheets", () => {
		const wb = book_new();
		book_append_sheet(
			wb,
			aoa_to_sheet([
				["A", "B"],
				[1, 2],
			]),
			"Sheet1",
		);
		book_append_sheet(
			wb,
			aoa_to_sheet([
				["X", "Y"],
				[3, 4],
			]),
			"Sheet2",
		);

		const wb2 = read(write(wb));
		expect(wb2.SheetNames).toEqual(["Sheet1", "Sheet2"]);

		const rows1 = sheet_to_json(wb2.Sheets["Sheet1"]);
		expect(rows1[0]["A"]).toBe(1);
		const rows2 = sheet_to_json(wb2.Sheets["Sheet2"]);
		expect(rows2[0]["X"]).toBe(3);
	});

	it("should roundtrip JSON data", () => {
		const data = [
			{ name: "Alpha", value: 100 },
			{ name: "Beta", value: 200 },
			{ name: "Gamma", value: 300 },
		];
		const ws = json_to_sheet(data);
		const wb = book_new();
		book_append_sheet(wb, ws, "Data");

		const wb2 = read(write(wb));
		const rows = sheet_to_json(wb2.Sheets["Data"]);
		expect(rows.length).toBe(3);
		expect(rows[0]["name"]).toBe("Alpha");
		expect(rows[0]["value"]).toBe(100);
		expect(rows[2]["name"]).toBe("Gamma");
	});

	it("should write base64 and re-read", () => {
		const ws = aoa_to_sheet([["Hello", "World"]]);
		const wb = book_new();
		book_append_sheet(wb, ws, "Test");

		const b64 = write(wb, { type: "base64" });
		expect(typeof b64).toBe("string");

		const wb2 = read(b64, { type: "base64" });
		expect(wb2.SheetNames).toEqual(["Test"]);
	});
});

describe("sheet_to_csv", () => {
	it("should produce CSV from worksheet", () => {
		const ws = aoa_to_sheet([
			["Name", "Score"],
			["Alice", 95],
			["Bob", 87],
		]);
		const csv = sheet_to_csv(ws);
		expect(csv).toContain("Name,Score");
		expect(csv).toContain("Alice,95");
		expect(csv).toContain("Bob,87");
	});

	it("should handle custom separators", () => {
		const ws = aoa_to_sheet([
			["A", "B"],
			[1, 2],
		]);
		const tsv = sheet_to_csv(ws, { FS: "\t" });
		expect(tsv).toContain("A\tB");
	});
});

describe("sheet_to_html", () => {
	it("should produce HTML table", () => {
		const ws = aoa_to_sheet([
			["X", "Y"],
			[1, 2],
		]);
		const html = sheet_to_html(ws);
		expect(html).toContain("<table");
		expect(html).toContain("<td");
		expect(html).toContain("X");
		expect(html).toContain("1");
	});
});

describe("sheet_to_json options", () => {
	it("should support header: 1 (array of arrays)", () => {
		const ws = aoa_to_sheet([
			["A", "B"],
			[1, 2],
			[3, 4],
		]);
		const aoa = sheet_to_json(ws, { header: 1 });
		expect(aoa).toEqual([
			["A", "B"],
			[1, 2],
			[3, 4],
		]);
	});

	it("should support header: 'A' (column letter keys)", () => {
		const ws = aoa_to_sheet([
			["Name", "Val"],
			["x", 1],
		]);
		const rows = sheet_to_json(ws, { header: "A" });
		expect(rows[0]["A"]).toBe("Name");
		expect(rows[1]["A"]).toBe("x");
	});

	it("should support custom header array", () => {
		const ws = aoa_to_sheet([
			["a", "b"],
			[1, 2],
		]);
		const rows = sheet_to_json(ws, { header: ["col1", "col2"] });
		expect(rows[0]["col1"]).toBe("a");
		expect(rows[1]["col1"]).toBe(1);
	});

	it("should support defval for missing cells", () => {
		const ws = aoa_to_sheet([["A", "B"], [1]]);
		const rows = sheet_to_json(ws, { defval: null });
		expect(rows[0]["B"]).toBe(null);
	});
});

describe("workbook utilities", () => {
	it("should create a new empty workbook", () => {
		const wb = book_new();
		expect(wb.SheetNames).toEqual([]);
		expect(wb.Sheets).toEqual({});
	});

	it("should create a workbook from a sheet", () => {
		const ws = aoa_to_sheet([[1, 2, 3]]);
		const wb = book_new(ws, "Data");
		expect(wb.SheetNames).toEqual(["Data"]);
		expect(wb.Sheets["Data"]).toBe(ws);
	});

	it("should append sheets", () => {
		const wb = book_new();
		book_append_sheet(wb, aoa_to_sheet([[1]]), "A");
		book_append_sheet(wb, aoa_to_sheet([[2]]), "B");
		expect(wb.SheetNames).toEqual(["A", "B"]);
	});

	it("sheet_new should create an empty worksheet", () => {
		const ws = sheet_new();
		expect(ws["!ref"]).toBeUndefined();
	});

	it("sheet_add_aoa should extend a worksheet", () => {
		const ws = aoa_to_sheet([["A"]]);
		sheet_add_aoa(ws, [["B"]], { origin: "A2" });
		const rows = sheet_to_json(ws, { header: 1 });
		expect(rows).toEqual([["A"], ["B"]]);
	});
});

describe("error handling", () => {
	it("should throw on PDF", () => {
		const pdf = new Uint8Array([0x25, 0x50, 0x44, 0x46]);
		expect(() => read(pdf)).toThrow("PDF");
	});

	it("should throw on PNG", () => {
		const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47]);
		expect(() => read(png)).toThrow("PNG");
	});

	it("should throw on unknown format", () => {
		const junk = new Uint8Array([0x00, 0x00, 0x00, 0x00]);
		expect(() => read(junk)).toThrow();
	});

	it("should throw on invalid workbook", () => {
		expect(() => write({} as any)).toThrow();
	});
});
