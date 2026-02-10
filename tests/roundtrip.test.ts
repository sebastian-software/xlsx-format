import { describe, it, expect } from "vitest";
import {
	read,
	write,
	createWorkbook,
	appendSheet,
	arrayToSheet,
	jsonToSheet,
	sheetToJson,
	sheetToCsv,
	sheetToHtml,
	addArrayToSheet,
	createSheet,
} from "../src/index.js";

describe("roundtrip", () => {
	it("should write and re-read a simple workbook", async () => {
		const data = [
			["Name", "Age", "Active"],
			["Alice", 30, true],
			["Bob", 25, false],
			["Charlie", 35, true],
		];
		const ws = arrayToSheet(data);
		const wb = createWorkbook();
		appendSheet(wb, ws, "People");

		const u8 = await write(wb);
		expect(u8).toBeInstanceOf(Uint8Array);
		expect(u8.length).toBeGreaterThan(0);

		// PK magic bytes
		expect(u8[0]).toBe(0x50);
		expect(u8[1]).toBe(0x4b);

		const wb2 = await read(u8);
		expect(wb2.SheetNames).toEqual(["People"]);

		const rows = sheetToJson(wb2.Sheets["People"]);
		expect(rows.length).toBe(3);
		expect(rows[0]["Name"]).toBe("Alice");
		expect(rows[0]["Age"]).toBe(30);
		expect(rows[0]["Active"]).toBe(true);
		expect(rows[1]["Name"]).toBe("Bob");
		expect(rows[2]["Name"]).toBe("Charlie");
	});

	it("should roundtrip multiple sheets", async () => {
		const wb = createWorkbook();
		appendSheet(
			wb,
			arrayToSheet([
				["A", "B"],
				[1, 2],
			]),
			"Sheet1",
		);
		appendSheet(
			wb,
			arrayToSheet([
				["X", "Y"],
				[3, 4],
			]),
			"Sheet2",
		);

		const wb2 = await read(await write(wb));
		expect(wb2.SheetNames).toEqual(["Sheet1", "Sheet2"]);

		const rows1 = sheetToJson(wb2.Sheets["Sheet1"]);
		expect(rows1[0]["A"]).toBe(1);
		const rows2 = sheetToJson(wb2.Sheets["Sheet2"]);
		expect(rows2[0]["X"]).toBe(3);
	});

	it("should roundtrip JSON data", async () => {
		const data = [
			{ name: "Alpha", value: 100 },
			{ name: "Beta", value: 200 },
			{ name: "Gamma", value: 300 },
		];
		const ws = jsonToSheet(data);
		const wb = createWorkbook();
		appendSheet(wb, ws, "Data");

		const wb2 = await read(await write(wb));
		const rows = sheetToJson(wb2.Sheets["Data"]);
		expect(rows.length).toBe(3);
		expect(rows[0]["name"]).toBe("Alpha");
		expect(rows[0]["value"]).toBe(100);
		expect(rows[2]["name"]).toBe("Gamma");
	});

	it("should write base64 and re-read", async () => {
		const ws = arrayToSheet([["Hello", "World"]]);
		const wb = createWorkbook();
		appendSheet(wb, ws, "Test");

		const b64 = await write(wb, { type: "base64" });
		expect(typeof b64).toBe("string");

		const wb2 = await read(b64, { type: "base64" });
		expect(wb2.SheetNames).toEqual(["Test"]);
	});
});

describe("sheetToCsv", () => {
	it("should produce CSV from worksheet", () => {
		const ws = arrayToSheet([
			["Name", "Score"],
			["Alice", 95],
			["Bob", 87],
		]);
		const csv = sheetToCsv(ws);
		expect(csv).toContain("Name,Score");
		expect(csv).toContain("Alice,95");
		expect(csv).toContain("Bob,87");
	});

	it("should handle custom separators", () => {
		const ws = arrayToSheet([
			["A", "B"],
			[1, 2],
		]);
		const tsv = sheetToCsv(ws, { FS: "\t" });
		expect(tsv).toContain("A\tB");
	});
});

describe("sheetToHtml", () => {
	it("should produce HTML table", () => {
		const ws = arrayToSheet([
			["X", "Y"],
			[1, 2],
		]);
		const html = sheetToHtml(ws);
		expect(html).toContain("<table");
		expect(html).toContain("<td");
		expect(html).toContain("X");
		expect(html).toContain("1");
	});
});

describe("sheetToJson options", () => {
	it("should support header: 1 (array of arrays)", () => {
		const ws = arrayToSheet([
			["A", "B"],
			[1, 2],
			[3, 4],
		]);
		const aoa = sheetToJson(ws, { header: 1 });
		expect(aoa).toEqual([
			["A", "B"],
			[1, 2],
			[3, 4],
		]);
	});

	it("should support header: 'A' (column letter keys)", () => {
		const ws = arrayToSheet([
			["Name", "Val"],
			["x", 1],
		]);
		const rows = sheetToJson(ws, { header: "A" });
		expect(rows[0]["A"]).toBe("Name");
		expect(rows[1]["A"]).toBe("x");
	});

	it("should support custom header array", () => {
		const ws = arrayToSheet([
			["a", "b"],
			[1, 2],
		]);
		const rows = sheetToJson(ws, { header: ["col1", "col2"] });
		expect(rows[0]["col1"]).toBe("a");
		expect(rows[1]["col1"]).toBe(1);
	});

	it("should support defval for missing cells", () => {
		const ws = arrayToSheet([["A", "B"], [1]]);
		const rows = sheetToJson(ws, { defval: null });
		expect(rows[0]["B"]).toBe(null);
	});
});

describe("workbook utilities", () => {
	it("should create a new empty workbook", () => {
		const wb = createWorkbook();
		expect(wb.SheetNames).toEqual([]);
		expect(wb.Sheets).toEqual({});
	});

	it("should create a workbook from a sheet", () => {
		const ws = arrayToSheet([[1, 2, 3]]);
		const wb = createWorkbook(ws, "Data");
		expect(wb.SheetNames).toEqual(["Data"]);
		expect(wb.Sheets["Data"]).toBe(ws);
	});

	it("should append sheets", () => {
		const wb = createWorkbook();
		appendSheet(wb, arrayToSheet([[1]]), "A");
		appendSheet(wb, arrayToSheet([[2]]), "B");
		expect(wb.SheetNames).toEqual(["A", "B"]);
	});

	it("createSheet should create an empty worksheet", () => {
		const ws = createSheet();
		expect(ws["!ref"]).toBeUndefined();
	});

	it("addArrayToSheet should extend a worksheet", () => {
		const ws = arrayToSheet([["A"]]);
		addArrayToSheet(ws, [["B"]], { origin: "A2" });
		const rows = sheetToJson(ws, { header: 1 });
		expect(rows).toEqual([["A"], ["B"]]);
	});
});

describe("error handling", () => {
	it("should throw on PDF", async () => {
		const pdf = new Uint8Array([0x25, 0x50, 0x44, 0x46]);
		await expect(read(pdf)).rejects.toThrow("PDF");
	});

	it("should throw on PNG", async () => {
		const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47]);
		await expect(read(png)).rejects.toThrow("PNG");
	});

	it("should throw on unknown format", async () => {
		const junk = new Uint8Array([0x00, 0x00, 0x00, 0x00]);
		await expect(read(junk)).rejects.toThrow();
	});

	it("should throw on invalid workbook", async () => {
		await expect(write({} as any)).rejects.toThrow();
	});
});
