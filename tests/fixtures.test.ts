import { describe, it, expect } from "vitest";
import * as fs from "node:fs";
import * as path from "node:path";
import { read, csvToSheet, sheetToJson } from "../src/index.js";

const fixturesDir = path.join(import.meta.dirname, "fixtures");
const csvDir = path.join(fixturesDir, "csv");
const xlsxDir = path.join(fixturesDir, "xlsx");

/** Read a CSV fixture and return header:1 array-of-arrays via csvToSheet */
function loadCsvAsAoA(name: string): any[][] {
	const text = fs.readFileSync(path.join(csvDir, name), "utf-8");
	const ws = csvToSheet(text);
	return sheetToJson(ws, { header: 1 });
}

/** Read an XLSX fixture and return header:1 array-of-arrays via read */
async function loadXlsxAsAoA(name: string): Promise<any[][]> {
	const data = fs.readFileSync(path.join(xlsxDir, name));
	const wb = await read(new Uint8Array(data));
	const ws = wb.Sheets[wb.SheetNames[0]];
	return sheetToJson(ws, { header: 1 });
}

/** Read an XLSX fixture and return the raw worksheet */
async function loadXlsxSheet(name: string) {
	const data = fs.readFileSync(path.join(xlsxDir, name));
	const wb = await read(new Uint8Array(data));
	return wb.Sheets[wb.SheetNames[0]];
}

describe("fixture roundtrip: basic-types", () => {
	it("should produce identical data from CSV and XLSX", async () => {
		const expected = loadCsvAsAoA("basic-types.csv");
		const actual = await loadXlsxAsAoA("basic-types.xlsx");
		expect(actual).toEqual(expected);
	});

	it("should have correct cell types", async () => {
		const ws = await loadXlsxSheet("basic-types.xlsx");
		// Row 2 = first data row (Alice, 30, 95.5, TRUE)
		expect(ws["A2"].t).toBe("s"); // string
		expect(ws["B2"].t).toBe("n"); // number
		expect(ws["C2"].t).toBe("n"); // number
		expect(ws["D2"].t).toBe("b"); // boolean
	});

	it("should have correct values", async () => {
		const ws = await loadXlsxSheet("basic-types.xlsx");
		expect(ws["A2"].v).toBe("Alice");
		expect(ws["B2"].v).toBe(30);
		expect(ws["C2"].v).toBe(95.5);
		expect(ws["D2"].v).toBe(true);
		expect(ws["D3"].v).toBe(false);
	});
});

describe("fixture roundtrip: unicode-data", () => {
	it("should produce identical data from CSV and XLSX", async () => {
		const expected = loadCsvAsAoA("unicode-data.csv");
		const actual = await loadXlsxAsAoA("unicode-data.xlsx");
		expect(actual).toEqual(expected);
	});

	it("should preserve CJK characters", async () => {
		const ws = await loadXlsxSheet("unicode-data.xlsx");
		expect(ws["A2"].v).toBe("田中太郎");
		expect(ws["B2"].v).toBe("東京");
		expect(ws["C2"].v).toBe("こんにちは");
	});

	it("should preserve accented characters", async () => {
		const ws = await loadXlsxSheet("unicode-data.xlsx");
		expect(ws["A3"].v).toBe("Müller");
		expect(ws["B3"].v).toBe("München");
		expect(ws["C3"].v).toBe("Grüße");
	});

	it("should preserve special Unicode characters", async () => {
		const ws = await loadXlsxSheet("unicode-data.xlsx");
		// José García
		expect(ws["A4"].v).toBe("José García");
		// Icelandic
		expect(ws["B5"].v).toBe("Reykjavík");
		// Polish
		expect(ws["A6"].v).toBe("Łukasz");
		expect(ws["B6"].v).toBe("Gdańsk");
	});
});

describe("fixture roundtrip: edge-cases", () => {
	it("should produce identical data from CSV and XLSX", async () => {
		const expected = loadCsvAsAoA("edge-cases.csv");
		const actual = await loadXlsxAsAoA("edge-cases.xlsx");
		expect(actual).toEqual(expected);
	});

	it("should preserve commas in quoted fields", async () => {
		const ws = await loadXlsxSheet("edge-cases.xlsx");
		expect(ws["A3"].v).toBe("Has, comma");
		expect(ws["C3"].v).toBe("Also, has comma");
	});

	it("should preserve empty string cells", async () => {
		const ws = await loadXlsxSheet("edge-cases.xlsx");
		// Row 4 has empty label, empty value, empty notes, "data"
		expect(ws["A4"].v).toBe("");
		expect(ws["C4"].v).toBe("");
		expect(ws["D4"].v).toBe("data");
	});

	it("should preserve escaped quotes", async () => {
		const ws = await loadXlsxSheet("edge-cases.xlsx");
		expect(ws["A5"].v).toBe('She said "hi"');
		expect(ws["C5"].v).toBe('Quotes "inside"');
		expect(ws["D5"].v).toBe('"quoted"');
	});

	it("should preserve long strings", async () => {
		const ws = await loadXlsxSheet("edge-cases.xlsx");
		const longValue = ws["A6"].v as string;
		expect(longValue).toContain("ABCDEFGHIJKLMNOPQRSTUVWXYZ");
		expect(longValue.length).toBeGreaterThan(100);
	});
});

describe("fixture roundtrip: large-dataset", () => {
	it("should produce identical data from CSV and XLSX", async () => {
		const expected = loadCsvAsAoA("large-dataset.csv");
		const actual = await loadXlsxAsAoA("large-dataset.xlsx");
		expect(actual).toEqual(expected);
	});

	it("should have 1001 rows (header + 1000 data)", async () => {
		const rows = await loadXlsxAsAoA("large-dataset.xlsx");
		expect(rows.length).toBe(1001);
	});

	it("should have correct first data row", async () => {
		const rows = await loadXlsxAsAoA("large-dataset.xlsx");
		// Row 0 = header [ID, Name, Value, Flag]
		expect(rows[0]).toEqual(["ID", "Name", "Value", "Flag"]);
		// Row 1 = first data (ID=1)
		expect(rows[1][0]).toBe(1);
		expect(typeof rows[1][1]).toBe("string");
		expect(typeof rows[1][2]).toBe("number");
		expect(typeof rows[1][3]).toBe("boolean");
	});

	it("should have correct row 500", async () => {
		const rows = await loadXlsxAsAoA("large-dataset.xlsx");
		// Row index 500 = data row with ID=500
		expect(rows[500][0]).toBe(500);
	});

	it("should have correct last row", async () => {
		const rows = await loadXlsxAsAoA("large-dataset.xlsx");
		// Last row = ID=1000
		expect(rows[1000][0]).toBe(1000);
	});
});
