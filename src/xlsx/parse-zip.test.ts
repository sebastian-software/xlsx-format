import { describe, it, expect } from "vitest";
import { read, write, createWorkbook, jsonToSheet } from "../index.js";

describe("parse-zip: bookSheets and bookProps options", () => {
	it("reads with bookSheets to get only sheet names", async () => {
		const ws = jsonToSheet([{ a: 1 }, { a: 2 }]);
		const wb = createWorkbook(ws, "TestSheet");
		const buf = await write(wb);
		const result = await read(buf, { bookSheets: true });
		expect(result.SheetNames).toBeDefined();
		expect(result.SheetNames).toContain("TestSheet");
		// Should not have full sheet data
	});

	it("reads with bookProps to get document properties", async () => {
		const ws = jsonToSheet([{ a: 1 }]);
		const wb = createWorkbook(ws, "S1");
		const buf = await write(wb);
		const result = await read(buf, { bookProps: true });
		expect(result.Props).toBeDefined();
	});

	it("reads with specific sheet by index", async () => {
		const wb = createWorkbook(jsonToSheet([{ a: 1 }]), "S1");
		const ws2 = jsonToSheet([{ b: 2 }]);
		wb.SheetNames.push("S2");
		wb.Sheets.S2 = ws2;
		const buf = await write(wb);
		const result = await read(buf, { sheets: 0 });
		expect(result.SheetNames).toContain("S1");
		// S2 should exist in SheetNames but may have empty data
	});

	it("reads with specific sheet by name", async () => {
		const wb = createWorkbook(jsonToSheet([{ a: 1 }]), "S1");
		const ws2 = jsonToSheet([{ b: 2 }]);
		wb.SheetNames.push("S2");
		wb.Sheets.S2 = ws2;
		const buf = await write(wb);
		const result = await read(buf, { sheets: "S2" });
		expect(result.SheetNames).toBeDefined();
	});

	it("reads with sheets as array of indices and names", async () => {
		const wb = createWorkbook(jsonToSheet([{ a: 1 }]), "S1");
		wb.SheetNames.push("S2");
		wb.Sheets.S2 = jsonToSheet([{ b: 2 }]);
		wb.SheetNames.push("S3");
		wb.Sheets.S3 = jsonToSheet([{ c: 3 }]);
		const buf = await write(wb);
		const result = await read(buf, { sheets: [0, "S3"] });
		expect(result.SheetNames).toBeDefined();
	});
});

describe("parse-zip: dense mode and cellStyles", () => {
	it("reads in dense mode", async () => {
		const ws = jsonToSheet([{ a: 1, b: "text" }]);
		const wb = createWorkbook(ws, "S1");
		const buf = await write(wb);
		const result = await read(buf, { dense: true });
		const sheet = result.Sheets.S1;
		expect(sheet["!data"]).toBeDefined();
	});

	it("reads with cellStyles to populate column info", async () => {
		const ws = jsonToSheet([{ a: 1, b: 2 }]);
		ws["!cols"] = [{ width: 15 }, { width: 20 }];
		const wb = createWorkbook(ws, "S1");
		const buf = await write(wb);
		const result = await read(buf, { cellStyles: true });
		const sheet = result.Sheets.S1;
		expect(sheet["!cols"]).toBeDefined();
	});
});
