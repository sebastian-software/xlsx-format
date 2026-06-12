import { describe, it, expect } from "vitest";
import { write, arrayToSheet, createWorkbook } from "./index.js";

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
