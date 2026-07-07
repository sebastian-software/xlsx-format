import { describe, it, expect } from "vitest";
import { read, write, arrayToSheet, sheetToJson, createWorkbook } from "./index.js";

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
		const arr = [...u8];
		const result = await read(arr);
		expect(result.SheetNames).toContain("Sheet1");
	});
});
