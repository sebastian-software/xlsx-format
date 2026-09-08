import { afterEach, describe, it, expect, vi } from "vitest";
import { write, arrayToSheet, createWorkbook } from "./index.js";

describe("write.ts — output types", () => {
	const simpleWb = () => createWorkbook(arrayToSheet([["A"]]), "Sheet1");

	afterEach(() => {
		vi.unstubAllGlobals();
	});

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

	it("should fall back to Uint8Array for buffer output outside Node.js", async () => {
		vi.stubGlobal("Buffer", undefined);
		const bytes = await write(simpleWb(), { bookType: "csv", type: "buffer" });
		expect(bytes).toBeInstanceOf(Uint8Array);
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

	it('should preserve XLSX bytes when type is "string"', async () => {
		const bytes = await write(simpleWb(), { type: "string" });
		expect(bytes).toBeInstanceOf(Uint8Array);
	});

	it("should reject password-protected writes before validation", async () => {
		await expect(write({} as any, { password: "secret" })).rejects.toMatchObject({
			name: "XlsxError",
			code: "UNSUPPORTED",
			message: "Password-protected workbooks are not supported",
		});
	});

	it("should preserve empty password behavior", async () => {
		await expect(write(simpleWb(), { bookType: "csv", type: "string", password: "" })).resolves.toContain("A");
	});

	it("should reject unsupported write options", async () => {
		await expect(write(simpleWb(), { bookVBA: true })).rejects.toMatchObject({
			code: "UNSUPPORTED",
			message: 'Write option "bookVBA" is not supported',
		});
		await expect(write(simpleWb(), { themeXLSX: "<theme/>" })).rejects.toMatchObject({
			code: "UNSUPPORTED",
			message: 'Write option "themeXLSX" is not supported',
		});
	});

	it.each([
		["string", "macro"],
		["array", [1]],
		["ArrayBuffer", new Uint8Array([1]).buffer],
		["DataView", new DataView(new Uint8Array([1]).buffer)],
		["Uint8Array", new Uint8Array([1])],
	])("should reject non-empty %s VBA payloads instead of dropping them", async (_kind, payload) => {
		const wb = simpleWb();
		wb.vbaraw = payload;
		await expect(write(wb, { bookType: "xlsm" })).rejects.toMatchObject({
			code: "UNSUPPORTED",
			message: "Workbooks containing VBA data cannot be written",
		});
	});

	it.each([
		["string", ""],
		["array", []],
		["ArrayBuffer", new ArrayBuffer(0)],
		["DataView", new DataView(new ArrayBuffer(0))],
	])("should preserve empty %s VBA compatibility payloads", async (_kind, payload) => {
		const wb = simpleWb();
		wb.vbaraw = payload;
		await expect(write(wb)).resolves.toBeInstanceOf(Uint8Array);
	});

	it.each([null, undefined, {}])(
		"should classify an invalid workbook before reading optional VBA metadata",
		async (wb) => {
			await expect(write(wb as any)).rejects.toMatchObject({
				name: "XlsxError",
				code: "INVALID_ARGUMENT",
				message: "Invalid Workbook",
			});
		},
	);

	it("should preserve empty compatibility values", async () => {
		const wb = simpleWb();
		wb.vbaraw = new Uint8Array();
		await expect(write(wb, { bookVBA: false, themeXLSX: "" })).resolves.toBeInstanceOf(Uint8Array);
	});

	it("should write empty workbook CSV", async () => {
		const emptyWb = { SheetNames: ["S1"], Sheets: { S1: {} } } as any;
		const csv = await write(emptyWb, { bookType: "csv", type: "string" });
		expect(csv).toBe("");
	});

	it("should pass CSV export options through write", async () => {
		const wb = createWorkbook(arrayToSheet([["=1+1"]]), "Sheet1");

		await expect(write(wb, { bookType: "csv", type: "string" })).resolves.toBe("'=1+1");
		await expect(write(wb, { bookType: "csv", type: "string", escapeFormulae: false })).resolves.toBe("=1+1");
	});

	it("should pass HTML export options through write", async () => {
		const ws = arrayToSheet([["Bad"]]);
		ws.A1.l = { Target: "javascript:alert(1)" };
		const wb = createWorkbook(ws, "Sheet1");

		await expect(write(wb, { bookType: "html", type: "string" })).resolves.not.toContain("href=");
		await expect(write(wb, { bookType: "html", type: "string", sanitizeLinks: false })).resolves.toContain(
			'href="javascript:alert(1)"',
		);
	});
});
