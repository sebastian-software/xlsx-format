import { describe, it, expect } from "vitest";
import { read, write, arrayToSheet, sheetToJson, createWorkbook } from "./index.js";
import { zipAddString, zipRead, zipReadString, zipWrite } from "./zip/index.js";

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

	it("should forward the field separator for plain text", async () => {
		const result = await read("A\tB\n1\t2", { type: "string", FS: "\t" });
		const rows = sheetToJson(result.Sheets[result.SheetNames[0]], { header: 1 });
		expect(rows).toStrictEqual([
			["A", "B"],
			[1, 2],
		]);
	});

	it("should apply metadata-only shapes to plain-text reads", async () => {
		const sheetNames = await read("A,B", { type: "string", bookSheets: true });
		expect(sheetNames).toStrictEqual({ SheetNames: ["Sheet1"] });

		const properties = await read("A,B", { type: "string", bookProps: true });
		expect(properties).toStrictEqual({ Props: {}, Custprops: {} });

		const combined = await read("A,B", { type: "string", bookSheets: true, bookProps: true });
		expect(combined).toStrictEqual({ SheetNames: ["Sheet1"], Props: {}, Custprops: {} });
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

	it("should reject password-protected reads before parsing", async () => {
		await expect(read(new Uint8Array([0x00]), { password: "secret" })).rejects.toMatchObject({
			name: "XlsxError",
			code: "UNSUPPORTED",
			message: "Password-protected workbooks are not supported",
		});
	});

	it.each(["bookFiles", "bookVBA", "bookDeps", "xlfn"] as const)(
		"should reject unsupported %s requests",
		async (option) => {
			await expect(read(new Uint8Array([0x00]), { [option]: true })).rejects.toMatchObject({
				name: "XlsxError",
				code: "UNSUPPORTED",
				message: `Read option "${option}" is not supported`,
			});
		},
	);

	it("should preserve false compatibility options", async () => {
		const bytes = await write(createWorkbook(arrayToSheet([["Data"]]), "Sheet1"));
		await expect(
			read(bytes, { bookFiles: false, bookVBA: false, bookDeps: false, xlfn: false }),
		).resolves.toMatchObject({ SheetNames: ["Sheet1"] });
	});

	it("should ignore stored dimensions and infer an A1-anchored range with nodim", async () => {
		const bytes = await write(createWorkbook(arrayToSheet([["Data"]]), "Sheet1"));
		const zip = await zipRead(bytes);
		const sheetXml = zipReadString(zip, "xl/worksheets/sheet1.xml");
		zipAddString(zip, "xl/worksheets/sheet1.xml", sheetXml.replace('ref="A1"', 'ref="A1:Z99"'));
		const modified = await zipWrite(zip);

		const stored = await read(modified);
		expect(stored.Sheets.Sheet1["!ref"]).toBe("A1:Z99");

		const inferred = await read(modified, { nodim: true });
		expect(inferred.Sheets.Sheet1["!ref"]).toBe("A1");
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
