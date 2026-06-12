import { describe, expect, it } from "vitest";
import {
	appendSheet,
	arrayToSheet,
	createWorkbook,
	read,
	sheetToCsv,
	sheetToHtml,
	sheetToJson,
	write,
} from "./index.js";
import type { WorkSheet } from "./types.js";
import { zipRead } from "./zip/index.js";

const encoder = new TextEncoder();
const decoder = new TextDecoder();

describe("export security guards", () => {
	it("preserves dangerous sheetToJson headers without prototype pollution", () => {
		const ws = arrayToSheet([
			["__proto__", "constructor", "prototype", "safe"],
			["polluted", "ctor", "proto", "ok"],
		]);
		const [row] = sheetToJson<Record<string, unknown>>(ws);

		expect(Object.prototype).not.toHaveProperty("polluted");
		expect(Object.hasOwn(row, "__proto__")).toBe(true);
		expect(Object.hasOwn(row, "constructor")).toBe(true);
		expect(Object.hasOwn(row, "prototype")).toBe(true);
		expect(row["__proto__"]).toBe("polluted");
		expect(row.constructor).toBe("ctor");
		expect(row.prototype).toBe("proto");
		expect(row.safe).toBe("ok");
	});

	it("keeps malicious workbook sheet names as own properties", async () => {
		const source = createWorkbook();
		appendSheet(source, arrayToSheet([["value"]]), "Safe");

		const zip = await zipRead(await write(source));
		zip.files["xl/workbook.xml"] = encoder.encode(
			decoder.decode(zip.files["xl/workbook.xml"]).replace('name="Safe"', 'name="__proto__"'),
		);

		const parsed = await read(await import("./zip/index.js").then(({ zipWrite }) => zipWrite(zip)));

		expect(parsed.SheetNames).toEqual(["__proto__"]);
		expect(Object.getPrototypeOf(parsed.Sheets)).toBeNull();
		expect(Object.hasOwn(parsed.Sheets, "__proto__")).toBe(true);
		expect(Object.prototype).not.toHaveProperty("value");
	});

	it("escapes formula-like CSV fields when requested", () => {
		const ws = arrayToSheet([["=1+1", "+SUM(A1)", "-cmd", "@handle", "plain"]]);

		expect(sheetToCsv(ws)).toBe("=1+1,+SUM(A1),-cmd,@handle,plain");
		expect(sheetToCsv(ws, { escapeFormulae: true })).toBe("'=1+1,'+SUM(A1),'-cmd,'@handle,plain");
	});

	it("escapes formula-only CSV cells when requested", () => {
		const ws = { A1: { t: "z", f: "SUM(B1:B10)" }, "!ref": "A1:A1" } as WorkSheet;

		expect(sheetToCsv(ws, { escapeFormulae: true })).toBe("'=SUM(B1:B10)");
	});

	it("sanitizes javascript links with embedded whitespace", () => {
		const ws = {
			"!ref": "A1",
			A1: { t: "s", v: "Bad", l: { Target: "java\nscript:alert(1)" } },
		} as WorkSheet;

		expect(sheetToHtml(ws, { sanitizeLinks: true })).not.toContain("href=");
	});

	it("clamps oversized declared ranges to occupied cells for exporters", () => {
		const ws = {
			"!ref": "A1:XFD1048576",
			A1: { t: "s", v: "Name" },
			B2: { t: "s", v: "Done" },
		} as WorkSheet;

		expect(sheetToCsv(ws)).toBe("Name,\n,Done");

		const rows = sheetToJson<unknown[]>(ws, { header: 1 });
		expect(rows).toHaveLength(2);
		expect(rows[0]).toEqual(["Name"]);
		expect(rows[1][1]).toBe("Done");

		const html = sheetToHtml(ws);
		expect(html.match(/<tr>/g)).toHaveLength(2);
		expect(html).toContain("Name");
		expect(html).toContain("Done");
	});
});
