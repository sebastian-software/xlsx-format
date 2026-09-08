import { describe, expect, it } from "vitest";
import {
	appendSheet,
	arrayToSheet,
	createWorkbook,
	read,
	sheetToCsv,
	sheetToFormulae,
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
		expect(Object.getOwnPropertyDescriptor(row, "__proto__")?.value).toBe("polluted");
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

		const parsed = await read(await import("./zip/index.js").then(async ({ zipWrite }) => zipWrite(zip)));

		expect(parsed.SheetNames).toStrictEqual(["__proto__"]);
		expect(Object.getPrototypeOf(parsed.Sheets)).toBeNull();
		expect(Object.hasOwn(parsed.Sheets, "__proto__")).toBe(true);
		expect(Object.prototype).not.toHaveProperty("value");
	});

	it("escapes formula-like CSV fields by default", () => {
		const ws = arrayToSheet([["=1+1", "+SUM(A1)", "-cmd", "@handle", "plain"]]);

		expect(sheetToCsv(ws)).toBe("'=1+1,'+SUM(A1),'-cmd,'@handle,plain");
		expect(sheetToCsv(ws, { escapeFormulae: false })).toBe("=1+1,+SUM(A1),-cmd,@handle,plain");
	});

	it("escapes formula-only CSV cells by default", () => {
		const ws = { A1: { t: "z", f: "SUM(B1:B10)" }, "!ref": "A1:A1" } as WorkSheet;

		expect(sheetToCsv(ws)).toBe("'=SUM(B1:B10)");
		expect(sheetToCsv(ws, { escapeFormulae: false })).toBe("=SUM(B1:B10)");
	});

	it("sanitizes javascript links with embedded whitespace", () => {
		const ws = {
			"!ref": "A1",
			A1: { t: "s", v: "Bad", l: { Target: "java\nscript:alert(1)" } },
		} as WorkSheet;
		const unsafeOptOutWs = {
			"!ref": "A1",
			A1: { t: "s", v: "Bad", l: { Target: "javascript:alert(1)" } },
		} as WorkSheet;

		expect(sheetToHtml(ws)).not.toContain("href=");
		expect(sheetToHtml(unsafeOptOutWs, { sanitizeLinks: false })).toContain('href="javascript:alert(1)"');
	});

	it("sanitizes script links with invisible unicode characters", () => {
		const ws = {
			"!ref": "A1",
			A1: { t: "s", v: "Bad", l: { Target: "java\u200bscript:alert(1)" } },
		} as WorkSheet;

		expect(sheetToHtml(ws)).not.toContain("href=");
	});

	it("sanitizes data and vbscript links", () => {
		for (const target of ["data:text/html,<script>alert(1)</script>", "vbscript:msgbox(1)"]) {
			const ws = {
				"!ref": "A1",
				A1: { t: "s", v: "Bad", l: { Target: target } },
			} as WorkSheet;

			expect(sheetToHtml(ws)).not.toContain("href=");
		}
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
		expect(rows[0]).toStrictEqual(["Name"]);
		expect(rows[1][1]).toBe("Done");

		const html = sheetToHtml(ws);
		expect(html.match(/<tr>/g)).toHaveLength(2);
		expect(html).toContain("Name");
		expect(html).toContain("Done");
	});

	it("clamps oversized dense worksheet ranges to occupied cells", () => {
		const ws = arrayToSheet([["Name"], [undefined, "Done"]], { dense: true });
		ws["!ref"] = "A1:XFD1048576";

		expect(sheetToCsv(ws)).toBe("Name,\n,Done");
		expect(sheetToJson<unknown[]>(ws, { header: 1 })[1][1]).toBe("Done");
		expect(sheetToHtml(ws)).toContain("Done");
	});

	it("ignores occupied cells outside a numeric oversized JSON range", () => {
		const ws = {
			"!ref": "A1:XFD1048576",
			A1: { t: "s", v: "Skip" },
			B2: { t: "s", v: "Done" },
		} as WorkSheet;

		const rows = sheetToJson<unknown[]>(ws, { header: 1, range: 1 });

		expect(rows).toHaveLength(1);
		expect(rows[0][0]).toBeUndefined();
		expect(rows[0][1]).toBe("Done");
	});

	it("skips oversized declared ranges with no occupied cells", () => {
		const ws = { "!ref": "A1:XFD1048576" } as WorkSheet;

		expect(sheetToCsv(ws)).toBe("");
		expect(sheetToJson(ws)).toStrictEqual([]);
		expect(sheetToHtml(ws)).not.toContain("<tr>");
	});

	it("rejects an occupied far edge under every default export budget", async () => {
		const ws = {
			"!ref": "A1:A1048576",
			A1048576: { t: "n", v: 7, f: "SUM(A1:A2)" },
		} as WorkSheet;

		for (const operation of [
			() => sheetToCsv(ws),
			() => sheetToJson(ws, { header: 1 }),
			() => sheetToHtml(ws),
			() => sheetToFormulae(ws),
		]) {
			expect(operation).toThrow(/worksheet export cell count 1048576 exceeds limit 1000000/);
		}
		await expect(write(createWorkbook(ws, "S"))).rejects.toThrow(
			/worksheet export cell count 1048576 exceeds limit 1000000/,
		);
	});

	it("preserves an occupied edge when the caller raises the export budget", async () => {
		const ws = { "!ref": "A1:B1", B1: { t: "n", v: 7, f: "1+6" } } as WorkSheet;

		expect(() => sheetToCsv(ws, { maxWorksheetCells: 1 })).toThrow(/exceeds limit 1/);
		expect(sheetToCsv(ws, { maxWorksheetCells: 2 })).toBe(",7");
		expect(sheetToJson<unknown[]>(ws, { header: 1, maxWorksheetCells: 2 })[0][1]).toBe(7);
		expect(sheetToHtml(ws, { maxWorksheetCells: 2 })).toContain('id="sjs-B1"');
		expect(sheetToFormulae(ws, { maxWorksheetCells: 2 })).toContain("B1=1+6");

		const written = await write(createWorkbook(ws, "S"), { maxWorksheetCells: 2 });
		expect((await read(written)).Sheets.S.B1?.v).toBe(7);
	});

	it("enforces explicit JSON ranges and validates range coordinates", () => {
		const ws = arrayToSheet([[1]]);

		expect(() => sheetToJson(ws, { header: 1, range: "A1:B1", maxWorksheetCells: 1 })).toThrow(/exceeds limit 1/);
		expect(() =>
			sheetToJson(ws, {
				header: 1,
				range: { s: { r: 0, c: 0 }, e: { r: 0, c: Number.POSITIVE_INFINITY } },
			}),
		).toThrow(/end column must be a non-negative safe integer/);
	});

	it("charges sparse column metadata without scanning array length", async () => {
		const ws = arrayToSheet([[1]]);
		ws["!cols"] = [];
		ws["!cols"][0] = { width: 10 };

		await expect(write(createWorkbook(ws, "S"), { maxWorksheetCells: 1 })).rejects.toThrow(
			/worksheet export cell count 2 exceeds limit 1/,
		);
		await expect(write(createWorkbook(ws, "S"), { maxWorksheetCells: 2 })).resolves.toBeInstanceOf(Uint8Array);

		ws["!cols"][1_000_000] = { width: 10 };
		await expect(write(createWorkbook(ws, "S"), { maxWorksheetCells: 3 })).rejects.toThrow(
			/column metadata exceeds XLSX column limit/,
		);
	});
});
