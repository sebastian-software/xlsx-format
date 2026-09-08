import { describe, expect, it } from "vitest";
import { arrayToSheet, createWorkbook, read, setRowHeight, write } from "../index.js";
import { BErr } from "../types.js";
import { zipRead, zipReadString } from "../zip/index.js";
import { parseWorksheetXml, writeWorksheetXml } from "./worksheet.js";

const worksheetXml = (cells: string, rows = ""): string => `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<dimension ref="A1:A3"/><sheetData>${rows || `<row r="1"><c r="A1">${cells}</c></row>`}</sheetData>
</worksheet>`;

describe("worksheet error values", () => {
	it("maps every supported error token from literal XML to its BErr code", () => {
		const entries = Object.entries(BErr);
		const maxRow = Math.max(...entries.map(([code]) => Number(code) + 1));
		const xml = `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<dimension ref="A1:A${maxRow}"/><sheetData>${entries
			.map(
				([code, token]) =>
					`<row r="${Number(code) + 1}"><c r="A${Number(code) + 1}" t="e"><v>${token}</v></c></row>`,
			)
			.join("")}</sheetData>
</worksheet>`;
		const sparse = parseWorksheetXml(xml);
		const dense = parseWorksheetXml(xml, { dense: true });

		for (const [code, token] of entries) {
			const row = Number(code);
			expect(sparse[`A${row + 1}`]).toMatchObject({ t: "e", v: row, w: token });
			expect(dense["!data"]?.[row]?.[0]).toMatchObject({ t: "e", v: row, w: token });
		}
	});

	it("preserves unknown error tokens instead of coercing them to zero", () => {
		const ws = parseWorksheetXml(worksheetXml("", '<row r="1"><c r="A1" t="e"><v>#SPILL!</v></c></row>'));
		expect(ws.A1).toMatchObject({ t: "e", v: "#SPILL!", w: "#SPILL!" });

		const xml = writeWorksheetXml({ A1: { t: "e", v: "#SPILL!" }, "!ref": "A1" }, {}, 0, {} as any, {});
		expect(xml).toContain('<c r="A1" t="e"><v>#SPILL!</v></c>');
	});

	it("serializes numeric error codes to XLSX error tokens", () => {
		const xml = writeWorksheetXml({ A1: { t: "e", v: 0x07 }, "!ref": "A1" }, {}, 0, {} as any, {});
		expect(xml).toContain('<c r="A1" t="e"><v>#DIV/0!</v></c>');
	});

	it("keeps standard and unknown error values through public write/read", async () => {
		const ws = arrayToSheet([
			[
				{ t: "e", v: 0x07 },
				{ t: "e", v: "#SPILL!" },
			],
		]);
		const bytes = await write(createWorkbook(ws, "Errors"));
		const xml = zipReadString(await zipRead(bytes), "xl/worksheets/sheet1.xml");
		expect(xml).toContain("#DIV/0!");
		expect(xml).toContain("#SPILL!");

		const result = await read(bytes);
		expect(result.Sheets.Errors.A1).toMatchObject({ t: "e", v: 0x07 });
		expect(result.Sheets.Errors.B1).toMatchObject({ t: "e", v: "#SPILL!" });
	});
});

describe("worksheet layout metadata", () => {
	it("preserves zero margins in literal XML parsing and writing", () => {
		const xml = `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData/><pageMargins left="0" right="0" top="0" bottom="0" header="0" footer="0"/>
</worksheet>`;
		const parsed = parseWorksheetXml(xml);
		expect(parsed["!margins"]).toStrictEqual({ left: 0, right: 0, top: 0, bottom: 0, header: 0, footer: 0 });
		const defaults = parseWorksheetXml(
			`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><pageMargins/></worksheet>`,
		);
		expect(defaults["!margins"]).toStrictEqual({
			left: 0.7,
			right: 0.7,
			top: 0.75,
			bottom: 0.75,
			header: 0.3,
			footer: 0.3,
		});

		const serialized = writeWorksheetXml(
			{ "!ref": "A1", "!margins": { left: 0, right: 0, top: 0, bottom: 0, header: 0, footer: 0 } },
			{},
			0,
			{} as any,
			{},
		);
		expect(serialized).toContain('left="0" right="0" top="0" bottom="0" header="0" footer="0"');
	});

	it("emits and reads metadata for blank rows outside the data range", async () => {
		const ws = arrayToSheet([["data"]]);
		setRowHeight(ws, 2, 30);
		ws["!rows"]![2].hidden = true;
		const bytes = await write(createWorkbook(ws, "Layout"));
		const xml = zipReadString(await zipRead(bytes), "xl/worksheets/sheet1.xml");
		expect(xml).toContain('<row r="3" ht="30" customHeight="1" hidden="1"></row>');

		const result = await read(bytes);
		expect(result.Sheets.Layout["!rows"]?.[2]).toMatchObject({ hpt: 30, hidden: true });
	});

	it("emits metadata for blank spacer rows inside a dense range", async () => {
		const ws = arrayToSheet([["top"], [], ["bottom"]], { dense: true });
		setRowHeight(ws, 1, 18);
		const bytes = await write(createWorkbook(ws, "Dense"));
		const xml = zipReadString(await zipRead(bytes), "xl/worksheets/sheet1.xml");
		expect(xml).toContain('<row r="2" ht="18" customHeight="1"></row>');

		const result = await read(bytes, { dense: true });
		expect(result.Sheets.Dense["!rows"]?.[1]).toMatchObject({ hpt: 18 });
	});

	it("keeps metadata rows before the data range ordered and preserves zero height", () => {
		const xml = writeWorksheetXml(
			{ A5: { t: "s", v: "data" }, "!ref": "A5:A5", "!rows": [{ hpt: 0 }] },
			{},
			0,
			{} as any,
			{},
		);
		expect(xml).toContain('<row r="1" ht="0" customHeight="1"></row>');
		expect(xml.indexOf('<row r="1"')).toBeLessThan(xml.indexOf('<row r="5"'));
	});
});
