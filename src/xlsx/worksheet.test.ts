import { describe, it, expect } from "vitest";
import { parseWorksheetXml, resolveSharedStrings } from "./worksheet.js";
import { read, write, createWorkbook, jsonToSheet, sheetToArray } from "../index.js";

describe("worksheet.ts: parsing edge cases", () => {
	it("empty data returns empty sheet", () => {
		const ws = parseWorksheetXml("");
		expect(ws).toBeDefined();
	});

	it("dense mode", () => {
		const xml = `<?xml version="1.0"?>
		<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
		<sheetData>
		<row r="1"><c r="A1" t="str"><v>hello</v></c></row>
		</sheetData>
		</worksheet>`;
		const ws = parseWorksheetXml(xml, { dense: true });
		expect(ws["!data"]).toBeDefined();
		expect(ws["!data"]![0][0].v).toBe("hello");
	});

	it("cell types: boolean, error, inline string", () => {
		const xml = `<?xml version="1.0"?>
		<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
		<sheetData>
		<row r="1">
			<c r="A1" t="b"><v>1</v></c>
			<c r="B1" t="e"><v>#DIV/0!</v></c>
			<c r="C1" t="inlineStr"><is><t>Inline</t></is></c>
		</row>
		</sheetData>
		</worksheet>`;
		const ws = parseWorksheetXml(xml);
		expect((ws as any).A1.t).toBe("b");
		expect((ws as any).A1.v).toBe(true);
		expect((ws as any).B1.t).toBe("e");
		expect((ws as any).C1.t).toBe("s");
		expect((ws as any).C1.v).toBe("Inline");
	});

	it("sheetRows limits parsing", () => {
		const xml = `<?xml version="1.0"?>
		<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
		<dimension ref="A1:A10"/>
		<sheetData>
		<row r="1"><c r="A1"><v>1</v></c></row>
		<row r="2"><c r="A2"><v>2</v></c></row>
		<row r="3"><c r="A3"><v>3</v></c></row>
		<row r="10"><c r="A10"><v>10</v></c></row>
		</sheetData>
		</worksheet>`;
		const ws = parseWorksheetXml(xml, { sheetRows: 2 });
		expect((ws as any).A1).toBeDefined();
		expect((ws as any).A3).toBeUndefined();
	});

	it("sheetStubs creates z-type cells", () => {
		const xml = `<?xml version="1.0"?>
		<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
		<sheetData>
		<row r="1"><c r="A1"></c></row>
		</sheetData>
		</worksheet>`;
		const ws = parseWorksheetXml(xml, { sheetStubs: true });
		expect((ws as any).A1).toBeDefined();
		expect((ws as any).A1.t).toBe("z");
	});

	it("preserves later cell positions after self-closing empty cells", () => {
		const xml = `<?xml version="1.0"?>
		<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
		<dimension ref="A1:C1"/>
		<sheetData>
		<row r="1">
			<c r="A1" t="str"><v>left</v></c>
			<c r="B1" s="0"/>
			<c r="C1" t="str"><v>right</v></c>
		</row>
		</sheetData>
		</worksheet>`;
		const ws = parseWorksheetXml(xml, { sheetStubs: true });
		expect((ws as any).A1).toMatchObject({ t: "s", v: "left" });
		expect((ws as any).B1).toMatchObject({ t: "z" });
		expect((ws as any).C1).toMatchObject({ t: "s", v: "right" });
		expect(sheetToArray(ws, { defval: null })).toStrictEqual([["left", null, "right"]]);
	});

	it("resolveSharedStrings works in sparse mode", () => {
		const ws: any = {
			A1: { t: "s", v: "", _sstIdx: 0 },
			B1: { t: "n", v: 42 },
			"!ref": "A1:B1",
		};
		const sst: any = [{ t: "Hello", h: "<b>Hello</b>", r: "<t>Hello</t>" }];
		resolveSharedStrings(ws, sst, {});
		expect(ws.A1.v).toBe("Hello");
	});

	it("resolveSharedStrings works in dense mode", () => {
		const ws: any = {
			"!data": [
				[
					{ t: "s", v: "", _sstIdx: 0 },
					{ t: "n", v: 42 },
				],
			],
			"!ref": "A1:B1",
		};
		const sst: any = [{ t: "World", h: "World", r: "<t>World</t>" }];
		resolveSharedStrings(ws, sst, {});
		expect(ws["!data"][0][0].v).toBe("World");
	});
});

describe("worksheet: margins and autofilter roundtrip", () => {
	it("roundtrips page margins", async () => {
		const ws = jsonToSheet([{ a: 1 }]);
		ws["!margins"] = {
			left: 0.5,
			right: 0.5,
			top: 1,
			bottom: 1,
			header: 0.25,
			footer: 0.25,
		};
		const wb = createWorkbook(ws, "S1");
		const buf = await write(wb);
		const result = await read(buf);
		expect(result.Sheets.S1["!margins"]).toBeDefined();
	});

	it("roundtrips autofilter", async () => {
		const ws = jsonToSheet([
			{ Name: "Alice", Age: 30 },
			{ Name: "Bob", Age: 25 },
		]);
		ws["!autofilter"] = { ref: "A1:B3" };
		const wb = createWorkbook(ws, "S1");
		const buf = await write(wb);
		const result = await read(buf);
		expect(result.Sheets.S1["!autofilter"]).toBeDefined();
		expect(result.Sheets.S1["!autofilter"]!.ref).toBe("A1:B3");
	});
});

describe("worksheet: hyperlink parsing via parseWorksheetXml", () => {
	it("parses external hyperlinks from XML", () => {
		const xml = `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetData>
<row r="1"><c r="A1" t="inlineStr"><is><t>Click</t></is></c></row>
</sheetData>
<hyperlinks>
<hyperlink ref="A1" r:id="rId1"/>
</hyperlinks>
</worksheet>`;
		const rels: any = {
			"!id": { rId1: { Target: "https://example.com", TargetMode: "External" } },
		};
		const ws = parseWorksheetXml(xml, {}, 0, rels);
		expect(ws.A1?.l?.Target).toBe("https://example.com");
	});

	it("parses hyperlinks with tooltip from XML", () => {
		const xml = `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetData>
<row r="1"><c r="A1" t="inlineStr"><is><t>Click</t></is></c></row>
</sheetData>
<hyperlinks>
<hyperlink ref="A1" r:id="rId1" tooltip="Visit site"/>
</hyperlinks>
</worksheet>`;
		const rels: any = {
			"!id": { rId1: { Target: "https://example.com", TargetMode: "External" } },
		};
		const ws = parseWorksheetXml(xml, {}, 0, rels);
		expect(ws.A1?.l?.Target).toBe("https://example.com");
		expect(ws.A1?.l?.Tooltip).toBe("Visit site");
	});
});

describe("parseWorksheetXml: direct XML parsing", () => {
	it("returns empty sheet for empty data", () => {
		const ws = parseWorksheetXml("");
		expect(ws).toBeDefined();
	});

	it("parses worksheet with pageMargins", () => {
		const xml = `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData></sheetData>
<pageMargins left="0.5" right="0.5" top="1.0" bottom="1.0" header="0.25" footer="0.25"/>
</worksheet>`;
		const ws = parseWorksheetXml(xml);
		expect(ws["!margins"]).toBeDefined();
		expect(ws["!margins"]!.left).toBe(0.5);
	});

	it("parses worksheet with autoFilter", () => {
		const xml = `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>
</sheetData>
<autoFilter ref="A1:B1"/>
</worksheet>`;
		const ws = parseWorksheetXml(xml);
		expect(ws["!autofilter"]).toBeDefined();
	});

	it("enforces worksheet row and cell scan limits", () => {
		const rowsXml = `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1"><c r="A1"><v>1</v></c></row>
<row r="2"><c r="A2"><v>2</v></c></row>
</sheetData>
</worksheet>`;
		const cellsXml = `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>
</sheetData>
</worksheet>`;
		expect(() => parseWorksheetXml(rowsXml, { maxWorksheetRows: 1 })).toThrow(
			/worksheet row count 2 exceeds limit 1/,
		);
		expect(() => parseWorksheetXml(cellsXml, { maxWorksheetCells: 1 })).toThrow(
			/worksheet cell count 2 exceeds limit 1/,
		);
	});

	it("parses worksheet with cols (cellStyles)", () => {
		const xml = `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<cols>
<col min="1" max="1" width="15" hidden="1"/>
<col min="2" max="3" width="20"/>
</cols>
<sheetData></sheetData>
</worksheet>`;
		const ws = parseWorksheetXml(xml, { cellStyles: true });
		expect(ws["!cols"]).toBeDefined();
		if (ws["!cols"]) {
			expect(ws["!cols"][0]?.width).toBe(15);
			expect(ws["!cols"][0]?.hidden).toBe(true);
			expect(ws["!cols"][1]?.width).toBe(20);
		}
	});

	it("parses worksheet with hyperlinks", () => {
		const xml = `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetData>
<row r="1"><c r="A1" t="inlineStr"><is><t>Click</t></is></c></row>
</sheetData>
<hyperlinks>
<hyperlink ref="A1" r:id="rId1" tooltip="Visit"/>
</hyperlinks>
</worksheet>`;
		const rels: any = {
			"!id": {
				rId1: { Target: "https://example.com", TargetMode: "External" },
			},
		};
		const ws = parseWorksheetXml(xml, {}, 0, rels);
		expect(ws.A1?.l).toBeDefined();
		expect(ws.A1?.l?.Target).toBe("https://example.com");
		expect(ws.A1?.l?.Tooltip).toBe("Visit");
	});

	it("parses worksheet with hyperlink location (internal link)", () => {
		const xml = `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1"><c r="A1" t="inlineStr"><is><t>Go</t></is></c></row>
</sheetData>
<hyperlinks>
<hyperlink ref="A1" location="Sheet2!A1"/>
</hyperlinks>
</worksheet>`;
		const ws = parseWorksheetXml(xml);
		expect(ws.A1?.l).toBeDefined();
		expect(ws.A1?.l?.Target).toContain("Sheet2");
	});

	it("parses worksheet with merge cells", () => {
		const xml = `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1"><c r="A1"><v>1</v></c></row>
</sheetData>
<mergeCells count="1"><mergeCell ref="A1:B2"/></mergeCells>
</worksheet>`;
		const ws = parseWorksheetXml(xml);
		expect(ws["!merges"]).toBeDefined();
		expect(ws["!merges"]!).toHaveLength(1);
	});
});
