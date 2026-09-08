import { describe, it, expect } from "vitest";
import {
	read,
	write,
	createWorkbook,
	appendSheet,
	arrayToSheet,
	sheetToJson,
	setArrayFormula,
	setSheetVisibility,
	jsonToSheet,
	formatCell,
	setCellStyle,
} from "../index.js";
import { is1904DateSystem } from "./workbook.js";
import { zipAddString, zipRead, zipReadString, zipWrite } from "../zip/index.js";

async function replaceZipPart(bytes: Uint8Array, path: string, content: string): Promise<Uint8Array> {
	const zip = await zipRead(bytes);
	zipAddString(zip, path, content);
	return zipWrite(zip);
}

describe("XLSX roundtrip: workbook features", () => {
	it("hidden sheets survive roundtrip", async () => {
		const ws1 = arrayToSheet([["Visible"]]);
		const ws2 = arrayToSheet([["Hidden"]]);
		const wb = createWorkbook(ws1, "Vis");
		appendSheet(wb, ws2, "Hid");
		setSheetVisibility(wb, "Hid", 1);

		const bytes = await write(wb);
		const wb2 = await read(bytes);
		expect(wb2.SheetNames).toContain("Vis");
		expect(wb2.SheetNames).toContain("Hid");
		// Check hidden state via Workbook metadata
		if (wb2.Workbook?.Sheets) {
			const hidSheet = wb2.Workbook.Sheets.find((s: any) => s.name === "Hid");
			if (hidSheet) {
				expect(hidSheet.Hidden).toBe(1);
			}
		}
	});

	it("very hidden sheets survive roundtrip", async () => {
		const ws1 = arrayToSheet([["A"]]);
		const ws2 = arrayToSheet([["B"]]);
		const wb = createWorkbook(ws1, "S1");
		appendSheet(wb, ws2, "S2");
		setSheetVisibility(wb, "S2", 2);

		const bytes = await write(wb);
		const wb2 = await read(bytes);
		expect(wb2.SheetNames).toContain("S2");
	});

	it("defined names survive roundtrip", async () => {
		const ws = arrayToSheet([["A"]]);
		const wb = createWorkbook(ws, "Sheet1");
		wb.Workbook = {
			WBProps: {},
			Sheets: [],
			Names: [
				{ Name: "MyRange", Ref: "Sheet1!$A$1:$B$2" },
				{ Name: "HiddenName", Ref: "Sheet1!$C$1", Hidden: true },
				{ Name: "LocalName", Ref: "Sheet1!$D$1", Sheet: 0 },
			],
		};

		const bytes = await write(wb);
		const wb2 = await read(bytes);
		expect(wb2.Workbook?.Names).toBeDefined();
		expect(wb2.Workbook!.Names!.length).toBeGreaterThanOrEqual(3);
		const myRange = wb2.Workbook!.Names!.find((n: any) => n.Name === "MyRange");
		expect(myRange).toBeDefined();
		expect(myRange!.Ref).toContain("Sheet1");
	});

	it("merge cells survive roundtrip", async () => {
		const ws = arrayToSheet([
			["Merged", null, null],
			[null, null, null],
			["Normal", "B", "C"],
		]);
		ws["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 1, c: 2 } }];
		const wb = createWorkbook(ws, "Sheet1");

		const bytes = await write(wb);
		const wb2 = await read(bytes);
		const ws2 = wb2.Sheets[wb2.SheetNames[0]];
		expect(ws2["!merges"]).toBeDefined();
		expect(ws2["!merges"]!).toHaveLength(1);
		expect(ws2["!merges"]![0].s.r).toBe(0);
		expect(ws2["!merges"]![0].e.r).toBe(1);
	});

	it("column widths survive roundtrip", async () => {
		const ws = arrayToSheet([["A", "B", "C"]]);
		ws["!cols"] = [{ width: 20 }, { width: 30 }, { width: 10 }];
		const wb = createWorkbook(ws, "Sheet1");

		const bytes = await write(wb);
		const wb2 = await read(bytes, { cellStyles: true });
		const ws2 = wb2.Sheets[wb2.SheetNames[0]];
		expect(ws2["!cols"]).toBeDefined();
		expect(ws2["!cols"]!.length).toBeGreaterThan(0);
	});

	it("row heights survive roundtrip", async () => {
		const ws = arrayToSheet([["A"], ["B"], ["C"]]);
		ws["!rows"] = [{ hpt: 30 }, undefined as any, { hpt: 40 }];
		const wb = createWorkbook(ws, "Sheet1");

		const bytes = await write(wb);
		const wb2 = await read(bytes);
		const ws2 = wb2.Sheets[wb2.SheetNames[0]];
		expect(ws2["!rows"]).toBeDefined();
		expect(ws2["!rows"]![0].hpt).toBe(30);
	});

	it("hidden rows survive roundtrip", async () => {
		const ws = arrayToSheet([["Visible"], ["Hidden"], ["Visible2"]]);
		ws["!rows"] = [undefined as any, { hidden: true }, undefined as any];
		const wb = createWorkbook(ws, "Sheet1");

		const bytes = await write(wb);
		const wb2 = await read(bytes);
		const ws2 = wb2.Sheets[wb2.SheetNames[0]];
		expect(ws2["!rows"]).toBeDefined();
		expect(ws2["!rows"]![1].hidden).toBe(true);
	});

	it("autofilter survives roundtrip", async () => {
		const ws = arrayToSheet([
			["Name", "Age"],
			["Alice", 30],
			["Bob", 25],
		]);
		ws["!autofilter"] = { ref: "A1:B3" };
		const wb = createWorkbook(ws, "Sheet1");

		const bytes = await write(wb);
		const wb2 = await read(bytes);
		const ws2 = wb2.Sheets[wb2.SheetNames[0]];
		expect(ws2["!autofilter"]).toBeDefined();
		expect(ws2["!autofilter"]!.ref).toBe("A1:B3");
	});

	it("page margins survive roundtrip", async () => {
		const ws = arrayToSheet([["A"]]);
		ws["!margins"] = { left: 1, right: 1, top: 1.5, bottom: 1.5, header: 0.5, footer: 0.5 };
		const wb = createWorkbook(ws, "Sheet1");

		const bytes = await write(wb);
		const wb2 = await read(bytes);
		const ws2 = wb2.Sheets[wb2.SheetNames[0]];
		expect(ws2["!margins"]).toBeDefined();
		expect(ws2["!margins"]!.left).toBe(1);
	});

	it("formulas survive roundtrip", async () => {
		const ws: any = {
			A1: { t: "n", v: 1 },
			A2: { t: "n", v: 2 },
			A3: { t: "n", v: 3, f: "SUM(A1:A2)" },
			"!ref": "A1:A3",
		};
		const wb = createWorkbook(ws, "Sheet1");

		const bytes = await write(wb);
		const wb2 = await read(bytes, { cellFormula: true });
		const ws2 = wb2.Sheets[wb2.SheetNames[0]];
		const a3 = (ws2 as any).A3;
		expect(a3).toBeDefined();
		expect(a3.f).toBe("SUM(A1:A2)");
	});

	it("array formulas survive roundtrip", async () => {
		const ws = arrayToSheet([
			[1, 10],
			[2, 20],
			[3, 30],
		]);
		setArrayFormula(ws, "C1:C3", "A1:A3*B1:B3");

		const wb = createWorkbook(ws, "Sheet1");
		const bytes = await write(wb);
		const wb2 = await read(bytes, { cellFormula: true });
		const ws2 = wb2.Sheets[wb2.SheetNames[0]];
		const c1 = (ws2 as any).C1;
		expect(c1).toBeDefined();
		expect(c1.f).toBe("A1:A3*B1:B3");
		expect(c1.F).toBe("C1:C3");
	});

	it("boolean cells survive roundtrip", async () => {
		const ws: any = { A1: { t: "b", v: true }, A2: { t: "b", v: false }, "!ref": "A1:A2" };
		const wb = createWorkbook(ws, "Sheet1");
		const bytes = await write(wb);
		const wb2 = await read(bytes);
		const ws2 = wb2.Sheets[wb2.SheetNames[0]];
		expect((ws2 as any).A1.v).toBe(true);
		expect((ws2 as any).A2.v).toBe(false);
	});

	it("error cells survive roundtrip", async () => {
		const ws: any = { A1: { t: "e", v: 0x07, w: "#DIV/0!" }, "!ref": "A1:A1" };
		const wb = createWorkbook(ws, "Sheet1");
		const bytes = await write(wb);
		const wb2 = await read(bytes);
		const ws2 = wb2.Sheets[wb2.SheetNames[0]];
		expect((ws2 as any).A1.t).toBe("e");
	});

	it("date cells with cellDates option", async () => {
		const ws: any = { A1: { t: "d", v: new Date("2021-06-15T00:00:00") }, "!ref": "A1:A1" };
		const wb = createWorkbook(ws, "Sheet1");
		const bytes = await write(wb, { cellDates: true });
		const wb2 = await read(bytes, { cellDates: true });
		const ws2 = wb2.Sheets[wb2.SheetNames[0]];
		const a1 = (ws2 as any).A1;
		expect(a1).toBeDefined();
	});

	it("multiple sheets", async () => {
		const ws1 = arrayToSheet([["First"]]);
		const ws2 = arrayToSheet([["Second"]]);
		const ws3 = arrayToSheet([["Third"]]);
		const wb = createWorkbook(ws1, "S1");
		appendSheet(wb, ws2, "S2");
		appendSheet(wb, ws3, "S3");

		const bytes = await write(wb);
		const wb2 = await read(bytes);
		expect(wb2.SheetNames).toStrictEqual(["S1", "S2", "S3"]);
	});

	it("inline strings survive roundtrip", async () => {
		const ws = arrayToSheet([["Hello world", "Test 123"]]);
		const wb = createWorkbook(ws, "Sheet1");

		const bytes = await write(wb);
		const wb2 = await read(bytes);
		const rows = sheetToJson(wb2.Sheets[wb2.SheetNames[0]], { header: 1 });
		expect(rows[0]).toContain("Hello world");
		expect(rows[0]).toContain("Test 123");
	});

	it("bookSheets option returns only sheet names", async () => {
		const ws = arrayToSheet([["A"]]);
		const wb = createWorkbook(ws, "MySheet");
		const bytes = await write(wb);
		const wb2 = await read(bytes, { bookSheets: true });
		expect(wb2.SheetNames).toContain("MySheet");
		expect(Object.keys(wb2.Sheets || {})).toHaveLength(0);
	});

	it("bookProps option returns properties", async () => {
		const ws = arrayToSheet([["A"]]);
		const wb = createWorkbook(ws, "Sheet1");
		wb.Props = { Title: "Test Title", Author: "Test Author" };
		const bytes = await write(wb);
		const wb2 = await read(bytes, { bookProps: true });
		expect(wb2.Props).toBeDefined();
	});

	it("sheetRows limits row count", async () => {
		const data = Array.from({ length: 50 }, (_, i) => [i]);
		const ws = arrayToSheet(data);
		const wb = createWorkbook(ws, "Sheet1");
		const bytes = await write(wb);
		const wb2 = await read(bytes, { sheetRows: 10 });
		const ws2 = wb2.Sheets[wb2.SheetNames[0]];
		const rows = sheetToJson(ws2, { header: 1 });
		expect(rows.length).toBeLessThanOrEqual(10);
	});

	it("sheets filter by name", async () => {
		const ws1 = arrayToSheet([["First"]]);
		const ws2 = arrayToSheet([["Second"]]);
		const wb = createWorkbook(ws1, "S1");
		appendSheet(wb, ws2, "S2");
		const bytes = await write(wb);
		const wb2 = await read(bytes, { sheets: "S2" });
		expect(wb2.SheetNames).toContain("S2");
		// S1 might still be in SheetNames (from workbook.xml) but its data shouldn't be loaded
		if (wb2.SheetNames.includes("S1")) {
			expect(wb2.Sheets.S1).toBeUndefined();
		}
	});

	it("sheets filter by index", async () => {
		const ws1 = arrayToSheet([["First"]]);
		const ws2 = arrayToSheet([["Second"]]);
		const wb = createWorkbook(ws1, "S1");
		appendSheet(wb, ws2, "S2");
		const bytes = await write(wb);
		const wb2 = await read(bytes, { sheets: 1 });
		expect(wb2.Sheets.S2).toBeDefined();
	});

	it("sheets filter by array", async () => {
		const ws1 = arrayToSheet([["First"]]);
		const ws2 = arrayToSheet([["Second"]]);
		const ws3 = arrayToSheet([["Third"]]);
		const wb = createWorkbook(ws1, "S1");
		appendSheet(wb, ws2, "S2");
		appendSheet(wb, ws3, "S3");
		const bytes = await write(wb);
		const wb2 = await read(bytes, { sheets: [0, "S3"] });
		expect(wb2.Sheets.S1).toBeDefined();
		expect(wb2.Sheets.S3).toBeDefined();
	});

	it("custom properties survive roundtrip", async () => {
		const ws = arrayToSheet([["A"]]);
		const wb = createWorkbook(ws, "Sheet1");
		wb.Custprops = { myKey: "myValue", myNumber: 42 };
		const bytes = await write(wb);
		const wb2 = await read(bytes);
		expect(wb2.Custprops).toBeDefined();
		expect(wb2.Custprops!.myKey).toBe("myValue");
	});
});

describe("XLSX roundtrip: defined names", () => {
	it("roundtrips defined names with comment, localSheetId, hidden", async () => {
		const wb = createWorkbook(jsonToSheet([{ Val: 1 }]), "Sheet1");
		wb.Workbook = wb.Workbook || { Sheets: [] };
		wb.Workbook.Names = [
			{ Name: "GlobalRange", Ref: "Sheet1!$A$1:$A$2" },
			{ Name: "LocalRange", Ref: "Sheet1!$B$1", Sheet: 0 },
			{ Name: "HiddenName", Ref: "Sheet1!$C$1", Hidden: true },
			{ Name: "Commented", Ref: "Sheet1!$D$1", Comment: "Note" },
		];

		const buf = await write(wb);
		const wb2 = await read(buf);

		expect(wb2.Workbook?.Names).toBeDefined();
		const names = wb2.Workbook!.Names!;
		expect(names.length).toBeGreaterThanOrEqual(4);

		const global = names.find((n: any) => n.Name === "GlobalRange");
		expect(global?.Ref).toBe("Sheet1!$A$1:$A$2");

		const local = names.find((n: any) => n.Name === "LocalRange");
		expect(local?.Sheet).toBe(0);

		const hidden = names.find((n: any) => n.Name === "HiddenName");
		expect(hidden?.Hidden).toBe(true);

		const commented = names.find((n: any) => n.Name === "Commented");
		expect(commented?.Comment).toBe("Note");
	});
});

describe("XLSX roundtrip: sheet visibility", () => {
	it("roundtrips hidden and veryHidden sheets", async () => {
		const wb = createWorkbook(jsonToSheet([{ a: 1 }]), "Visible");
		appendSheet(wb, jsonToSheet([{ b: 2 }]), "Hidden");
		appendSheet(wb, jsonToSheet([{ c: 3 }]), "VeryHidden");
		setSheetVisibility(wb, 1, 1);
		setSheetVisibility(wb, 2, 2);

		const buf = await write(wb);
		const wb2 = await read(buf);

		expect(wb2.Workbook?.Sheets?.[0]?.Hidden).toBe(0);
		expect(wb2.Workbook?.Sheets?.[1]?.Hidden).toBe(1);
		expect(wb2.Workbook?.Sheets?.[2]?.Hidden).toBe(2);
	});
});

describe("XLSX roundtrip: workbook properties", () => {
	it("roundtrips date1904 mode", async () => {
		const wb = createWorkbook(jsonToSheet([{ a: 1 }]), "S1");
		wb.Workbook = wb.Workbook || { Sheets: [] };
		wb.Workbook.WBProps = { date1904: true };

		const buf = await write(wb);
		const wb2 = await read(buf);

		expect(is1904DateSystem(wb2)).toBe("true");
	});

	it.each([
		{ date1904: false, expectedSerial: 46037.52425925926 },
		{ date1904: true, expectedSerial: 44575.52425925926 },
	])("serializes Date cells with the $date1904 workbook date system", async ({ date1904, expectedSerial }) => {
		const date = new Date("2026-01-15T12:34:56.000Z");
		const ws = arrayToSheet([[date]], { cellDates: true, UTC: true });
		const wb = createWorkbook(ws, "S");
		wb.Workbook = { WBProps: { date1904 } };

		const bytes = await write(wb, { cellStyles: true });
		const numeric = await read(bytes, { cellText: false });
		const dated = await read(bytes, { cellDates: true, cellText: false });

		expect(numeric.Sheets.S.A1.v).toBeCloseTo(expectedSerial, 8);
		expect(dated.Sheets.S.A1).toMatchObject({ t: "d" });
		expect((dated.Sheets.S.A1.v as Date).toISOString()).toBe(date.toISOString());
	});

	it("keeps explicit ISO date cells independent of the workbook date system", async () => {
		const date = new Date("2026-01-15T12:34:56.000Z");
		const ws = arrayToSheet([[date]], { cellDates: true, UTC: true });
		const wb = createWorkbook(ws, "S");
		wb.Workbook = { WBProps: { date1904: true } };

		const bytes = await write(wb, { cellDates: true });
		const zip = await zipRead(bytes);
		const sheetXml = zipReadString(zip, "xl/worksheets/sheet1.xml");
		const result = await read(bytes, { cellDates: true, cellText: false });

		expect(sheetXml).toContain('<c r="A1" t="d"><v>2026-01-15T12:34:56.000Z</v></c>');
		expect((result.Sheets.S.A1.v as Date).toISOString()).toBe(date.toISOString());
	});
});

describe("XLSX read formatting", () => {
	it("does not retain redundant XF metadata for a default-style cell", async () => {
		const bytes = await write(createWorkbook(arrayToSheet([[42]]), "S"));
		const result = await read(bytes);

		expect(result.Sheets.S.A1).toMatchObject({ t: "n", v: 42, w: "42" });
		expect(result.Sheets.S.A1.XF).toBeUndefined();
	});

	it("applies implicit style index zero and explicit positive style indexes", async () => {
		const source = await write(createWorkbook(arrayToSheet([[45292, 12.5]]), "S"));
		const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFF0000"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="2"><xf numFmtId="14" fontId="1" fillId="0" borderId="0" applyNumberFormat="1" applyFont="1"/><xf numFmtId="2" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/></cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;
		const withStyles = await replaceZipPart(source, "xl/styles.xml", styles);
		const zip = await zipRead(withStyles);
		const sheetXml = zipReadString(zip, "xl/worksheets/sheet1.xml")!.replace('<c r="B1">', '<c r="B1" s="1">');
		const fixture = await replaceZipPart(withStyles, "xl/worksheets/sheet1.xml", sheetXml);

		const result = await read(fixture, { cellDates: true, cellNF: true, cellStyles: true });

		expect(result.Sheets.S.A1.t).toBe("d");
		expect((result.Sheets.S.A1.v as Date).toISOString()).toBe("2024-01-01T00:00:00.000Z");
		expect(result.Sheets.S.A1).toMatchObject({
			z: "m/d/yy",
			s: { font: { bold: true, color: { rgb: "FFFF0000" } } },
		});
		expect(result.Sheets.S.B1).toMatchObject({ t: "n", v: 12.5, z: "0.00", w: "12.50" });
	});

	it("keeps date conversion independent from cell text and honors dateNF", async () => {
		const ws = arrayToSheet([[45292, 45292, 0.5]]);
		setCellStyle(ws.A1, { numFmt: 14 });
		setCellStyle(ws.B1, { numFmt: "yyyy-mm-dd" });
		setCellStyle(ws.C1, { numFmt: 20 });
		const bytes = await write(createWorkbook(ws, "S"), { cellStyles: true });

		const withoutText = await read(bytes, { cellDates: true, cellText: false });
		const withText = await read(bytes, { cellDates: true, dateNF: "yyyy-mm-dd" });
		const numeric = await read(bytes, { cellDates: false, cellText: false });

		expect(withoutText.Sheets.S.A1.t).toBe("d");
		expect(withoutText.Sheets.S.B1.t).toBe("d");
		expect(withoutText.Sheets.S.C1).toMatchObject({ t: "n", v: 0.5 });
		expect(withoutText.Sheets.S.A1.w).toBeUndefined();
		expect(withoutText.Sheets.S.B1.w).toBeUndefined();
		expect(withText.Sheets.S.A1).toMatchObject({ t: "d", w: "2024-01-01" });
		expect(withText.Sheets.S.B1).toMatchObject({ t: "d", w: "2024-01-01" });
		expect(withText.Sheets.S.C1).toMatchObject({ t: "n", v: 0.5, w: "12:00" });
		expect(numeric.Sheets.S.A1).toMatchObject({ t: "n", v: 45292 });
		expect(numeric.Sheets.S.B1).toMatchObject({ t: "n", v: 45292 });
	});

	it.each([
		{ date1904: false, serial: 45292, expected: "2024-01-01T00:00:00.000Z" },
		{ date1904: true, serial: 0, expected: "1904-01-01T00:00:00.000Z" },
	])("converts styled dates using the $date1904 workbook date system", async ({ date1904, serial, expected }) => {
		const ws = arrayToSheet([[serial]]);
		setCellStyle(ws.A1, { numFmt: 14 });
		const wb = createWorkbook(ws, "S");
		wb.Workbook = { WBProps: { date1904 } };
		const result = await read(await write(wb, { cellStyles: true }), { cellDates: true, cellText: false });

		expect(result.Sheets.S.A1.t).toBe("d");
		expect((result.Sheets.S.A1.v as Date).toISOString()).toBe(expected);
	});

	it("keeps custom formats stable across interleaved workbook operations", async () => {
		const percentage = arrayToSheet([[0.25]]);
		setCellStyle(percentage.A1, { numFmt: "0.0%" });
		const first = await read(await write(createWorkbook(percentage, "S"), { cellStyles: true }), {
			cellText: false,
		});
		const retainedCell = { ...first.Sheets.S.A1 };

		const measurement = arrayToSheet([[0.25]]);
		setCellStyle(measurement.A1, { numFmt: '0.00 "kg"' });
		await read(await write(createWorkbook(measurement, "S"), { cellStyles: true }), { cellText: false });

		expect(formatCell(retainedCell)).toBe("25.0%");
		const rewritten = await read(await write(first, { cellStyles: true }), { cellText: false });
		expect(formatCell({ ...rewritten.Sheets.S.A1 })).toBe("25.0%");
	});
});

describe("XLSX roundtrip", () => {
	it("should preserve number values", async () => {
		const ws = arrayToSheet([[42, 3.14, -100]]);
		const wb = createWorkbook(ws, "Sheet1");
		const u8 = await write(wb, { type: "array" });
		const wb2 = await read(u8);
		const s = wb2.Sheets.Sheet1;
		expect((s as any).A1.v).toBe(42);
		expect((s as any).B1.v).toBeCloseTo(3.14);
		expect((s as any).C1.v).toBe(-100);
	});

	it("should preserve boolean values", async () => {
		const ws = arrayToSheet([[true, false]]);
		const wb = createWorkbook(ws, "Sheet1");
		const u8 = await write(wb, { type: "array" });
		const wb2 = await read(u8);
		const s = wb2.Sheets.Sheet1;
		expect((s as any).A1.v).toBe(true);
		expect((s as any).B1.v).toBe(false);
	});
});
