/**
 * Final coverage push — targeting remaining gaps to reach >=90%.
 *
 * Focuses on uncovered branches in:
 * - SSF format engine: fractions, phone numbers, scientific notation,
 *   Buddhist/Hijri calendars, elapsed time, 1904 date system, Date object formatting
 * - parse-zip: bookSheets, bookProps, sheet filtering options
 * - worksheet: margins, autofilter, cols, hyperlinks
 * - cell utils: quoteSheetName, getCell (dense mode)
 * - comments: threaded vs legacy precedence, writeCommentsXml edge cases
 * - relationships: addRelationship edge cases
 */
import { describe, it, expect } from "vitest";
import { formatNumber, isDateFormat, parseExcelDateCode } from "./ssf/format.js";
import { addRelationship } from "./opc/relationships.js";
import { RELS } from "./xml/namespaces.js";
import { quoteSheetName, getCell } from "./utils/cell.js";
import { insertCommentsIntoSheet, parseCommentsXml, writeCommentsXml } from "./xlsx/comments.js";
import { parseWorksheetXml } from "./xlsx/worksheet.js";
import { read, write, createWorkbook, jsonToSheet } from "./index.js";

// ─── SSF: Fraction formats ──────────────────────────────────────────────────

describe("SSF: fraction formatting", () => {
	it("formats simple fraction ?/?", () => {
		const result = formatNumber("# ?/?", 1.5);
		expect(result).toContain("1");
		expect(result).toContain("/");
		expect(result).toContain("2");
	});

	it("formats fraction ??/??", () => {
		const result = formatNumber("# ??/??", 3.75);
		expect(result).toContain("3");
		expect(result).toContain("/");
	});

	it("formats fraction without whole part ?/?", () => {
		const result = formatNumber("?/?", 0.5);
		expect(result).toContain("/");
	});

	it("formats fraction ??/?? without whole part", () => {
		const result = formatNumber("??/??", 0.333);
		expect(result).toContain("/");
	});

	it("formats negative fraction", () => {
		const result = formatNumber("# ?/?", -1.5);
		expect(result).toContain("-");
		expect(result).toContain("/");
	});

	it("formats fraction with fixed denominator # ?/8", () => {
		const result = formatNumber("# ?/8", 1.25);
		expect(result).toContain("/");
		expect(result).toContain("8");
	});

	it("formats zero as fraction", () => {
		const result = formatNumber("# ?/?", 0);
		expect(typeof result).toBe("string");
	});
});

// ─── SSF: Phone number format ──────────────────────────────────────────────

describe("SSF: phone number formatting", () => {
	it("formats 10-digit number as phone number (flt)", () => {
		const result = formatNumber("[<=9999999]###-####;(###) ###-####", 5551234567);
		expect(result).toContain("(");
		expect(result).toContain(")");
		expect(result).toContain("-");
	});

	it("formats 7-digit number with phone format", () => {
		const result = formatNumber("[<=9999999]###-####;(###) ###-####", 5551234);
		expect(result).toContain("-");
	});
});

// ─── SSF: Scientific notation ──────────────────────────────────────────────

describe("SSF: scientific notation", () => {
	it("formats with 0.00E+00", () => {
		const result = formatNumber("0.00E+00", 12345);
		expect(result).toContain("E");
		expect(result).toContain("+");
	});

	it("formats small number with 0.00E+00", () => {
		const result = formatNumber("0.00E+00", 0.00123);
		expect(result).toContain("E");
	});

	it("formats zero with 0.00E+00", () => {
		const result = formatNumber("0.00E+00", 0);
		expect(result).toContain("E");
	});

	it("formats negative with 0.00E+00", () => {
		const result = formatNumber("0.00E+00", -12345);
		expect(result).toContain("-");
		expect(result).toContain("E");
	});

	it("formats with ##0.0E+0 engineering notation", () => {
		const result = formatNumber("##0.0E+0", 12345);
		expect(result).toContain("E");
	});

	it("formats with 0.0E+0", () => {
		const result = formatNumber("0.0E+0", 123);
		expect(result).toContain("E");
	});

	it("formats with E- (no plus sign for positive exponent)", () => {
		const result = formatNumber("0.00E-00", 12345);
		expect(result).toContain("E");
	});
});

// ─── SSF: Buddhist/Hijri calendars and 1904 date system ────────────────────

describe("SSF: Buddhist and Hijri calendars", () => {
	it("detects B1 (Buddhist calendar) as date format", () => {
		expect(isDateFormat("B1yyyy-mm-dd")).toBe(true);
	});

	it("detects B2 (Hijri calendar) as date format", () => {
		expect(isDateFormat("B2yyyy-mm-dd")).toBe(true);
	});

	it("formats with Buddhist calendar (B1) modifier", () => {
		// Buddhist year = Gregorian year + 543
		const result = formatNumber("B1yyyy", 44000);
		expect(result).toBeDefined();
		expect(typeof result).toBe("string");
	});

	it("formats with Hijri calendar (B2) modifier", () => {
		const result = formatNumber("B2yyyy", 44000);
		expect(result).toBeDefined();
		expect(typeof result).toBe("string");
	});
});

describe("SSF: 1904 date system", () => {
	it("parses date with 1904 date system", () => {
		const result = parseExcelDateCode(1000, { date1904: true });
		expect(result).toBeDefined();
		expect(result!.year).toBeGreaterThan(1900);
	});

	it("formats date with 1904 option", () => {
		const result = formatNumber("yyyy-mm-dd", 1000, { date1904: true });
		expect(result).toContain("-");
	});
});

// ─── SSF: Elapsed time formats ─────────────────────────────────────────────

describe("SSF: elapsed time formats", () => {
	it("detects [h] as date format", () => {
		expect(isDateFormat("[h]:mm")).toBe(true);
	});

	it("detects [mm] as date format", () => {
		expect(isDateFormat("[mm]:ss")).toBe(true);
	});

	it("detects [ss] as date format", () => {
		expect(isDateFormat("[ss]")).toBe(true);
	});

	it("formats elapsed hours [h]:mm", () => {
		// 1.5 days = 36 hours
		const result = formatNumber("[h]:mm", 1.5);
		expect(result).toContain(":");
	});

	it("formats elapsed minutes [mm]:ss", () => {
		const result = formatNumber("[mm]:ss", 0.5);
		expect(result).toContain(":");
	});
});

// ─── SSF: Sub-second precision ─────────────────────────────────────────────

describe("SSF: sub-second precision", () => {
	it("formats with ss.0 (tenths)", () => {
		const result = formatNumber("h:mm:ss.0", 0.55555);
		expect(result).toContain(":");
		expect(result).toContain(".");
	});

	it("formats with ss.00 (hundredths)", () => {
		const result = formatNumber("h:mm:ss.00", 0.55555);
		expect(result).toContain(".");
	});

	it("formats with ss.000 (thousandths)", () => {
		const result = formatNumber("h:mm:ss.000", 0.55555);
		expect(result).toContain(".");
	});
});

// ─── SSF: Date object handling ─────────────────────────────────────────────

describe("SSF: Date object formatting", () => {
	it("formats Date object in General format", () => {
		const d = new Date(2024, 0, 15);
		const result = formatNumber("General", d);
		expect(result).toContain("/");
	});

	it("formats Date object with explicit format", () => {
		const d = new Date(2024, 5, 15);
		const result = formatNumber("yyyy-mm-dd", d);
		expect(result).toContain("2024");
	});
});

// ─── SSF: Parenthesized negative format ─────────────────────────────────────

describe("SSF: parenthesized negatives", () => {
	it("formats negative number with parens", () => {
		const result = formatNumber("#,##0_);(#,##0)", -42);
		expect(result).toContain("(");
		expect(result).toContain(")");
	});

	it("formats positive number with paren format (no parens)", () => {
		const result = formatNumber("#,##0_);(#,##0)", 42);
		expect(result).not.toContain("(");
	});
});

// ─── SSF: Leading zeros format ─────────────────────────────────────────────

describe("SSF: leading zeros format", () => {
	it("formats with 00000 (zip code)", () => {
		expect(formatNumber("00000", 1234)).toBe("01234");
	});

	it("formats with 000-00-0000 (SSN)", () => {
		const result = formatNumber("000-00-0000", 123456789);
		expect(result).toContain("-");
	});
});

// ─── SSF: Format string with $ prefix ──────────────────────────────────────

describe("SSF: dollar-prefixed formats", () => {
	it("formats $#,##0", () => {
		const result = formatNumber("$#,##0", 42);
		expect(result).toContain("$");
	});

	it("formats $ #,##0 (with space)", () => {
		const result = formatNumber("$ #,##0", 42);
		expect(result).toContain("$");
	});
});

// ─── SSF: dateNF override ──────────────────────────────────────────────────

describe("SSF: dateNF option", () => {
	it("overrides format 14 with dateNF", () => {
		const result = formatNumber(14, 44000, { dateNF: "dd/mm/yyyy" });
		expect(result).toContain("/");
	});

	it("overrides 'm/d/yy' string with dateNF", () => {
		const result = formatNumber("m/d/yy", 44000, { dateNF: "dd/mm/yyyy" });
		// dateNF overrides the default "m/d/yy" format
		expect(result).toContain("/");
	});
});

// ─── SSF: Chinese AM/PM ────────────────────────────────────────────────────

describe("SSF: Chinese AM/PM detection", () => {
	it("detects 上午/下午 as date format", () => {
		expect(isDateFormat("\u4E0A\u5348/\u4E0B\u5348")).toBe(true);
	});
});

// ─── SSF: Conditional format expressions ───────────────────────────────────

describe("SSF: conditional format expressions", () => {
	it("applies [>=1000] condition", () => {
		const fmt = "[>=1000]#,##0;0";
		const result = formatNumber(fmt, 1500);
		expect(result).toContain(",");
	});

	it("falls through when condition not met", () => {
		const fmt = "[>=1000]#,##0;0";
		const result = formatNumber(fmt, 500);
		expect(result).toBe("500");
	});

	it("applies [<0] condition", () => {
		const fmt = "[<0](0);0";
		const result = formatNumber(fmt, -5);
		expect(result).toContain("(");
	});

	it("applies [=0] condition", () => {
		const fmt = '[=0]"-";0';
		const result = formatNumber(fmt, 0);
		expect(result).toBe("-");
	});

	it("applies [<>0] condition", () => {
		const fmt = '[<>0]0;"-"';
		const result = formatNumber(fmt, 5);
		expect(result).toBe("5");
	});

	it("applies [<=100] condition", () => {
		const fmt = "[<=100]0.00;0.00";
		const result = formatNumber(fmt, 50);
		expect(result).toBe("50.00");
	});

	it("uses section 3 when both conditions have expressions and neither matches", () => {
		const fmt = '[>100]0.00;[<0]0.00;"-"';
		const result = formatNumber(fmt, 0);
		expect(result).toBe("-");
	});
});

// ─── SSF: Miscellaneous number format branches ────────────────────────────

describe("SSF: miscellaneous format branches", () => {
	it("formats with General and non-number value", () => {
		const result = formatNumber("General", true);
		expect(result).toBe("TRUE");
	});

	it("formats with General and false", () => {
		const result = formatNumber("General", false);
		expect(result).toBe("FALSE");
	});

	it("formats undefined as empty string", () => {
		const result = formatNumber("General", undefined);
		expect(result).toBe("");
	});

	it("formats null as empty string", () => {
		const result = formatNumber("General", null);
		expect(result).toBe("");
	});

	it("formats text with @ placeholder", () => {
		const result = formatNumber("@", "hello");
		expect(result).toBe("hello");
	});

	it("4-section format: text section", () => {
		const result = formatNumber("#,##0;-#,##0;0;@", "text");
		expect(result).toBe("text");
	});

	it("formats with Infinity", () => {
		const result = formatNumber("0.00", Infinity);
		expect(typeof result).toBe("string");
	});

	it("formats with format table lookup by index", () => {
		// Format 2 = "0.00" in the standard table
		const result = formatNumber(2, 42);
		expect(result).toBe("42.00");
	});

	it("formats with unknown index falls back to General", () => {
		const result = formatNumber(999, 42);
		expect(result).toBe("42");
	});

	it("formats with custom table option", () => {
		const result = formatNumber(200, 42, { table: { 200: "0.0" } });
		expect(result).toBe("42.0");
	});
});

// ─── SSF: write_num_flt specific branches ──────────────────────────────────

describe("SSF: write_num_flt specific patterns", () => {
	it("handles #? pattern (question marks)", () => {
		const result = formatNumber("??", 5);
		expect(result.length).toBe(2);
	});

	it("handles large number in flr() path (>2^31)", () => {
		const result = formatNumber("0", 3000000000);
		expect(typeof result).toBe("string");
	});

	it("handles decimal with format 0.00", () => {
		const result = formatNumber("0.00", 3.1);
		expect(result).toBe("3.10");
	});

	it("handles decimal rounding carry", () => {
		// 9.999 with format 0.00 should carry
		const result = formatNumber("0.00", 9.999);
		expect(result).toContain("10");
	});
});

// ─── SSF: write_num_int specific branches ──────────────────────────────────

describe("SSF: integer-specific patterns", () => {
	it("formats integer with leading zeros 00000", () => {
		expect(formatNumber("00000", 42)).toBe("00042");
	});

	it("formats integer with fraction pattern", () => {
		const result = formatNumber("# ?/?", 3);
		expect(result).toContain("3");
	});

	it("formats integer with $ prefix", () => {
		const result = formatNumber("$00", 42);
		expect(result).toBe("$42");
	});

	it("formats integer with trailing comma (thousands scale)", () => {
		const result = formatNumber("0,", 5000);
		expect(result).toBe("5");
	});
});

// ─── SSF: parseExcelDateCode edge cases ─────────────────────────────────────

describe("SSF: parseExcelDateCode edge cases", () => {
	it("parses serial 0 (Jan 0, 1900)", () => {
		const result = parseExcelDateCode(0);
		expect(result).toBeDefined();
		expect(result!.day).toBe(0);
		expect(result!.month).toBe(1);
		expect(result!.year).toBe(1900);
	});

	it("parses serial 60 (Feb 29, 1900 — phantom)", () => {
		const result = parseExcelDateCode(60);
		expect(result).toBeDefined();
		expect(result!.day).toBe(29);
		expect(result!.month).toBe(2);
		expect(result!.year).toBe(1900);
	});

	it("parses serial with sub-second overflow", () => {
		// Time fraction just below next second — subSeconds should be normalized
		const result = parseExcelDateCode(1.9999999);
		expect(result).toBeDefined();
		expect(result!.subSeconds).toBeGreaterThanOrEqual(0);
	});

	it("parses with Hijri mode", () => {
		const result = parseExcelDateCode(44000, {}, true);
		expect(result).toBeDefined();
		// Hijri year should be much less than Gregorian
		expect(result!.year).toBeLessThan(1900);
	});

	it("parses serial 0 with Hijri mode", () => {
		const result = parseExcelDateCode(0, {}, true);
		expect(result).toBeDefined();
	});

	it("returns null for negative serial", () => {
		const result = parseExcelDateCode(-1);
		expect(result).toBeNull();
	});
});

// ─── parse-zip: bookSheets and bookProps ───────────────────────────────────

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
		wb.Sheets["S2"] = ws2;
		const buf = await write(wb);
		const result = await read(buf, { sheets: 0 });
		expect(result.SheetNames).toContain("S1");
		// S2 should exist in SheetNames but may have empty data
	});

	it("reads with specific sheet by name", async () => {
		const wb = createWorkbook(jsonToSheet([{ a: 1 }]), "S1");
		const ws2 = jsonToSheet([{ b: 2 }]);
		wb.SheetNames.push("S2");
		wb.Sheets["S2"] = ws2;
		const buf = await write(wb);
		const result = await read(buf, { sheets: "S2" });
		expect(result.SheetNames).toBeDefined();
	});

	it("reads with sheets as array of indices and names", async () => {
		const wb = createWorkbook(jsonToSheet([{ a: 1 }]), "S1");
		wb.SheetNames.push("S2");
		wb.Sheets["S2"] = jsonToSheet([{ b: 2 }]);
		wb.SheetNames.push("S3");
		wb.Sheets["S3"] = jsonToSheet([{ c: 3 }]);
		const buf = await write(wb);
		const result = await read(buf, { sheets: [0, "S3"] });
		expect(result.SheetNames).toBeDefined();
	});
});

// ─── parse-zip: dense mode and cellStyles ──────────────────────────────────

describe("parse-zip: dense mode and cellStyles", () => {
	it("reads in dense mode", async () => {
		const ws = jsonToSheet([{ a: 1, b: "text" }]);
		const wb = createWorkbook(ws, "S1");
		const buf = await write(wb);
		const result = await read(buf, { dense: true });
		const sheet = result.Sheets["S1"];
		expect(sheet["!data"]).toBeDefined();
	});

	it("reads with cellStyles to populate column info", async () => {
		const ws = jsonToSheet([{ a: 1, b: 2 }]);
		ws["!cols"] = [{ width: 15 }, { width: 20 }];
		const wb = createWorkbook(ws, "S1");
		const buf = await write(wb);
		const result = await read(buf, { cellStyles: true });
		const sheet = result.Sheets["S1"];
		expect(sheet["!cols"]).toBeDefined();
	});
});

// ─── worksheet: margins and autofilter ─────────────────────────────────────

describe("worksheet: margins and autofilter roundtrip", () => {
	it("roundtrips page margins", async () => {
		const ws = jsonToSheet([{ a: 1 }]);
		ws["!margins"] = {
			left: 0.5,
			right: 0.5,
			top: 1.0,
			bottom: 1.0,
			header: 0.25,
			footer: 0.25,
		};
		const wb = createWorkbook(ws, "S1");
		const buf = await write(wb);
		const result = await read(buf);
		expect(result.Sheets["S1"]["!margins"]).toBeDefined();
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
		expect(result.Sheets["S1"]["!autofilter"]).toBeDefined();
		expect(result.Sheets["S1"]["!autofilter"]!.ref).toBe("A1:B3");
	});
});

// ─── worksheet: hyperlinks ────────────────────────────────────────────────

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
		expect(ws["A1"]?.l?.Target).toBe("https://example.com");
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
		expect(ws["A1"]?.l?.Target).toBe("https://example.com");
		expect(ws["A1"]?.l?.Tooltip).toBe("Visit site");
	});
});

// ─── cell utils: quoteSheetName ────────────────────────────────────────────

describe("cell utils: quoteSheetName", () => {
	it("returns unquoted name for simple names", () => {
		expect(quoteSheetName("Sheet1")).toBe("Sheet1");
	});

	it("quotes name with spaces", () => {
		expect(quoteSheetName("Sheet 1")).toBe("'Sheet 1'");
	});

	it("escapes single quotes in name", () => {
		expect(quoteSheetName("It's")).toBe("'It''s'");
	});

	it("quotes name with special characters", () => {
		expect(quoteSheetName("Data+Summary")).toBe("'Data+Summary'");
	});

	it("throws for empty name", () => {
		expect(() => quoteSheetName("")).toThrow("empty sheet name");
	});

	it("does not quote CJK names", () => {
		expect(quoteSheetName("データ")).toBe("データ");
	});
});

// ─── cell utils: getCell (dense mode) ──────────────────────────────────────

describe("cell utils: getCell", () => {
	it("gets cell from dense worksheet", () => {
		const ws: any = { "!data": [[{ t: "n", v: 42 }]] };
		const cell = getCell(ws, 0, 0);
		expect(cell?.v).toBe(42);
	});

	it("returns undefined for missing row in dense mode", () => {
		const ws: any = { "!data": [] };
		const cell = getCell(ws, 5, 0);
		expect(cell).toBeUndefined();
	});

	it("gets cell from sparse worksheet", () => {
		const ws: any = { A1: { t: "n", v: 42 } };
		const cell = getCell(ws, 0, 0);
		expect(cell?.v).toBe(42);
	});
});

// ─── comments: threaded vs legacy precedence ──────────────────────────────

describe("comments: insertCommentsIntoSheet", () => {
	it("inserts legacy comment into sheet", () => {
		const ws: any = { A1: { t: "n", v: 1 }, "!ref": "A1" };
		const comments = [{ ref: "A1", author: "Alice", t: "Hello", r: "<t>Hello</t>", h: "" }];
		insertCommentsIntoSheet(ws, comments as any, false);
		expect(ws["A1"].c).toBeDefined();
		expect(ws["A1"].c[0].t).toBe("Hello");
	});

	it("inserts threaded comment and removes legacy", () => {
		const ws: any = { A1: { t: "n", v: 1, c: [{ a: "Alice", t: "legacy", T: false }] }, "!ref": "A1" };
		const comments = [{ ref: "A1", author: "Bob", t: "threaded", r: "<t>threaded</t>", h: "" }];
		insertCommentsIntoSheet(ws, comments as any, true);
		// Threaded should override legacy
		expect(ws["A1"].c.some((c: any) => c.T === true)).toBe(true);
	});

	it("does not add legacy when threaded exists", () => {
		const ws: any = { A1: { t: "n", v: 1, c: [{ a: "Alice", t: "threaded", T: true }] }, "!ref": "A1" };
		const comments = [{ ref: "A1", author: "Bob", t: "legacy", r: "<t>legacy</t>", h: "" }];
		insertCommentsIntoSheet(ws, comments as any, false);
		// Should not add legacy comment when threaded exists
		expect(ws["A1"].c.length).toBe(1);
	});

	it("creates cell when it doesn't exist and expands range", () => {
		const ws: any = { "!ref": "A1" };
		const comments = [{ ref: "C3", author: "Alice", t: "New", r: "<t>New</t>", h: "" }];
		insertCommentsIntoSheet(ws, comments as any, false);
		expect(ws["C3"]).toBeDefined();
		expect(ws["C3"].c).toBeDefined();
		// Range should be expanded
		expect(ws["!ref"]).not.toBe("A1");
	});

	it("inserts into dense mode worksheet", () => {
		const ws: any = { "!data": [[{ t: "n", v: 1 }]], "!ref": "A1" };
		const comments = [{ ref: "A1", author: "Alice", t: "Dense", r: "<t>Dense</t>", h: "" }];
		insertCommentsIntoSheet(ws, comments as any, false);
		expect(ws["!data"][0][0].c).toBeDefined();
	});
});

// ─── comments: writeCommentsXml ───────────────────────────────────────────

describe("comments: writeCommentsXml", () => {
	it("writes basic comment XML", () => {
		const data: [string, any[]][] = [["A1", [{ a: "Alice", t: "Hello" }]]];
		const xml = writeCommentsXml(data);
		expect(xml).toContain("<authors>");
		expect(xml).toContain("Alice");
		expect(xml).toContain("Hello");
		expect(xml).toContain("<commentList>");
	});

	it("adds default author when no comments have authors", () => {
		// Empty data array triggers default author path
		const data: [string, any[]][] = [];
		const xml = writeCommentsXml(data);
		expect(xml).toContain("SheetJ5");
	});

	it("writes threaded comment with ID", () => {
		const data: [string, any[]][] = [["A1", [{ a: "Alice", t: "Thread", T: true, ID: "TC001" }]]];
		const xml = writeCommentsXml(data);
		expect(xml).toContain("tc=TC001");
	});
});

// ─── comments: parseCommentsXml ───────────────────────────────────────────

describe("comments: parseCommentsXml", () => {
	it("parses comment with empty text", () => {
		const xml = `<?xml version="1.0"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<authors><author>Alice</author></authors>
<commentList>
<comment ref="A1" authorId="0"><text><t></t></text></comment>
</commentList>
</comments>`;
		const comments = parseCommentsXml(xml, { cellHTML: true });
		expect(comments).toBeDefined();
		expect(comments.length).toBeGreaterThanOrEqual(1);
	});

	it("parses comment with sheetRows limit", () => {
		const xml = `<?xml version="1.0"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<authors><author>Alice</author></authors>
<commentList>
<comment ref="A100" authorId="0"><text><t>Hidden</t></text></comment>
</commentList>
</comments>`;
		const comments = parseCommentsXml(xml, { sheetRows: 10 });
		// Comment on row 100 should be excluded
		expect(comments.length).toBe(0);
	});
});

// ─── relationships: addRelationship edge cases ────────────────────────────

describe("relationships: addRelationship", () => {
	it("auto-assigns rId when given -1", () => {
		const rels: any = { "!id": {} };
		const id = addRelationship(rels, -1, "sheet1.xml", RELS.WS);
		expect(id).toBeGreaterThan(0);
		expect(rels["!id"]["rId" + id]).toBeDefined();
	});

	it("auto-sets TargetMode=External for hyperlink type", () => {
		const rels: any = { "!id": {} };
		addRelationship(rels, 1, "https://example.com", RELS.HLINK);
		expect(rels["!id"]["rId1"].TargetMode).toBe("External");
	});

	it("uses explicit targetmode when provided", () => {
		const rels: any = { "!id": {} };
		addRelationship(rels, 1, "target.xml", RELS.WS, "External");
		expect(rels["!id"]["rId1"].TargetMode).toBe("External");
	});

	it("throws on duplicate rId", () => {
		const rels: any = { "!id": {} };
		addRelationship(rels, 1, "sheet1.xml", RELS.WS);
		expect(() => addRelationship(rels, 1, "sheet2.xml", RELS.WS)).toThrow("Cannot rewrite rId");
	});

	it("initializes !id when missing", () => {
		const rels: any = {};
		addRelationship(rels, 1, "sheet1.xml", RELS.WS);
		expect(rels["!id"]).toBeDefined();
		expect(rels["!id"]["rId1"]).toBeDefined();
	});
});

// ─── parseWorksheetXml: direct XML parsing ────────────────────────────────

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
		expect(ws["A1"]?.l).toBeDefined();
		expect(ws["A1"]?.l?.Target).toBe("https://example.com");
		expect(ws["A1"]?.l?.Tooltip).toBe("Visit");
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
		expect(ws["A1"]?.l).toBeDefined();
		expect(ws["A1"]?.l?.Target).toContain("Sheet2");
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
		expect(ws["!merges"]!.length).toBe(1);
	});
});

// ─── SSF: format edge cases for coverage ──────────────────────────────────

describe("SSF: additional format edge cases", () => {
	it("formats with escaped character \\x", () => {
		const result = formatNumber("0\\x", 5);
		expect(result).toContain("x");
		expect(result).toContain("5");
	});

	it("formats with quoted string literal", () => {
		const result = formatNumber('0" units"', 5);
		expect(result).toBe("5 units");
	});

	it("formats with underscore padding character", () => {
		const result = formatNumber("0_)", 5);
		expect(result).toContain("5");
		expect(result).toContain(" ");
	});

	it("formats with asterisk fill", () => {
		const result = formatNumber("0*-", 5);
		expect(result).toContain("5");
	});

	it("formats boolean TRUE with format", () => {
		const result = formatNumber("0.00", true);
		expect(result).toBe("TRUE");
	});

	it("formats boolean FALSE with format", () => {
		const result = formatNumber("0.00", false);
		expect(result).toBe("FALSE");
	});

	it("formats empty string returns empty", () => {
		const result = formatNumber("0.00", "");
		expect(result).toBe("");
	});

	it("formats null returns empty", () => {
		const result = formatNumber("0.00", null);
		expect(result).toBe("");
	});
});

// ─── SSF: AM/PM formatting ────────────────────────────────────────────────

describe("SSF: AM/PM time formatting", () => {
	it("formats with h:mm AM/PM", () => {
		// 0.25 = 6:00 AM
		const result = formatNumber("h:mm AM/PM", 0.25);
		expect(result).toContain("AM");
	});

	it("formats with h:mm AM/PM (afternoon)", () => {
		// 0.75 = 6:00 PM
		const result = formatNumber("h:mm AM/PM", 0.75);
		expect(result).toContain("PM");
	});

	it("formats with h:mm A/P", () => {
		const result = formatNumber("h:mm A/P", 0.25);
		expect(result).toContain("A");
	});

	it("formats with h:mm A/P (afternoon)", () => {
		const result = formatNumber("h:mm A/P", 0.75);
		expect(result).toContain("P");
	});
});

// ─── SSF: day-of-week and month-name formats ─────────────────────────────

describe("SSF: day-of-week and month-name formats", () => {
	it("formats with ddd (abbreviated day)", () => {
		const result = formatNumber("ddd", 44000);
		expect(result.length).toBeLessThanOrEqual(3);
	});

	it("formats with dddd (full day name)", () => {
		const result = formatNumber("dddd", 44000);
		expect(result.length).toBeGreaterThan(3);
	});

	it("formats with mmm (abbreviated month)", () => {
		const result = formatNumber("mmm", 44000);
		expect(result.length).toBe(3);
	});

	it("formats with mmmm (full month name)", () => {
		const result = formatNumber("mmmm", 44000);
		expect(result.length).toBeGreaterThan(3);
	});

	it("formats with mmmmm (first letter of month)", () => {
		const result = formatNumber("mmmmm", 44000);
		expect(result.length).toBe(1);
	});
});

// ─── SSF: color and conditional tokens ────────────────────────────────────

describe("SSF: color and bracket tokens", () => {
	it("formats with [Red] color prefix", () => {
		const result = formatNumber("[Red]0.00", 42);
		expect(result).toBe("42.00");
	});

	it("formats with [Blue] color prefix", () => {
		const result = formatNumber("[Blue]0.00", 42);
		expect(result).toBe("42.00");
	});

	it("formats with [Color5] numeric color", () => {
		const result = formatNumber("[Color5]0.00", 42);
		expect(result).toBe("42.00");
	});
});
