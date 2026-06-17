import { describe, it, expect } from "vitest";
import {
	validateSheetName,
	validateWorkbook,
	is1904DateSystem,
	parseWorkbookXml,
	writeWorkbookXml,
} from "./workbook.js";
import { createWorkbook, appendSheet, jsonToSheet, setSheetVisibility } from "../index.js";
import type { WorkBook } from "../types.js";

describe("workbook.ts: validation", () => {
	it("validateSheetName rejects empty name", () => {
		expect(() => validateSheetName("")).toThrow("blank");
	});

	it("validateSheetName rejects too-long name", () => {
		expect(() => validateSheetName("a".repeat(32))).toThrow("exceed 31");
	});

	it("validateSheetName rejects apostrophe boundaries", () => {
		expect(() => validateSheetName("'Name")).toThrow("apostrophe");
		expect(() => validateSheetName("Name'")).toThrow("apostrophe");
	});

	it("validateSheetName rejects 'History'", () => {
		expect(() => validateSheetName("history")).toThrow("History");
	});

	it("validateSheetName rejects forbidden characters", () => {
		expect(() => validateSheetName("Sheet:1")).toThrow();
		expect(() => validateSheetName("Sheet[1]")).toThrow();
		expect(() => validateSheetName("Sheet*")).toThrow();
		expect(() => validateSheetName("Sheet?")).toThrow();
		expect(() => validateSheetName("Sheet/1")).toThrow();
		expect(() => validateSheetName("Sheet\\1")).toThrow();
	});

	it("validateSheetName safe mode returns false", () => {
		expect(validateSheetName("", true)).toBe(false);
		expect(validateSheetName("a".repeat(32), true)).toBe(false);
	});

	it("validateWorkbook rejects invalid structure", () => {
		expect(() => {
			validateWorkbook(null as any);
		}).toThrow("Invalid");
		expect(() => {
			validateWorkbook({ SheetNames: [], Sheets: {} });
		}).toThrow("empty");
		expect(() => {
			validateWorkbook({ SheetNames: ["A", "A"], Sheets: {} });
		}).toThrow("Duplicate");
	});

	it("is1904DateSystem returns false for non-1904 workbooks", () => {
		expect(is1904DateSystem({} as any)).toBe("false");
		expect(is1904DateSystem({ Workbook: {} } as any)).toBe("false");
		expect(is1904DateSystem({ Workbook: { WBProps: { date1904: false } } } as any)).toBe("false");
	});

	it("is1904DateSystem returns true for 1904 workbooks", () => {
		expect(is1904DateSystem({ Workbook: { WBProps: { date1904: true } } } as any)).toBe("true");
	});
});

describe("parseWorkbookXml: comprehensive", () => {
	const xmlns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

	it("parses basic workbook with sheets", () => {
		const xml = `<?xml version="1.0"?>
<workbook xmlns="${xmlns}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
<sheet name="Sheet1" sheetId="1" r:id="rId1"/>
</sheets>
</workbook>`;
		const wb = parseWorkbookXml(xml);
		expect(wb.Sheets).toHaveLength(1);
		expect(wb.Sheets[0].name).toBe("Sheet1");
		expect(wb.Sheets[0].Hidden).toBe(0);
	});

	it("parses hidden and veryHidden sheets", () => {
		const xml = `<?xml version="1.0"?>
<workbook xmlns="${xmlns}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
<sheet name="Visible" sheetId="1" r:id="rId1"/>
<sheet name="Hidden" sheetId="2" r:id="rId2" state="hidden"/>
<sheet name="VeryHidden" sheetId="3" r:id="rId3" state="veryHidden"/>
</sheets>
</workbook>`;
		const wb = parseWorkbookXml(xml);
		expect(wb.Sheets[0].Hidden).toBe(0);
		expect(wb.Sheets[1].Hidden).toBe(1);
		expect(wb.Sheets[2].Hidden).toBe(2);
	});

	it("parses defined names with comment, localSheetId, hidden", () => {
		const xml = `<?xml version="1.0"?>
<workbook xmlns="${xmlns}">
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
<definedNames>
<definedName name="GlobalRange">Sheet1!$A$1:$A$10</definedName>
<definedName name="LocalRange" localSheetId="0">Sheet1!$B$1</definedName>
<definedName name="HiddenName" hidden="1">Sheet1!$C$1</definedName>
<definedName name="CommentedName" comment="A note">Sheet1!$D$1</definedName>
</definedNames>
</workbook>`;
		const wb = parseWorkbookXml(xml);
		expect(wb.Names).toHaveLength(4);
		expect(wb.Names[0].Name).toBe("GlobalRange");
		expect(wb.Names[0].Ref).toBe("Sheet1!$A$1:$A$10");
		expect(wb.Names[1].Sheet).toBe(0);
		expect(wb.Names[2].Hidden).toBe(true);
		expect(wb.Names[3].Comment).toBe("A note");
	});

	it("parses workbookPr with bool and int types", () => {
		const xml = `<?xml version="1.0"?>
<workbook xmlns="${xmlns}">
<workbookPr date1904="true" defaultThemeVersion="164011" filterPrivacy="true" codeName="MyWorkbook"/>
<sheets><sheet name="S1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`;
		const wb = parseWorkbookXml(xml);
		expect(wb.WBProps.date1904).toBe(true);
		expect(wb.WBProps.defaultThemeVersion).toBe(164011);
		expect(wb.WBProps.filterPrivacy).toBe(true);
		expect(wb.WBProps.CodeName).toBe("MyWorkbook");
	});

	it("parses fileVersion element", () => {
		const xml = `<?xml version="1.0"?>
<workbook xmlns="${xmlns}">
<fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="12345"/>
<sheets><sheet name="S1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`;
		const wb = parseWorkbookXml(xml);
		expect(wb.AppVersion.appname).toBe("xl");
	});

	it("parses calcPr element", () => {
		const xml = `<?xml version="1.0"?>
<workbook xmlns="${xmlns}">
<sheets><sheet name="S1" sheetId="1" r:id="rId1"/></sheets>
<calcPr calcId="191029" fullCalcOnLoad="true"/>
</workbook>`;
		const wb = parseWorkbookXml(xml);
		expect(wb.CalcPr.calcid).toBeTruthy();
	});

	it("parses workbookView with defaults applied", () => {
		const xml = `<?xml version="1.0"?>
<workbook xmlns="${xmlns}">
<bookViews><workbookView activeTab="1" firstSheet="1"/></bookViews>
<sheets><sheet name="S1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`;
		const wb = parseWorkbookXml(xml);
		expect(wb.WBView).toHaveLength(1);
		// After defaults, activeTab is coerced to int
		expect(wb.WBView[0].activeTab).toBe(1);
	});

	it("throws on empty data", () => {
		expect(() => parseWorkbookXml("")).toThrow("Could not find file");
	});

	it("throws on unknown namespace", () => {
		const xml = `<?xml version="1.0"?><workbook xmlns="http://example.com/unknown"><sheets><sheet name="S1" sheetId="1"/></sheets></workbook>`;
		expect(() => parseWorkbookXml(xml)).toThrow("Unknown Namespace");
	});

	it("parses with namespace prefix (<x:workbook>)", () => {
		const xml = `<?xml version="1.0"?>
<x:workbook xmlns:x="${xmlns}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<x:sheets>
<x:sheet name="S1" sheetId="1" r:id="rId1"/>
</x:sheets>
</x:workbook>`;
		const wb = parseWorkbookXml(xml);
		expect(wb.xmlns).toBe(xmlns);
		expect(wb.Sheets).toHaveLength(1);
	});
});

describe("writeWorkbookXml: comprehensive", () => {
	it("writes hidden sheets with state attribute", () => {
		const wb: WorkBook = createWorkbook(jsonToSheet([{ a: 1 }]), "Visible");
		appendSheet(wb, jsonToSheet([{ b: 2 }]), "Hidden");
		appendSheet(wb, jsonToSheet([{ c: 3 }]), "VeryHidden");
		// setSheetVisibility takes numeric values: 0=visible, 1=hidden, 2=veryHidden
		setSheetVisibility(wb, 1, 1);
		setSheetVisibility(wb, 2, 2);
		const xml = writeWorkbookXml(wb);
		expect(xml).toContain('state="hidden"');
		expect(xml).toContain('state="veryHidden"');
	});

	it("emits bookViews when first sheet is hidden", () => {
		const wb: WorkBook = createWorkbook(jsonToSheet([{ a: 1 }]), "Hidden1");
		appendSheet(wb, jsonToSheet([{ b: 2 }]), "Visible");
		setSheetVisibility(wb, 0, 1);
		const xml = writeWorkbookXml(wb);
		expect(xml).toContain("<bookViews>");
		expect(xml).toContain('activeTab="1"');
		expect(xml).toContain('firstSheet="1"');
	});

	it("handles all sheets hidden (activeTab=0 fallback)", () => {
		const wb: WorkBook = createWorkbook(jsonToSheet([{ a: 1 }]), "H1");
		appendSheet(wb, jsonToSheet([{ b: 2 }]), "H2");
		setSheetVisibility(wb, 0, 1);
		setSheetVisibility(wb, 1, 1);
		const xml = writeWorkbookXml(wb);
		expect(xml).toContain("<bookViews>");
		expect(xml).toContain('activeTab="0"');
	});

	it("writes defined names with comment", () => {
		const wb: WorkBook = createWorkbook(jsonToSheet([{ a: 1 }]), "Data");
		wb.Workbook = wb.Workbook || { Sheets: [] };
		wb.Workbook.Names = [{ Name: "TestName", Ref: "Data!$A$1:$A$10", Comment: "A comment" }];
		const xml = writeWorkbookXml(wb);
		expect(xml).toContain('comment="A comment"');
		expect(xml).toContain("TestName");
	});

	it("writes defined names with localSheetId", () => {
		const wb: WorkBook = createWorkbook(jsonToSheet([{ a: 1 }]), "Sheet1");
		wb.Workbook = wb.Workbook || { Sheets: [] };
		wb.Workbook.Names = [{ Name: "LocalName", Ref: "$A$1", Sheet: 0 }];
		const xml = writeWorkbookXml(wb);
		expect(xml).toContain('localSheetId="0"');
	});

	it("writes defined names with hidden flag", () => {
		const wb: WorkBook = createWorkbook(jsonToSheet([{ a: 1 }]), "Data");
		wb.Workbook = wb.Workbook || { Sheets: [] };
		wb.Workbook.Names = [{ Name: "HiddenName", Ref: "Data!$A$1", Hidden: true }];
		const xml = writeWorkbookXml(wb);
		expect(xml).toContain('hidden="1"');
	});

	it("skips defined names without Ref", () => {
		const wb: WorkBook = createWorkbook(jsonToSheet([{ a: 1 }]), "Data");
		wb.Workbook = wb.Workbook || { Sheets: [] };
		wb.Workbook.Names = [{ Name: "Valid", Ref: "Data!$A$1" }, { Name: "NoRef" } as any];
		const xml = writeWorkbookXml(wb);
		expect(xml).toContain("Valid");
		expect(xml).not.toContain("NoRef");
	});

	it("writes workbookPr with non-default properties", () => {
		const wb: WorkBook = createWorkbook(jsonToSheet([{ a: 1 }]), "S1");
		wb.Workbook = wb.Workbook || { Sheets: [] };
		wb.Workbook.WBProps = { date1904: true, CodeName: "CustomCode" };
		const xml = writeWorkbookXml(wb);
		expect(xml).toContain("date1904");
		expect(xml).toContain("CustomCode");
	});
});

describe("is1904DateSystem", () => {
	it("returns false when no Workbook", () => {
		expect(is1904DateSystem({} as any)).toBe("false");
	});

	it("returns false when no WBProps", () => {
		expect(is1904DateSystem({ Workbook: {} } as any)).toBe("false");
	});

	it("returns true when date1904 is true", () => {
		expect(is1904DateSystem({ Workbook: { WBProps: { date1904: true } } } as any)).toBe("true");
	});
});

describe("validateSheetName edge cases", () => {
	it("rejects names starting with apostrophe", () => {
		expect(() => validateSheetName("'Sheet")).toThrow(/apostrophe/);
	});

	it("rejects names ending with apostrophe", () => {
		expect(() => validateSheetName("Sheet'")).toThrow(/apostrophe/);
	});

	it("rejects History", () => {
		expect(() => validateSheetName("History")).toThrow(/History/);
	});

	it("returns false in safe mode for invalid names", () => {
		expect(validateSheetName("", true)).toBe(false);
		expect(validateSheetName("History", true)).toBe(false);
	});
});

describe("validateWorkbook edge cases", () => {
	it("throws on null workbook", () => {
		expect(() => {
			validateWorkbook(null as any);
		}).toThrow("Invalid Workbook");
	});

	it("throws on empty sheet names", () => {
		expect(() => {
			validateWorkbook({ SheetNames: [], Sheets: {} });
		}).toThrow("empty");
	});
});
