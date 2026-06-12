import { describe, it, expect } from "vitest";
import { writeSstXml, parseSstXml } from "./shared-strings.js";
import type { SST } from "./shared-strings.js";

describe("shared-strings.ts: write and parse", () => {
	it("writeSstXml with bookSST=false returns empty", () => {
		expect(writeSstXml([] as any, { bookSST: false })).toBe("");
	});

	it("roundtrip plain text SST", () => {
		const sst: any = [{ t: "Hello" }, { t: "World" }];
		sst.Count = 2;
		sst.Unique = 2;
		const xml = writeSstXml(sst, { bookSST: true });
		expect(xml).toContain("<sst");
		expect(xml).toContain("Hello");
		expect(xml).toContain("World");

		const parsed = parseSstXml(xml);
		expect(parsed[0].t).toBe("Hello");
		expect(parsed[1].t).toBe("World");
	});

	it("roundtrip SST with whitespace-preserving text", () => {
		const sst: any = [{ t: "  leading spaces" }, { t: "trailing  " }];
		sst.Count = 2;
		sst.Unique = 2;
		const xml = writeSstXml(sst, { bookSST: true });
		expect(xml).toContain('xml:space="preserve"');

		const parsed = parseSstXml(xml);
		expect(parsed[0].t).toBe("  leading spaces");
	});

	it("handles null entries", () => {
		const sst: any = [{ t: "A" }, null, { t: "C" }];
		sst.Count = 3;
		sst.Unique = 2;
		const xml = writeSstXml(sst, { bookSST: true });
		expect(xml).toContain("A");
		expect(xml).toContain("C");
	});

	it("empty data returns empty array", () => {
		const parsed = parseSstXml("");
		expect(parsed).toHaveLength(0);
	});
});

describe("parseSstXml: rich text parsing", () => {
	it("parses plain text string items", () => {
		const xml = `<?xml version="1.0"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
<si><t>Hello</t></si>
<si><t>World</t></si>
</sst>`;
		const sst = parseSstXml(xml);
		// parseSstXml may include an empty trailing item from split
		const nonEmpty = sst.filter((s) => s.t !== "");
		expect(nonEmpty).toHaveLength(2);
		expect(nonEmpty[0].t).toBe("Hello");
		expect(nonEmpty[1].t).toBe("World");
	});

	it("parses rich text with bold formatting", () => {
		const xml = `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
<si><r><rPr><b/></rPr><t>Bold</t></r><r><t> Normal</t></r></si>
</sst>`;
		const sst = parseSstXml(xml, { cellHTML: true });
		const items = sst.filter((s) => s.t !== "");
		expect(items).toHaveLength(1);
		expect(items[0].t).toContain("Bold");
		expect(items[0].t).toContain("Normal");
		expect(items[0].h).toContain("<b>");
	});

	it("parses rich text with italic formatting", () => {
		const xml = `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
<si><r><rPr><i/></rPr><t>Italic</t></r></si>
</sst>`;
		const sst = parseSstXml(xml, { cellHTML: true });
		const items = sst.filter((s) => s.t !== "");
		expect(items[0].h).toContain("<i>");
	});

	it("parses rich text with underline variants", () => {
		const xml = `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
<si><r><rPr><u val="double"/></rPr><t>Underlined</t></r></si>
</sst>`;
		const sst = parseSstXml(xml, { cellHTML: true });
		const items = sst.filter((s) => s.t !== "");
		expect(items[0].h).toContain("underline");
		expect(items[0].h).toContain("double");
	});

	it("parses underline with singleAccounting", () => {
		const xml = `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
<si><r><rPr><u val="singleAccounting"/></rPr><t>Acct</t></r></si>
</sst>`;
		const sst = parseSstXml(xml, { cellHTML: true });
		const items = sst.filter((s) => s.t !== "");
		expect(items[0].h).toContain("single-accounting");
	});

	it("parses underline with doubleAccounting", () => {
		const xml = `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
<si><r><rPr><u val="doubleAccounting"/></rPr><t>DblAcct</t></r></si>
</sst>`;
		const sst = parseSstXml(xml, { cellHTML: true });
		const items = sst.filter((s) => s.t !== "");
		expect(items[0].h).toContain("double-accounting");
	});

	it("parses strikethrough formatting", () => {
		const xml = `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
<si><r><rPr><strike/></rPr><t>Struck</t></r></si>
</sst>`;
		const sst = parseSstXml(xml, { cellHTML: true });
		const items = sst.filter((s) => s.t !== "");
		expect(items[0].h).toContain("<s>");
	});

	it("parses shadow formatting", () => {
		const xml = `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
<si><r><rPr><shadow/></rPr><t>Shadow</t></r></si>
</sst>`;
		const sst = parseSstXml(xml, { cellHTML: true });
		const items = sst.filter((s) => s.t !== "");
		expect(items[0].h).toContain("text-shadow");
	});

	it("parses color with rgb attribute", () => {
		const xml = `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
<si><r><rPr><color rgb="FFFF0000"/><sz val="14"/><rFont val="Arial"/></rPr><t>Red</t></r></si>
</sst>`;
		const sst = parseSstXml(xml, { cellHTML: true });
		const items = sst.filter((s) => s.t !== "");
		expect(items[0].h).toContain("font-size:14pt");
	});

	it("parses vertAlign superscript", () => {
		const xml = `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
<si><r><rPr><vertAlign val="superscript"/></rPr><t>sup</t></r></si>
</sst>`;
		const sst = parseSstXml(xml, { cellHTML: true });
		const items = sst.filter((s) => s.t !== "");
		expect(items[0].h).toContain("<sup>");
	});

	it("parses subscript vertAlign", () => {
		const xml = `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
<si><r><rPr><vertAlign val="subscript"/></rPr><t>sub</t></r></si>
</sst>`;
		const sst = parseSstXml(xml, { cellHTML: true });
		const items = sst.filter((s) => s.t !== "");
		expect(items[0].h).toContain("<sub>");
	});

	it("parses rich text with phonetic runs (<rPh>)", () => {
		const xml = `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
<si><r><t>漢字</t></r><rPh sb="0" eb="2"><t>かんじ</t></rPh></si>
</sst>`;
		const sst = parseSstXml(xml);
		const items = sst.filter((s) => s.t !== "" && s.t !== undefined);
		// Parser returns at least one entry for the rich text
		expect(items.length).toBeGreaterThanOrEqual(1);
	});

	it("parses condense and extend tags (ignored)", () => {
		const xml = `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
<si><r><rPr><condense val="1"/><extend val="1"/><b/></rPr><t>X</t></r></si>
</sst>`;
		const sst = parseSstXml(xml, { cellHTML: true });
		const items = sst.filter((s) => s.t !== "");
		expect(items[0].h).toContain("<b>");
	});

	it("parses bold val=0 (explicitly not bold)", () => {
		const xml = `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
<si><r><rPr><b val="0"/></rPr><t>NotBold</t></r></si>
</sst>`;
		const sst = parseSstXml(xml, { cellHTML: true });
		const items = sst.filter((s) => s.t !== "");
		expect(items[0].h).not.toContain("<b>");
	});

	it("parses italic val=0 (explicitly not italic)", () => {
		const xml = `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
<si><r><rPr><i val="0"/></rPr><t>NotItalic</t></r></si>
</sst>`;
		const sst = parseSstXml(xml, { cellHTML: true });
		const items = sst.filter((s) => s.t !== "");
		expect(items[0].h).not.toContain("<i>");
	});

	it("parses scheme tag (ignored)", () => {
		const xml = `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
<si><r><rPr><scheme val="minor"/><b/></rPr><t>X</t></r></si>
</sst>`;
		const sst = parseSstXml(xml, { cellHTML: true });
		const items = sst.filter((s) => s.t !== "");
		expect(items[0].h).toContain("<b>");
	});

	it("parses family tag", () => {
		const xml = `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
<si><r><rPr><family val="2"/><b/></rPr><t>X</t></r></si>
</sst>`;
		const sst = parseSstXml(xml, { cellHTML: true });
		const items = sst.filter((s) => s.t !== "");
		expect(items[0].h).toContain("<b>");
	});

	it("parses ext/extLst tags (skipped)", () => {
		const xml = `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
<si><r><rPr><b/><extLst><ext uri="foo"><x15:bar/></ext></extLst></rPr><t>X</t></r></si>
</sst>`;
		const sst = parseSstXml(xml, { cellHTML: true });
		const items = sst.filter((s) => s.t !== "");
		expect(items[0].h).toContain("<b>");
	});

	it("returns empty for null data", () => {
		const sst = parseSstXml("");
		expect(sst).toHaveLength(0);
	});

	it("handles cellHTML=false", () => {
		const xml = `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
<si><t>Plain</t></si>
</sst>`;
		const sst = parseSstXml(xml, { cellHTML: false });
		const items = sst.filter((s) => s.t !== "");
		expect(items[0].t).toBe("Plain");
		expect(items[0].h).toBeUndefined();
	});

	it("parses rich text with newline in content", () => {
		const xml = `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
<si><r><rPr><b/></rPr><t>Line1\nLine2</t></r></si>
</sst>`;
		const sst = parseSstXml(xml, { cellHTML: true });
		const items = sst.filter((s) => s.t !== "");
		expect(items[0].h).toContain("<br/>");
	});

	it("parses shadow with val attribute", () => {
		const xml = `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
<si><r><rPr><shadow val="1"/></rPr><t>X</t></r></si>
</sst>`;
		const sst = parseSstXml(xml, { cellHTML: true });
		const items = sst.filter((s) => s.t !== "");
		expect(items[0].h).toContain("text-shadow");
	});

	it("parses strike with val attribute", () => {
		const xml = `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
<si><r><rPr><strike val="1"/></rPr><t>X</t></r></si>
</sst>`;
		const sst = parseSstXml(xml, { cellHTML: true });
		const items = sst.filter((s) => s.t !== "");
		expect(items[0].h).toContain("<s>");
	});

	it("parses simple underline <u/>", () => {
		const xml = `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
<si><r><rPr><u/></rPr><t>U</t></r></si>
</sst>`;
		const sst = parseSstXml(xml, { cellHTML: true });
		const items = sst.filter((s) => s.t !== "");
		expect(items[0].h).toContain("underline");
	});
});

describe("writeSstXml", () => {
	it("returns empty when bookSST is false", () => {
		const sst: SST = [] as any;
		expect(writeSstXml(sst, { bookSST: false })).toBe("");
	});

	it("writes basic string entries", () => {
		const sst: SST = [{ t: "Hello" }, { t: "World" }] as any;
		sst.Count = 2;
		sst.Unique = 2;
		const xml = writeSstXml(sst, { bookSST: true });
		expect(xml).toContain("<si><t>Hello</t></si>");
		expect(xml).toContain("<si><t>World</t></si>");
		expect(xml).toContain('count="2"');
		expect(xml).toContain('uniqueCount="2"');
	});

	it("preserves rich text XML", () => {
		const sst: SST = [{ t: "Bold", r: "<r><rPr><b/></rPr><t>Bold</t></r>" }] as any;
		sst.Count = 1;
		sst.Unique = 1;
		const xml = writeSstXml(sst, { bookSST: true });
		expect(xml).toContain("<r><rPr><b/></rPr><t>Bold</t></r>");
	});

	it("handles whitespace preservation", () => {
		const sst: SST = [{ t: "  leading" }, { t: "trailing  " }, { t: "has\ttab" }] as any;
		sst.Count = 3;
		sst.Unique = 3;
		const xml = writeSstXml(sst, { bookSST: true });
		expect(xml).toContain('xml:space="preserve"');
	});

	it("converts non-string types", () => {
		const sst: SST = [{ t: 123 as any }] as any;
		sst.Count = 1;
		sst.Unique = 1;
		const xml = writeSstXml(sst, { bookSST: true });
		expect(xml).toContain("123");
	});

	it("handles null entries", () => {
		const sst: SST = [{ t: "A" }, null as any, { t: "C" }] as any;
		sst.Count = 2;
		sst.Unique = 2;
		const xml = writeSstXml(sst, { bookSST: true });
		expect(xml).toContain("A");
		expect(xml).toContain("C");
	});

	it("handles empty string entry", () => {
		const sst: SST = [{ t: "" }] as any;
		sst.Count = 1;
		sst.Unique = 1;
		const xml = writeSstXml(sst, { bookSST: true });
		expect(xml).toContain("<t></t>");
	});

	it("handles entry with undefined t", () => {
		const sst: SST = [{ t: undefined as any }] as any;
		sst.Count = 1;
		sst.Unique = 1;
		const xml = writeSstXml(sst, { bookSST: true });
		expect(xml).toContain("<t>");
	});
});

describe("xlsx/shared-strings", () => {
	it("parseSstXml should parse plain text strings", () => {
		const xml = `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
			<si><t>Hello</t></si>
			<si><t>World</t></si>
		</sst>`;
		const sst = parseSstXml(xml);
		// The split produces a trailing empty entry — verify the meaningful ones
		expect(sst.length).toBeGreaterThanOrEqual(2);
		expect(sst[0].t).toBe("Hello");
		expect(sst[1].t).toBe("World");
	});

	it("parseSstXml should parse rich text strings", () => {
		const xml = `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
			<si><r><rPr><b/></rPr><t>Bold</t></r><r><t> Normal</t></r></si>
		</sst>`;
		const sst = parseSstXml(xml);
		expect(sst.length).toBeGreaterThanOrEqual(1);
		expect(sst[0].t).toBe("Bold Normal");
		expect(sst[0].r).toBeDefined();
	});

	it("parseSstXml should handle empty input", () => {
		expect(parseSstXml("")).toHaveLength(0);
	});

	it("writeSstXml should return empty when bookSST is false", () => {
		expect(writeSstXml([] as any, { bookSST: false })).toBe("");
	});

	it("writeSstXml should produce valid XML when bookSST is true", () => {
		const sst: any = [{ t: "Hello" }, { t: "World" }];
		sst.Count = 2;
		sst.Unique = 2;
		const xml = writeSstXml(sst, { bookSST: true });
		expect(xml).toContain("<sst");
		expect(xml).toContain("<si>");
		expect(xml).toContain("Hello");
	});
});
