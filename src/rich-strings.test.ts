import { describe, expect, it } from "vitest";
import { read, write } from "./index.js";
import type { WorkBook, WorkSheet } from "./types.js";
import { zipRead, zipReadString, zipWrite } from "./zip/index.js";

const main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const officeRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const packageRel = "http://schemas.openxmlformats.org/package/2006/relationships";

const inlineRich =
	'<r><rPr><b/></rPr><t xml:space="preserve"> First </t></r><r><t>second</t></r><rPh sb="0" eb="6"><t>phonetic</t></rPh>';
const namespacedRich = "<x:r><x:rPr><x:i/></x:rPr><x:t>Namespaced</x:t></x:r>";

async function inlineStringWorkbook(): Promise<Uint8Array> {
	const parts: Record<string, string> = {
		"[Content_Types].xml": `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
		"_rels/.rels": `<Relationships xmlns="${packageRel}">
<Relationship Id="rId1" Type="${officeRel}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
		"xl/workbook.xml": `<workbook xmlns="${main}" xmlns:r="${officeRel}">
<sheets><sheet name="Rich" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
		"xl/_rels/workbook.xml.rels": `<Relationships xmlns="${packageRel}">
<Relationship Id="rId1" Type="${officeRel}/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
		"xl/worksheets/sheet1.xml": `<worksheet xmlns="${main}" xmlns:x="urn:test"><dimension ref="A1:C1"/><sheetData><row r="1">
<c r="A1" t="inlineStr"><is>${inlineRich}</is></c>
<c r="B1" t="inlineStr"><x:is>${namespacedRich}</x:is></c>
<c r="C1" t="inlineStr"><is><t xml:space="preserve"> plain </t></is></c>
</row></sheetData></worksheet>`,
	};
	const encoder = new TextEncoder();
	return zipWrite({
		files: Object.fromEntries(Object.entries(parts).map(([path, xml]) => [path, encoder.encode(xml)])),
	});
}

function richWorkbook(): WorkBook {
	const bold = "<r><rPr><b/></rPr><t>styled</t></r>";
	const italic = "<r><rPr><i/></rPr><t>styled</t></r>";
	const first: WorkSheet = {
		"!ref": "A1:C1",
		A1: { t: "s", v: "repeat" },
		B1: { t: "s", v: "repeat" },
		C1: { t: "s", v: "styled", r: bold },
	};
	const second: WorkSheet = {
		"!ref": "A1:C1",
		A1: { t: "s", v: "styled", r: bold },
		B1: { t: "s", v: "styled", r: italic },
		C1: { t: "s", v: "other" },
	};
	return { SheetNames: ["First", "Second"], Sheets: { First: first, Second: second } };
}

describe("rich strings", () => {
	it("reads every visible inline-string run and preserves rich representations", async () => {
		const bytes = await inlineStringWorkbook();
		const withHtml = await read(bytes, { WTF: true });
		const ws = withHtml.Sheets.Rich;
		expect(ws.A1).toMatchObject({
			t: "s",
			v: " First second",
			r: inlineRich,
			h: '<span style=""><b> First </b></span>second',
		});
		expect(ws.B1).toMatchObject({
			t: "s",
			v: "Namespaced",
			r: namespacedRich,
			h: '<span style=""><i>Namespaced</i></span>',
		});
		expect(ws.C1).toMatchObject({ t: "s", v: " plain ", r: '<t xml:space="preserve"> plain </t>', h: " plain " });

		const withoutHtml = await read(bytes, { cellHTML: false, WTF: true });
		expect(withoutHtml.Sheets.Rich.A1).toMatchObject({ t: "s", v: " First second", r: inlineRich });
		expect(withoutHtml.Sheets.Rich.A1.h).toBeUndefined();
	});

	it("writes a workbook-wide SST with rich-format-aware deduplication", async () => {
		const bytes = await write(richWorkbook(), { bookSST: true });
		const zip = await zipRead(bytes);
		const sst = zipReadString(zip, "xl/sharedStrings.xml")!;
		const first = zipReadString(zip, "xl/worksheets/sheet1.xml")!;
		const second = zipReadString(zip, "xl/worksheets/sheet2.xml")!;
		const rels = zipReadString(zip, "xl/_rels/workbook.xml.rels")!;
		const contentTypes = zipReadString(zip, "[Content_Types].xml")!;

		expect(sst).toContain('count="6" uniqueCount="4"');
		expect(sst).toContain("<si><t>repeat</t></si>");
		expect(sst).toContain("<si><r><rPr><b/></rPr><t>styled</t></r></si>");
		expect(sst).toContain("<si><r><rPr><i/></rPr><t>styled</t></r></si>");
		expect(first).toContain('<c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>0</v></c><c r="C1" t="s"><v>1</v></c>');
		expect(second).toContain(
			'<c r="A1" t="s"><v>1</v></c><c r="B1" t="s"><v>2</v></c><c r="C1" t="s"><v>3</v></c>',
		);
		expect(rels).toContain("/sharedStrings");
		expect(contentTypes).toContain("sharedStrings+xml");

		const roundtrip = await read(bytes, { WTF: true });
		expect(roundtrip.Sheets.First.A1.v).toBe("repeat");
		expect(roundtrip.Sheets.First.C1).toMatchObject({
			v: "styled",
			r: "<r><rPr><b/></rPr><t>styled</t></r>",
			h: '<span style=""><b>styled</b></span>',
		});
		expect(roundtrip.Sheets.Second.B1.r).toBe("<r><rPr><i/></rPr><t>styled</t></r>");
	});

	it("preserves rich runs inline when bookSST is disabled", async () => {
		const bytes = await write(richWorkbook(), { bookSST: false });
		const zip = await zipRead(bytes);
		const first = zipReadString(zip, "xl/worksheets/sheet1.xml")!;
		expect(zipReadString(zip, "xl/sharedStrings.xml")).toBeNull();
		expect(zipReadString(zip, "xl/_rels/workbook.xml.rels")).not.toContain("/sharedStrings");
		expect(zipReadString(zip, "[Content_Types].xml")).not.toContain("sharedStrings+xml");
		expect(first).toContain('<c r="A1" t="str"><v>repeat</v></c>');
		expect(first).toContain('<c r="C1" t="inlineStr"><is><r><rPr><b/></rPr><t>styled</t></r></is></c>');

		const roundtrip = await read(bytes, { WTF: true });
		expect(roundtrip.Sheets.First.C1).toMatchObject({
			v: "styled",
			r: "<r><rPr><b/></rPr><t>styled</t></r>",
			h: '<span style=""><b>styled</b></span>',
		});
	});
});
