import { describe, expect, it } from "vitest";
import { read, sheetToCsv, write } from "./index.js";
import { zipWrite } from "./zip/index.js";

const sheetName = "Büro Süd";
const plainText = "Grüße Straße café 日本語 😀";
const richXml = "<r><rPr><b/></rPr><t>Crème </t></r><r><t>brûlée 東京 🎵</t></r>";

/** Literal OOXML keeps the reader regression independent of our XML writer. */
async function unicodeWorkbook(): Promise<Uint8Array> {
	const main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
	const rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
	const parts: Record<string, string> = {
		"[Content_Types].xml": `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`,
		"_rels/.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="${rel}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
		"xl/workbook.xml": `<workbook xmlns="${main}" xmlns:r="${rel}">
<workbookPr codeName="Büro"/>
<sheets><sheet name="Büro Süd" sheetId="1" r:id="rId1"/></sheets>
<definedNames><definedName name="Größe">'Büro Süd'!$A$1</definedName></definedNames>
</workbook>`,
		"xl/_rels/workbook.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="${rel}/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="${rel}/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`,
		"xl/worksheets/sheet1.xml": `<worksheet xmlns="${main}"><dimension ref="A1:C1"/><sheetData><row r="1">
<c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c>
</row></sheetData></worksheet>`,
		"xl/sharedStrings.xml": `<sst xmlns="${main}" count="3" uniqueCount="3">
<si><t>${plainText}</t></si><si>${richXml}</si><si><t>café &amp; th&#233; _x00FC_</t></si>
</sst>`,
	};
	const encoder = new TextEncoder();
	return zipWrite({
		files: Object.fromEntries(Object.entries(parts).map(([path, xml]) => [path, encoder.encode(xml)])),
	});
}

describe("Unicode XLSX input", () => {
	it.each([true, false])("preserves shared strings and rich text with cellHTML=%s", async (cellHTML) => {
		const wb = await read(await unicodeWorkbook(), { cellHTML, WTF: true });
		const ws = wb.Sheets[wb.SheetNames[0]];
		expect(ws.A1.v).toBe(plainText);
		expect(ws.A1.r).toBe(`<t>${plainText}</t>`);
		expect(ws.B1.v).toBe("Crème brûlée 東京 🎵");
		expect(ws.B1.r).toBe(richXml);
		expect(ws.C1.v).toBe("café & thé ü");
		if (cellHTML) {
			expect(ws.A1.h).toBe(plainText);
			expect(ws.B1.h).toBe('<span style=""><b>Crème </b></span>brûlée 東京 🎵');
		} else {
			expect(ws.A1.h).toBeUndefined();
			expect(ws.B1.h).toBeUndefined();
		}
		expect(sheetToCsv(ws)).toBe(`${plainText},Crème brûlée 東京 🎵,café & thé ü`);
	});

	it("preserves sheet names, workbook code names, and defined names", async () => {
		const wb = await read(await unicodeWorkbook(), { WTF: true });
		expect(wb.SheetNames).toStrictEqual([sheetName]);
		expect(wb.Sheets[sheetName]).toBeDefined();
		expect(wb.Workbook?.WBProps?.CodeName).toBe("Büro");
		expect(wb.Workbook?.Names).toStrictEqual([{ Name: "Größe", Ref: "'Büro Süd'!$A$1" }]);
	});

	it("preserves names in the sheet-only read path", async () => {
		const wb = await read(await unicodeWorkbook(), { bookSheets: true });
		expect(wb.SheetNames).toStrictEqual([sheetName]);
	});

	it.each([true, false])("preserves Unicode on re-export with bookSST=%s", async (bookSST) => {
		const original = await read(await unicodeWorkbook(), { WTF: true });
		const wb = await read(await write(original, { bookSST }), { WTF: true });
		expect(wb.SheetNames).toStrictEqual([sheetName]);
		expect(sheetToCsv(wb.Sheets[sheetName])).toBe(`${plainText},Crème brûlée 東京 🎵,café & thé ü`);
	});
});
