import { describe, expect, it } from "vitest";
import { addCellComment, arrayToSheet, createWorkbook, read, write } from "./index.js";
import { zipHas, zipRead, zipReadString } from "./zip/index.js";

function denseCell(sheet: any, row = 0, column = 0): any {
	return sheet["!data"]?.[row]?.[column];
}

describe("comment XLSX roundtrips", () => {
	it.each([false, true])("writes public sparse comments with bookSST=%s", async (bookSST) => {
		const sheet = arrayToSheet([["Value"]]);
		addCellComment(sheet.A1, "<review> & café\nsecond line", "Alice & Co");
		sheet.A1.c!.hidden = false;
		const workbook = createWorkbook(sheet, "Comments");

		const bytes = await write(workbook, { bookSST });
		const zip = await zipRead(bytes);
		const worksheetXml = zipReadString(zip, "xl/worksheets/sheet1.xml")!;
		const relationshipsXml = zipReadString(zip, "xl/worksheets/_rels/sheet1.xml.rels")!;
		const commentsXml = zipReadString(zip, "xl/comments1.xml")!;
		const vml = zipReadString(zip, "xl/drawings/vmlDrawing1.vml")!;

		expect(worksheetXml).toMatch(/<legacyDrawing r:id="rId\d+"\/>/);
		expect(relationshipsXml).toContain("/comments");
		expect(relationshipsXml).toContain("/vmlDrawing");
		expect(commentsXml).toContain("<author>Alice &amp; Co</author>");
		expect(commentsXml).toContain("&lt;review&gt; &amp; café\nsecond line");
		expect(commentsXml).not.toContain("&amp;lt;review&amp;gt;");
		expect(vml).toContain("<x:Visible/>");

		const roundtripped = await read(bytes);
		const comment = roundtripped.Sheets.Comments.A1.c!;
		expect(comment).toHaveLength(1);
		expect(comment[0]).toMatchObject({ a: "Alice & Co", t: "<review> & café\nsecond line" });
		expect(comment.hidden).toBe(false);
	});

	it("collects comments from dense cells without scanning sparse array lengths", async () => {
		const sheet = arrayToSheet([["Dense"]], { dense: true });
		const cell = denseCell(sheet);
		addCellComment(cell, "Dense note", "Dana");
		cell.c.hidden = true;
		// A sparse array length alone must not trigger a linear scan.
		(sheet["!data"] as any[]).length = 1_000_000;
		const workbook = createWorkbook(sheet, "Dense");

		const bytes = await write(workbook);
		const result = await read(bytes, { dense: true });
		const comments = denseCell(result.Sheets.Dense).c;
		expect(comments).toHaveLength(1);
		expect(comments[0]).toMatchObject({ a: "Dana", t: "Dense note" });
		expect(comments.hidden).toBe(true);
	});

	it("preserves imported threaded comments across repeated saves without mutating the cells", async () => {
		const sheet = arrayToSheet([["Thread"]]);
		const author = 'A & "B" <C>';
		addCellComment(sheet.A1, "<root> & Grüß dich\nnext", author);
		addCellComment(sheet.A1, "reply & <more>", "Rémy");
		for (const comment of sheet.A1.c!) {
			comment.T = true;
		}
		sheet.A1.c!.hidden = true;

		const imported = await read(await write(createWorkbook(sheet, "Threaded")));
		const importedComments = imported.Sheets.Threaded.A1.c!;
		const before = importedComments.map((comment) => ({ ...comment }));
		const beforeHidden = importedComments.hidden;

		const secondBytes = await write(imported);
		const thirdBytes = await write(imported);
		expect(importedComments.map((comment) => ({ ...comment }))).toStrictEqual(before);
		expect(importedComments.hidden).toBe(beforeHidden);
		expect(importedComments.every((comment: any) => comment.ID == null)).toBe(true);

		const zip = await zipRead(secondBytes);
		const threadedXml = zipReadString(zip, "xl/threadedComments/threadedComment1.xml")!;
		const fallbackXml = zipReadString(zip, "xl/comments1.xml")!;
		const peopleXml = zipReadString(zip, "xl/persons/person.xml")!;
		expect(zipHas(zip, "xl/drawings/vmlDrawing1.vml")).toBe(true);
		expect(threadedXml).toContain("&lt;root&gt; &amp; Grüß dich\nnext");
		expect(threadedXml).not.toContain("&amp;lt;root&amp;gt;");
		expect(fallbackXml).toContain("&lt;root&gt; &amp; Grüß dich\nnext");
		expect(fallbackXml).not.toContain("&amp;lt;root&amp;gt;");
		expect(peopleXml).toContain('displayName="A &amp; &quot;B&quot; &lt;C&gt;"');

		for (const bytes of [secondBytes, thirdBytes]) {
			const result = await read(bytes);
			const comments = result.Sheets.Threaded.A1.c!;
			expect(comments).toHaveLength(2);
			expect(comments[0]).toMatchObject({ a: author, t: "<root> & Grüß dich\nnext", T: true });
			expect(comments[1]).toMatchObject({ a: "Rémy", t: "reply & <more>", T: true });
			expect(comments.hidden).toBe(true);
		}
	});

	it("ignores non-cell worksheet properties that happen to contain comments", async () => {
		const sheet: any = arrayToSheet([["Value"]]);
		sheet.custom = { c: [{ a: "Wrong", t: "metadata" }] };
		sheet.XFE1 = { t: "s", v: "outside columns", c: [{ a: "Wrong", t: "outside columns" }] };
		sheet.A1048577 = { t: "s", v: "outside rows", c: [{ a: "Wrong", t: "outside rows" }] };
		const zip = await zipRead(await write(createWorkbook(sheet, "NoComments"), { unsafe: true } as any));
		expect(zipHas(zip, "xl/comments1.xml")).toBe(false);
		expect(zipHas(zip, "xl/drawings/vmlDrawing1.vml")).toBe(false);
	});
});
