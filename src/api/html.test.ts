import { describe, it, expect } from "vitest";
import { arrayToSheet, sheetToHtml, sheetToJson, htmlToSheet } from "../index.js";

describe("html.ts — sheetToHtml", () => {
	it("should handle merged cells", () => {
		const ws: any = {
			"!ref": "A1:B2",
			"!merges": [{ s: { r: 0, c: 0 }, e: { r: 1, c: 1 } }],
			A1: { t: "s", v: "Merged" },
		};
		const html = sheetToHtml(ws);
		expect(html).toContain('rowspan="2"');
		expect(html).toContain('colspan="2"');
	});

	it("should handle NaN as #NUM! and Infinity as #DIV/0!", () => {
		const ws: any = {
			"!ref": "A1:B1",
			A1: { t: "n", v: NaN },
			B1: { t: "n", v: Infinity },
		};
		const html = sheetToHtml(ws);
		// NaN → error 0x24 = #NUM!, Infinity → error 0x07 = #DIV/0!
		expect(html).toContain("#NUM!");
		expect(html).toContain("#DIV/0!");
	});

	it("should support editable mode", () => {
		const ws: any = {
			"!ref": "A1",
			A1: { t: "s", v: "text" },
		};
		const html = sheetToHtml(ws, { editable: true });
		expect(html).toContain('contenteditable="true"');
	});

	it("should include data attributes", () => {
		const ws: any = {
			"!ref": "A1",
			A1: { t: "n", v: 42, z: "#,##0", f: "=6*7" },
		};
		const html = sheetToHtml(ws);
		expect(html).toContain('data-v="42"');
		expect(html).toContain('data-f="=6*7"');
		expect(html).toContain('data-z="#,##0"');
	});

	it("should render hyperlinks", () => {
		const ws: any = {
			"!ref": "A1",
			A1: { t: "s", v: "Link", l: { Target: "https://example.com" } },
		};
		const html = sheetToHtml(ws);
		expect(html).toContain('href="https://example.com"');
	});

	it("should sanitize javascript: links", () => {
		const ws: any = {
			"!ref": "A1",
			A1: { t: "s", v: "Bad", l: { Target: "javascript:alert(1)" } },
		};
		const html = sheetToHtml(ws, { sanitizeLinks: true });
		expect(html).not.toContain("javascript:");
	});

	it("should skip internal links (#)", () => {
		const ws: any = {
			"!ref": "A1",
			A1: { t: "s", v: "Internal", l: { Target: "#Sheet2!A1" } },
		};
		const html = sheetToHtml(ws);
		expect(html).not.toContain("href=");
	});

	it("should support custom header/footer", () => {
		const ws = arrayToSheet([["x"]]);
		const html = sheetToHtml(ws, { header: "<div>", footer: "</div>" });
		expect(html.startsWith("<div>")).toBe(true);
		expect(html.endsWith("</div>")).toBe(true);
	});

	it("should support table id", () => {
		const ws = arrayToSheet([["x"]]);
		const html = sheetToHtml(ws, { id: "mytable" });
		expect(html).toContain('id="mytable"');
	});
});

describe("html.ts — htmlToSheet", () => {
	it("should handle rowspan", () => {
		const html = `<table>
			<tr><td rowspan="2">A</td><td>B</td></tr>
			<tr><td>C</td></tr>
		</table>`;
		const ws = htmlToSheet(html);
		const rows = sheetToJson(ws, { header: 1 });
		expect(rows[0][0]).toBe("A");
		expect(rows[0][1]).toBe("B");
		expect(rows[1][1]).toBe("C");
	});

	it("should handle colspan", () => {
		const html = `<table>
			<tr><td colspan="3">Wide</td></tr>
			<tr><td>A</td><td>B</td><td>C</td></tr>
		</table>`;
		const ws = htmlToSheet(html);
		const rows = sheetToJson(ws, { header: 1 });
		expect(rows[0][0]).toBe("Wide");
	});

	it("should coerce types from text", () => {
		const html = `<table>
			<tr><td>42</td><td>TRUE</td><td>hello</td></tr>
		</table>`;
		const ws = htmlToSheet(html);
		expect((ws as any).A1.v).toBe(42);
		expect((ws as any).B1.v).toBe(true);
		expect((ws as any).C1.v).toBe("hello");
	});

	it("should handle data-t and data-v attributes", () => {
		const html = `<table>
			<tr><td data-t="n" data-v="99">formatted</td></tr>
		</table>`;
		const ws = htmlToSheet(html);
		expect((ws as any).A1.v).toBe(99);
	});

	it("should return empty sheet for no table", () => {
		const ws = htmlToSheet("<div>no table</div>");
		expect(ws["!ref"]).toBeUndefined();
	});

	it("should handle combined rowspan and colspan", () => {
		const html = `<table>
			<tr><td rowspan="2" colspan="2">Big</td><td>C</td></tr>
			<tr><td>D</td></tr>
			<tr><td>E</td><td>F</td><td>G</td></tr>
		</table>`;
		const ws = htmlToSheet(html);
		const rows = sheetToJson(ws, { header: 1 });
		expect(rows[0][0]).toBe("Big");
		expect(rows[2][0]).toBe("E");
		expect(rows[2][2]).toBe("G");
	});

	it("should unescape HTML entities", () => {
		const html = `<table><tr><td>&lt;b&gt;bold&lt;/b&gt;</td></tr></table>`;
		const ws = htmlToSheet(html);
		expect((ws as any).A1.v).toBe("<b>bold</b>");
	});
});
