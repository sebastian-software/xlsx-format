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

	it("escapes values in every generated HTML attribute context", () => {
		const ws: any = {
			"!ref": "A1",
			A1: {
				t: "s",
				v: 'line 1\nline 2 "quoted" & <tag>',
				z: '0 "units"\nnext',
				f: 'IF(A1="quoted","<&","")',
				l: { Target: 'https://example.test/?q="marker"&value=<tag>' },
			},
		};
		const html = sheetToHtml(ws, { id: 'grid" data-marker="table' });

		expect(html).toContain('data-v="line 1&#x000a;line 2 &quot;quoted&quot; &amp; &lt;tag&gt;"');
		expect(html).toContain('data-z="0 &quot;units&quot;&#x000a;next"');
		expect(html).toContain('data-f="IF(A1=&quot;quoted&quot;,&quot;&lt;&amp;&quot;,&quot;&quot;)"');
		expect(html).toContain('href="https://example.test/?q=&quot;marker&quot;&amp;value=&lt;tag&gt;"');
		expect(html).toContain('id="grid&quot; data-marker=&quot;table"');
		expect(html).toContain('id="grid&quot; data-marker=&quot;table-A1"');
		expect(html).not.toContain('id="grid" data-marker="table"');
	});

	it("keeps generated rich formatting and renders unsupported cell.h tags as text", () => {
		const ws: any = {
			"!ref": "A1",
			A1: {
				t: "s",
				v: "Bold Sized literal",
				h: '<b>Bold</b> <span style="font-size:12pt;position:fixed">Sized</span> <mark data-marker="x">literal</mark>',
			},
		};

		const html = sheetToHtml(ws);
		expect(html).toContain('<b>Bold</b> <span style="font-size:12pt;">Sized</span>');
		expect(html).toContain("&lt;mark data-marker=&quot;x&quot;&gt;literal&lt;/mark&gt;");
		expect(html).not.toContain("position:fixed");
		expect(html).not.toContain("<mark");
	});

	it("confines unbalanced and malformed cell.h markup to its cell", () => {
		const ws: any = {
			"!ref": "A1:B1",
			A1: { t: "s", v: "One", h: "<b>One</i><broken" },
			B1: { t: "s", v: "Two" },
		};

		const html = sheetToHtml(ws);
		expect(html).toContain("<b>One&lt;/i&gt;&lt;broken</b></td>");
		expect(html).toContain(">Two</td>");
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

	it("reconstructs typed values, formulas, and number formats", () => {
		const original: any = {
			"!ref": "A1:E1",
			A1: { t: "d", v: new Date("2026-09-08T12:34:56.000Z"), z: "yyyy-mm-dd" },
			B1: { t: "e", v: 0x07 },
			C1: { t: "n", v: 42, f: "SUM(40,2)", z: '0.00 "units"' },
			D1: { t: "s", v: "café &lt;\t" },
			E1: { t: "n", f: "SUM(40,2)", z: "0.00" },
		};

		const parsed: any = htmlToSheet(sheetToHtml(original));
		expect(parsed.A1).toMatchObject({ t: "d", z: "yyyy-mm-dd" });
		expect(parsed.A1.v).toBeInstanceOf(Date);
		expect(parsed.A1.v.toISOString()).toBe("2026-09-08T12:34:56.000Z");
		expect(parsed.B1).toMatchObject({ t: "e", v: 0x07 });
		expect(parsed.C1).toMatchObject({ t: "n", v: 42, f: "SUM(40,2)", z: '0.00 "units"' });
		expect(parsed.D1).toMatchObject({ t: "s", v: "café &lt;\t" });
		expect(parsed.E1).toMatchObject({ t: "n", f: "SUM(40,2)", z: "0.00" });
		expect("v" in parsed.E1).toBe(false);
	});

	it("decodes named and numeric entities once and preserves br line breaks", () => {
		const html = `<table><tr>
			<td data-t="s" data-v="caf&#233; &#x0009; &amp;lt;">ignored</td>
			<td>line 1<br>line 2 &amp;amp;</td>
		</tr></table>`;
		const ws: any = htmlToSheet(html);

		expect(ws.A1).toMatchObject({ t: "s", v: "café \t &lt;" });
		expect(ws.B1.v).toBe("line 1\nline 2 &amp;");
	});

	it("supports single-quoted typed attributes and falls back for invalid metadata", () => {
		const html = `<table><tr>
			<td data-t='b' data-v='1'>ignored</td>
			<td data-t="d" data-v="not-a-date">42</td>
			<td data-t="unknown" data-v="99">TRUE</td>
			<td data-t="z"></td>
		</tr></table>`;
		const ws: any = htmlToSheet(html);

		expect(ws.A1).toMatchObject({ t: "b", v: true });
		expect(ws.B1).toMatchObject({ t: "n", v: 42 });
		expect(ws.C1).toMatchObject({ t: "b", v: true });
		expect(ws.D1).toMatchObject({ t: "z" });
	});
});
