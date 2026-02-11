import { describe, it, expect } from "vitest";
import { escapeXml, unescapeXml, escapeHtml, htmlDecode, escapeXmlTag } from "./escape.js";
import { parseXmlTag, parseXmlBoolean, stripNamespace } from "./parser.js";

describe("escapeXml", () => {
	it("should escape ampersand", () => {
		expect(escapeXml("a & b")).toBe("a &amp; b");
	});

	it("should escape angle brackets", () => {
		expect(escapeXml("<div>")).toBe("&lt;div&gt;");
	});

	it("should escape quotes", () => {
		expect(escapeXml(`"hello" 'world'`)).toBe("&quot;hello&quot; &apos;world&apos;");
	});

	it("should encode control characters as _xHHHH_", () => {
		expect(escapeXml("\x01")).toBe("_x0001_");
		expect(escapeXml("\x1F")).toBe("_x001f_");
	});

	it("should pass through normal text unchanged", () => {
		expect(escapeXml("Hello World 123")).toBe("Hello World 123");
	});
});

describe("unescapeXml", () => {
	it("should unescape named entities", () => {
		expect(unescapeXml("&amp;")).toBe("&");
		expect(unescapeXml("&lt;")).toBe("<");
		expect(unescapeXml("&gt;")).toBe(">");
		expect(unescapeXml("&quot;")).toBe('"');
		expect(unescapeXml("&apos;")).toBe("'");
	});

	it("should unescape numeric character references", () => {
		expect(unescapeXml("&#65;")).toBe("A");
		expect(unescapeXml("&#x41;")).toBe("A");
	});

	it("should unescape _xHHHH_ OOXML escapes", () => {
		expect(unescapeXml("_x0041_")).toBe("A");
		expect(unescapeXml("_x0020_")).toBe(" ");
	});

	it("should handle CDATA sections", () => {
		expect(unescapeXml("<![CDATA[<raw>&text]]>")).toBe("<raw>&text");
	});

	it("should normalize line endings when xlsx=true", () => {
		expect(unescapeXml("a\r\nb", true)).toBe("a\nb");
	});

	it("should keep line endings when xlsx=false", () => {
		expect(unescapeXml("a\r\nb", false)).toBe("a\r\nb");
	});

	it("should roundtrip with escapeXml", () => {
		const input = "Tom & Jerry <\"friends\"> & 'rivals'";
		expect(unescapeXml(escapeXml(input))).toBe(input);
	});
});

describe("escapeXmlTag", () => {
	it("should escape spaces as _x0020_", () => {
		expect(escapeXmlTag("my tag")).toBe("my_x0020_tag");
	});
});

describe("escapeHtml", () => {
	it("should escape special characters", () => {
		expect(escapeHtml("<b>bold</b>")).toBe("&lt;b&gt;bold&lt;/b&gt;");
	});

	it("should convert newlines to <br/>", () => {
		expect(escapeHtml("line1\nline2")).toBe("line1<br/>line2");
	});

	it("should encode control characters as hex references", () => {
		expect(escapeHtml("\x01")).toBe("&#x0001;");
	});
});

describe("htmlDecode", () => {
	it("should strip HTML tags", () => {
		expect(htmlDecode("<b>bold</b>")).toBe("bold");
	});

	it("should decode named entities", () => {
		expect(htmlDecode("&amp;")).toBe("&");
		expect(htmlDecode("&lt;")).toBe("<");
		expect(htmlDecode("&gt;")).toBe(">");
		expect(htmlDecode("&nbsp;")).toBe(" ");
	});

	it("should convert <br> to newlines", () => {
		expect(htmlDecode("a<br>b")).toBe("a\nb");
		expect(htmlDecode("a<BR/>b")).toBe("a\nb");
		expect(htmlDecode("a<br />b")).toBe("a\nb");
	});
});

describe("parseXmlTag", () => {
	it("should parse tag name into key 0", () => {
		const result = parseXmlTag("foo");
		expect(result[0]).toBe("foo");
	});

	it("should parse attributes", () => {
		const result = parseXmlTag('tag id="123" name="test"');
		expect(result["id"]).toBe("123");
		expect(result["name"]).toBe("test");
	});

	it("should handle single-quoted attributes", () => {
		const result = parseXmlTag("tag id='123'");
		expect(result["id"]).toBe("123");
	});

	it("should skip root tag name when skip_root=true", () => {
		const result = parseXmlTag('tag id="1"', true);
		expect(result[0]).toBeUndefined();
		expect(result["id"]).toBe("1");
	});

	it("should strip namespace prefixes from attributes", () => {
		const result = parseXmlTag('tag r:id="rId1"');
		expect(result["id"]).toBe("rId1");
	});
});

describe("parseXmlBoolean", () => {
	it("should return true for truthy values", () => {
		expect(parseXmlBoolean(1)).toBe(true);
		expect(parseXmlBoolean(true)).toBe(true);
		expect(parseXmlBoolean("1")).toBe(true);
		expect(parseXmlBoolean("true")).toBe(true);
	});

	it("should return false for falsy values", () => {
		expect(parseXmlBoolean(0)).toBe(false);
		expect(parseXmlBoolean(false)).toBe(false);
		expect(parseXmlBoolean("0")).toBe(false);
		expect(parseXmlBoolean("false")).toBe(false);
	});

	it("should default to false for unknown values", () => {
		expect(parseXmlBoolean("yes")).toBe(false);
		expect(parseXmlBoolean(null)).toBe(false);
	});
});

describe("stripNamespace", () => {
	it("should strip namespace prefixes from tags", () => {
		expect(stripNamespace("<a:foo>")).toBe("<foo>");
		expect(stripNamespace("</a:foo>")).toBe("</foo>");
	});

	it("should leave tags without namespace unchanged", () => {
		expect(stripNamespace("<foo>")).toBe("<foo>");
	});
});
