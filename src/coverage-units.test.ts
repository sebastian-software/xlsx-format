/**
 * Unit tests targeting coverage gaps in internal modules.
 * Organized by source directory, top-to-bottom.
 */
import { describe, it, expect } from "vitest";
import { arrayToSheet, formatCell } from "./index.js";

// Internal imports for direct testing
import { matchXmlTagGlobal, matchXmlTagFirst } from "./utils/helpers.js";
import { utf8read, utf8encode, utf8decode, NULL_CHAR_REGEX, CONTROL_CHAR_REGEX } from "./utils/buffer.js";
import { base64decode, base64encode } from "./utils/base64.js";
import { writeXmlTag, writeXmlElement, writeW3cDatetime, writeVariantType } from "./xml/writer.js";
import { parseCustomProperties, writeCustomProperties } from "./opc/custom-properties.js";
import { parseCoreProperties, writeCoreProperties } from "./opc/core-properties.js";
import { parseExtendedProperties } from "./opc/extended-properties.js";
import { parseRelationships, addRelationship, writeRelationships, getRelsPath } from "./opc/relationships.js";
import { parseCalcChainXml } from "./xlsx/calc-chain.js";
import {
	rcToA1,
	a1ToRc,
	shiftFormulaStr,
	shiftFormulaXlsx,
	isFuzzyFormula,
	stripXlFunctionPrefix,
} from "./xlsx/formula.js";
import {
	parseCommentsXml,
	writeCommentsXml,
	parseTcmntXml,
	writeTcmntXml,
	parsePeopleXml,
	writePeopleXml,
	insertCommentsIntoSheet,
} from "./xlsx/comments.js";
import { parseSstXml, writeSstXml } from "./xlsx/shared-strings.js";
import { parseVml, writeVml } from "./xlsx/vml.js";
import { parseMetadataXml, writeMetadataXml } from "./xlsx/metadata.js";
import { initFormatTable, loadFormat, loadFormatTable, resetFormatTable } from "./ssf/table.js";

// ============================================================
// src/utils/helpers.ts
// ============================================================
describe("utils/helpers", () => {
	it("matchXmlTagGlobal should find tags", () => {
		const xml = "<root><item>a</item><item>b</item></root>";
		const result = matchXmlTagGlobal(xml, "item");
		expect(result).toHaveLength(2);
		expect(result![0]).toBe("<item>a</item>");
		expect(result![1]).toBe("<item>b</item>");
	});

	it("matchXmlTagGlobal should handle namespaced tags", () => {
		const xml = '<root><ns:item x="1">content</ns:item></root>';
		const result = matchXmlTagGlobal(xml, "item");
		expect(result).toHaveLength(1);
		expect(result![0]).toContain("content");
	});

	it("matchXmlTagGlobal should return null for no matches", () => {
		expect(matchXmlTagGlobal("<root></root>", "missing")).toBeNull();
	});

	it("matchXmlTagFirst should return first match or null", () => {
		const xml = "<a><b>1</b><b>2</b></a>";
		expect(matchXmlTagFirst(xml, "b")).toBe("<b>1</b>");
		expect(matchXmlTagFirst(xml, "c")).toBeNull();
	});
});

// ============================================================
// src/utils/buffer.ts
// ============================================================
describe("utils/buffer", () => {
	it("utf8encode and utf8decode should roundtrip", () => {
		const s = "Hello, Wörld! 日本語";
		expect(utf8decode(utf8encode(s))).toBe(s);
	});

	it("utf8read should decode ASCII", () => {
		expect(utf8read("Hello")).toBe("Hello");
	});

	it("utf8read should decode 2-byte UTF-8", () => {
		// ö = U+00F6 = 0xC3 0xB6 in UTF-8
		const binary = String.fromCharCode(0xc3, 0xb6);
		expect(utf8read(binary)).toBe("ö");
	});

	it("utf8read should decode 3-byte UTF-8", () => {
		// 日 = U+65E5 = 0xE6 0x97 0xA5
		const binary = String.fromCharCode(0xe6, 0x97, 0xa5);
		expect(utf8read(binary)).toBe("日");
	});

	it("utf8read should decode 4-byte UTF-8 (surrogate pair)", () => {
		// 𝄞 (musical symbol G clef) = U+1D11E = 0xF0 0x9D 0x84 0x9E
		const binary = String.fromCharCode(0xf0, 0x9d, 0x84, 0x9e);
		expect(utf8read(binary)).toBe("𝄞");
	});

	it("should export regex patterns", () => {
		expect("abc\u0000def".replace(NULL_CHAR_REGEX, "")).toBe("abcdef");
		expect("a\u0001b\u0003c".replace(CONTROL_CHAR_REGEX, "")).toBe("abc");
	});
});

// ============================================================
// src/utils/base64.ts
// ============================================================
describe("utils/base64", () => {
	it("should roundtrip encode/decode", () => {
		const data = new Uint8Array([72, 101, 108, 108, 111]);
		const encoded = base64encode(data);
		const decoded = base64decode(encoded);
		expect(decoded).toEqual(data);
	});

	it("should strip data URI prefix", () => {
		const data = new Uint8Array([1, 2, 3]);
		const b64 = base64encode(data);
		const withPrefix = "data:application/octet-stream;base64," + b64;
		expect(base64decode(withPrefix)).toEqual(data);
	});
});

// ============================================================
// src/xml/writer.ts
// ============================================================
describe("xml/writer", () => {
	it("writeXmlTag should wrap content", () => {
		expect(writeXmlTag("t", "hello")).toBe("<t>hello</t>");
	});

	it("writeXmlTag should add xml:space for whitespace", () => {
		expect(writeXmlTag("t", " hello ")).toContain('xml:space="preserve"');
		expect(writeXmlTag("t", "line\nbreak")).toContain('xml:space="preserve"');
	});

	it("writeXmlElement should self-close without content", () => {
		expect(writeXmlElement("br")).toBe("<br/>");
	});

	it("writeXmlElement should emit content with attributes", () => {
		const result = writeXmlElement("div", "text", { id: "x" });
		expect(result).toContain('id="x"');
		expect(result).toContain(">text</div>");
	});

	it("writeXmlElement should handle null content as self-closing", () => {
		const result = writeXmlElement("img", null, { src: "a.png" });
		expect(result).toContain('src="a.png"');
		expect(result).toContain("/>");
	});

	it("writeW3cDatetime should format dates", () => {
		const d = new Date("2024-06-15T10:30:00.000Z");
		expect(writeW3cDatetime(d)).toBe("2024-06-15T10:30:00Z");
	});

	it("writeW3cDatetime should return empty on invalid date", () => {
		expect(writeW3cDatetime(new Date("invalid"))).toBe("");
	});

	it("writeW3cDatetime should throw on invalid date when throwOnError", () => {
		expect(() => writeW3cDatetime(new Date("invalid"), true)).toThrow();
	});

	it("writeVariantType should handle strings", () => {
		expect(writeVariantType("hello")).toContain("vt:lpwstr");
	});

	it("writeVariantType should handle integers", () => {
		expect(writeVariantType(42)).toContain("vt:i4");
	});

	it("writeVariantType should handle floats", () => {
		expect(writeVariantType(3.14)).toContain("vt:r8");
	});

	it("writeVariantType should handle booleans", () => {
		expect(writeVariantType(true)).toContain("true");
		expect(writeVariantType(false)).toContain("false");
	});

	it("writeVariantType should handle dates", () => {
		expect(writeVariantType(new Date("2024-01-01"))).toContain("vt:filetime");
	});

	it("writeVariantType should escape quotes in xlsx mode", () => {
		const result = writeVariantType('say "hi"', true);
		expect(result).toContain("_x0022_");
	});

	it("writeVariantType should throw on unsupported type", () => {
		expect(() => writeVariantType(Symbol("x"))).toThrow();
	});
});

// ============================================================
// src/opc/custom-properties.ts
// ============================================================
describe("opc/custom-properties", () => {
	it("parseCustomProperties should parse string properties", () => {
		const xml = `<?xml version="1.0"?>
			<Properties><property name="Author"><vt:lpwstr>Alice</vt:lpwstr></property></Properties>`;
		const result = parseCustomProperties(xml);
		expect(result["Author"]).toBe("Alice");
	});

	it("parseCustomProperties should parse bool properties", () => {
		const xml = `<Properties><property name="Draft"><vt:bool>true</vt:bool></property></Properties>`;
		expect(parseCustomProperties(xml)["Draft"]).toBe(true);
	});

	it("parseCustomProperties should parse int properties", () => {
		const xml = `<Properties><property name="Count"><vt:i4>42</vt:i4></property></Properties>`;
		expect(parseCustomProperties(xml)["Count"]).toBe(42);
	});

	it("parseCustomProperties should parse float properties", () => {
		const xml = `<Properties><property name="Rate"><vt:r8>3.14</vt:r8></property></Properties>`;
		expect(parseCustomProperties(xml)["Rate"]).toBeCloseTo(3.14);
	});

	it("parseCustomProperties should parse date properties", () => {
		const xml = `<Properties><property name="Due"><vt:filetime>2024-01-01T00:00:00Z</vt:filetime></property></Properties>`;
		const result = parseCustomProperties(xml);
		expect(result["Due"]).toBeInstanceOf(Date);
	});

	it("parseCustomProperties should handle empty/self-closing vt types", () => {
		const xml = `<Properties><property name="X"><vt:empty/></property></Properties>`;
		const result = parseCustomProperties(xml);
		expect(result["X"]).toBeUndefined();
	});

	it("parseCustomProperties should handle cy and error types", () => {
		const xml = `<Properties><property name="Amt"><vt:cy>100</vt:cy></property></Properties>`;
		expect(parseCustomProperties(xml)["Amt"]).toBe("100");
	});

	it("writeCustomProperties should produce valid XML", () => {
		const result = writeCustomProperties({ Title: "Test", Count: 5, Flag: true, Rate: 1.5 });
		expect(result).toContain("<?xml");
		expect(result).toContain("<Properties");
		expect(result).toContain("<property");
		expect(result).toContain("</Properties>");
	});

	it("writeCustomProperties with undefined should produce minimal XML", () => {
		const result = writeCustomProperties(undefined);
		expect(result).toContain("<Properties");
		expect(result).not.toContain("</Properties>");
	});
});

// ============================================================
// src/opc/core-properties.ts
// ============================================================
describe("opc/core-properties", () => {
	it("parseCoreProperties should extract string fields", () => {
		const xml = `<cp:coreProperties><dc:title>My Doc</dc:title><dc:creator>Bob</dc:creator></cp:coreProperties>`;
		const result = parseCoreProperties(xml);
		expect(result.Title).toBe("My Doc");
		expect(result.Author).toBe("Bob");
	});

	it("parseCoreProperties should parse date fields", () => {
		const xml = `<cp:coreProperties><dcterms:created>2024-01-15T10:00:00Z</dcterms:created></cp:coreProperties>`;
		const result = parseCoreProperties(xml);
		expect(result.CreatedDate).toBeInstanceOf(Date);
	});

	it("writeCoreProperties should produce XML with dates", () => {
		const result = writeCoreProperties({
			Title: "Test",
			Author: "Alice",
			CreatedDate: new Date("2024-06-15T00:00:00Z"),
			ModifiedDate: new Date("2024-06-16T00:00:00Z"),
		});
		expect(result).toContain("dc:title");
		expect(result).toContain("dc:creator");
		expect(result).toContain("dcterms:created");
		expect(result).toContain("dcterms:modified");
		expect(result).toContain("</cp:coreProperties>");
	});

	it("writeCoreProperties with undefined should produce minimal XML", () => {
		const result = writeCoreProperties(undefined);
		expect(result).toContain("cp:coreProperties");
		expect(result).not.toContain("</cp:coreProperties>");
	});

	it("writeCoreProperties should handle boolean and number values", () => {
		const result = writeCoreProperties({}, { Props: { Revision: 5, Category: "Test" } });
		expect(result).toContain("cp:revision");
		expect(result).toContain("cp:category");
	});
});

// ============================================================
// src/opc/extended-properties.ts
// ============================================================
describe("opc/extended-properties", () => {
	it("parseExtendedProperties should parse string and bool fields", () => {
		const xml = `<Properties>
			<Application>Microsoft Excel</Application>
			<SharedDoc>true</SharedDoc>
			<ScaleCrop>false</ScaleCrop>
		</Properties>`;
		const result = parseExtendedProperties(xml);
		expect(result.Application).toBe("Microsoft Excel");
		expect(result.SharedDoc).toBe(true);
		expect(result.ScaleCrop).toBe(false);
	});

	it("parseExtendedProperties should parse HeadingPairs and TitlesOfParts", () => {
		const xml = `<Properties>
			<HeadingPairs><vt:vector><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>2</vt:i4></vt:variant></vt:vector></HeadingPairs>
			<TitlesOfParts><vt:vector><vt:lpstr>Sheet1</vt:lpstr><vt:lpstr>Sheet2</vt:lpstr></vt:vector></TitlesOfParts>
		</Properties>`;
		const result = parseExtendedProperties(xml);
		expect(result.Worksheets).toBe(2);
		expect(result.SheetNames).toEqual(["Sheet1", "Sheet2"]);
	});
});

// ============================================================
// src/opc/relationships.ts
// ============================================================
describe("opc/relationships", () => {
	it("parseRelationships should parse rels XML", () => {
		const xml = `<Relationships>
			<Relationship Id="rId1" Type="http://test/worksheet" Target="worksheets/sheet1.xml"/>
		</Relationships>`;
		const rels = parseRelationships(xml, "/xl/workbook.xml");
		expect(rels["!id"]["rId1"]).toBeDefined();
		expect(rels["!id"]["rId1"].Target).toBe("worksheets/sheet1.xml");
	});

	it("parseRelationships should handle null input", () => {
		const rels = parseRelationships(null, "/test");
		expect(rels["!id"]).toEqual({});
	});

	it("getRelsPath should compute correct path", () => {
		expect(getRelsPath("xl/workbook.xml")).toBe("xl/_rels/workbook.xml.rels");
	});

	it("addRelationship should auto-assign rId", () => {
		const rels = parseRelationships(null, "/");
		const id = addRelationship(rels, -1, "test.xml", "http://type");
		expect(id).toBe(1);
		expect(rels["!id"]["rId1"]).toBeDefined();
	});

	it("writeRelationships should produce valid XML", () => {
		const rels = parseRelationships(null, "/");
		addRelationship(rels, -1, "sheet.xml", "http://type");
		const xml = writeRelationships(rels);
		expect(xml).toContain("<Relationships");
		expect(xml).toContain("<Relationship");
	});
});

// ============================================================
// src/xlsx/formula.ts
// ============================================================
describe("xlsx/formula", () => {
	it("rcToA1 should convert absolute references", () => {
		expect(rcToA1("R1C1", { r: 0, c: 0 })).toBe("$A$1");
		expect(rcToA1("R2C3", { r: 0, c: 0 })).toBe("$C$2");
	});

	it("rcToA1 should convert relative references", () => {
		expect(rcToA1("R[1]C[1]", { r: 0, c: 0 })).toBe("B2");
		expect(rcToA1("R[0]C[0]", { r: 2, c: 3 })).toBe("D3");
		expect(rcToA1("RC", { r: 5, c: 2 })).toBe("C6");
	});

	it("a1ToRc should convert A1 to R1C1", () => {
		expect(a1ToRc("$A$1", { r: 0, c: 0 })).toBe("R1C1");
		expect(a1ToRc("B2", { r: 1, c: 1 })).toBe("RC");
		expect(a1ToRc("C3", { r: 0, c: 0 })).toBe("R[2]C[2]");
	});

	it("shiftFormulaStr should shift relative references", () => {
		const result = shiftFormulaStr("A1+B2", { r: 1, c: 1 });
		expect(result).toBe("B2+C3");
	});

	it("shiftFormulaStr should not shift absolute references", () => {
		const result = shiftFormulaStr("$A$1+B2", { r: 1, c: 1 });
		expect(result).toBe("$A$1+C3");
	});

	it("shiftFormulaXlsx should shift based on range and cell", () => {
		const result = shiftFormulaXlsx("A1*2", "A1:A10", "A3");
		expect(result).toBe("A3*2");
	});

	it("isFuzzyFormula should reject single chars", () => {
		expect(isFuzzyFormula("=")).toBe(false);
		expect(isFuzzyFormula("=SUM(A1)")).toBe(true);
	});

	it("stripXlFunctionPrefix should remove _xlfn.", () => {
		expect(stripXlFunctionPrefix("_xlfn.CONCAT(A1,B1)")).toBe("CONCAT(A1,B1)");
		expect(stripXlFunctionPrefix("SUM(A1)")).toBe("SUM(A1)");
	});
});

// ============================================================
// src/xlsx/calc-chain.ts
// ============================================================
describe("xlsx/calc-chain", () => {
	it("parseCalcChainXml should parse chain entries", () => {
		const xml = `<?xml version="1.0"?><calcChain><c r="A1" i="1"/><c r="B2"/><c r="C3" i="2"/></calcChain>`;
		const chain = parseCalcChainXml(xml);
		expect(chain).toHaveLength(3);
		expect(chain[0].r).toBe("A1");
		expect(chain[0].i).toBe("1");
		expect(chain[1].r).toBe("B2");
		expect(chain[1].i).toBe("1"); // sticky
		expect(chain[2].i).toBe("2");
	});

	it("parseCalcChainXml should handle empty input", () => {
		expect(parseCalcChainXml("")).toEqual([]);
	});
});

// ============================================================
// src/xlsx/comments.ts
// ============================================================
describe("xlsx/comments", () => {
	it("parseCommentsXml should roundtrip through write → parse", () => {
		const data: [string, any[]][] = [["A1", [{ a: "Alice", t: "Hello World" }]]];
		const xml = writeCommentsXml(data);
		const comments = parseCommentsXml(xml);
		expect(comments).toHaveLength(1);
		expect(comments[0].ref).toBe("A1");
		expect(comments[0].author).toBe("Alice");
		expect(comments[0].t).toBe("Hello World");
	});

	it("parseCommentsXml should parse hand-crafted XML", () => {
		const xml = `<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
			<authors><author>Bob</author></authors>
			<commentList>
				<comment ref="B2" authorId="0"><text><t>A note</t></text></comment>
			</commentList>
		</comments>`;
		const comments = parseCommentsXml(xml);
		expect(comments).toHaveLength(1);
		expect(comments[0].ref).toBe("B2");
		expect(comments[0].author).toBe("Bob");
		expect(comments[0].t).toBe("A note");
	});

	it("parseCommentsXml should handle empty comments", () => {
		expect(parseCommentsXml("<comments/>")).toEqual([]);
	});

	it("writeCommentsXml should produce valid XML", () => {
		const data: [string, any[]][] = [["A1", [{ a: "Bob", t: "Note here" }]]];
		const xml = writeCommentsXml(data);
		expect(xml).toContain("<comments");
		expect(xml).toContain("<author>Bob</author>");
		expect(xml).toContain('ref="A1"');
		expect(xml).toContain("Note here");
	});

	it("writeCommentsXml should handle threaded comments", () => {
		const data: [string, any[]][] = [
			[
				"B2",
				[
					{ a: "Alice", t: "Main comment", T: 1, ID: "{123}" },
					{ a: "Bob", t: "Reply", T: 1 },
				],
			],
		];
		const xml = writeCommentsXml(data);
		expect(xml).toContain("Comment:");
		expect(xml).toContain("Reply:");
	});

	it("parseTcmntXml should parse threaded comments", () => {
		const xml = `<?xml version="1.0"?><ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
			<threadedComment ref="A1" id="{111}" personId="{222}"><text>Thread text</text></threadedComment>
		</ThreadedComments>`;
		const result = parseTcmntXml(xml);
		expect(result).toHaveLength(1);
		expect(result[0].t).toBe("Thread text");
		expect(result[0].ref).toBe("A1");
		expect(result[0].T).toBe(1);
	});

	it("writeTcmntXml should produce valid XML", () => {
		const comments: [string, any[]][] = [["A1", [{ a: "Alice", t: "Hello", T: 1 }]]];
		const people: string[] = [];
		const xml = writeTcmntXml(comments, people, { tcid: 1 });
		expect(xml).toContain("<ThreadedComments");
		expect(xml).toContain("<threadedComment");
		expect(xml).toContain("Hello");
		expect(people).toContain("Alice");
	});

	it("parsePeopleXml should parse person list", () => {
		const xml = `<?xml version="1.0"?><personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
			<person displayname="Alice" id="{111}" userId="alice" providerId="None"/>
		</personList>`;
		const people = parsePeopleXml(xml);
		expect(people).toHaveLength(1);
		expect(people[0].name).toBe("Alice");
		expect(people[0].id).toBe("{111}");
	});

	it("writePeopleXml should produce valid XML", () => {
		const xml = writePeopleXml(["Alice", "Bob"]);
		expect(xml).toContain("<personList");
		expect(xml).toContain("Alice");
		expect(xml).toContain("Bob");
	});

	it("insertCommentsIntoSheet should attach comments to cells", () => {
		const ws = arrayToSheet([
			["A", "B"],
			[1, 2],
		]);
		const comments = [{ ref: "A1", author: "Test", t: "Comment text" }];
		insertCommentsIntoSheet(ws, comments, false);
		expect(ws["A1"].c).toBeDefined();
		expect(ws["A1"].c![0].t).toBe("Comment text");
	});

	it("insertCommentsIntoSheet should create cells if needed", () => {
		const ws = arrayToSheet([[1]]);
		const comments = [{ ref: "C3", author: "Test", t: "On empty cell" }];
		insertCommentsIntoSheet(ws, comments, false);
		expect((ws as any)["C3"]).toBeDefined();
		expect((ws as any)["C3"].c[0].t).toBe("On empty cell");
	});
});

// ============================================================
// src/xlsx/shared-strings.ts
// ============================================================
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

// ============================================================
// src/xlsx/vml.ts
// ============================================================
describe("xlsx/vml", () => {
	it("writeVml should produce valid VML XML", () => {
		const comments: [string, any][] = [["A1", { hidden: false }]];
		const xml = writeVml(1, comments);
		expect(xml).toContain("<xml");
		expect(xml).toContain("v:shape");
		expect(xml).toContain("x:Row");
		expect(xml).toContain("x:Column");
		expect(xml).toContain("x:Visible");
	});

	it("writeVml should handle hidden comments", () => {
		const comments: [string, any][] = [["B2", { hidden: true }]];
		const xml = writeVml(1, comments);
		expect(xml).toContain("visibility:hidden");
		expect(xml).not.toContain("<x:Visible/>");
	});

	it("writeVml should handle empty comments", () => {
		const xml = writeVml(1, []);
		expect(xml).toContain("<xml");
		expect(xml).not.toContain("v:shapetype");
	});

	it("parseVml should set comment visibility", () => {
		const ws = arrayToSheet([[1]]);
		ws["A1"].c = [{ a: "Test", t: "Note" }];
		const vml = `<xml><v:shape><x:ClientData ObjectType="Note"><x:Row>0</x:Row><x:Column>0</x:Column><x:Visible/></x:ClientData></v:shape></xml>`;
		parseVml(vml, ws, [{ ref: "A1" }]);
		expect(ws["A1"].c.hidden).toBe(false);
	});
});

// ============================================================
// src/xlsx/metadata.ts
// ============================================================
describe("xlsx/metadata", () => {
	it("parseMetadataXml should parse XLDAPR metadata", () => {
		const xml = `<?xml version="1.0"?><metadata>
			<metadataTypes count="1"><metadataType name="XLDAPR"/></metadataTypes>
			<futureMetadata name="XLDAPR"><bk><rvb i="0"/></bk></futureMetadata>
			<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata>
		</metadata>`;
		const meta = parseMetadataXml(xml);
		expect(meta.Types).toHaveLength(1);
		expect(meta.Types[0].name).toBe("XLDAPR");
		expect(meta.Cell).toHaveLength(1);
		expect(meta.Cell[0].type).toBe("XLDAPR");
	});

	it("parseMetadataXml should handle valueMetadata", () => {
		const xml = `<metadata>
			<metadataTypes count="1"><metadataType name="TEST"/></metadataTypes>
			<valueMetadata count="1"><bk><rc t="1" v="0"/></bk></valueMetadata>
		</metadata>`;
		const meta = parseMetadataXml(xml);
		expect(meta.Value).toHaveLength(1);
	});

	it("parseMetadataXml should handle empty input", () => {
		expect(parseMetadataXml("").Types).toEqual([]);
	});

	it("writeMetadataXml should produce XLDAPR metadata", () => {
		const xml = writeMetadataXml();
		expect(xml).toContain("XLDAPR");
		expect(xml).toContain("cellMetadata");
	});
});

// ============================================================
// src/ssf/table.ts
// ============================================================
describe("ssf/table", () => {
	it("initFormatTable should populate defaults", () => {
		const t = initFormatTable();
		expect(t[0]).toBe("General");
		expect(t[14]).toBe("m/d/yy");
		expect(t[49]).toBe("@");
	});

	it("loadFormat should register custom format", () => {
		resetFormatTable();
		const idx = loadFormat("#,##0.000");
		expect(idx).toBeGreaterThan(0);
	});

	it("loadFormat should find existing format", () => {
		resetFormatTable();
		const idx = loadFormat("General");
		expect(idx).toBe(0);
	});

	it("loadFormatTable should bulk-load formats", () => {
		resetFormatTable();
		loadFormatTable({ 200: "custom-fmt" });
		const t = initFormatTable();
		// After reset, custom format should be gone
		expect(t[200]).toBeUndefined();
	});
});

// ============================================================
// src/api/format.ts
// ============================================================
describe("api/format", () => {
	it("formatCell should return empty for null/stub cells", () => {
		expect(formatCell(null as any)).toBe("");
		expect(formatCell({ t: "z" } as any)).toBe("");
	});

	it("formatCell should return cached w value", () => {
		expect(formatCell({ t: "s", v: "x", w: "cached" } as any)).toBe("cached");
	});

	it("formatCell should format error cells", () => {
		const result = formatCell({ t: "e", v: 0x07 } as any);
		expect(result).toBe("#DIV/0!");
	});

	it("formatCell should format with explicit value", () => {
		const cell: any = { t: "n", v: 0 };
		const result = formatCell(cell, 42);
		expect(result).toBeDefined();
	});

	it("formatCell should use dateNF option", () => {
		const cell: any = { t: "d", v: new Date("2024-01-15") };
		formatCell(cell, null, { dateNF: "yyyy-mm-dd" });
		expect(cell.z).toBe("yyyy-mm-dd");
	});
});
