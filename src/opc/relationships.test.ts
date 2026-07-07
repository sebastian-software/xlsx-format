import { describe, it, expect } from "vitest";
import { addRelationship, parseRelationships, writeRelationships, getRelsPath } from "./relationships.js";
import { RELS } from "../xml/namespaces.js";

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
		expect(rels["!id"].rId1.TargetMode).toBe("External");
	});

	it("uses explicit targetmode when provided", () => {
		const rels: any = { "!id": {} };
		addRelationship(rels, 1, "target.xml", RELS.WS, "External");
		expect(rels["!id"].rId1.TargetMode).toBe("External");
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
		expect(rels["!id"].rId1).toBeDefined();
	});
});

describe("opc/relationships", () => {
	it("parseRelationships should parse rels XML", () => {
		const xml = `<Relationships>
			<Relationship Id="rId1" Type="http://test/worksheet" Target="worksheets/sheet1.xml"/>
		</Relationships>`;
		const rels = parseRelationships(xml, "/xl/workbook.xml");
		expect(rels["!id"].rId1).toBeDefined();
		expect(rels["!id"].rId1.Target).toBe("worksheets/sheet1.xml");
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
		expect(rels["!id"].rId1).toBeDefined();
	});

	it("writeRelationships should produce valid XML", () => {
		const rels = parseRelationships(null, "/");
		addRelationship(rels, -1, "sheet.xml", "http://type");
		const xml = writeRelationships(rels);
		expect(xml).toContain("<Relationships");
		expect(xml).toContain("<Relationship");
	});
});
