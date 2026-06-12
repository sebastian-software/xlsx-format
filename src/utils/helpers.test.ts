import { describe, it, expect } from "vitest";
import { matchXmlTagGlobal, matchXmlTagFirst } from "./helpers.js";

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
