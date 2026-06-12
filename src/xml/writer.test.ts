import { describe, it, expect } from "vitest";
import { writeXmlTag, writeXmlElement, writeW3cDatetime, writeVariantType } from "./writer.js";

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
