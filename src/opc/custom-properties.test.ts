import { describe, it, expect } from "vitest";
import { parseCustomProperties, writeCustomProperties } from "./custom-properties.js";

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
