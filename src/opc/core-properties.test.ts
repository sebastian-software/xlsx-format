import { describe, it, expect } from "vitest";
import { parseCoreProperties, writeCoreProperties } from "./core-properties.js";

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
