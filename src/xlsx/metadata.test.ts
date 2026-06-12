import { describe, it, expect } from "vitest";
import { parseMetadataXml, writeMetadataXml } from "./metadata.js";

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
