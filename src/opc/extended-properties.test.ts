import { describe, it, expect } from "vitest";
import { parseExtendedProperties } from "./extended-properties.js";

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
		expect(result.SheetNames).toStrictEqual(["Sheet1", "Sheet2"]);
	});
});
