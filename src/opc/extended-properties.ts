import type { FullProperties } from "../types.js";
import { XML_HEADER } from "../xml/parser.js";
import { unescapeXml, escapeXml } from "../xml/escape.js";
import { writeXmlElement } from "../xml/writer.js";
import { XMLNS } from "../xml/namespaces.js";

/**
 * Mapping of extended property XML element names to FullProperties keys and value types.
 * Each entry is [xmlElementName, propertyKey, valueType].
 */
const EXT_PROPS: [string, string, string][] = [
	["Application", "Application", "string"],
	["AppVersion", "AppVersion", "string"],
	["Company", "Company", "string"],
	["DocSecurity", "DocSecurity", "string"],
	["Manager", "Manager", "string"],
	["HyperlinksChanged", "HyperlinksChanged", "bool"],
	["SharedDoc", "SharedDoc", "bool"],
	["LinksUpToDate", "LinksUpToDate", "bool"],
	["ScaleCrop", "ScaleCrop", "bool"],
];

/**
 * Extract text content from an XML tag, allowing for optional namespace prefixes.
 * Builds a regex that matches both `<tag>` and `<ns:tag>` forms.
 * @param data - the XML string to search
 * @param tag - the local element name (without namespace prefix)
 * @returns the text content, or null if not found
 */
function xml_extract_ns(data: string, tag: string): string | null {
	const re = new RegExp("<(?:\\w+:)?" + tag + "[\\s>]([\\s\\S]*?)<\\/(?:\\w+:)?" + tag + ">");
	const m = data.match(re);
	return m ? m[1] : null;
}

/**
 * Parse OPC extended properties XML (Application, AppVersion, HeadingPairs, etc.)
 * into a partial FullProperties object.
 * @param data - raw XML string of the extended properties part
 * @param props - optional existing properties object to merge into
 * @returns the populated properties object
 */
export function parseExtendedProperties(data: string, props?: Partial<FullProperties>): Partial<FullProperties> {
	if (!props) {
		props = {};
	}

	for (const propDef of EXT_PROPS) {
		const xml = xml_extract_ns(data, propDef[0]);
		switch (propDef[2]) {
			case "string":
				if (xml) {
					(props as any)[propDef[1]] = unescapeXml(xml);
				}
				break;
			case "bool":
				(props as any)[propDef[1]] = xml === "true";
				break;
		}
	}

	// Parse HeadingPairs + TitlesOfParts to extract sheet names and count.
	// HeadingPairs contains category labels and counts (e.g., "Worksheets" -> 3),
	// TitlesOfParts contains all part titles in order.
	const hpMatch = data.match(/<HeadingPairs>([\s\S]*?)<\/HeadingPairs>/);
	const topMatch = data.match(/<TitlesOfParts>([\s\S]*?)<\/TitlesOfParts>/);
	if (hpMatch && topMatch) {
		const lpstrs = topMatch[1].match(/<vt:lpstr>([\s\S]*?)<\/vt:lpstr>/g);
		if (lpstrs) {
			const parts = lpstrs.map((lpstr) => {
				const match = lpstr.match(/<vt:lpstr>([\s\S]*?)<\/vt:lpstr>/);
				return match ? unescapeXml(match[1]) : "";
			});
			// The i4 value in HeadingPairs tells us how many of the parts are worksheets
			const i4match = hpMatch[1].match(/<vt:i4>(\d+)<\/vt:i4>/);
			if (i4match) {
				(props as any).Worksheets = parseInt(i4match[1], 10);
				(props as any).SheetNames = parts.slice(0, (props as any).Worksheets);
			}
		}
	}

	return props;
}

/**
 * Serialize OPC extended properties to XML.
 * Produces the Properties element with application metadata and sheet information.
 * @param cp - record of property values (may be undefined; defaults are applied)
 * @returns the complete XML string for the extended properties part
 */
export function writeExtendedProperties(cp: Record<string, any> | undefined): string {
	const lines: string[] = [];
	const writeElement = writeXmlElement;
	if (!cp) {
		cp = {};
	}
	cp.Application = "xlsx-format";

	lines.push(XML_HEADER);
	lines.push(
		writeXmlElement("Properties", null, {
			xmlns: XMLNS.EXT_PROPS,
			"xmlns:vt": XMLNS.vt,
		}),
	);

	// Write simple scalar properties (strings and booleans)
	for (const propDef of EXT_PROPS) {
		if (cp[propDef[1]] === undefined) {
			continue;
		}
		let propValue: string | undefined;
		switch (propDef[2]) {
			case "string":
				propValue = escapeXml(String(cp[propDef[1]]));
				break;
			case "bool":
				propValue = cp[propDef[1]] ? "true" : "false";
				break;
		}
		if (propValue !== undefined) {
			lines.push(writeElement(propDef[0], propValue));
		}
	}

	// HeadingPairs: a vt:vector of variant pairs (label, count) describing part categories.
	// For spreadsheets this is always ["Worksheets", <count>].
	lines.push(
		writeElement(
			"HeadingPairs",
			writeElement(
				"vt:vector",
				writeElement("vt:variant", "<vt:lpstr>Worksheets</vt:lpstr>") +
					writeElement("vt:variant", writeElement("vt:i4", String(cp.Worksheets))),
				{ size: "2", baseType: "variant" },
			),
		),
	);

	// TitlesOfParts: a vt:vector listing all sheet names as vt:lpstr elements
	lines.push(
		writeElement(
			"TitlesOfParts",
			writeElement(
				"vt:vector",
				(cp.SheetNames as string[])
					.map((sheetName: string) => "<vt:lpstr>" + escapeXml(sheetName) + "</vt:lpstr>")
					.join(""),
				{ size: String(cp.Worksheets), baseType: "lpstr" },
			),
		),
	);

	// Only close the root element if child elements were added
	if (lines.length > 2) {
		lines.push("</Properties>");
		// Convert the self-closing root tag to an opening tag
		lines[1] = lines[1].replace("/>", ">");
	}
	return lines.join("");
}
