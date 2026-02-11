import type { FullProperties } from "../types.js";
import { XML_HEADER } from "../xml/parser.js";
import { unescapeXml, escapeXml } from "../xml/escape.js";
import { writeXmlTag, writeXmlElement, writeW3cDatetime } from "../xml/writer.js";
import { XMLNS } from "../xml/namespaces.js";

/**
 * Mapping of OPC core property XML tag names to their FullProperties keys.
 * Each entry is [xmlTagName, propertyKey, optionalType].
 * When the optional type is "date", the parsed value is converted to a Date object.
 */
const CORE_PROPS: [string, string, string?][] = [
	["cp:category", "Category"],
	["cp:contentStatus", "ContentStatus"],
	["cp:keywords", "Keywords"],
	["cp:lastModifiedBy", "LastAuthor"],
	["cp:lastPrinted", "LastPrinted"],
	["cp:revision", "Revision"],
	["cp:version", "Version"],
	["dc:creator", "Author"],
	["dc:description", "Comments"],
	["dc:identifier", "Identifier"],
	["dc:language", "Language"],
	["dc:subject", "Subject"],
	["dc:title", "Title"],
	["dcterms:created", "CreatedDate", "date"],
	["dcterms:modified", "ModifiedDate", "date"],
];

/**
 * Extract text content between an XML open/close tag pair using simple string search.
 * Does not handle nested tags of the same name -- sufficient for flat OPC property elements.
 * @param data - the XML string to search
 * @param tag - the fully-qualified tag name (e.g., "dc:title")
 * @returns the text content between the tags, or null if not found
 */
function xml_extract(data: string, tag: string): string | null {
	const open = "<" + tag;
	const close = "</" + tag + ">";
	const openTagIdx = data.indexOf(open);
	if (openTagIdx === -1) {
		return null;
	}
	// Find the end of the opening tag (handles tags with attributes)
	const closeAngleIdx = data.indexOf(">", openTagIdx);
	if (closeAngleIdx === -1) {
		return null;
	}
	const closeTagIdx = data.indexOf(close, closeAngleIdx);
	if (closeTagIdx === -1) {
		return null;
	}
	return data.slice(closeAngleIdx + 1, closeTagIdx);
}

/**
 * Parse OPC core properties XML (dc:title, dc:creator, dcterms:created, etc.)
 * into a partial FullProperties object.
 * @param data - raw XML string of the core properties part
 * @returns parsed property values
 */
export function parseCoreProperties(data: string): Partial<FullProperties> {
	const p: Record<string, any> = {};

	for (const propDef of CORE_PROPS) {
		const content = xml_extract(data, propDef[0]);
		if (content != null && content.length > 0) {
			p[propDef[1]] = unescapeXml(content);
		}
		// Convert date-typed properties from string to Date objects
		if (propDef[2] === "date" && p[propDef[1]]) {
			p[propDef[1]] = new Date(p[propDef[1]]);
		}
	}

	return p as Partial<FullProperties>;
}

/**
 * Write a single property field to the output lines array, avoiding duplicates.
 * @param tagName - the XML tag name (e.g., "dc:title")
 * @param value - the text value to write (null/undefined/empty are skipped)
 * @param attributes - optional attributes for the element (e.g., xsi:type for dates)
 * @param lines - output array to append the XML element to
 * @param written - tracks already-written tags to prevent duplicate entries
 */
function writePropertyField(
	tagName: string,
	value: string | null | undefined,
	attributes: Record<string, string> | null,
	lines: string[],
	written: Record<string, any>,
): void {
	if (written[tagName] != null || value == null || value === "") {
		return;
	}
	written[tagName] = value;
	value = escapeXml(value);
	lines.push(attributes ? writeXmlElement(tagName, value, attributes) : writeXmlTag(tagName, value));
}

/**
 * Serialize OPC core properties to XML.
 * Produces the cp:coreProperties element with Dublin Core metadata fields.
 * @param cp - the properties to serialize (may be undefined)
 * @param opts - optional settings: WTF enables strict date errors, Props provides overrides
 * @returns the complete XML string for the core properties part
 */
export function writeCoreProperties(
	cp: Partial<FullProperties> | undefined,
	opts?: { WTF?: boolean; Props?: Record<string, any> },
): string {
	const lines: string[] = [
		XML_HEADER,
		writeXmlElement("cp:coreProperties", null, {
			"xmlns:cp": XMLNS.CORE_PROPS,
			"xmlns:dc": XMLNS.dc,
			"xmlns:dcterms": XMLNS.dcterms,
			"xmlns:dcmitype": XMLNS.dcmitype,
			"xmlns:xsi": XMLNS.xsi,
		}),
	];
	const written: Record<string, any> = {};
	if (!cp && !opts?.Props) {
		return lines.join("");
	}

	// Date fields require special handling: they need xsi:type attribute
	if (cp) {
		if (cp.CreatedDate != null) {
			writePropertyField(
				"dcterms:created",
				typeof cp.CreatedDate === "string" ? cp.CreatedDate : writeW3cDatetime(cp.CreatedDate, opts?.WTF),
				{ "xsi:type": "dcterms:W3CDTF" },
				lines,
				written,
			);
		}
		if (cp.ModifiedDate != null) {
			writePropertyField(
				"dcterms:modified",
				typeof cp.ModifiedDate === "string" ? cp.ModifiedDate : writeW3cDatetime(cp.ModifiedDate, opts?.WTF),
				{ "xsi:type": "dcterms:W3CDTF" },
				lines,
				written,
			);
		}
	}

	// Write remaining properties, with opts.Props taking priority over cp values
	for (const propDef of CORE_PROPS) {
		let propValue: any =
			opts?.Props?.[propDef[1]] != null ? opts.Props[propDef[1]] : cp ? (cp as any)[propDef[1]] : null;
		// Coerce booleans and numbers to strings for XML serialization
		if (propValue === true) {
			propValue = "1";
		} else if (propValue === false) {
			propValue = "0";
		} else if (typeof propValue === "number") {
			propValue = String(propValue);
		}
		if (propValue != null) {
			writePropertyField(propDef[0], propValue, null, lines, written);
		}
	}
	// Only close the root element if child elements were added
	if (lines.length > 2) {
		lines.push("</cp:coreProperties>");
		// Convert the self-closing root tag to an opening tag
		lines[1] = lines[1].replace("/>", ">");
	}
	return lines.join("");
}
