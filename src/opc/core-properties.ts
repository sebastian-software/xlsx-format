import type { FullProperties } from "../types.js";
import { XML_HEADER } from "../xml/parser.js";
import { unescapeXml, escapeXml } from "../xml/escape.js";
import { writeXmlTag, writeXmlElement, writeW3cDatetime } from "../xml/writer.js";
import { XMLNS } from "../xml/namespaces.js";

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

/** Simple XML tag content extraction */
function xml_extract(data: string, tag: string): string | null {
	const open = "<" + tag;
	const close = "</" + tag + ">";
	const openTagIdx = data.indexOf(open);
	if (openTagIdx === -1) {
		// Try with namespace prefix removed for tags like "dc:creator"
		return null;
	}
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

export function parseCoreProperties(data: string): Partial<FullProperties> {
	const p: Record<string, any> = {};

	for (const propDef of CORE_PROPS) {
		const content = xml_extract(data, propDef[0]);
		if (content != null && content.length > 0) {
			p[propDef[1]] = unescapeXml(content);
		}
		if (propDef[2] === "date" && p[propDef[1]]) {
			p[propDef[1]] = new Date(p[propDef[1]]);
		}
	}

	return p as Partial<FullProperties>;
}

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

	for (const propDef of CORE_PROPS) {
		let propValue: any = opts?.Props?.[propDef[1]] != null ? opts.Props[propDef[1]] : cp ? (cp as any)[propDef[1]] : null;
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
	if (lines.length > 2) {
		lines.push("</cp:coreProperties>");
		lines[1] = lines[1].replace("/>", ">");
	}
	return lines.join("");
}
