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
	const si = data.indexOf(open);
	if (si === -1) {
		// Try with namespace prefix removed for tags like "dc:creator"
		return null;
	}
	const gt = data.indexOf(">", si);
	if (gt === -1) {
		return null;
	}
	const ei = data.indexOf(close, gt);
	if (ei === -1) {
		return null;
	}
	return data.slice(gt + 1, ei);
}

export function parseCoreProperties(data: string): Partial<FullProperties> {
	const p: Record<string, any> = {};

	for (const f of CORE_PROPS) {
		const content = xml_extract(data, f[0]);
		if (content != null && content.length > 0) {
			p[f[1]] = unescapeXml(content);
		}
		if (f[2] === "date" && p[f[1]]) {
			p[f[1]] = new Date(p[f[1]]);
		}
	}

	return p as Partial<FullProperties>;
}

function writePropertyField(
	f: string,
	g: string | null | undefined,
	h: Record<string, string> | null,
	lines: string[],
	p: Record<string, any>,
): void {
	if (p[f] != null || g == null || g === "") {
		return;
	}
	p[f] = g;
	g = escapeXml(g);
	lines.push(h ? writeXmlElement(f, g, h) : writeXmlTag(f, g));
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
	const p: Record<string, any> = {};
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
				p,
			);
		}
		if (cp.ModifiedDate != null) {
			writePropertyField(
				"dcterms:modified",
				typeof cp.ModifiedDate === "string" ? cp.ModifiedDate : writeW3cDatetime(cp.ModifiedDate, opts?.WTF),
				{ "xsi:type": "dcterms:W3CDTF" },
				lines,
				p,
			);
		}
	}

	for (const f of CORE_PROPS) {
		let v: any = opts?.Props?.[f[1]] != null ? opts.Props[f[1]] : cp ? (cp as any)[f[1]] : null;
		if (v === true) {
			v = "1";
		} else if (v === false) {
			v = "0";
		} else if (typeof v === "number") {
			v = String(v);
		}
		if (v != null) {
			writePropertyField(f[0], v, null, lines, p);
		}
	}
	if (lines.length > 2) {
		lines.push("</cp:coreProperties>");
		lines[1] = lines[1].replace("/>", ">");
	}
	return lines.join("");
}
