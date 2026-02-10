import type { FullProperties } from "../types.js";
import { XML_HEADER } from "../xml/parser.js";
import { unescapeXml, escapeXml } from "../xml/escape.js";
import { writeXmlElement } from "../xml/writer.js";
import { XMLNS } from "../xml/namespaces.js";

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

/** Simple namespace-aware XML tag content extraction */
function xml_extract_ns(data: string, tag: string): string | null {
	// Match <tag>...</tag> or <ns:tag>...</ns:tag>
	const re = new RegExp("<(?:\\w+:)?" + tag + "[\\s>]([\\s\\S]*?)<\\/(?:\\w+:)?" + tag + ">");
	const m = data.match(re);
	return m ? m[1] : null;
}

export function parseExtendedProperties(data: string, p?: Partial<FullProperties>): Partial<FullProperties> {
	if (!p) {
		p = {};
	}

	for (const f of EXT_PROPS) {
		const xml = xml_extract_ns(data, f[0]);
		switch (f[2]) {
			case "string":
				if (xml) {
					(p as any)[f[1]] = unescapeXml(xml);
				}
				break;
			case "bool":
				(p as any)[f[1]] = xml === "true";
				break;
		}
	}

	// Parse HeadingPairs and TitlesOfParts for sheet names
	const hpMatch = data.match(/<HeadingPairs>([\s\S]*?)<\/HeadingPairs>/);
	const topMatch = data.match(/<TitlesOfParts>([\s\S]*?)<\/TitlesOfParts>/);
	if (hpMatch && topMatch) {
		const lpstrs = topMatch[1].match(/<vt:lpstr>([\s\S]*?)<\/vt:lpstr>/g);
		if (lpstrs) {
			const parts = lpstrs.map((s) => {
				const m = s.match(/<vt:lpstr>([\s\S]*?)<\/vt:lpstr>/);
				return m ? unescapeXml(m[1]) : "";
			});
			// Try to extract Worksheets count from HeadingPairs
			const i4match = hpMatch[1].match(/<vt:i4>(\d+)<\/vt:i4>/);
			if (i4match) {
				(p as any).Worksheets = parseInt(i4match[1], 10);
				(p as any).SheetNames = parts.slice(0, (p as any).Worksheets);
			}
		}
	}

	return p;
}

export function writeExtendedProperties(cp: Record<string, any> | undefined): string {
	const o: string[] = [];
	const W = writeXmlElement;
	if (!cp) {
		cp = {};
	}
	cp.Application = "xlsx-format";

	o.push(XML_HEADER);
	o.push(
		writeXmlElement("Properties", null, {
			xmlns: XMLNS.EXT_PROPS,
			"xmlns:vt": XMLNS.vt,
		}),
	);

	for (const f of EXT_PROPS) {
		if (cp[f[1]] === undefined) {
			continue;
		}
		let v: string | undefined;
		switch (f[2]) {
			case "string":
				v = escapeXml(String(cp[f[1]]));
				break;
			case "bool":
				v = cp[f[1]] ? "true" : "false";
				break;
		}
		if (v !== undefined) {
			o.push(W(f[0], v));
		}
	}

	o.push(
		W(
			"HeadingPairs",
			W(
				"vt:vector",
				W("vt:variant", "<vt:lpstr>Worksheets</vt:lpstr>") + W("vt:variant", W("vt:i4", String(cp.Worksheets))),
				{ size: "2", baseType: "variant" },
			),
		),
	);

	o.push(
		W(
			"TitlesOfParts",
			W(
				"vt:vector",
				(cp.SheetNames as string[]).map((s: string) => "<vt:lpstr>" + escapeXml(s) + "</vt:lpstr>").join(""),
				{ size: String(cp.Worksheets), baseType: "lpstr" },
			),
		),
	);

	if (o.length > 2) {
		o.push("</Properties>");
		o[1] = o[1].replace("/>", ">");
	}
	return o.join("");
}
