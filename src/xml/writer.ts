import { escapeXml } from "./escape.js";
import { objectKeys } from "../utils/helpers.js";

const wtregex = /(^\s|\s$|\n)/;

/** Write an XML tag with text content */
export function writeXmlTag(f: string, g: string): string {
	return "<" + f + (g.match(wtregex) ? ' xml:space="preserve"' : "") + ">" + g + "</" + f + ">";
}

function formatXmlAttributes(h: Record<string, string>): string {
	return objectKeys(h)
		.map((k) => " " + k + '="' + h[k] + '"')
		.join("");
}

/** Write an XML tag with optional attributes and content */
export function writeXmlElement(f: string, g?: string | null, h?: Record<string, string> | null): string {
	return (
		"<" +
		f +
		(h != null ? formatXmlAttributes(h) : "") +
		(g != null ? (g.match(wtregex) ? ' xml:space="preserve"' : "") + ">" + g + "</" + f : "/") +
		">"
	);
}

/** Write a W3C datetime string from a Date */
export function writeW3cDatetime(d: Date, t?: boolean): string {
	try {
		return d.toISOString().replace(/\.\d*/, "");
	} catch (e) {
		if (t) {
			throw e;
		}
	}
	return "";
}

/** Write a vt: (variant type) XML element */
export function writeVariantType(s: any, xlsx?: boolean): string {
	switch (typeof s) {
		case "string": {
			let o = writeXmlElement("vt:lpwstr", escapeXml(s));
			if (xlsx) {
				o = o.replace(/&quot;/g, "_x0022_");
			}
			return o;
		}
		case "number":
			return writeXmlElement((s | 0) === s ? "vt:i4" : "vt:r8", escapeXml(String(s)));
		case "boolean":
			return writeXmlElement("vt:bool", s ? "true" : "false");
	}
	if (s instanceof Date) {
		return writeXmlElement("vt:filetime", writeW3cDatetime(s));
	}
	throw new Error("Unable to serialize " + s);
}
