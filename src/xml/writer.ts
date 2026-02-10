import { escapexml } from "./escape.js";
import { keys } from "../utils/helpers.js";

const wtregex = /(^\s|\s$|\n)/;

/** Write an XML tag with text content */
export function writetag(f: string, g: string): string {
	return "<" + f + (g.match(wtregex) ? ' xml:space="preserve"' : "") + ">" + g + "</" + f + ">";
}

function wxt_helper(h: Record<string, string>): string {
	return keys(h)
		.map((k) => " " + k + '="' + h[k] + '"')
		.join("");
}

/** Write an XML tag with optional attributes and content */
export function writextag(f: string, g?: string | null, h?: Record<string, string> | null): string {
	return (
		"<" +
		f +
		(h != null ? wxt_helper(h) : "") +
		(g != null ? (g.match(wtregex) ? ' xml:space="preserve"' : "") + ">" + g + "</" + f : "/") +
		">"
	);
}

/** Write a W3C datetime string from a Date */
export function write_w3cdtf(d: Date, t?: boolean): string {
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
export function write_vt(s: any, xlsx?: boolean): string {
	switch (typeof s) {
		case "string": {
			let o = writextag("vt:lpwstr", escapexml(s));
			if (xlsx) {
				o = o.replace(/&quot;/g, "_x0022_");
			}
			return o;
		}
		case "number":
			return writextag((s | 0) === s ? "vt:i4" : "vt:r8", escapexml(String(s)));
		case "boolean":
			return writextag("vt:bool", s ? "true" : "false");
	}
	if (s instanceof Date) {
		return writextag("vt:filetime", write_w3cdtf(s));
	}
	throw new Error("Unable to serialize " + s);
}
