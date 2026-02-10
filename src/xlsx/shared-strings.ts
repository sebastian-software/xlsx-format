import { parseXmlTag, XML_TAG_REGEX } from "../xml/parser.js";
import { unescapeXml, escapeXml, escapeHtml } from "../xml/escape.js";
import { writeXmlElement } from "../xml/writer.js";
import { XML_HEADER } from "../xml/parser.js";
import { XMLNS_main } from "../xml/namespaces.js";
import { utf8read } from "../utils/buffer.js";

export interface XLString {
	t: string;
	r?: string;
	h?: string;
	s?: any;
}

export interface SST extends Array<XLString> {
	Count?: number;
	Unique?: number;
}

/** Parse rich-text run properties */
function parseRunProperties(rpr: string): Record<string, any> {
	const font: Record<string, any> = {};
	const m = rpr.match(XML_TAG_REGEX);
	let pass = false;
	if (m) {
		for (let i = 0; i < m.length; ++i) {
			const y = parseXmlTag(m[i]);
			switch ((y[0] as string).replace(/<\w*:/g, "<")) {
				case "<condense":
				case "<extend":
					break;
				case "<shadow":
					if (!y.val) {
						break;
					}
				case "<shadow>":
				case "<shadow/>":
					font.shadow = 1;
					break;
				case "</shadow>":
					break;
				case "<rFont":
					font.name = y.val;
					break;
				case "<sz":
					font.sz = y.val;
					break;
				case "<strike":
					if (!y.val) {
						break;
					}
				case "<strike>":
				case "<strike/>":
					font.strike = 1;
					break;
				case "</strike>":
					break;
				case "<u":
					if (!y.val) {
						break;
					}
					switch (y.val) {
						case "double":
							font.uval = "double";
							break;
						case "singleAccounting":
							font.uval = "single-accounting";
							break;
						case "doubleAccounting":
							font.uval = "double-accounting";
							break;
					}
				case "<u>":
				case "<u/>":
					font.u = 1;
					break;
				case "</u>":
					break;
				case "<b":
					if (y.val === "0") {
						break;
					}
				case "<b>":
				case "<b/>":
					font.b = 1;
					break;
				case "</b>":
					break;
				case "<i":
					if (y.val === "0") {
						break;
					}
				case "<i>":
				case "<i/>":
					font.i = 1;
					break;
				case "</i>":
					break;
				case "<color":
					if (y.rgb) {
						font.color = y.rgb.slice(2, 8);
					}
					break;
				case "<color>":
				case "<color/>":
				case "</color>":
					break;
				case "<family":
					font.family = y.val;
					break;
				case "<vertAlign":
					font.valign = y.val;
					break;
				case "<scheme":
					break;
				case "<extLst":
				case "<extLst>":
				case "</extLst>":
					break;
				case "<ext":
					pass = true;
					break;
				case "</ext>":
					pass = false;
					break;
				default:
					if ((y[0] as string).charCodeAt(1) !== 47 && !pass) {
						throw new Error("Unrecognized rich format " + y[0]);
					}
			}
		}
	}
	return font;
}

const rregex = /<(?:\w+:)?r>/g;
const rend = /<\/(?:\w+:)?r>/;

function str_match_xml_ns_local(str: string, tag: string): [string, string] | null {
	const re = new RegExp("<(?:\\w+:)?" + tag + "\\b[^<>]*>", "g");
	const reEnd = new RegExp("<\\/(?:\\w+:)?" + tag + ">", "g");
	const m = re.exec(str);
	if (!m) {
		return null;
	}
	const si = m.index;
	const sf = re.lastIndex;
	reEnd.lastIndex = re.lastIndex;
	const m2 = reEnd.exec(str);
	if (!m2) {
		return null;
	}
	const ei = m2.index;
	const ef = reEnd.lastIndex;
	return [str.slice(si, ef), str.slice(sf, ei)];
}

function str_remove_xml_ns_g_local(str: string, tag: string): string {
	const re = new RegExp("<(?:\\w+:)?" + tag + "\\b[^<>]*>", "g");
	const reEnd = new RegExp("<\\/(?:\\w+:)?" + tag + ">", "g");
	const out: string[] = [];
	let lastEnd = 0;
	let m;
	while ((m = re.exec(str))) {
		out.push(str.slice(lastEnd, m.index));
		reEnd.lastIndex = re.lastIndex;
		const m2 = reEnd.exec(str);
		if (!m2) {
			break;
		}
		lastEnd = reEnd.lastIndex;
		re.lastIndex = reEnd.lastIndex;
	}
	out.push(str.slice(lastEnd));
	return out.join("");
}

function parseRichTextRun(r: string): { t: string; v: string; s?: any } {
	const t = str_match_xml_ns_local(r, "t");
	if (!t) {
		return { t: "s", v: "" };
	}
	const o: any = { t: "s", v: unescapeXml(t[1]) };
	const rpr = str_match_xml_ns_local(r, "rPr");
	if (rpr) {
		o.s = parseRunProperties(rpr[1]);
	}
	return o;
}

function parseRichTextRuns(rs: string): { t: string; v: string; s?: any }[] {
	return rs
		.replace(rregex, "")
		.split(rend)
		.map(parseRichTextRun)
		.filter((r) => r.v);
}

function richTextToHtml(rs: { t: string; v: string; s?: any }[]): string {
	const nlregex = /(\r\n|\n)/g;
	return rs
		.map((r) => {
			if (!r.v) {
				return "";
			}
			const intro: string[] = [];
			const outro: string[] = [];
			if (r.s) {
				const font = r.s;
				const style: string[] = [];
				if (font.u) {
					style.push("text-decoration: underline;");
				}
				if (font.uval) {
					style.push("text-underline-style:" + font.uval + ";");
				}
				if (font.sz) {
					style.push("font-size:" + font.sz + "pt;");
				}
				if (font.outline) {
					style.push("text-effect: outline;");
				}
				if (font.shadow) {
					style.push("text-shadow: auto;");
				}
				intro.push('<span style="' + style.join("") + '">');
				if (font.b) {
					intro.push("<b>");
					outro.push("</b>");
				}
				if (font.i) {
					intro.push("<i>");
					outro.push("</i>");
				}
				if (font.strike) {
					intro.push("<s>");
					outro.push("</s>");
				}
				let align = font.valign || "";
				if (align === "superscript" || align === "super") {
					align = "sup";
				} else if (align === "subscript") {
					align = "sub";
				}
				if (align !== "") {
					intro.push("<" + align + ">");
					outro.push("</" + align + ">");
				}
				outro.push("</span>");
			}
			return intro.join("") + r.v.replace(nlregex, "<br/>") + outro.join("");
		})
		.join("");
}

const sitregex = /<(?:\w+:)?t\b[^<>]*>([^<]*)<\/(?:\w+:)?t>/g;
const sirregex = /<(?:\w+:)?r\b[^<>]*>/;

function parseStringItem(x: string, opts?: { cellHTML?: boolean }): XLString {
	const html = opts ? opts.cellHTML !== false : true;
	const z: any = {};
	if (!x) {
		return { t: "" };
	}

	if (x.match(/^\s*<(?:\w+:)?t[^>]*>/)) {
		z.t = unescapeXml(utf8read(x.slice(x.indexOf(">") + 1).split(/<\/(?:\w+:)?t>/)[0] || ""), true);
		z.r = utf8read(x);
		if (html) {
			z.h = escapeHtml(z.t);
		}
	} else if (x.match(sirregex)) {
		z.r = utf8read(x);
		const stripped = str_remove_xml_ns_g_local(x, "rPh");
		sitregex.lastIndex = 0;
		const matches = stripped.match(sitregex) || [];
		z.t = unescapeXml(utf8read(matches.join("").replace(XML_TAG_REGEX, "")), true);
		if (html) {
			z.h = richTextToHtml(parseRichTextRuns(z.r));
		}
	}
	return z;
}

const sstr1 = /<(?:\w+:)?(?:si|sstItem)>/g;
const sstr2 = /<\/(?:\w+:)?(?:si|sstItem)>/;

/** Parse the Shared String Table XML */
export function parseSstXml(data: string, opts?: { cellHTML?: boolean }): SST {
	const s: SST = [] as any;
	if (!data) {
		return s;
	}

	const sst = str_match_xml_ns_local(data, "sst");
	if (sst) {
		const ss = sst[1].replace(sstr1, "").split(sstr2);
		for (let i = 0; i < ss.length; ++i) {
			const o = parseStringItem(ss[i].trim(), opts);
			if (o != null) {
				s[s.length] = o;
			}
		}
		const tag = parseXmlTag(sst[0].slice(0, sst[0].indexOf(">")));
		s.Count = tag.count;
		s.Unique = tag.uniquecount;
	}
	return s;
}

const straywsregex = /^\s|\s$|[\t\n\r]/;

/** Write the Shared String Table XML */
export function writeSstXml(sst: SST, opts: { bookSST?: boolean }): string {
	if (!opts.bookSST) {
		return "";
	}
	const o: string[] = [XML_HEADER];
	o.push(
		writeXmlElement("sst", null, {
			xmlns: XMLNS_main[0],
			count: String(sst.Count),
			uniqueCount: String(sst.Unique),
		}),
	);
	for (let i = 0; i !== sst.length; ++i) {
		if (sst[i] == null) {
			continue;
		}
		const s = sst[i];
		let sitag = "<si>";
		if (s.r) {
			sitag += s.r;
		} else {
			sitag += "<t";
			if (!s.t) {
				s.t = "";
			}
			if (typeof s.t !== "string") {
				s.t = String(s.t);
			}
			if (s.t.match(straywsregex)) {
				sitag += ' xml:space="preserve"';
			}
			sitag += ">" + escapeXml(s.t) + "</t>";
		}
		sitag += "</si>";
		o.push(sitag);
	}
	if (o.length > 2) {
		o.push("</sst>");
		o[1] = o[1].replace("/>", ">");
	}
	return o.join("");
}
