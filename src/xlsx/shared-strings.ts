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
	const matches = rpr.match(XML_TAG_REGEX);
	let pass = false;
	if (matches) {
		for (let i = 0; i < matches.length; ++i) {
			const parsedTag = parseXmlTag(matches[i]);
			switch ((parsedTag[0] as string).replace(/<\w*:/g, "<")) {
				case "<condense":
				case "<extend":
					break;
				case "<shadow":
					if (!parsedTag.val) {
						break;
					}
				case "<shadow>":
				case "<shadow/>":
					font.shadow = 1;
					break;
				case "</shadow>":
					break;
				case "<rFont":
					font.name = parsedTag.val;
					break;
				case "<sz":
					font.sz = parsedTag.val;
					break;
				case "<strike":
					if (!parsedTag.val) {
						break;
					}
				case "<strike>":
				case "<strike/>":
					font.strike = 1;
					break;
				case "</strike>":
					break;
				case "<u":
					if (!parsedTag.val) {
						break;
					}
					switch (parsedTag.val) {
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
					if (parsedTag.val === "0") {
						break;
					}
				case "<b>":
				case "<b/>":
					font.b = 1;
					break;
				case "</b>":
					break;
				case "<i":
					if (parsedTag.val === "0") {
						break;
					}
				case "<i>":
				case "<i/>":
					font.i = 1;
					break;
				case "</i>":
					break;
				case "<color":
					if (parsedTag.rgb) {
						font.color = parsedTag.rgb.slice(2, 8);
					}
					break;
				case "<color>":
				case "<color/>":
				case "</color>":
					break;
				case "<family":
					font.family = parsedTag.val;
					break;
				case "<vertAlign":
					font.valign = parsedTag.val;
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
					if ((parsedTag[0] as string).charCodeAt(1) !== 47 && !pass) {
						throw new Error("Unrecognized rich format " + parsedTag[0]);
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
	const openMatch = re.exec(str);
	if (!openMatch) {
		return null;
	}
	const startIdx = openMatch.index;
	const contentStart = re.lastIndex;
	reEnd.lastIndex = re.lastIndex;
	const closeMatch = reEnd.exec(str);
	if (!closeMatch) {
		return null;
	}
	const endIdx = closeMatch.index;
	const contentEnd = reEnd.lastIndex;
	return [str.slice(startIdx, contentEnd), str.slice(contentStart, endIdx)];
}

function str_remove_xml_ns_g_local(str: string, tag: string): string {
	const re = new RegExp("<(?:\\w+:)?" + tag + "\\b[^<>]*>", "g");
	const reEnd = new RegExp("<\\/(?:\\w+:)?" + tag + ">", "g");
	const out: string[] = [];
	let lastEnd = 0;
	let openMatch;
	while ((openMatch = re.exec(str))) {
		out.push(str.slice(lastEnd, openMatch.index));
		reEnd.lastIndex = re.lastIndex;
		const closeMatch = reEnd.exec(str);
		if (!closeMatch) {
			break;
		}
		lastEnd = reEnd.lastIndex;
		re.lastIndex = reEnd.lastIndex;
	}
	out.push(str.slice(lastEnd));
	return out.join("");
}

function parseRichTextRun(r: string): { t: string; v: string; s?: any } {
	const textMatch = str_match_xml_ns_local(r, "t");
	if (!textMatch) {
		return { t: "s", v: "" };
	}
	const runObj: any = { t: "s", v: unescapeXml(textMatch[1]) };
	const rpr = str_match_xml_ns_local(r, "rPr");
	if (rpr) {
		runObj.s = parseRunProperties(rpr[1]);
	}
	return runObj;
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
	const result: any = {};
	if (!x) {
		return { t: "" };
	}

	if (x.match(/^\s*<(?:\w+:)?t[^>]*>/)) {
		result.t = unescapeXml(utf8read(x.slice(x.indexOf(">") + 1).split(/<\/(?:\w+:)?t>/)[0] || ""), true);
		result.r = utf8read(x);
		if (html) {
			result.h = escapeHtml(result.t);
		}
	} else if (x.match(sirregex)) {
		result.r = utf8read(x);
		const stripped = str_remove_xml_ns_g_local(x, "rPh");
		sitregex.lastIndex = 0;
		const matches = stripped.match(sitregex) || [];
		result.t = unescapeXml(utf8read(matches.join("").replace(XML_TAG_REGEX, "")), true);
		if (html) {
			result.h = richTextToHtml(parseRichTextRuns(result.r));
		}
	}
	return result;
}

const sstr1 = /<(?:\w+:)?(?:si|sstItem)>/g;
const sstr2 = /<\/(?:\w+:)?(?:si|sstItem)>/;

/** Parse the Shared String Table XML */
export function parseSstXml(data: string, opts?: { cellHTML?: boolean }): SST {
	const strings: SST = [] as any;
	if (!data) {
		return strings;
	}

	const sst = str_match_xml_ns_local(data, "sst");
	if (sst) {
		const stringItems = sst[1].replace(sstr1, "").split(sstr2);
		for (let i = 0; i < stringItems.length; ++i) {
			const parsedItem = parseStringItem(stringItems[i].trim(), opts);
			if (parsedItem != null) {
				strings[strings.length] = parsedItem;
			}
		}
		const tag = parseXmlTag(sst[0].slice(0, sst[0].indexOf(">")));
		strings.Count = tag.count;
		strings.Unique = tag.uniquecount;
	}
	return strings;
}

const straywsregex = /^\s|\s$|[\t\n\r]/;

/** Write the Shared String Table XML */
export function writeSstXml(sst: SST, opts: { bookSST?: boolean }): string {
	if (!opts.bookSST) {
		return "";
	}
	const lines: string[] = [XML_HEADER];
	lines.push(
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
		const entry = sst[i];
		let sitag = "<si>";
		if (entry.r) {
			sitag += entry.r;
		} else {
			sitag += "<t";
			if (!entry.t) {
				entry.t = "";
			}
			if (typeof entry.t !== "string") {
				entry.t = String(entry.t);
			}
			if (entry.t.match(straywsregex)) {
				sitag += ' xml:space="preserve"';
			}
			sitag += ">" + escapeXml(entry.t) + "</t>";
		}
		sitag += "</si>";
		lines.push(sitag);
	}
	if (lines.length > 2) {
		lines.push("</sst>");
		lines[1] = lines[1].replace("/>", ">");
	}
	return lines.join("");
}
