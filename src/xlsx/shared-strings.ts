import { parseXmlTag, XML_TAG_REGEX } from "../xml/parser.js";
import { unescapeXml, escapeXml, escapeHtml } from "../xml/escape.js";
import { writeXmlElement } from "../xml/writer.js";
import { XML_HEADER } from "../xml/parser.js";
import { XMLNS_main } from "../xml/namespaces.js";
import { utf8read } from "../utils/buffer.js";

/** A single string entry from the shared string table */
export interface XLString {
	/** Plain text content */
	t: string;
	/** Raw rich-text XML */
	r?: string;
	/** HTML representation */
	h?: string;
	/** Style/font properties for rich text */
	s?: any;
}

/** Shared String Table - an array of XLString entries with total/unique counts */
export interface SST extends Array<XLString> {
	Count?: number;
	Unique?: number;
}

/** Parse rich-text run properties (<rPr>) into a font descriptor object */
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
				// Note: intentional fall-through from <shadow with no val to <shadow/>
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
				// Note: intentional fall-through from <strike with no val to <strike/>
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
				// Underline: various styles (double, singleAccounting, doubleAccounting)
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
				// Intentional fall-through: set underline flag
				case "<u>":
				case "<u/>":
					font.u = 1;
					break;
				case "</u>":
					break;
				// Bold: val="0" means explicitly not bold
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
				// Italic: val="0" means explicitly not italic
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
						// Strip the alpha channel prefix (first 2 hex chars) from ARGB
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
				// Skip extension lists
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
					// charCodeAt(1) !== 47 means it's not a closing tag (not '/')
					if ((parsedTag[0] as string).charCodeAt(1) !== 47 && !pass) {
						throw new Error("Unrecognized rich format " + parsedTag[0]);
					}
			}
		}
	}
	return font;
}

/** Regex to match opening <r> tags (rich-text run boundaries) */
const rregex = /<(?:\w+:)?r>/g;
/** Regex to match closing </r> tags */
const rend = /<\/(?:\w+:)?r>/;

/**
 * Find the first occurrence of a namespace-agnostic XML tag and return its
 * full outer content and inner content as a tuple.
 */
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
	// [0] = full outer XML (open tag through close tag), [1] = inner content only
	return [str.slice(startIdx, contentEnd), str.slice(contentStart, endIdx)];
}

/**
 * Remove all occurrences of a namespace-agnostic XML element (including content)
 * from the string. Used to strip <rPh> (phonetic run) elements.
 */
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

/** Parse a single rich-text run (<r>) element, extracting text and optional style */
function parseRichTextRun(r: string): { t: string; v: string; s?: any } {
	const textMatch = str_match_xml_ns_local(r, "t");
	if (!textMatch) {
		return { t: "s", v: "" };
	}
	const runObj: any = { t: "s", v: unescapeXml(textMatch[1]) };
	// Parse run properties (font/style) if present
	const rpr = str_match_xml_ns_local(r, "rPr");
	if (rpr) {
		runObj.s = parseRunProperties(rpr[1]);
	}
	return runObj;
}

/** Split rich-text XML into individual runs and parse each one */
function parseRichTextRuns(rs: string): { t: string; v: string; s?: any }[] {
	return rs
		.replace(rregex, "")
		.split(rend)
		.map(parseRichTextRun)
		.filter((r) => r.v);
}

/** Convert an array of parsed rich-text runs into an HTML string */
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
				// Map OOXML vertical alignment names to HTML elements
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

/** Regex to extract text from <t> elements */
const sitregex = /<(?:\w+:)?t\b[^<>]*>([^<]*)<\/(?:\w+:)?t>/g;
/** Regex to detect if a string item contains rich-text runs (<r>) */
const sirregex = /<(?:\w+:)?r\b[^<>]*>/;

/**
 * Parse a single string item (<si>) from the shared string table.
 * Handles both plain text (<t>) and rich text (<r>) formats.
 */
function parseStringItem(x: string, opts?: { cellHTML?: boolean }): XLString {
	const html = opts ? opts.cellHTML !== false : true;
	const result: any = {};
	if (!x) {
		return { t: "" };
	}

	if (x.match(/^\s*<(?:\w+:)?t[^>]*>/)) {
		// Plain text: extract content between <t> and </t>
		result.t = unescapeXml(utf8read(x.slice(x.indexOf(">") + 1).split(/<\/(?:\w+:)?t>/)[0] || ""), true);
		result.r = utf8read(x);
		if (html) {
			result.h = escapeHtml(result.t);
		}
	} else if (x.match(sirregex)) {
		// Rich text: concatenate text from all runs
		result.r = utf8read(x);
		// Strip phonetic run (<rPh>) elements before extracting text
		const stripped = str_remove_xml_ns_g_local(x, "rPh");
		sitregex.lastIndex = 0;
		const matches = stripped.match(sitregex) || [];
		// Join all <t> content and strip tags to get plain text
		result.t = unescapeXml(utf8read(matches.join("").replace(XML_TAG_REGEX, "")), true);
		if (html) {
			result.h = richTextToHtml(parseRichTextRuns(result.r));
		}
	}
	return result;
}

/** Regex to match opening <si> or <sstItem> tags */
const sstr1 = /<(?:\w+:)?(?:si|sstItem)>/g;
/** Regex to match closing </si> or </sstItem> tags */
const sstr2 = /<\/(?:\w+:)?(?:si|sstItem)>/;

/**
 * Parse the Shared String Table (SST) XML into an array of string entries.
 *
 * The SST is a deduplicated table of all string values used across the workbook.
 * Each cell with type "s" references an index into this table.
 *
 * @param data - Raw XML string of sharedStrings.xml
 * @param opts - Options controlling HTML generation (cellHTML)
 * @returns Array of parsed string entries with Count and Unique metadata
 */
export function parseSstXml(data: string, opts?: { cellHTML?: boolean }): SST {
	const strings: SST = [] as any;
	if (!data) {
		return strings;
	}

	const sst = str_match_xml_ns_local(data, "sst");
	if (sst) {
		// Split the SST content by </si> boundaries to get individual string items
		const stringItems = sst[1].replace(sstr1, "").split(sstr2);
		for (let i = 0; i < stringItems.length; ++i) {
			const parsedItem = parseStringItem(stringItems[i].trim(), opts);
			if (parsedItem != null) {
				strings[strings.length] = parsedItem;
			}
		}
		// Extract count/uniqueCount attributes from the <sst> opening tag
		const tag = parseXmlTag(sst[0].slice(0, sst[0].indexOf(">")));
		strings.Count = tag.count;
		strings.Unique = tag.uniquecount;
	}
	return strings;
}

/** Matches strings with leading/trailing whitespace or internal whitespace chars that need xml:space="preserve" */
const straywsregex = /^\s|\s$|[\t\n\r]/;

/**
 * Write the Shared String Table (SST) as XML.
 *
 * @param sst - Array of string entries to serialize
 * @param opts - Options; bookSST must be true to produce output
 * @returns Complete sharedStrings.xml string, or empty string if bookSST is false
 */
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
			// Preserve original rich-text XML
			sitag += entry.r;
		} else {
			sitag += "<t";
			if (!entry.t) {
				entry.t = "";
			}
			if (typeof entry.t !== "string") {
				entry.t = String(entry.t);
			}
			// Add xml:space="preserve" for strings with significant whitespace
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
		// Convert self-closing <sst .../> to opening tag <sst ...>
		lines[1] = lines[1].replace("/>", ">");
	}
	return lines.join("");
}
