import { parseXmlTag, XML_TAG_REGEX, XML_HEADER, stripNamespace } from "../xml/parser.js";
import { unescapeXml, escapeXml } from "../xml/escape.js";
import { writeXmlTag, writeXmlElement } from "../xml/writer.js";
import { XMLNS_main, XMLNS } from "../xml/namespaces.js";
import { decodeCell, encodeRange, safeDecodeRange } from "../utils/cell.js";
import { matchXmlTagFirst } from "../utils/helpers.js";
import type { WorkSheet } from "../types.js";

/** A parsed comment entry with cell reference, author, and text content */
export interface RawComment {
	ref: string;
	author: string;
	t: string;
	/** Raw rich-text XML */
	r?: string;
	/** HTML representation */
	h?: string;
	/** GUID for threaded comments */
	guid?: string;
	/** Whether this is a threaded comment (1 = threaded) */
	T?: number;
}

/**
 * Insert parsed comments into a worksheet, attaching them to the appropriate cells.
 *
 * Creates empty cells if needed and expands the sheet range to include comment cells.
 * When inserting threaded comments, any existing legacy comments on the same cell are removed.
 *
 * @param sheet - Target worksheet
 * @param comments - Array of parsed comment entries
 * @param threaded - Whether these are threaded (modern) comments
 * @param people - Optional people list for resolving threaded comment author IDs to display names
 */
export function insertCommentsIntoSheet(
	sheet: WorkSheet,
	comments: RawComment[],
	threaded: boolean,
	people?: { name: string; id: string }[],
): void {
	const dense = (sheet as any)["!data"] != null;
	for (const comment of comments) {
		const r = decodeCell(comment.ref);
		if (r.r < 0 || r.c < 0) {
			continue;
		}
		let cell: any;
		if (dense) {
			if (!(sheet as any)["!data"][r.r]) {
				(sheet as any)["!data"][r.r] = [];
			}
			cell = (sheet as any)["!data"][r.r][r.c];
		} else {
			cell = (sheet as any)[comment.ref];
		}

		if (!cell) {
			cell = { t: "z" };
			if (dense) {
				(sheet as any)["!data"][r.r][r.c] = cell;
			} else {
				(sheet as any)[comment.ref] = cell;
			}
			// Expand sheet range to include the cell with the comment
			// "BDWGO1000001:A1" is a sentinel range that decodes to a maximally inverted range
			const range = safeDecodeRange(sheet["!ref"] || "BDWGO1000001:A1");
			if (range.s.r > r.r) {
				range.s.r = r.r;
			}
			if (range.e.r < r.r) {
				range.e.r = r.r;
			}
			if (range.s.c > r.c) {
				range.s.c = r.c;
			}
			if (range.e.c < r.c) {
				range.e.c = r.c;
			}
			sheet["!ref"] = encodeRange(range);
		}

		if (!cell.c) {
			cell.c = [];
		}
		const o: any = { a: comment.author, t: comment.t, r: comment.r, T: threaded };
		if (comment.h) {
			o.h = comment.h;
		}

		/* Threaded comments always override legacy comments on the same cell */
		for (let i = cell.c.length - 1; i >= 0; --i) {
			if (!threaded && cell.c[i].T) {
				return;
			}
			if (threaded && !cell.c[i].T) {
				cell.c.splice(i, 1);
			}
		}
		// Resolve threaded comment author IDs to display names
		if (threaded && people) {
			for (let i = 0; i < people.length; ++i) {
				if (o.a === people[i].id) {
					o.a = people[i].name || o.a;
					break;
				}
			}
		}
		cell.c.push(o);
	}
}

/** Parse a simple inline string item for comment text content */
function parse_si_simple(x: string): { t: string; r: string; h: string } {
	if (!x) {
		return { t: "", r: "", h: "" };
	}
	const tMatch = x.match(/<(?:\w+:)?t[^>]*>([^<]*)<\/(?:\w+:)?t>/);
	const t = tMatch ? unescapeXml(tMatch[1]) : "";
	return { t, r: x, h: t };
}

/**
 * Parse comments XML (ECMA-376 18.7 Comments).
 *
 * Extracts the author list and comment entries from a comments.xml part.
 *
 * @param data - Raw XML string of the comments file
 * @param opts - Parsing options (sheetRows, cellHTML)
 * @returns Array of parsed comment entries
 */
export function parseCommentsXml(data: string, opts?: any): RawComment[] {
	// Handle empty/self-closing <comments/>
	if (data.match(/<(?:\w+:)?comments\s*\/>/)) {
		return [];
	}
	const authors: string[] = [];
	const commentList: RawComment[] = [];

	// Parse <authors> section
	const authtag = matchXmlTagFirst(data, "authors");
	if (authtag) {
		authtag.split(/<\/\w*:?author>/).forEach((x) => {
			if (x === "" || x.trim() === "") {
				return;
			}
			const a = x.match(/<(?:\w+:)?author[^<>]*>(.*)/);
			if (a) {
				authors.push(a[1]);
			}
		});
	}

	// Parse <commentList> section
	const cmnttag = matchXmlTagFirst(data, "commentList");
	if (cmnttag) {
		cmnttag.split(/<\/\w*:?comment>/).forEach((x) => {
			if (x === "" || x.trim() === "") {
				return;
			}
			const cm = x.match(/<(?:\w+:)?comment[^<>]*>/);
			if (!cm) {
				return;
			}
			const y = parseXmlTag(cm[0]);
			const comment: RawComment = {
				author: (y.authorId && authors[y.authorId]) || "sheetjsghost",
				ref: y.ref,
				guid: y.guid,
				t: "",
			};
			const cell = decodeCell(y.ref);
			// Respect the sheetRows limit
			if (opts && opts.sheetRows && opts.sheetRows <= cell.r) {
				return;
			}
			const textMatch = matchXmlTagFirst(x, "text");
			const rt = textMatch ? parse_si_simple(textMatch) : { r: "", t: "", h: "" };
			comment.r = rt.r;
			if (rt.r === "<t></t>") {
				rt.t = "";
				rt.h = "";
			}
			// Normalize line endings to Unix style
			comment.t = (rt.t || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n");
			if (opts && opts.cellHTML) {
				comment.h = rt.h;
			}
			commentList.push(comment);
		});
	}
	return commentList;
}

/**
 * Write comments XML (ECMA-376 18.7).
 *
 * Serializes comment data into the legacy comments.xml format. For threaded comments,
 * the text is flattened into "Comment:/Reply:" format for backward compatibility.
 *
 * @param data - Array of [cell_ref, comments_array] tuples
 * @returns Complete comments.xml string
 */
export function writeCommentsXml(data: [string, any[]][]): string {
	const o: string[] = [XML_HEADER, writeXmlElement("comments", null, { xmlns: XMLNS_main[0] })];

	// Build unique author list
	const iauthor: string[] = [];
	o.push("<authors>");
	data.forEach((x) => {
		x[1].forEach((w: any) => {
			const a = escapeXml(w.a);
			if (iauthor.indexOf(a) === -1) {
				iauthor.push(a);
				o.push("<author>" + a + "</author>");
			}
			// For threaded comments, add a "tc=<ID>" pseudo-author for the root comment
			if (w.T && w.ID && iauthor.indexOf("tc=" + w.ID) === -1) {
				iauthor.push("tc=" + w.ID);
				o.push("<author>" + "tc=" + w.ID + "</author>");
			}
		});
	});
	if (iauthor.length === 0) {
		iauthor.push("SheetJ5");
		o.push("<author>SheetJ5</author>");
	}
	o.push("</authors>");
	o.push("<commentList>");

	data.forEach((d) => {
		let lastauthor = 0;
		const ts: string[] = [];
		let tcnt = 0;
		if (d[1][0] && d[1][0].T && d[1][0].ID) {
			lastauthor = iauthor.indexOf("tc=" + d[1][0].ID);
		}
		d[1].forEach((c: any) => {
			if (c.a) {
				lastauthor = iauthor.indexOf(escapeXml(c.a));
			}
			if (c.T) {
				++tcnt;
			}
			ts.push(c.t == null ? "" : escapeXml(c.t));
		});
		if (tcnt === 0) {
			// Non-threaded: each comment gets its own <comment> element
			d[1].forEach((c: any) => {
				o.push('<comment ref="' + d[0] + '" authorId="' + iauthor.indexOf(escapeXml(c.a)) + '"><text>');
				o.push(writeXmlTag("t", c.t == null ? "" : escapeXml(c.t)));
				o.push("</text></comment>");
			});
		} else {
			// Threaded: merge all replies into a single <comment> with formatted text
			if (d[1][0] && d[1][0].T && d[1][0].ID) {
				lastauthor = iauthor.indexOf("tc=" + d[1][0].ID);
			}
			o.push('<comment ref="' + d[0] + '" authorId="' + lastauthor + '"><text>');
			let t = "Comment:\n    " + ts[0] + "\n";
			for (let i = 1; i < ts.length; ++i) {
				t += "Reply:\n    " + ts[i] + "\n";
			}
			o.push(writeXmlTag("t", escapeXml(t)));
			o.push("</text></comment>");
		}
	});
	o.push("</commentList>");
	if (o.length > 2) {
		o.push("</comments>");
		// Convert self-closing <comments .../> to opening tag
		o[1] = o[1].replace("/>", ">");
	}
	return o.join("");
}

/**
 * Parse threaded comments XML (MS-XLSX 2.1.17).
 *
 * Threaded comments are a modern Excel feature that supports reply chains.
 * Each comment has a personId (author), optional parentId (reply), and text.
 *
 * @param data - Raw XML string of the threadedComment file
 * @param opts - Parsing options
 * @returns Array of parsed threaded comment entries
 */
export function parseTcmntXml(data: string, _opts?: any): RawComment[] {
	const out: RawComment[] = [];
	let comment: any = {};
	// Track the character offset where the <text> content starts
	let tidx = 0;

	const ignoredTags = new Set([
		"<?xml",
		"<ThreadedComments",
		"</ThreadedComments>",
		"<mentions",
		"<mentions>",
		"</mentions>",
		"<extLst",
		"<extLst>",
		"</extLst>",
		"<extLst/>",
		"<ext",
		"</ext>",
	]);

	data.replace(XML_TAG_REGEX, function xml_tcmnt(x: string, idx: number): string {
		const y: any = parseXmlTag(x);
		const tag = stripNamespace(y[0]);
		if (ignoredTags.has(tag)) {
			return x;
		}
		switch (tag) {
			case "<threadedComment":
				comment = { author: y.personId, guid: y.id, ref: y.ref, T: 1 };
				break;
			case "</threadedComment>":
				if (comment.t != null) {
					out.push(comment);
				}
				break;
			case "<text>":
			case "<text":
				tidx = idx + x.length;
				break;
			case "</text>":
				// Normalize line endings
				comment.t = data.slice(tidx, idx).replace(/\r\n/g, "\n").replace(/\r/g, "\n");
				break;
		}
		return x;
	});
	return out;
}

/**
 * Write threaded comments XML (MS-XLSX 2.1.17).
 *
 * Generates GUIDs for each threaded comment using a deterministic counter.
 * The first comment in a chain is the root; subsequent comments reference it via parentId.
 *
 * @param comments - Array of [cell_ref, comments_array] tuples
 * @param people - Mutable people list (new authors are appended)
 * @param opts - Options with tcid counter for generating unique GUIDs
 * @returns Complete threadedComment XML string
 */
export function writeTcmntXml(comments: [string, any[]][], people: string[], opts: { tcid: number }): string {
	const o: string[] = [
		XML_HEADER,
		writeXmlElement("ThreadedComments", null, { xmlns: XMLNS.TCMNT }).replace(/[/]>/, ">"),
	];
	comments.forEach((carr) => {
		let rootid = "";
		(carr[1] || []).forEach((c: any, idx: number) => {
			if (!c.T) {
				delete c.ID;
				return;
			}
			if (c.a && people.indexOf(c.a) === -1) {
				people.push(c.a);
			}
			// Generate a deterministic GUID using the tcid counter (zero-padded to 12 digits)
			const tcopts: any = {
				ref: carr[0],
				id: "{54EE7951-7262-4200-6969-" + ("000000000000" + opts.tcid++).slice(-12) + "}",
			};
			if (idx === 0) {
				rootid = tcopts.id; // First comment is the root of the thread
			} else {
				tcopts.parentId = rootid; // Replies reference the root
			}
			c.ID = tcopts.id;
			if (c.a) {
				// Person GUID uses a different prefix than comment GUID
				tcopts.personId = "{54EE7950-7262-4200-6969-" + ("000000000000" + people.indexOf(c.a)).slice(-12) + "}";
			}
			o.push(writeXmlElement("threadedComment", writeXmlTag("text", c.t || ""), tcopts));
		});
	});
	o.push("</ThreadedComments>");
	return o.join("");
}

/**
 * Parse people XML (MS-XLSX 2.1.18).
 *
 * The people list maps person GUIDs to display names for threaded comment authorship.
 *
 * @param data - Raw XML string of the person.xml file
 * @returns Array of person entries with name and id
 */
export function parsePeopleXml(data: string): { name: string; id: string }[] {
	const out: { name: string; id: string }[] = [];
	const ignoredTags = new Set([
		"<?xml",
		"<personList",
		"</personList>",
		"</person>",
		"<extLst",
		"<extLst>",
		"</extLst>",
		"<extLst/>",
		"<ext",
		"</ext>",
	]);

	data.replace(XML_TAG_REGEX, function xml_people(x: string): string {
		const y: any = parseXmlTag(x);
		const tag = stripNamespace(y[0]);
		if (ignoredTags.has(tag)) {
			return x;
		}
		switch (tag) {
			case "<person":
				out.push({ name: y.displayname, id: y.id });
				break;
		}
		return x;
	});
	return out;
}

/**
 * Write people XML for threaded comments authorship.
 *
 * @param people - Array of author display names
 * @returns Complete person.xml string
 */
export function writePeopleXml(people: string[]): string {
	const o: string[] = [
		XML_HEADER,
		writeXmlElement("personList", null, {
			xmlns: XMLNS.TCMNT,
			"xmlns:x": XMLNS_main[0],
		}).replace(/[/]>/, ">"),
	];
	people.forEach((person, idx) => {
		o.push(
			writeXmlElement("person", null, {
				displayName: person,
				// Deterministic GUID based on index
				id: "{54EE7950-7262-4200-6969-" + ("000000000000" + idx).slice(-12) + "}",
				userId: person,
				providerId: "None",
			}),
		);
	});
	o.push("</personList>");
	return o.join("");
}
