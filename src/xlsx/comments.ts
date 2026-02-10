import { parsexmltag, tagregex, XML_HEADER, strip_ns } from "../xml/parser.js";
import { unescapexml, escapexml } from "../xml/escape.js";
import { writetag, writextag } from "../xml/writer.js";
import { XMLNS_main, XMLNS } from "../xml/namespaces.js";
import { decode_cell, encode_cell, encode_range, safe_decode_range } from "../utils/cell.js";
import { str_match_xml_ns } from "../utils/helpers.js";
import type { WorkSheet } from "../types.js";

export interface RawComment {
	ref: string;
	author: string;
	t: string;
	r?: string;
	h?: string;
	guid?: string;
	T?: number;
}

/** Insert parsed comments into a worksheet */
export function sheet_insert_comments(
	sheet: WorkSheet,
	comments: RawComment[],
	threaded: boolean,
	people?: { name: string; id: string }[],
): void {
	const dense = (sheet as any)["!data"] != null;
	for (const comment of comments) {
		const r = decode_cell(comment.ref);
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
			const range = safe_decode_range(sheet["!ref"] || "BDWGO1000001:A1");
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
			sheet["!ref"] = encode_range(range);
		}

		if (!cell.c) {
			cell.c = [];
		}
		const o: any = { a: comment.author, t: comment.t, r: comment.r, T: threaded };
		if (comment.h) {
			o.h = comment.h;
		}

		/* threaded comments always override */
		for (let i = cell.c.length - 1; i >= 0; --i) {
			if (!threaded && cell.c[i].T) {
				return;
			}
			if (threaded && !cell.c[i].T) {
				cell.c.splice(i, 1);
			}
		}
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

/** Parse the si (string item) for comment text - simplified inline version */
function parse_si_simple(x: string): { t: string; r: string; h: string } {
	if (!x) {
		return { t: "", r: "", h: "" };
	}
	const tMatch = x.match(/<(?:\w+:)?t[^>]*>([^<]*)<\/(?:\w+:)?t>/);
	const t = tMatch ? unescapexml(tMatch[1]) : "";
	return { t, r: x, h: t };
}

/** Parse comments XML (18.7) */
export function parse_comments_xml(data: string, opts?: any): RawComment[] {
	if (data.match(/<(?:\w+:)?comments\s*\/>/)) {
		return [];
	}
	const authors: string[] = [];
	const commentList: RawComment[] = [];

	const authtag = str_match_xml_ns(data, "authors");
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

	const cmnttag = str_match_xml_ns(data, "commentList");
	if (cmnttag) {
		cmnttag.split(/<\/\w*:?comment>/).forEach((x) => {
			if (x === "" || x.trim() === "") {
				return;
			}
			const cm = x.match(/<(?:\w+:)?comment[^<>]*>/);
			if (!cm) {
				return;
			}
			const y = parsexmltag(cm[0]);
			const comment: RawComment = {
				author: (y.authorId && authors[y.authorId]) || "sheetjsghost",
				ref: y.ref,
				guid: y.guid,
				t: "",
			};
			const cell = decode_cell(y.ref);
			if (opts && opts.sheetRows && opts.sheetRows <= cell.r) {
				return;
			}
			const textMatch = str_match_xml_ns(x, "text");
			const rt = textMatch ? parse_si_simple(textMatch) : { r: "", t: "", h: "" };
			comment.r = rt.r;
			if (rt.r === "<t></t>") {
				rt.t = "";
				rt.h = "";
			}
			comment.t = (rt.t || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n");
			if (opts && opts.cellHTML) {
				comment.h = rt.h;
			}
			commentList.push(comment);
		});
	}
	return commentList;
}

/** Write comments XML */
export function write_comments_xml(data: [string, any[]][]): string {
	const o: string[] = [XML_HEADER, writextag("comments", null, { xmlns: XMLNS_main[0] })];

	const iauthor: string[] = [];
	o.push("<authors>");
	data.forEach((x) => {
		x[1].forEach((w: any) => {
			const a = escapexml(w.a);
			if (iauthor.indexOf(a) === -1) {
				iauthor.push(a);
				o.push("<author>" + a + "</author>");
			}
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
				lastauthor = iauthor.indexOf(escapexml(c.a));
			}
			if (c.T) {
				++tcnt;
			}
			ts.push(c.t == null ? "" : escapexml(c.t));
		});
		if (tcnt === 0) {
			d[1].forEach((c: any) => {
				o.push('<comment ref="' + d[0] + '" authorId="' + iauthor.indexOf(escapexml(c.a)) + '"><text>');
				o.push(writetag("t", c.t == null ? "" : escapexml(c.t)));
				o.push("</text></comment>");
			});
		} else {
			if (d[1][0] && d[1][0].T && d[1][0].ID) {
				lastauthor = iauthor.indexOf("tc=" + d[1][0].ID);
			}
			o.push('<comment ref="' + d[0] + '" authorId="' + lastauthor + '"><text>');
			let t = "Comment:\n    " + ts[0] + "\n";
			for (let i = 1; i < ts.length; ++i) {
				t += "Reply:\n    " + ts[i] + "\n";
			}
			o.push(writetag("t", escapexml(t)));
			o.push("</text></comment>");
		}
	});
	o.push("</commentList>");
	if (o.length > 2) {
		o.push("</comments>");
		o[1] = o[1].replace("/>", ">");
	}
	return o.join("");
}

/** Parse threaded comments XML [MS-XLSX] 2.1.17 */
export function parse_tcmnt_xml(data: string, opts?: any): RawComment[] {
	const out: RawComment[] = [];
	let pass = false;
	let comment: any = {};
	let tidx = 0;

	data.replace(tagregex, function xml_tcmnt(x: string, idx: number): string {
		const y: any = parsexmltag(x);
		switch (strip_ns(y[0])) {
			case "<?xml":
				break;
			case "<ThreadedComments":
			case "</ThreadedComments>":
				break;
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
				comment.t = data.slice(tidx, idx).replace(/\r\n/g, "\n").replace(/\r/g, "\n");
				break;
			case "<mentions":
			case "<mentions>":
				pass = true;
				break;
			case "</mentions>":
				pass = false;
				break;
			case "<extLst":
			case "<extLst>":
			case "</extLst>":
			case "<extLst/>":
				break;
			case "<ext":
				pass = true;
				break;
			case "</ext>":
				pass = false;
				break;
			default:
				break;
		}
		return x;
	});
	return out;
}

/** Write threaded comments XML */
export function write_tcmnt_xml(comments: [string, any[]][], people: string[], opts: { tcid: number }): string {
	const o: string[] = [XML_HEADER, writextag("ThreadedComments", null, { xmlns: XMLNS.TCMNT }).replace(/[/]>/, ">")];
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
			const tcopts: any = {
				ref: carr[0],
				id: "{54EE7951-7262-4200-6969-" + ("000000000000" + opts.tcid++).slice(-12) + "}",
			};
			if (idx === 0) {
				rootid = tcopts.id;
			} else {
				tcopts.parentId = rootid;
			}
			c.ID = tcopts.id;
			if (c.a) {
				tcopts.personId = "{54EE7950-7262-4200-6969-" + ("000000000000" + people.indexOf(c.a)).slice(-12) + "}";
			}
			o.push(writextag("threadedComment", writetag("text", c.t || ""), tcopts));
		});
	});
	o.push("</ThreadedComments>");
	return o.join("");
}

/** Parse people XML [MS-XLSX] 2.1.18 */
export function parse_people_xml(data: string): { name: string; id: string }[] {
	const out: { name: string; id: string }[] = [];
	let pass = false;
	data.replace(tagregex, function xml_people(x: string): string {
		const y: any = parsexmltag(x);
		switch (strip_ns(y[0])) {
			case "<?xml":
				break;
			case "<personList":
			case "</personList>":
				break;
			case "<person":
				out.push({ name: y.displayname, id: y.id });
				break;
			case "</person>":
				break;
			case "<extLst":
			case "<extLst>":
			case "</extLst>":
			case "<extLst/>":
				break;
			case "<ext":
				pass = true;
				break;
			case "</ext>":
				pass = false;
				break;
			default:
				break;
		}
		return x;
	});
	return out;
}

/** Write people XML */
export function write_people_xml(people: string[]): string {
	const o: string[] = [
		XML_HEADER,
		writextag("personList", null, {
			xmlns: XMLNS.TCMNT,
			"xmlns:x": XMLNS_main[0],
		}).replace(/[/]>/, ">"),
	];
	people.forEach((person, idx) => {
		o.push(
			writextag("person", null, {
				displayName: person,
				id: "{54EE7950-7262-4200-6969-" + ("000000000000" + idx).slice(-12) + "}",
				userId: person,
				providerId: "None",
			}),
		);
	});
	o.push("</personList>");
	return o.join("");
}
