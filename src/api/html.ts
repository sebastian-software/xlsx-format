import type { WorkSheet, Sheet2HTMLOpts, Range } from "../types.js";
import { BErr } from "../types.js";
import { encode_col, encode_row, decode_range } from "../utils/cell.js";
import { escapehtml } from "../xml/escape.js";
import { writextag } from "../xml/writer.js";
import { format_cell } from "./format.js";

const HTML_BEGIN = '<html><head><meta charset="utf-8"/><title>SheetJS Table Export</title></head><body>';
const HTML_END = "</body></html>";

function make_html_row(ws: WorkSheet, r: Range, R: number, o: Sheet2HTMLOpts): string {
	const M = ws["!merges"] || [];
	const oo: string[] = [];
	const dense = (ws as any)["!data"] != null;

	for (let C = r.s.c; C <= r.e.c; ++C) {
		let RS = 0,
			CS = 0;
		for (let j = 0; j < M.length; ++j) {
			if (M[j].s.r > R || M[j].s.c > C) {
				continue;
			}
			if (M[j].e.r < R || M[j].e.c < C) {
				continue;
			}
			if (M[j].s.r < R || M[j].s.c < C) {
				RS = -1;
				break;
			}
			RS = M[j].e.r - M[j].s.r + 1;
			CS = M[j].e.c - M[j].s.c + 1;
			break;
		}
		if (RS < 0) {
			continue;
		}

		const coord = encode_col(C) + encode_row(R);
		let cell: any = dense ? ((ws as any)["!data"][R] || [])[C] : (ws as any)[coord];

		if (cell && cell.t === "n" && cell.v != null && !isFinite(cell.v)) {
			if (isNaN(cell.v)) {
				cell = { t: "e", v: 0x24, w: BErr[0x24] };
			} else {
				cell = { t: "e", v: 0x07, w: BErr[0x07] };
			}
		}

		let w = (cell && cell.v != null && (cell.h || escapehtml(cell.w || (format_cell(cell), cell.w) || ""))) || "";

		const sp: Record<string, any> = {};
		if (RS > 1) {
			sp.rowspan = String(RS);
		}
		if (CS > 1) {
			sp.colspan = String(CS);
		}

		if (o.editable) {
			w = '<span contenteditable="true">' + w + "</span>";
		} else if (cell) {
			sp["data-t"] = (cell && cell.t) || "z";
			if (cell.v != null) {
				sp["data-v"] = escapehtml(cell.v instanceof Date ? cell.v.toISOString() : String(cell.v));
			}
			if (cell.z != null) {
				sp["data-z"] = String(cell.z);
			}
			if (cell.f != null) {
				sp["data-f"] = escapehtml(cell.f);
			}
			if (
				cell.l &&
				(cell.l.Target || "#").charAt(0) !== "#" &&
				(!o.sanitizeLinks || (cell.l.Target || "").slice(0, 11).toLowerCase() !== "javascript:")
			) {
				w = '<a href="' + escapehtml(cell.l.Target) + '">' + w + "</a>";
			}
		}
		sp.id = (o.id || "sjs") + "-" + coord;
		oo.push(writextag("td", w, sp));
	}

	return "<tr>" + oo.join("") + "</tr>";
}

function make_html_preamble(_ws: WorkSheet, _r: Range, o: Sheet2HTMLOpts): string {
	return "<table" + (o && o.id ? ' id="' + o.id + '"' : "") + ">";
}

/** Convert a worksheet to an HTML table string */
export function sheet_to_html(ws: WorkSheet, opts?: Sheet2HTMLOpts): string {
	const o: Sheet2HTMLOpts = opts || {};
	const header = o.header != null ? o.header : HTML_BEGIN;
	const footer = o.footer != null ? o.footer : HTML_END;
	const out: string[] = [header];
	const r = decode_range(ws["!ref"] || "A1");
	out.push(make_html_preamble(ws, r, o));
	if (ws["!ref"]) {
		for (let R = r.s.r; R <= r.e.r; ++R) {
			out.push(make_html_row(ws, r, R, o));
		}
	}
	out.push("</table>" + footer);
	return out.join("");
}
