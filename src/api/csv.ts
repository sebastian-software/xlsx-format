import type { WorkSheet, Sheet2CSVOpts, Range } from "../types.js";
import { encode_col, encode_row, safe_decode_range } from "../utils/cell.js";
import { format_cell } from "./format.js";

const qreg = /"/g;

function make_csv_row(
	sheet: WorkSheet,
	r: Range,
	R: number,
	cols: string[],
	fs: number,
	rs: number,
	FS: string,
	w: number,
	o: any,
): string | null {
	let isempty = true;
	const row: string[] = [];
	const rr = encode_row(R);
	const dense = (sheet as any)["!data"] != null;
	const datarow = dense ? (sheet as any)["!data"][R] || [] : [];

	for (let C = r.s.c; C <= r.e.c; ++C) {
		if (!cols[C]) {
			continue;
		}
		const val = dense ? datarow[C] : (sheet as any)[cols[C] + rr];
		let txt = "";
		if (val == null) {
			txt = "";
		} else if (val.v != null) {
			isempty = false;
			txt = "" + (o.rawNumbers && val.t === "n" ? val.v : format_cell(val, null, o));
			for (let i = 0, cc = 0; i !== txt.length; ++i) {
				if (
					(cc = txt.charCodeAt(i)) === fs ||
					cc === rs ||
					cc === 10 ||
					cc === 13 ||
					cc === 34 ||
					o.forceQuotes
				) {
					txt = '"' + txt.replace(qreg, '""') + '"';
					break;
				}
			}
			if (txt === "ID" && w === 0 && row.length === 0) {
				txt = '"ID"';
			}
		} else if (val.f != null && !val.F) {
			isempty = false;
			txt = "=" + val.f;
			if (txt.indexOf(",") >= 0) {
				txt = '"' + txt.replace(qreg, '""') + '"';
			}
		} else {
			txt = "";
		}
		row.push(txt);
	}
	if (o.strip) {
		while (row[row.length - 1] === "") {
			--row.length;
		}
	}
	if (o.blankrows === false && isempty) {
		return null;
	}
	return row.join(FS);
}

/** Convert a worksheet to CSV string */
export function sheet_to_csv(sheet: WorkSheet, opts?: Sheet2CSVOpts): string {
	const out: string[] = [];
	const o: any = opts == null ? {} : opts;
	if (sheet == null || sheet["!ref"] == null) {
		return "";
	}
	const r = safe_decode_range(sheet["!ref"]);
	const FS = o.FS !== undefined ? o.FS : ",";
	const fs = FS.charCodeAt(0);
	const RS = o.RS !== undefined ? o.RS : "\n";
	const rs = RS.charCodeAt(0);
	const cols: string[] = [];
	const colinfo: any[] = (o.skipHidden && sheet["!cols"]) || [];
	const rowinfo: any[] = (o.skipHidden && sheet["!rows"]) || [];

	for (let C = r.s.c; C <= r.e.c; ++C) {
		if (!(colinfo[C] || {}).hidden) {
			cols[C] = encode_col(C);
		}
	}

	let w = 0;
	for (let R = r.s.r; R <= r.e.r; ++R) {
		if ((rowinfo[R] || {}).hidden) {
			continue;
		}
		const row = make_csv_row(sheet, r, R, cols, fs, rs, FS, w, o);
		if (row == null) {
			continue;
		}
		if (row || o.blankrows !== false) {
			out.push((w++ ? RS : "") + row);
		}
	}
	return out.join("");
}

/** Convert a worksheet to tab-separated text */
export function sheet_to_txt(sheet: WorkSheet, opts?: Sheet2CSVOpts): string {
	const o: any = opts || {};
	o.FS = "\t";
	o.RS = "\n";
	return sheet_to_csv(sheet, o);
}
