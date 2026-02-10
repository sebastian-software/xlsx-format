import type { WorkSheet, Sheet2CSVOpts, Range } from "../types.js";
import { encodeCol, encodeRow, safeDecodeRange } from "../utils/cell.js";
import { formatCell } from "./format.js";

const qreg = /"/g;

function buildCsvRow(
	sheet: WorkSheet,
	r: Range,
	R: number,
	cols: string[],
	fieldSepCode: number,
	recordSepCode: number,
	FS: string,
	rowCount: number,
	o: any,
): string | null {
	let isempty = true;
	const row: string[] = [];
	const rr = encodeRow(R);
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
			txt = "" + (o.rawNumbers && val.t === "n" ? val.v : formatCell(val, null, o));
			for (let i = 0, cc = 0; i !== txt.length; ++i) {
				if (
					(cc = txt.charCodeAt(i)) === fieldSepCode ||
					cc === recordSepCode ||
					cc === 10 ||
					cc === 13 ||
					cc === 34 ||
					o.forceQuotes
				) {
					txt = '"' + txt.replace(qreg, '""') + '"';
					break;
				}
			}
			if (txt === "ID" && rowCount === 0 && row.length === 0) {
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
export function sheetToCsv(sheet: WorkSheet, opts?: Sheet2CSVOpts): string {
	const out: string[] = [];
	const o: any = opts == null ? {} : opts;
	if (sheet == null || sheet["!ref"] == null) {
		return "";
	}
	const r = safeDecodeRange(sheet["!ref"]);
	const FS = o.FS !== undefined ? o.FS : ",";
	const fieldSepCode = FS.charCodeAt(0);
	const RS = o.RS !== undefined ? o.RS : "\n";
	const recordSepCode = RS.charCodeAt(0);
	const cols: string[] = [];
	const colinfo: any[] = (o.skipHidden && sheet["!cols"]) || [];
	const rowinfo: any[] = (o.skipHidden && sheet["!rows"]) || [];

	for (let C = r.s.c; C <= r.e.c; ++C) {
		if (!(colinfo[C] || {}).hidden) {
			cols[C] = encodeCol(C);
		}
	}

	let rowCount = 0;
	for (let R = r.s.r; R <= r.e.r; ++R) {
		if ((rowinfo[R] || {}).hidden) {
			continue;
		}
		const row = buildCsvRow(sheet, r, R, cols, fieldSepCode, recordSepCode, FS, rowCount, o);
		if (row == null) {
			continue;
		}
		if (row || o.blankrows !== false) {
			out.push((rowCount++ ? RS : "") + row);
		}
	}
	return out.join("");
}

/** Convert a worksheet to tab-separated text */
export function sheetToTxt(sheet: WorkSheet, opts?: Sheet2CSVOpts): string {
	const o: any = opts || {};
	o.FS = "\t";
	o.RS = "\n";
	return sheetToCsv(sheet, o);
}
