import type { WorkSheet, Sheet2CSVOpts, Range } from "../types.js";
import { encodeCol, safeDecodeRange, getCell } from "../utils/cell.js";
import { formatCell } from "./format.js";

const qreg = /"/g;

function buildCsvRow(
	sheet: WorkSheet,
	range: Range,
	rowIndex: number,
	cols: string[],
	fieldSepCode: number,
	recordSepCode: number,
	fieldSeparator: string,
	rowCount: number,
	options: any,
): string | null {
	let isempty = true;
	const row: string[] = [];

	for (let colIdx = range.s.c; colIdx <= range.e.c; ++colIdx) {
		if (!cols[colIdx]) {
			continue;
		}
		const val = getCell(sheet, rowIndex, colIdx);
		let txt = "";
		if (val == null) {
			txt = "";
		} else if (val.v != null) {
			isempty = false;
			txt = "" + (options.rawNumbers && val.t === "n" ? val.v : formatCell(val, null, options));
			for (let i = 0, charCode = 0; i !== txt.length; ++i) {
				if (
					(charCode = txt.charCodeAt(i)) === fieldSepCode ||
					charCode === recordSepCode ||
					charCode === 10 ||
					charCode === 13 ||
					charCode === 34 ||
					options.forceQuotes
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
	if (options.strip) {
		while (row[row.length - 1] === "") {
			--row.length;
		}
	}
	if (options.blankrows === false && isempty) {
		return null;
	}
	return row.join(fieldSeparator);
}

/** Convert a worksheet to CSV string */
export function sheetToCsv(sheet: WorkSheet, opts?: Sheet2CSVOpts): string {
	const out: string[] = [];
	const options: any = opts == null ? {} : opts;
	if (sheet == null || sheet["!ref"] == null) {
		return "";
	}
	const range = safeDecodeRange(sheet["!ref"]);
	const fieldSeparator = options.FS !== undefined ? options.FS : ",";
	const fieldSepCode = fieldSeparator.charCodeAt(0);
	const recordSeparator = options.RS !== undefined ? options.RS : "\n";
	const recordSepCode = recordSeparator.charCodeAt(0);
	const cols: string[] = [];
	const colinfo: any[] = (options.skipHidden && sheet["!cols"]) || [];
	const rowinfo: any[] = (options.skipHidden && sheet["!rows"]) || [];

	for (let colIdx = range.s.c; colIdx <= range.e.c; ++colIdx) {
		if (!(colinfo[colIdx] || {}).hidden) {
			cols[colIdx] = encodeCol(colIdx);
		}
	}

	let rowCount = 0;
	for (let rowIdx = range.s.r; rowIdx <= range.e.r; ++rowIdx) {
		if ((rowinfo[rowIdx] || {}).hidden) {
			continue;
		}
		const row = buildCsvRow(sheet, range, rowIdx, cols, fieldSepCode, recordSepCode, fieldSeparator, rowCount, options);
		if (row == null) {
			continue;
		}
		if (row || options.blankrows !== false) {
			out.push((rowCount++ ? recordSeparator : "") + row);
		}
	}
	return out.join("");
}

/** Convert a worksheet to tab-separated text */
export function sheetToTxt(sheet: WorkSheet, opts?: Sheet2CSVOpts): string {
	const options: any = opts || {};
	options.FS = "\t";
	options.RS = "\n";
	return sheetToCsv(sheet, options);
}
