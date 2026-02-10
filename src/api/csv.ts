import type { WorkSheet, Sheet2CSVOpts, Range } from "../types.js";
import { encodeCol, safeDecodeRange, getCell } from "../utils/cell.js";
import { formatCell } from "./format.js";

/** Regex to match double-quote characters for CSV escaping (doubled inside quoted fields) */
const qreg = /"/g;

/**
 * Build a single CSV row string from a worksheet row.
 *
 * Handles value quoting (when field/record separators, newlines, or double
 * quotes appear in the text), the special "ID" SYLK-avoidance quoting,
 * formula-only cells, and the `strip`/`blankrows` options.
 *
 * @returns The joined CSV row string, or `null` if the row is blank and blankrows is disabled
 */
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
		// Skip hidden columns (cols[colIdx] is undefined for hidden ones)
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
			// Check each character: if the text contains the field separator,
			// record separator, LF (10), CR (13), or double-quote (34), wrap in quotes
			for (let i = 0, charCode = 0; i !== txt.length; ++i) {
				if (
					(charCode = txt.charCodeAt(i)) === fieldSepCode ||
					charCode === recordSepCode ||
					charCode === 10 ||  // LF
					charCode === 13 ||  // CR
					charCode === 34 ||  // double-quote
					options.forceQuotes
				) {
					txt = '"' + txt.replace(qreg, '""') + '"';
					break;
				}
			}
			// Quote bare "ID" in the first cell to avoid misdetection as a SYLK file
			if (txt === "ID" && rowCount === 0 && row.length === 0) {
				txt = '"ID"';
			}
		} else if (val.f != null && !val.F) {
			// Cell has a formula but no cached value and is not part of an array formula
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
	// Strip trailing empty cells from the row if requested
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

/**
 * Convert a worksheet to a CSV string.
 *
 * Supports customizable field and record separators, hidden row/column
 * skipping, blank-row suppression, raw number output, and forced quoting.
 *
 * @param sheet - The worksheet to convert
 * @param opts - Optional CSV generation options (FS, RS, skipHidden, strip, blankrows, rawNumbers, forceQuotes)
 * @returns The CSV string representation of the worksheet
 */
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

	// Build column-letter lookup, skipping hidden columns when skipHidden is set
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
		// Prepend the record separator for all rows after the first
		if (row || options.blankrows !== false) {
			out.push((rowCount++ ? recordSeparator : "") + row);
		}
	}
	return out.join("");
}

/**
 * Convert a worksheet to a tab-separated values (TSV) string.
 *
 * This is a convenience wrapper around `sheetToCsv` with tab as the field
 * separator and newline as the record separator.
 *
 * @param sheet - The worksheet to convert
 * @param opts - Optional CSV/TSV generation options (same as `sheetToCsv`)
 * @returns The TSV string representation of the worksheet
 */
export function sheetToTxt(sheet: WorkSheet, opts?: Sheet2CSVOpts): string {
	const options: any = opts || {};
	options.FS = "\t";
	options.RS = "\n";
	return sheetToCsv(sheet, options);
}
