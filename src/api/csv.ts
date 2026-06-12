import type { WorkSheet, Sheet2CSVOpts, Range } from "../types.js";
import { encodeCol, safeDecodeRange, getCell } from "../utils/cell.js";
import { formatCell } from "./format.js";
import { arrayToSheet } from "./aoa.js";

/** Regex to match double-quote characters for CSV escaping (doubled inside quoted fields) */
const qreg = /"/g;
const CELL_REF_RE = /^[A-Z]+[1-9][0-9]*$/;
const MAX_EXPORT_CELLS = 1000000;

function rangeCellCount(range: Range): number {
	const rows = range.e.r - range.s.r + 1;
	const cols = range.e.c - range.s.c + 1;
	return rows > 0 && cols > 0 ? rows * cols : 0;
}

function occupiedRangeEnd(sheet: WorkSheet, range: Range): { r: number; c: number } | null {
	let maxRow = -1;
	let maxCol = -1;
	const data = (sheet as any)["!data"];
	if (data != null) {
		for (const rowKey of Object.keys(data)) {
			const rowIdx = Number(rowKey);
			if (!Number.isInteger(rowIdx) || rowIdx < range.s.r || rowIdx > range.e.r) {
				continue;
			}
			const row = data[rowIdx];
			if (!row) {
				continue;
			}
			for (const colKey of Object.keys(row)) {
				const colIdx = Number(colKey);
				if (!Number.isInteger(colIdx) || colIdx < range.s.c || colIdx > range.e.c || row[colIdx] == null) {
					continue;
				}
				if (rowIdx > maxRow) {
					maxRow = rowIdx;
				}
				if (colIdx > maxCol) {
					maxCol = colIdx;
				}
			}
		}
	} else {
		for (const ref of Object.keys(sheet)) {
			if (!CELL_REF_RE.test(ref) || (sheet as any)[ref] == null) {
				continue;
			}
			let rowIdx = 0;
			let colIdx = 0;
			for (let i = 0; i < ref.length; ++i) {
				const charCode = ref.charCodeAt(i);
				if (charCode >= 65 && charCode <= 90) {
					colIdx = 26 * colIdx + charCode - 64;
				} else {
					rowIdx = 10 * rowIdx + charCode - 48;
				}
			}
			--rowIdx;
			--colIdx;
			if (rowIdx < range.s.r || rowIdx > range.e.r || colIdx < range.s.c || colIdx > range.e.c) {
				continue;
			}
			if (rowIdx > maxRow) {
				maxRow = rowIdx;
			}
			if (colIdx > maxCol) {
				maxCol = colIdx;
			}
		}
	}
	return maxRow === -1 ? null : { r: maxRow, c: maxCol };
}

function clampLargeExportRange(sheet: WorkSheet, range: Range): Range | null {
	if (rangeCellCount(range) <= MAX_EXPORT_CELLS) {
		return range;
	}
	const end = occupiedRangeEnd(sheet, range);
	if (!end) {
		return null;
	}
	return {
		s: { r: range.s.r, c: range.s.c },
		e: {
			r: Math.max(range.s.r, Math.min(range.e.r, end.r)),
			c: Math.max(range.s.c, Math.min(range.e.c, end.c)),
		},
	};
}

function escapeFormulaText(txt: string, options: any): string {
	if (!options.escapeFormulae || txt.length === 0) {
		return txt;
	}
	switch (txt.charCodeAt(0)) {
		case 0x09:
		case 0x0d:
		case 0x2b:
		case 0x2d:
		case 0x3d:
		case 0x40:
			return "'" + txt;
		default:
			return txt;
	}
}

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
			txt = escapeFormulaText(txt, options);
			// Check each character: if the text contains the field separator,
			// record separator, LF (10), CR (13), or double-quote (34), wrap in quotes
			for (let i = 0, charCode = 0; i !== txt.length; ++i) {
				if (
					(charCode = txt.charCodeAt(i)) === fieldSepCode ||
					charCode === recordSepCode ||
					charCode === 10 || // LF
					charCode === 13 || // CR
					charCode === 34 || // double-quote
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
			txt = escapeFormulaText(txt, options);
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
	const range = clampLargeExportRange(sheet, safeDecodeRange(sheet["!ref"]));
	if (!range) {
		return "";
	}
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
		const row = buildCsvRow(
			sheet,
			range,
			rowIdx,
			cols,
			fieldSepCode,
			recordSepCode,
			fieldSeparator,
			rowCount,
			options,
		);
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

/**
 * Parse an RFC 4180 CSV string into a 2D array of values.
 *
 * Handles quoted fields, escaped double-quotes, and newlines within quotes.
 */
function parseCsv(text: string, sep: string): any[][] {
	const rows: any[][] = [];
	let row: any[] = [];
	let i = 0;
	const len = text.length;

	while (i <= len) {
		if (i === len) {
			// End of input — push final row if it has content or there are already rows
			if (row.length > 0 || rows.length > 0) {
				row.push("");
				rows.push(row);
			}
			break;
		}

		if (text[i] === '"') {
			// Quoted field
			let val = "";
			i++; // skip opening quote
			while (i < len) {
				if (text[i] === '"') {
					if (i + 1 < len && text[i + 1] === '"') {
						// Escaped double-quote
						val += '"';
						i += 2;
					} else {
						// Closing quote
						i++; // skip closing quote
						break;
					}
				} else {
					val += text[i];
					i++;
				}
			}
			row.push(val);
			// After closing quote, expect separator, newline, or end
			if (i < len && text[i] === sep) {
				i++;
			} else if (i < len && (text[i] === "\r" || text[i] === "\n")) {
				if (text[i] === "\r" && i + 1 < len && text[i + 1] === "\n") {
					i++;
				}
				i++;
				rows.push(row);
				row = [];
			}
		} else if (text[i] === sep) {
			row.push("");
			i++;
		} else if (text[i] === "\r" || text[i] === "\n") {
			if (text[i] === "\r" && i + 1 < len && text[i + 1] === "\n") {
				i++;
			}
			i++;
			rows.push(row);
			row = [];
		} else {
			// Unquoted field
			let val = "";
			while (i < len && text[i] !== sep && text[i] !== "\r" && text[i] !== "\n") {
				val += text[i];
				i++;
			}
			row.push(val);
			if (i < len && text[i] === sep) {
				i++;
			} else if (i < len && (text[i] === "\r" || text[i] === "\n")) {
				if (text[i] === "\r" && i + 1 < len && text[i + 1] === "\n") {
					i++;
				}
				i++;
				rows.push(row);
				row = [];
			}
		}
	}

	return rows;
}

/** Try to coerce a string value to a number or boolean */
function coerceValue(val: string): string | number | boolean {
	if (val === "") {
		return val;
	}
	if (val === "TRUE" || val === "true") {
		return true;
	}
	if (val === "FALSE" || val === "false") {
		return false;
	}
	const num = Number(val);
	if (val.length > 0 && !isNaN(num) && isFinite(num)) {
		return num;
	}
	return val;
}

/**
 * Parse a CSV string into a WorkSheet.
 *
 * @param text - CSV text to parse
 * @param opts - Optional: { FS: field separator (default ",") }
 * @returns A WorkSheet with the parsed data
 */
export function csvToSheet(text: string, opts?: { FS?: string }): WorkSheet {
	const sep = (opts && opts.FS) || ",";
	const rows = parseCsv(text, sep);
	const data: any[][] = rows.map((row) => row.map(coerceValue));
	return arrayToSheet(data);
}
