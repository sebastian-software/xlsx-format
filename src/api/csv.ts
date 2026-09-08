import type { WorkSheet, Sheet2CSVOpts, Range, CSV2SheetOpts } from "../types.js";
import { XlsxError } from "../errors.js";
import { encodeCol, safeDecodeRange, getCell } from "../utils/cell.js";
import { clampLargeExportRange } from "../utils/export-range.js";
import { formatCellForOutput, getCellDateTimeFormatKind } from "./format.js";
import { arrayToSheet } from "./aoa.js";
import {
	DEFAULT_MAX_EXPORT_CELLS,
	DEFAULT_MAX_IMPORT_CELLS,
	WorksheetCellBudget,
	worksheetOptionLimit,
	XLSX_MAX_COLUMNS,
	XLSX_MAX_ROWS,
} from "../utils/worksheet-budget.js";

/** Regex to match double-quote characters for CSV escaping (doubled inside quoted fields) */
const qreg = /"/g;

function containsSeparatorCharacter(text: string, separator: string): boolean {
	return separator.length > 0 && text.includes(separator.charAt(0));
}

function quoteCsvField(text: string, fieldSeparator: string, recordSeparator: string, forceQuotes?: boolean): string {
	if (
		forceQuotes ||
		text.includes('"') ||
		text.includes("\r") ||
		text.includes("\n") ||
		containsSeparatorCharacter(text, fieldSeparator) ||
		containsSeparatorCharacter(text, recordSeparator)
	) {
		return '"' + text.replace(qreg, '""') + '"';
	}
	return text;
}

function escapeFormulaText(txt: string, options: any): string {
	if (options.escapeFormulae === false || txt.length === 0) {
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
	fieldSeparator: string,
	recordSeparator: string,
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
			const fmtKind = getCellDateTimeFormatKind(val, options);
			const useDateOutput =
				options.dateOutput === "iso" && (val.t === "d" || fmtKind === "date" || fmtKind === "datetime");
			txt =
				"" +
				(options.rawNumbers && val.t === "n" && !useDateOutput
					? val.v
					: formatCellForOutput(val, null, options));
			txt = escapeFormulaText(txt, options);
			txt = quoteCsvField(txt, fieldSeparator, recordSeparator, options.forceQuotes);
			// Quote bare "ID" in the first cell to avoid misdetection as a SYLK file
			if (txt === "ID" && rowCount === 0 && row.length === 0) {
				txt = '"ID"';
			}
		} else if (val.f != null && !val.F) {
			// Cell has a formula but no cached value and is not part of an array formula
			isempty = false;
			txt = "=" + val.f;
			txt = escapeFormulaText(txt, options);
			txt = quoteCsvField(txt, fieldSeparator, recordSeparator, options.forceQuotes);
		} else {
			txt = "";
		}
		row.push(txt);
	}
	// Strip trailing empty cells from the row if requested
	if (options.strip) {
		while (row.at(-1) === "") {
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
	const budget = new WorksheetCellBudget(options.maxWorksheetCells, DEFAULT_MAX_EXPORT_CELLS);
	const range = clampLargeExportRange(sheet, safeDecodeRange(sheet["!ref"]), budget);
	if (!range) {
		return "";
	}
	const fieldSeparator = options.FS !== undefined ? options.FS : ",";
	const recordSeparator = options.RS !== undefined ? options.RS : "\n";

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
		const row = buildCsvRow(sheet, range, rowIdx, cols, fieldSeparator, recordSeparator, rowCount, options);
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
function parseCsv(text: string, sep: string, opts: CSV2SheetOpts): any[][] {
	const rows: any[][] = [];
	let row: any[] = [];
	let i = 0;
	const len = text.length;
	let afterSeparator = false;
	let rowActive = false;
	let rowCount = 0;
	const maxRows = worksheetOptionLimit(opts.maxWorksheetRows, XLSX_MAX_ROWS, "maxWorksheetRows");
	const sheetRows = worksheetOptionLimit(opts.sheetRows, 0, "sheetRows");
	const budget = new WorksheetCellBudget(opts.maxWorksheetCells, DEFAULT_MAX_IMPORT_CELLS);

	const beginRow = (): void => {
		if (rowActive) {
			return;
		}
		++rowCount;
		if (rowCount > XLSX_MAX_ROWS) {
			throw new XlsxError("MALFORMED", `CSV data exceeds XLSX row limit ${XLSX_MAX_ROWS}`);
		}
		if (rowCount > maxRows) {
			throw new XlsxError("LIMIT_EXCEEDED", `worksheet row count ${rowCount} exceeds limit ${maxRows}`);
		}
		rowActive = true;
	};
	const pushField = (value: string): void => {
		if (row.length >= XLSX_MAX_COLUMNS) {
			throw new XlsxError("MALFORMED", `CSV row exceeds XLSX column limit ${XLSX_MAX_COLUMNS}`);
		}
		budget.charge("worksheet cell", 1);
		row.push(value);
	};
	const finishRow = (): boolean => {
		beginRow();
		if (row.length === 0) {
			budget.charge("worksheet cell", 1);
		}
		rows.push(row);
		row = [];
		rowActive = false;
		afterSeparator = false;
		return sheetRows > 0 && rows.length >= sheetRows;
	};

	while (i < len) {
		beginRow();
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
			pushField(val);
			afterSeparator = false;
			// After closing quote, expect separator, newline, or end
			if (i < len && text[i] === sep) {
				i++;
				afterSeparator = true;
			} else if (i < len && (text[i] === "\r" || text[i] === "\n")) {
				if (text[i] === "\r" && i + 1 < len && text[i + 1] === "\n") {
					i++;
				}
				i++;
				if (finishRow()) {
					return rows;
				}
			}
		} else if (text[i] === sep) {
			pushField("");
			i++;
			afterSeparator = true;
		} else if (text[i] === "\r" || text[i] === "\n") {
			if (afterSeparator) {
				pushField("");
			}
			if (text[i] === "\r" && i + 1 < len && text[i + 1] === "\n") {
				i++;
			}
			i++;
			if (finishRow()) {
				return rows;
			}
		} else {
			// Unquoted field
			let val = "";
			while (i < len && text[i] !== sep && text[i] !== "\r" && text[i] !== "\n") {
				val += text[i];
				i++;
			}
			pushField(val);
			afterSeparator = false;
			if (i < len && text[i] === sep) {
				i++;
				afterSeparator = true;
			} else if (i < len && (text[i] === "\r" || text[i] === "\n")) {
				if (text[i] === "\r" && i + 1 < len && text[i + 1] === "\n") {
					i++;
				}
				i++;
				if (finishRow()) {
					return rows;
				}
			}
		}
	}
	if (afterSeparator) {
		pushField("");
	}
	if (row.length > 0) {
		finishRow();
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
export function csvToSheet(text: string, opts?: CSV2SheetOpts): WorkSheet {
	const sep = (opts && opts.FS) || ",";
	const rows = parseCsv(text, sep, opts || {});
	// Keep blank records visible in the worksheet by giving them a stub cell.
	const data: any[][] = rows.map((row) => (row.length === 0 ? [null] : row.map((value) => coerceValue(value))));
	return arrayToSheet(data, { sheetStubs: true });
}
