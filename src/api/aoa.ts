import type { WorkSheet, AOA2SheetOpts, Range, Sheet2JSONOpts } from "../types.js";
import { decodeCell, encodeRange, safeDecodeRange, getCell, setCell } from "../utils/cell.js";
import { dateToSerialNumber, localToUtc } from "../utils/date.js";
import { formatNumber } from "../ssf/format.js";
import { formatTable } from "../ssf/table.js";
import { sheetToJson } from "./json.js";

/**
 * Add an array-of-arrays to an existing worksheet, or create a new one.
 *
 * Each inner array represents a row, and each element within it a cell value.
 * Supports dense and sparse storage modes, origin offsets, date handling,
 * and automatic type detection (number, boolean, string, date, error).
 *
 * @param worksheet - An existing worksheet to append to, or `null` to create a new one
 * @param data - The array-of-arrays containing raw cell values
 * @param opts - Optional settings (origin, dense, dateNF, cellDates, UTC, date1904, nullError, sheetStubs)
 * @returns The updated or newly created worksheet
 */
export function addArrayToSheet(worksheet: WorkSheet | null, data: any[][], opts?: AOA2SheetOpts): WorkSheet {
	const options = opts || {};
	const dense = worksheet ? worksheet["!data"] != null : !!options.dense;
	const ws: WorkSheet = worksheet || (dense ? { "!data": [] } : {});
	if (dense && !ws["!data"]) {
		ws["!data"] = [];
	}

	// Determine the insertion origin (top-left cell for the data)
	let originRow = 0,
		originCol = 0;
	if (ws && options.origin != null) {
		if (typeof options.origin === "number") {
			originRow = options.origin;
		} else {
			const parsedOrigin = typeof options.origin === "string" ? decodeCell(options.origin) : options.origin;
			originRow = parsedOrigin.r;
			originCol = parsedOrigin.c;
		}
	}

	// Initialize range with sentinel values that will be narrowed during iteration
	const range: Range = { s: { c: 10000000, r: 10000000 }, e: { c: 0, r: 0 } };
	if (ws["!ref"]) {
		// Merge with the existing range so we don't shrink the sheet
		const existingRange = safeDecodeRange(ws["!ref"]);
		range.s.c = existingRange.s.c;
		range.s.r = existingRange.s.r;
		range.e.c = Math.max(range.e.c, existingRange.e.c);
		range.e.r = Math.max(range.e.r, existingRange.e.r);
		// origin of -1 means "append after the last row"
		if (originRow === -1) {
			range.e.r = originRow = ws["!ref"] ? existingRange.e.r + 1 : 0;
		}
	} else {
		range.s.c = range.e.c = range.s.r = range.e.r = 0;
	}

	let seen = false;
	for (let rowIdx = 0; rowIdx < data.length; ++rowIdx) {
		if (!data[rowIdx]) {
			continue;
		}
		if (!Array.isArray(data[rowIdx])) {
			throw new Error("arrayToSheet expects an array of arrays");
		}
		const targetRow = originRow + rowIdx;
		const rowData = data[rowIdx];
		for (let colIdx = 0; colIdx < rowData.length; ++colIdx) {
			if (typeof rowData[colIdx] === "undefined") {
				continue;
			}
			let cell: any = { v: rowData[colIdx], t: "" };
			const targetCol = originCol + colIdx;

			// Expand the tracked range to include this cell
			if (range.s.r > targetRow) {
				range.s.r = targetRow;
			}
			if (range.s.c > targetCol) {
				range.s.c = targetCol;
			}
			if (range.e.r < targetRow) {
				range.e.r = targetRow;
			}
			if (range.e.c < targetCol) {
				range.e.c = targetCol;
			}
			seen = true;

			// If the value is a plain object (not an array or Date), treat it as a pre-built cell object
			if (
				rowData[colIdx] &&
				typeof rowData[colIdx] === "object" &&
				!Array.isArray(rowData[colIdx]) &&
				!(rowData[colIdx] instanceof Date)
			) {
				cell = rowData[colIdx];
			} else {
				// Array values encode [value, formula]
				if (Array.isArray(cell.v)) {
					cell.f = rowData[colIdx][1];
					cell.v = cell.v[0];
				}
				if (cell.v === null) {
					if (cell.f) {
						cell.t = "n";
					} else if (options.nullError) {
						cell.t = "e";
						cell.v = 0;
					} else if (!options.sheetStubs) {
						continue;
					} else {
						cell.t = "z";
					}
				} else if (typeof cell.v === "number") {
					if (isFinite(cell.v)) {
						cell.t = "n";
					} else if (isNaN(cell.v)) {
						cell.t = "e";
						cell.v = 0x0f; // #VALUE! error code
					} else {
						cell.t = "e";
						cell.v = 0x07; // #DIV/0! error code
					}
				} else if (typeof cell.v === "boolean") {
					cell.t = "b";
				} else if (cell.v instanceof Date) {
					cell.z = options.dateNF || formatTable[14]; // default short date format
					if (!options.UTC) {
						cell.v = localToUtc(cell.v);
					}
					if (options.cellDates) {
						cell.t = "d";
						cell.w = formatNumber(cell.z, dateToSerialNumber(cell.v, options.date1904));
					} else {
						// Store dates as serial numbers by default
						cell.t = "n";
						cell.v = dateToSerialNumber(cell.v, options.date1904);
						cell.w = formatNumber(cell.z, cell.v);
					}
				} else {
					cell.t = "s";
				}
			}

			// Preserve any existing number format from a prior cell at this position
			const existingCell = getCell(ws, targetRow, targetCol);
			if (existingCell?.z && !cell.z) {
				cell.z = existingCell.z;
			}
			setCell(ws, targetRow, targetCol, cell);
		}
	}
	// Only write the ref if we actually saw data (sentinel check: 10000000 > max valid column)
	if (seen && range.s.c < 10400000) {
		ws["!ref"] = encodeRange(range);
	}
	return ws;
}

/**
 * Create a new worksheet from an array-of-arrays.
 *
 * This is a convenience wrapper around `addArrayToSheet` that always creates
 * a fresh worksheet.
 *
 * @param data - The array-of-arrays containing raw cell values
 * @param opts - Optional settings (same as `addArrayToSheet`)
 * @returns A new worksheet populated with the given data
 */
export function arrayToSheet(data: any[][], opts?: AOA2SheetOpts): WorkSheet {
	return addArrayToSheet(null, data, opts);
}

/**
 * Convert a worksheet to an array-of-arrays.
 *
 * This is a convenience wrapper around `sheetToJson(sheet, { header: 1 })`
 * that mirrors `arrayToSheet` for callers that prefer explicit conversion
 * pairs.
 *
 * @param sheet - The worksheet to convert
 * @param opts - Optional conversion settings, except `header` which is fixed to array output
 * @returns A two-dimensional array of worksheet values
 */
export function sheetToArray<T = unknown>(sheet: WorkSheet, opts?: Omit<Sheet2JSONOpts, "header">): T[][] {
	return sheetToJson<T[]>(sheet, { ...opts, header: 1 });
}
