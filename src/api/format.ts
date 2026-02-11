import type { CellObject } from "../types.js";
import { BErr } from "../types.js";
import { formatNumber } from "../ssf/format.js";
import { dateToSerialNumber } from "../utils/date.js";

/**
 * Attempt to format a cell value using the cell's number format or XF record.
 * Falls back to a plain string coercion if all formatting attempts fail.
 */
function safeFormatCell(cell: CellObject, value: any): string {
	const isDateCell = cell.t === "d" && value instanceof Date;
	// First try the explicit format string stored on the cell (cell.z)
	if (cell.z != null) {
		try {
			return (cell.w = formatNumber(cell.z, isDateCell ? dateToSerialNumber(value) : value));
		} catch {}
	}
	// Fall back to the numFmtId from the cell's XF style record;
	// default to format 14 (short date) for date cells, or 0 (General) otherwise
	try {
		return (cell.w = formatNumber(
			(cell.XF || {}).numFmtId || (isDateCell ? 14 : 0),
			isDateCell ? dateToSerialNumber(value) : value,
		));
	} catch {
		return "" + value;
	}
}

/**
 * Format a cell's value into its display string representation.
 *
 * Returns the cached `cell.w` if already computed, otherwise formats the value
 * using the cell's number format string or XF style information.
 *
 * @param cell - The cell object to format
 * @param value - Optional override value; if omitted, uses `cell.v`
 * @param options - Optional settings (e.g. `dateNF` for a default date format)
 * @returns The formatted display string, or empty string for null/blank cells
 */
export function formatCell(cell: CellObject, value?: any, options?: any): string {
	// Null cells or cells with type "z" (stub/blank) produce empty strings
	if (cell == null || cell.t == null || cell.t === "z") {
		return "";
	}
	// Return the cached formatted string if already computed
	if (cell.w !== undefined) {
		return cell.w;
	}
	// Apply the caller-supplied date format if the cell is a date without its own format
	if (cell.t === "d" && !cell.z && options && options.dateNF) {
		cell.z = options.dateNF;
	}
	// Error cells: look up the error code in BErr for a human-readable string
	if (cell.t === "e") {
		return BErr[cell.v as number] || String(cell.v);
	}
	if (value == null) {
		return safeFormatCell(cell, cell.v);
	}
	return safeFormatCell(cell, value);
}
