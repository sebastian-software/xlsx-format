import type { CellObject } from "../types.js";
import { BErr } from "../types.js";
import { formatNumber, getDateTimeFormatKind, type DateTimeFormatKind } from "../ssf/format.js";
import { DEFAULT_FORMAT_MAP, DEFAULT_FORMAT_STRINGS, formatTable } from "../ssf/table.js";
import { dateToSerialNumber, serialNumberToDate } from "../utils/date.js";

function resolveNumberFormat(fmt: unknown, options?: any): string | undefined {
	if (typeof fmt === "string") {
		return fmt === "m/d/yy" && options?.dateNF ? String(options.dateNF) : fmt;
	}
	if (typeof fmt !== "number") {
		return undefined;
	}
	if (fmt === 14 && options?.dateNF) {
		return String(options.dateNF);
	}
	const table = options?.table || formatTable;
	return table[fmt] || table[DEFAULT_FORMAT_MAP[fmt]] || DEFAULT_FORMAT_STRINGS[fmt];
}

function resolveCellNumberFormat(cell: CellObject, options?: any): string | undefined {
	return resolveNumberFormat(cell.z ?? cell.XF?.numFmtId, options);
}

/** Return the date/time classification for a cell's number format. */
export function getCellDateTimeFormatKind(cell: CellObject, options?: any): DateTimeFormatKind {
	const fmt = resolveCellNumberFormat(cell, options);
	return fmt ? getDateTimeFormatKind(fmt) : "none";
}

function pad2(value: number): string {
	return value < 10 ? "0" + value : "" + value;
}

function formatDateIso(date: Date, kind: DateTimeFormatKind, options?: any): string {
	const useUtc = options?.UTC !== false;
	const year = useUtc ? date.getUTCFullYear() : date.getFullYear();
	const month = (useUtc ? date.getUTCMonth() : date.getMonth()) + 1;
	const day = useUtc ? date.getUTCDate() : date.getDate();
	const hours = useUtc ? date.getUTCHours() : date.getHours();
	const minutes = useUtc ? date.getUTCMinutes() : date.getMinutes();
	const seconds = useUtc ? date.getUTCSeconds() : date.getSeconds();
	const datePart = `${year}-${pad2(month)}-${pad2(day)}`;
	if (kind === "date") {
		return datePart;
	}
	return `${datePart}T${pad2(hours)}:${pad2(minutes)}:${pad2(seconds)}`;
}

function inferDateKind(date: Date, options?: any): DateTimeFormatKind {
	const useUtc = options?.UTC !== false;
	const hours = useUtc ? date.getUTCHours() : date.getHours();
	const minutes = useUtc ? date.getUTCMinutes() : date.getMinutes();
	const seconds = useUtc ? date.getUTCSeconds() : date.getSeconds();
	const ms = useUtc ? date.getUTCMilliseconds() : date.getMilliseconds();
	return hours || minutes || seconds || ms ? "datetime" : "date";
}

/**
 * Attempt to format a cell value using the cell's number format or XF record.
 * Falls back to a plain string coercion if all formatting attempts fail.
 */
function safeFormatCell(cell: CellObject, value: any): string {
	const isDateCell = cell.t === "d" && value instanceof Date;
	// First try the explicit format string stored on the cell (cell.z)
	if (cell.z != null) {
		try {
			cell.w = formatNumber(cell.z, isDateCell ? dateToSerialNumber(value) : value);
			return cell.w;
		} catch {}
	}
	// Fall back to the numFmtId from the cell's XF style record;
	// default to format 14 (short date) for date cells, or 0 (General) otherwise
	try {
		cell.w = formatNumber(
			(cell.XF || {}).numFmtId || (isDateCell ? 14 : 0),
			isDateCell ? dateToSerialNumber(value) : value,
		);
		return cell.w;
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

/**
 * Format a cell for worksheet export APIs.
 *
 * This mirrors `formatCell` by default. With `dateOutput: "iso"`, date and
 * datetime number formats are rendered as stable ISO-like strings, while
 * time-only formats keep their display value instead of becoming epoch dates.
 */
export function formatCellForOutput(cell: CellObject, value?: any, options?: any): string {
	if (options?.dateOutput !== "iso") {
		return formatCell(cell, value, options);
	}
	const cellValue = value == null ? cell.v : value;
	const kind = getCellDateTimeFormatKind(cell, options);
	if (cell.t === "n" && typeof cellValue === "number") {
		if (kind === "date" || kind === "datetime") {
			return formatDateIso(serialNumberToDate(cellValue, options?.date1904), kind, options);
		}
		return formatCell(cell, value, options);
	}
	if (cellValue instanceof Date) {
		return formatDateIso(
			cellValue,
			kind === "none" || kind === "time" ? inferDateKind(cellValue, options) : kind,
			options,
		);
	}
	return formatCell(cell, value, options);
}
