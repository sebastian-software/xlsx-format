import type { CellAddress, CellObject, Range, WorkSheet } from "../types.js";

/**
 * Decode a row string (1-based) to a zero-based row index.
 * @param rowstr - Row string, possibly with a "$" absolute marker (e.g. "5" or "$5")
 * @returns Zero-based row index
 */
export function decodeRow(rowstr: string): number {
	return parseInt(removeRowAbsolute(rowstr), 10) - 1;
}

/**
 * Encode a zero-based row index to a 1-based row string.
 * @param row - Zero-based row index
 * @returns 1-based row string (e.g. "1" for row index 0)
 */
export function encodeRow(row: number): string {
	return "" + (row + 1);
}

/**
 * Make the row portion of a cell reference absolute by prefixing "$".
 * @param cstr - Cell reference string (e.g. "A5")
 * @returns Cell reference with absolute row (e.g. "A$5")
 */
export function makeRowAbsolute(cstr: string): string {
	return cstr.replace(/([A-Z]|^)(\d+)$/, "$1$$$2");
}

/**
 * Remove the "$" absolute marker from the row portion of a cell reference.
 * @param cstr - Cell reference string (e.g. "A$5")
 * @returns Cell reference with relative row (e.g. "A5")
 */
export function removeRowAbsolute(cstr: string): string {
	return cstr.replace(/\$(\d+)$/, "$1");
}

/**
 * Decode a column label (e.g. "A", "AA") to a zero-based column index.
 *
 * Treats column letters as a base-26 number where A=1, B=2, ..., Z=26.
 *
 * @param colstr - Column label string, possibly with "$" prefix
 * @returns Zero-based column index (A=0, B=1, ..., Z=25, AA=26, ...)
 */
export function decodeCol(colstr: string): number {
	const c = removeColAbsolute(colstr);
	let d = 0;
	for (let i = 0; i < c.length; ++i) {
		// 'A' is charCode 65; subtract 64 so A=1, B=2, ..., Z=26
		d = 26 * d + c.charCodeAt(i) - 64;
	}
	return d - 1;
}

/**
 * Encode a zero-based column index to an Excel column label (A, B, ..., Z, AA, AB, ...).
 *
 * Uses bijective base-26 numeration: col 0 = "A", col 25 = "Z", col 26 = "AA".
 *
 * @param col - Zero-based column index
 * @returns Column label string
 * @throws Error if col is negative
 */
export function encodeCol(col: number): string {
	if (col < 0) {
		throw new Error("invalid column " + col);
	}
	let result = "";
	// Convert to 1-based, then repeatedly extract bijective base-26 digits
	for (++col; col; col = Math.floor((col - 1) / 26)) {
		// 65 = 'A'; ((col - 1) % 26) gives 0-25 for A-Z
		result = String.fromCharCode(((col - 1) % 26) + 65) + result;
	}
	return result;
}

/**
 * Make the column portion of a cell reference absolute by prefixing "$".
 * @param cstr - Cell reference string (e.g. "A5")
 * @returns Cell reference with absolute column (e.g. "$A5")
 */
export function makeColAbsolute(cstr: string): string {
	return cstr.replace(/^([A-Z])/, "$$$1");
}

/**
 * Remove the "$" absolute marker from the column portion of a cell reference.
 * @param cstr - Cell reference string (e.g. "$A5")
 * @returns Cell reference with relative column (e.g. "A5")
 */
export function removeColAbsolute(cstr: string): string {
	return cstr.replace(/^\$([A-Z])/, "$1");
}

/**
 * Split a cell reference into its column and row parts.
 * @param cstr - Cell reference string (e.g. "A1", "$B$2")
 * @returns Two-element array: [column part, row part]
 */
export function splitCellReference(cstr: string): string[] {
	return cstr.replace(/(\$?[A-Z]*)(\$?\d*)/, "$1,$2").split(",");
}

/**
 * Decode an A1-style cell reference to a numeric {c, r} address (zero-based).
 *
 * Hand-optimized parser that processes characters by charCode for performance:
 * digits (48-57) accumulate into the row, uppercase letters (65-90) into the column.
 *
 * @param cstr - Cell reference string (e.g. "A1", "AB12")
 * @returns Zero-based cell address {c: column, r: row}
 */
export function decodeCell(cstr: string): CellAddress {
	let R = 0,
		C = 0;
	for (let i = 0; i < cstr.length; ++i) {
		const charCode = cstr.charCodeAt(i);
		if (charCode >= 48 && charCode <= 57) {
			// '0'-'9': accumulate row number
			R = 10 * R + (charCode - 48);
		} else if (charCode >= 65 && charCode <= 90) {
			// 'A'-'Z': accumulate column in base-26 (A=1)
			C = 26 * C + (charCode - 64);
		}
	}
	// Convert from 1-based to 0-based
	return { c: C - 1, r: R - 1 };
}

/**
 * Encode a zero-based {c, r} cell address to an A1-style reference string.
 * @param cell - Zero-based cell address
 * @returns A1-style cell reference (e.g. "A1" for {c:0, r:0})
 */
export function encodeCell(cell: CellAddress): string {
	let col = cell.c + 1;
	let result = "";
	// Bijective base-26 conversion; `| 0` is a fast Math.floor for positive numbers
	for (; col; col = ((col - 1) / 26) | 0) {
		result = String.fromCharCode(((col - 1) % 26) + 65) + result;
	}
	return result + (cell.r + 1);
}

/**
 * Decode a range string (e.g. "A1:B2") to a Range object with start and end addresses.
 *
 * If no colon is present, the range is a single cell (start equals end).
 *
 * @param range - Range string in A1 notation
 * @returns Range object with start (s) and end (e) addresses
 */
export function decodeRange(range: string): Range {
	const idx = range.indexOf(":");
	if (idx === -1) {
		return { s: decodeCell(range), e: decodeCell(range) };
	}
	return { s: decodeCell(range.slice(0, idx)), e: decodeCell(range.slice(idx + 1)) };
}

/**
 * Encode a Range or pair of CellAddresses to an A1:B2 range string.
 *
 * Can be called as:
 * - `encodeRange(range)` with a Range object
 * - `encodeRange(start, end)` with two CellAddress objects
 *
 * If start and end are the same cell, returns a single cell reference (no colon).
 *
 * @param cs - A Range object, or the start CellAddress
 * @param ce - Optional end CellAddress (when cs is a CellAddress)
 * @returns Range string in A1 notation (e.g. "A1:B2" or "A1")
 */
export function encodeRange(cs: CellAddress | Range, ce?: CellAddress): string {
	if (typeof ce === "undefined" || typeof ce === "number") {
		return encodeRange((cs as Range).s, (cs as Range).e);
	}
	const s = typeof cs === "string" ? cs : encodeCell(cs as CellAddress);
	const e = typeof ce === "string" ? ce : encodeCell(ce);
	return s === e ? s : s + ":" + e;
}

/**
 * Performance-optimized range decoder that parses directly by charCode.
 *
 * Unlike {@link decodeRange}, this avoids creating intermediate strings/objects.
 * Used on hot paths where many ranges must be parsed quickly.
 *
 * @param range - Range string in A1 notation (e.g. "A1:B2")
 * @returns Range object with start (s) and end (e) addresses (zero-based)
 */
export function safeDecodeRange(range: string): Range {
	const result = { s: { c: 0, r: 0 }, e: { c: 0, r: 0 } };
	let idx = 0,
		i = 0,
		charCode = 0;
	const len = range.length;

	// Parse start column letters (A-Z are charCodes 65-90, subtract 64 for base-26)
	for (idx = 0; i < len; ++i) {
		if ((charCode = range.charCodeAt(i) - 64) < 1 || charCode > 26) {
			break;
		}
		idx = 26 * idx + charCode;
	}
	result.s.c = --idx;

	// Parse start row digits (0-9 are charCodes 48-57, subtract 48)
	for (idx = 0; i < len; ++i) {
		if ((charCode = range.charCodeAt(i) - 48) < 0 || charCode > 9) {
			break;
		}
		idx = 10 * idx + charCode;
	}
	result.s.r = --idx;

	// charCode 10 corresponds to ':' - 48 = 10 (i.e. ':' has charCode 58)
	if (i === len || charCode !== 10) {
		result.e.c = result.s.c;
		result.e.r = result.s.r;
		return result;
	}
	++i;

	// Parse end column letters
	for (idx = 0; i !== len; ++i) {
		if ((charCode = range.charCodeAt(i) - 64) < 1 || charCode > 26) {
			break;
		}
		idx = 26 * idx + charCode;
	}
	result.e.c = --idx;

	// Parse end row digits
	for (idx = 0; i !== len; ++i) {
		if ((charCode = range.charCodeAt(i) - 48) < 0 || charCode > 9) {
			break;
		}
		idx = 10 * idx + charCode;
	}
	result.e.r = --idx;
	return result;
}

/** Retrieve a cell from a worksheet, handling both dense (array) and sparse (object) storage */
export function getCell(sheet: WorkSheet, row: number, col: number): CellObject | undefined {
	const data = (sheet as any)["!data"];
	if (data != null) {
		return (data[row] || [])[col];
	}
	return (sheet as any)[encodeCol(col) + encodeRow(row)];
}

/**
 * Quote a sheet name for safe use in formulas (e.g. "'Sheet 1'!A1").
 *
 * Wraps the name in single quotes if it contains characters outside
 * word characters, CJK unified ideographs, or Japanese Hiragana/Katakana.
 * Single quotes within the name are escaped by doubling ('').
 *
 * @param sname - Sheet name to quote
 * @returns Quoted sheet name safe for formula references
 * @throws Error if the sheet name is empty
 */
export function quoteSheetName(sname: string): string {
	if (!sname) {
		throw new Error("empty sheet name");
	}
	// Match any character not in: word chars, CJK Unified Ideographs (4E00-9FFF), Hiragana/Katakana (3040-30FF)
	if (/[^\w\u4E00-\u9FFF\u3040-\u30FF]/.test(sname)) {
		return "'" + sname.replace(/'/g, "''") + "'";
	}
	return sname;
}
