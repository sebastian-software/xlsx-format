import { decodeRow, encodeRow, decodeCol, encodeCol } from "../utils/cell.js";
import { decodeRange, decodeCell } from "../utils/cell.js";
import type { CellAddress } from "../types.js";

/**
 * Regex to match R1C1-style cell references in formulas.
 *
 * Matches patterns like R1C1 (absolute), R[1]C[1] (relative), RC (current cell).
 * The negative lookbehind prevents matching inside identifiers. The three capture
 * groups are: $1 = prefix char, $2 = row part, $3 = column part.
 */
const rcregex = /(^|[^A-Za-z_])R(\[?-?\d+\]|[1-9]\d*|)C(\[?-?\d+\]|[1-9]\d*|)(?![A-Za-z0-9_])/g;

/**
 * Convert R1C1-style formula references to A1-style.
 *
 * R1C1 references can be absolute (R1C1) or relative (R[1]C[1]).
 * Relative references use brackets and are offset from the base cell.
 *
 * @param fstr - Formula string containing R1C1 references
 * @param base - Base cell address for resolving relative references
 * @returns Formula string with A1-style references
 */
export function rcToA1(fstr: string, base: CellAddress): string {
	return fstr.replace(rcregex, ($$, $1, $2, $3) => {
		let cRel = false,
			rRel = false;

		// Empty row/col part means relative offset of 0 (e.g. "RC" = current row/col)
		if ($2.length === 0) {
			rRel = true;
		} else if ($2.charAt(0) === "[") {
			rRel = true;
			$2 = $2.slice(1, -1); // Strip brackets
		}

		if ($3.length === 0) {
			cRel = true;
		} else if ($3.charAt(0) === "[") {
			cRel = true;
			$3 = $3.slice(1, -1); // Strip brackets
		}

		const R = $2.length > 0 ? parseInt($2, 10) | 0 : 0;
		const C = $3.length > 0 ? parseInt($3, 10) | 0 : 0;

		// Relative: offset from base; Absolute: R1C1 is 1-based so subtract 1
		const cFinal = cRel ? C + base.c : C - 1;
		const rFinal = rRel ? R + base.r : R - 1;

		// Absolute references get a "$" prefix in A1 notation
		return $1 + (cRel ? "" : "$") + encodeCol(cFinal) + (rRel ? "" : "$") + encodeRow(rFinal);
	});
}

/**
 * Regex to match A1-style cell references in formulas.
 *
 * Matches column letters (A-XFD), optional $ anchors, and row numbers (1-1048576).
 * The lookbehind and lookahead prevent matching inside identifiers or function names.
 */
const crefregex =
	/(^|[^._A-Za-z0-9])([$]?)([A-Z]{1,2}|[A-W][A-Z]{2}|X[A-E][A-Z]|XF[A-D])([$]?)(10[0-3]\d{4}|104[0-7]\d{3}|1048[0-4]\d{2}|10485[0-6]\d|104857[0-6]|[1-9]\d{0,5})(?![_.(A-Za-z0-9])/g;

/**
 * Convert A1-style formula references to R1C1-style.
 *
 * @param fstr - Formula string containing A1 references
 * @param base - Base cell address for computing relative offsets
 * @returns Formula string with R1C1-style references
 */
export function a1ToRc(fstr: string, base: CellAddress): string {
	return fstr.replace(crefregex, ($0, $1, $2, $3, $4, $5) => {
		const c = decodeCol($3) - ($2 ? 0 : base.c);
		const r = decodeRow($5) - ($4 ? 0 : base.r);
		// Absolute ($): use 1-based index; Relative: use [offset] or empty for 0
		const R = $4 === "$" ? r + 1 : r === 0 ? "" : "[" + r + "]";
		const C = $2 === "$" ? c + 1 : c === 0 ? "" : "[" + c + "]";
		return $1 + "R" + R + "C" + C;
	});
}

/**
 * Shift all cell references in a formula string by a given row/column delta.
 *
 * Absolute references (prefixed with $) are left unchanged.
 *
 * @param f - Formula string to transform
 * @param delta - Row and column offsets to apply
 * @returns Formula string with shifted references
 */
export function shiftFormulaStr(f: string, delta: CellAddress): string {
	return f.replace(crefregex, ($0, $1, $2, $3, $4, $5) => {
		return (
			$1 +
			($2 === "$" ? $2 + $3 : encodeCol(decodeCol($3) + delta.c)) +
			($4 === "$" ? $4 + $5 : encodeRow(decodeRow($5) + delta.r))
		);
	});
}

/**
 * Shift a shared formula from its master cell to a target cell.
 *
 * Computes the delta between the shared formula's origin (start of range)
 * and the target cell, then applies that shift to all references.
 *
 * @param f - The shared formula string (from the master cell)
 * @param range - The shared formula range (e.g. "A1:C5")
 * @param cell - The target cell address (e.g. "B3")
 * @returns Formula string adjusted for the target cell
 */
export function shiftFormulaXlsx(f: string, range: string, cell: string): string {
	const r = decodeRange(range);
	const s = r.s;
	const c = decodeCell(cell);
	const delta = { r: c.r - s.r, c: c.c - s.c };
	return shiftFormulaStr(f, delta);
}

/**
 * Heuristic check: is this string likely a formula?
 *
 * Single-character strings are not treated as formulas (e.g. "=" alone).
 *
 * @param f - String to test
 * @returns true if the string looks like it could be a formula
 */
export function isFuzzyFormula(f: string): boolean {
	if (f.length === 1) {
		return false;
	}
	return true;
}

/**
 * Strip the _xlfn. prefix from Excel function names.
 *
 * Excel uses _xlfn. to prefix newer function names for backward compatibility
 * (e.g. _xlfn.CONCAT becomes CONCAT).
 *
 * @param f - Formula string potentially containing _xlfn. prefixes
 * @returns Formula string with prefixes removed
 */
export function stripXlFunctionPrefix(f: string): string {
	return f.replace(/_xlfn\./g, "");
}
