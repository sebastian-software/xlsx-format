import { decodeRow, encodeRow, decodeCol, encodeCol } from "../utils/cell.js";
import { decodeRange, decodeCell } from "../utils/cell.js";
import type { CellAddress } from "../types.js";

const rcregex = /(^|[^A-Za-z_])R(\[?-?\d+\]|[1-9]\d*|)C(\[?-?\d+\]|[1-9]\d*|)(?![A-Za-z0-9_])/g;

/** Convert R1C1-style formula to A1-style */
export function rcToA1(fstr: string, base: CellAddress): string {
	return fstr.replace(rcregex, ($$, $1, $2, $3) => {
		let cRel = false,
			rRel = false;

		if ($2.length === 0) {
			rRel = true;
		} else if ($2.charAt(0) === "[") {
			rRel = true;
			$2 = $2.slice(1, -1);
		}

		if ($3.length === 0) {
			cRel = true;
		} else if ($3.charAt(0) === "[") {
			cRel = true;
			$3 = $3.slice(1, -1);
		}

		const R = $2.length > 0 ? parseInt($2, 10) | 0 : 0;
		const C = $3.length > 0 ? parseInt($3, 10) | 0 : 0;

		const cFinal = cRel ? C + base.c : C - 1;
		const rFinal = rRel ? R + base.r : R - 1;

		return $1 + (cRel ? "" : "$") + encodeCol(cFinal) + (rRel ? "" : "$") + encodeRow(rFinal);
	});
}

const crefregex =
	/(^|[^._A-Za-z0-9])([$]?)([A-Z]{1,2}|[A-W][A-Z]{2}|X[A-E][A-Z]|XF[A-D])([$]?)(10[0-3]\d{4}|104[0-7]\d{3}|1048[0-4]\d{2}|10485[0-6]\d|104857[0-6]|[1-9]\d{0,5})(?![_.(A-Za-z0-9])/g;

/** Convert A1-style formula to R1C1-style */
export function a1ToRc(fstr: string, base: CellAddress): string {
	return fstr.replace(crefregex, ($0, $1, $2, $3, $4, $5) => {
		const c = decodeCol($3) - ($2 ? 0 : base.c);
		const r = decodeRow($5) - ($4 ? 0 : base.r);
		const R = $4 === "$" ? r + 1 : r === 0 ? "" : "[" + r + "]";
		const C = $2 === "$" ? c + 1 : c === 0 ? "" : "[" + c + "]";
		return $1 + "R" + R + "C" + C;
	});
}

/** Shift cell references in a formula string by delta */
export function shiftFormulaStr(f: string, delta: CellAddress): string {
	return f.replace(crefregex, ($0, $1, $2, $3, $4, $5) => {
		return (
			$1 +
			($2 === "$" ? $2 + $3 : encodeCol(decodeCol($3) + delta.c)) +
			($4 === "$" ? $4 + $5 : encodeRow(decodeRow($5) + delta.r))
		);
	});
}

/** Shift formula for shared formulas in XLSX */
export function shiftFormulaXlsx(f: string, range: string, cell: string): string {
	const r = decodeRange(range);
	const s = r.s;
	const c = decodeCell(cell);
	const delta = { r: c.r - s.r, c: c.c - s.c };
	return shiftFormulaStr(f, delta);
}

/** Heuristic: is this string a formula? */
export function isFuzzyFormula(f: string): boolean {
	if (f.length === 1) {
		return false;
	}
	return true;
}

/** Strip _xlfn. prefix from function names */
export function stripXlFunctionPrefix(f: string): string {
	return f.replace(/_xlfn\./g, "");
}
