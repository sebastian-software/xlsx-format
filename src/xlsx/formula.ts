import { decode_row, encode_row, decode_col, encode_col } from "../utils/cell.js";
import { decode_range, decode_cell } from "../utils/cell.js";
import type { CellAddress } from "../types.js";

const rcregex = /(^|[^A-Za-z_])R(\[?-?\d+\]|[1-9]\d*|)C(\[?-?\d+\]|[1-9]\d*|)(?![A-Za-z0-9_])/g;

/** Convert R1C1-style formula to A1-style */
export function rc_to_a1(fstr: string, base: CellAddress): string {
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

		return $1 + (cRel ? "" : "$") + encode_col(cFinal) + (rRel ? "" : "$") + encode_row(rFinal);
	});
}

const crefregex =
	/(^|[^._A-Za-z0-9])([$]?)([A-Z]{1,2}|[A-W][A-Z]{2}|X[A-E][A-Z]|XF[A-D])([$]?)(10[0-3]\d{4}|104[0-7]\d{3}|1048[0-4]\d{2}|10485[0-6]\d|104857[0-6]|[1-9]\d{0,5})(?![_.(A-Za-z0-9])/g;

/** Convert A1-style formula to R1C1-style */
export function a1_to_rc(fstr: string, base: CellAddress): string {
	return fstr.replace(crefregex, ($0, $1, $2, $3, $4, $5) => {
		const c = decode_col($3) - ($2 ? 0 : base.c);
		const r = decode_row($5) - ($4 ? 0 : base.r);
		const R = $4 === "$" ? r + 1 : r === 0 ? "" : "[" + r + "]";
		const C = $2 === "$" ? c + 1 : c === 0 ? "" : "[" + c + "]";
		return $1 + "R" + R + "C" + C;
	});
}

/** Shift cell references in a formula string by delta */
export function shift_formula_str(f: string, delta: CellAddress): string {
	return f.replace(crefregex, ($0, $1, $2, $3, $4, $5) => {
		return (
			$1 +
			($2 === "$" ? $2 + $3 : encode_col(decode_col($3) + delta.c)) +
			($4 === "$" ? $4 + $5 : encode_row(decode_row($5) + delta.r))
		);
	});
}

/** Shift formula for shared formulas in XLSX */
export function shift_formula_xlsx(f: string, range: string, cell: string): string {
	const r = decode_range(range);
	const s = r.s;
	const c = decode_cell(cell);
	const delta = { r: c.r - s.r, c: c.c - s.c };
	return shift_formula_str(f, delta);
}

/** Heuristic: is this string a formula? */
export function fuzzyfmla(f: string): boolean {
	if (f.length === 1) {
		return false;
	}
	return true;
}

/** Strip _xlfn. prefix from function names */
export function _xlfn(f: string): string {
	return f.replace(/_xlfn\./g, "");
}
