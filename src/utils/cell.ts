import type { CellAddress, Range } from "../types.js";

export function decode_row(rowstr: string): number {
	return parseInt(unfix_row(rowstr), 10) - 1;
}

export function encode_row(row: number): string {
	return "" + (row + 1);
}

export function fix_row(cstr: string): string {
	return cstr.replace(/([A-Z]|^)(\d+)$/, "$1$$$2");
}

export function unfix_row(cstr: string): string {
	return cstr.replace(/\$(\d+)$/, "$1");
}

export function decode_col(colstr: string): number {
	const c = unfix_col(colstr);
	let d = 0;
	for (let i = 0; i < c.length; ++i) {
		d = 26 * d + c.charCodeAt(i) - 64;
	}
	return d - 1;
}

export function encode_col(col: number): string {
	if (col < 0) {
		throw new Error("invalid column " + col);
	}
	let s = "";
	for (++col; col; col = Math.floor((col - 1) / 26)) {
		s = String.fromCharCode(((col - 1) % 26) + 65) + s;
	}
	return s;
}

export function fix_col(cstr: string): string {
	return cstr.replace(/^([A-Z])/, "$$$1");
}

export function unfix_col(cstr: string): string {
	return cstr.replace(/^\$([A-Z])/, "$1");
}

export function split_cell(cstr: string): string[] {
	return cstr.replace(/(\$?[A-Z]*)(\$?\d*)/, "$1,$2").split(",");
}

export function decode_cell(cstr: string): CellAddress {
	let R = 0,
		C = 0;
	for (let i = 0; i < cstr.length; ++i) {
		const cc = cstr.charCodeAt(i);
		if (cc >= 48 && cc <= 57) {
			R = 10 * R + (cc - 48);
		} else if (cc >= 65 && cc <= 90) {
			C = 26 * C + (cc - 64);
		}
	}
	return { c: C - 1, r: R - 1 };
}

export function encode_cell(cell: CellAddress): string {
	let col = cell.c + 1;
	let s = "";
	for (; col; col = ((col - 1) / 26) | 0) {
		s = String.fromCharCode(((col - 1) % 26) + 65) + s;
	}
	return s + (cell.r + 1);
}

export function decode_range(range: string): Range {
	const idx = range.indexOf(":");
	if (idx === -1) {
		return { s: decode_cell(range), e: decode_cell(range) };
	}
	return { s: decode_cell(range.slice(0, idx)), e: decode_cell(range.slice(idx + 1)) };
}

export function encode_range(cs: CellAddress | Range, ce?: CellAddress): string {
	if (typeof ce === "undefined" || typeof ce === "number") {
		return encode_range((cs as Range).s, (cs as Range).e);
	}
	const s = typeof cs === "string" ? cs : encode_cell(cs as CellAddress);
	const e = typeof ce === "string" ? ce : encode_cell(ce);
	return s === e ? s : s + ":" + e;
}

export function safe_decode_range(range: string): Range {
	const o = { s: { c: 0, r: 0 }, e: { c: 0, r: 0 } };
	let idx = 0,
		i = 0,
		cc = 0;
	const len = range.length;
	for (idx = 0; i < len; ++i) {
		if ((cc = range.charCodeAt(i) - 64) < 1 || cc > 26) {
			break;
		}
		idx = 26 * idx + cc;
	}
	o.s.c = --idx;

	for (idx = 0; i < len; ++i) {
		if ((cc = range.charCodeAt(i) - 48) < 0 || cc > 9) {
			break;
		}
		idx = 10 * idx + cc;
	}
	o.s.r = --idx;

	if (i === len || cc !== 10) {
		o.e.c = o.s.c;
		o.e.r = o.s.r;
		return o;
	}
	++i;

	for (idx = 0; i !== len; ++i) {
		if ((cc = range.charCodeAt(i) - 64) < 1 || cc > 26) {
			break;
		}
		idx = 26 * idx + cc;
	}
	o.e.c = --idx;

	for (idx = 0; i !== len; ++i) {
		if ((cc = range.charCodeAt(i) - 48) < 0 || cc > 9) {
			break;
		}
		idx = 10 * idx + cc;
	}
	o.e.r = --idx;
	return o;
}

export function formula_quote_sheet_name(sname: string): string {
	if (!sname) {
		throw new Error("empty sheet name");
	}
	if (/[^\w\u4E00-\u9FFF\u3040-\u30FF]/.test(sname)) {
		return "'" + sname.replace(/'/g, "''") + "'";
	}
	return sname;
}
