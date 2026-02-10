import type { CellAddress, Range } from "../types.js";

export function decodeRow(rowstr: string): number {
	return parseInt(removeRowAbsolute(rowstr), 10) - 1;
}

export function encodeRow(row: number): string {
	return "" + (row + 1);
}

export function makeRowAbsolute(cstr: string): string {
	return cstr.replace(/([A-Z]|^)(\d+)$/, "$1$$$2");
}

export function removeRowAbsolute(cstr: string): string {
	return cstr.replace(/\$(\d+)$/, "$1");
}

export function decodeCol(colstr: string): number {
	const c = removeColAbsolute(colstr);
	let d = 0;
	for (let i = 0; i < c.length; ++i) {
		d = 26 * d + c.charCodeAt(i) - 64;
	}
	return d - 1;
}

export function encodeCol(col: number): string {
	if (col < 0) {
		throw new Error("invalid column " + col);
	}
	let result = "";
	for (++col; col; col = Math.floor((col - 1) / 26)) {
		result = String.fromCharCode(((col - 1) % 26) + 65) + result;
	}
	return result;
}

export function makeColAbsolute(cstr: string): string {
	return cstr.replace(/^([A-Z])/, "$$$1");
}

export function removeColAbsolute(cstr: string): string {
	return cstr.replace(/^\$([A-Z])/, "$1");
}

export function splitCellReference(cstr: string): string[] {
	return cstr.replace(/(\$?[A-Z]*)(\$?\d*)/, "$1,$2").split(",");
}

export function decodeCell(cstr: string): CellAddress {
	let R = 0,
		C = 0;
	for (let i = 0; i < cstr.length; ++i) {
		const charCode = cstr.charCodeAt(i);
		if (charCode >= 48 && charCode <= 57) {
			R = 10 * R + (charCode - 48);
		} else if (charCode >= 65 && charCode <= 90) {
			C = 26 * C + (charCode - 64);
		}
	}
	return { c: C - 1, r: R - 1 };
}

export function encodeCell(cell: CellAddress): string {
	let col = cell.c + 1;
	let result = "";
	for (; col; col = ((col - 1) / 26) | 0) {
		result = String.fromCharCode(((col - 1) % 26) + 65) + result;
	}
	return result + (cell.r + 1);
}

export function decodeRange(range: string): Range {
	const idx = range.indexOf(":");
	if (idx === -1) {
		return { s: decodeCell(range), e: decodeCell(range) };
	}
	return { s: decodeCell(range.slice(0, idx)), e: decodeCell(range.slice(idx + 1)) };
}

export function encodeRange(cs: CellAddress | Range, ce?: CellAddress): string {
	if (typeof ce === "undefined" || typeof ce === "number") {
		return encodeRange((cs as Range).s, (cs as Range).e);
	}
	const s = typeof cs === "string" ? cs : encodeCell(cs as CellAddress);
	const e = typeof ce === "string" ? ce : encodeCell(ce);
	return s === e ? s : s + ":" + e;
}

export function safeDecodeRange(range: string): Range {
	const result = { s: { c: 0, r: 0 }, e: { c: 0, r: 0 } };
	let idx = 0,
		i = 0,
		charCode = 0;
	const len = range.length;
	for (idx = 0; i < len; ++i) {
		if ((charCode = range.charCodeAt(i) - 64) < 1 || charCode > 26) {
			break;
		}
		idx = 26 * idx + charCode;
	}
	result.s.c = --idx;

	for (idx = 0; i < len; ++i) {
		if ((charCode = range.charCodeAt(i) - 48) < 0 || charCode > 9) {
			break;
		}
		idx = 10 * idx + charCode;
	}
	result.s.r = --idx;

	if (i === len || charCode !== 10) {
		result.e.c = result.s.c;
		result.e.r = result.s.r;
		return result;
	}
	++i;

	for (idx = 0; i !== len; ++i) {
		if ((charCode = range.charCodeAt(i) - 64) < 1 || charCode > 26) {
			break;
		}
		idx = 26 * idx + charCode;
	}
	result.e.c = --idx;

	for (idx = 0; i !== len; ++i) {
		if ((charCode = range.charCodeAt(i) - 48) < 0 || charCode > 9) {
			break;
		}
		idx = 10 * idx + charCode;
	}
	result.e.r = --idx;
	return result;
}

export function quoteSheetName(sname: string): string {
	if (!sname) {
		throw new Error("empty sheet name");
	}
	if (/[^\w\u4E00-\u9FFF\u3040-\u30FF]/.test(sname)) {
		return "'" + sname.replace(/'/g, "''") + "'";
	}
	return sname;
}
