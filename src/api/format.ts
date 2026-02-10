import type { CellObject } from "../types.js";
import { BErr } from "../types.js";
import { formatNumber } from "../ssf/format.js";
import { dateToSerialNumber } from "../utils/date.js";

function safeFormatCell(cell: CellObject, v: any): string {
	const q = cell.t === "d" && v instanceof Date;
	if (cell.z != null) {
		try {
			return (cell.w = formatNumber(cell.z, q ? dateToSerialNumber(v) : v));
		} catch {}
	}
	try {
		return (cell.w = formatNumber((cell.XF || {}).numFmtId || (q ? 14 : 0), q ? dateToSerialNumber(v) : v));
	} catch {
		return "" + v;
	}
}

export function formatCell(cell: CellObject, v?: any, o?: any): string {
	if (cell == null || cell.t == null || cell.t === "z") {
		return "";
	}
	if (cell.w !== undefined) {
		return cell.w;
	}
	if (cell.t === "d" && !cell.z && o && o.dateNF) {
		cell.z = o.dateNF;
	}
	if (cell.t === "e") {
		return BErr[cell.v as number] || String(cell.v);
	}
	if (v == null) {
		return safeFormatCell(cell, cell.v);
	}
	return safeFormatCell(cell, v);
}
