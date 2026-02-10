import type { CellObject } from "../types.js";
import { BErr } from "../types.js";
import { SSF_format } from "../ssf/format.js";
import { datenum } from "../utils/date.js";

function safe_format_cell(cell: CellObject, v: any): string {
	const q = cell.t === "d" && v instanceof Date;
	if (cell.z != null) {
		try {
			return (cell.w = SSF_format(cell.z, q ? datenum(v) : v));
		} catch {}
	}
	try {
		return (cell.w = SSF_format((cell.XF || {}).numFmtId || (q ? 14 : 0), q ? datenum(v) : v));
	} catch {
		return "" + v;
	}
}

export function format_cell(cell: CellObject, v?: any, o?: any): string {
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
		return safe_format_cell(cell, cell.v);
	}
	return safe_format_cell(cell, v);
}
