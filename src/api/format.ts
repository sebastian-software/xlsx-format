import type { CellObject } from "../types.js";
import { BErr } from "../types.js";
import { formatNumber } from "../ssf/format.js";
import { dateToSerialNumber } from "../utils/date.js";

function safeFormatCell(cell: CellObject, value: any): string {
	const isDateCell = cell.t === "d" && value instanceof Date;
	if (cell.z != null) {
		try {
			return (cell.w = formatNumber(cell.z, isDateCell ? dateToSerialNumber(value) : value));
		} catch {}
	}
	try {
		return (cell.w = formatNumber((cell.XF || {}).numFmtId || (isDateCell ? 14 : 0), isDateCell ? dateToSerialNumber(value) : value));
	} catch {
		return "" + value;
	}
}

export function formatCell(cell: CellObject, value?: any, options?: any): string {
	if (cell == null || cell.t == null || cell.t === "z") {
		return "";
	}
	if (cell.w !== undefined) {
		return cell.w;
	}
	if (cell.t === "d" && !cell.z && options && options.dateNF) {
		cell.z = options.dateNF;
	}
	if (cell.t === "e") {
		return BErr[cell.v as number] || String(cell.v);
	}
	if (value == null) {
		return safeFormatCell(cell, cell.v);
	}
	return safeFormatCell(cell, value);
}
