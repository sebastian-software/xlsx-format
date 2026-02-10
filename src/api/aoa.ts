import type { WorkSheet, AOA2SheetOpts, Range } from "../types.js";
import { decode_cell, encode_col, encode_range, safe_decode_range } from "../utils/cell.js";
import { datenum, local_to_utc } from "../utils/date.js";
import { SSF_format } from "../ssf/format.js";
import { table_fmt } from "../ssf/table.js";

/** Add an array of arrays to an existing (or new) worksheet */
export function sheet_add_aoa(_ws: WorkSheet | null, data: any[][], opts?: AOA2SheetOpts): WorkSheet {
	const o = opts || ({} as any);
	const dense = _ws ? (_ws as any)["!data"] != null : !!o.dense;
	const ws: any = _ws || (dense ? { "!data": [] } : {});
	if (dense && !ws["!data"]) {
		ws["!data"] = [];
	}

	let _R = 0,
		_C = 0;
	if (ws && o.origin != null) {
		if (typeof o.origin === "number") {
			_R = o.origin;
		} else {
			const _origin = typeof o.origin === "string" ? decode_cell(o.origin) : o.origin;
			_R = _origin.r;
			_C = _origin.c;
		}
	}

	const range: Range = { s: { c: 10000000, r: 10000000 }, e: { c: 0, r: 0 } };
	if (ws["!ref"]) {
		const _range = safe_decode_range(ws["!ref"]);
		range.s.c = _range.s.c;
		range.s.r = _range.s.r;
		range.e.c = Math.max(range.e.c, _range.e.c);
		range.e.r = Math.max(range.e.r, _range.e.r);
		if (_R === -1) {
			range.e.r = _R = ws["!ref"] ? _range.e.r + 1 : 0;
		}
	} else {
		range.s.c = range.e.c = range.s.r = range.e.r = 0;
	}

	let row: any[] = [];
	let seen = false;
	for (let R = 0; R < data.length; ++R) {
		if (!data[R]) {
			continue;
		}
		if (!Array.isArray(data[R])) {
			throw new Error("aoa_to_sheet expects an array of arrays");
		}
		const __R = _R + R;
		if (dense) {
			if (!ws["!data"][__R]) {
				ws["!data"][__R] = [];
			}
			row = ws["!data"][__R];
		}
		const data_R = data[R];
		for (let C = 0; C < data_R.length; ++C) {
			if (typeof data_R[C] === "undefined") {
				continue;
			}
			let cell: any = { v: data_R[C], t: "" };
			const __C = _C + C;
			if (range.s.r > __R) {
				range.s.r = __R;
			}
			if (range.s.c > __C) {
				range.s.c = __C;
			}
			if (range.e.r < __R) {
				range.e.r = __R;
			}
			if (range.e.c < __C) {
				range.e.c = __C;
			}
			seen = true;

			if (
				data_R[C] &&
				typeof data_R[C] === "object" &&
				!Array.isArray(data_R[C]) &&
				!(data_R[C] instanceof Date)
			) {
				cell = data_R[C];
			} else {
				if (Array.isArray(cell.v)) {
					cell.f = data_R[C][1];
					cell.v = cell.v[0];
				}
				if (cell.v === null) {
					if (cell.f) {
						cell.t = "n";
					} else if (o.nullError) {
						cell.t = "e";
						cell.v = 0;
					} else if (!o.sheetStubs) {
						continue;
					} else {
						cell.t = "z";
					}
				} else if (typeof cell.v === "number") {
					if (isFinite(cell.v)) {
						cell.t = "n";
					} else if (isNaN(cell.v)) {
						cell.t = "e";
						cell.v = 0x0f;
					} else {
						cell.t = "e";
						cell.v = 0x07;
					}
				} else if (typeof cell.v === "boolean") {
					cell.t = "b";
				} else if (cell.v instanceof Date) {
					cell.z = o.dateNF || table_fmt[14];
					if (!o.UTC) {
						cell.v = local_to_utc(cell.v);
					}
					if (o.cellDates) {
						cell.t = "d";
						cell.w = SSF_format(cell.z, datenum(cell.v, o.date1904));
					} else {
						cell.t = "n";
						cell.v = datenum(cell.v, o.date1904);
						cell.w = SSF_format(cell.z, cell.v);
					}
				} else {
					cell.t = "s";
				}
			}

			if (dense) {
				if (row[__C] && row[__C].z) {
					cell.z = row[__C].z;
				}
				row[__C] = cell;
			} else {
				const cell_ref = encode_col(__C) + (__R + 1);
				if (ws[cell_ref] && ws[cell_ref].z) {
					cell.z = ws[cell_ref].z;
				}
				ws[cell_ref] = cell;
			}
		}
	}
	if (seen && range.s.c < 10400000) {
		ws["!ref"] = encode_range(range);
	}
	return ws as WorkSheet;
}

/** Create a new worksheet from an array of arrays */
export function aoa_to_sheet(data: any[][], opts?: AOA2SheetOpts): WorkSheet {
	return sheet_add_aoa(null, data, opts);
}
