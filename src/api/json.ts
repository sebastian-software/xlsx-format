import type { WorkSheet, Sheet2JSONOpts, JSON2SheetOpts, CellObject, Range } from "../types.js";
import { decode_cell, encode_col, encode_row, encode_range, safe_decode_range } from "../utils/cell.js";
import { datenum, numdate, utc_to_local, local_to_utc } from "../utils/date.js";
import { fmt_is_date } from "../ssf/format.js";
import { table_fmt } from "../ssf/table.js";
import { keys } from "../utils/helpers.js";
import { format_cell } from "./format.js";

function make_json_row(
	sheet: WorkSheet,
	r: Range,
	R: number,
	cols: string[],
	header: number,
	hdr: any[],
	o: any,
): { row: any; isempty: boolean } {
	const rr = encode_row(R);
	const defval = o.defval;
	const raw = o.raw || !Object.prototype.hasOwnProperty.call(o, "raw");
	let isempty = true;
	const dense = (sheet as any)["!data"] != null;
	const row: any = header === 1 ? [] : {};

	if (header !== 1) {
		try {
			Object.defineProperty(row, "__rowNum__", { value: R, enumerable: false });
		} catch {
			row.__rowNum__ = R;
		}
	}

	if (!dense || (sheet as any)["!data"][R]) {
		for (let C = r.s.c; C <= r.e.c; ++C) {
			const val: CellObject | undefined = dense
				? ((sheet as any)["!data"][R] || [])[C]
				: (sheet as any)[cols[C] + rr];
			if (val == null || val.t === undefined) {
				if (defval === undefined) {
					continue;
				}
				if (hdr[C] != null) {
					row[hdr[C]] = defval;
				}
				continue;
			}
			let v: any = val.v;
			switch (val.t) {
				case "z":
					if (v == null) {
						break;
					}
					continue;
				case "e":
					v = v === 0 ? null : undefined;
					break;
				case "s":
				case "b":
					break;
				case "n":
					if (!val.z || !fmt_is_date(String(val.z))) {
						break;
					}
					v = numdate(v as number);
					if (typeof v === "number") {
						break;
					}
				/* falls through */
				case "d":
					if (!(o && (o.UTC || o.raw === false))) {
						v = utc_to_local(new Date(v));
					}
					break;
				default:
					throw new Error("unrecognized type " + val.t);
			}
			if (hdr[C] != null) {
				if (v == null) {
					if (val.t === "e" && v === null) {
						row[hdr[C]] = null;
					} else if (defval !== undefined) {
						row[hdr[C]] = defval;
					} else if (raw && v === null) {
						row[hdr[C]] = null;
					} else {
						continue;
					}
				} else {
					row[hdr[C]] = (val.t === "n" && typeof o.rawNumbers === "boolean" ? o.rawNumbers : raw)
						? v
						: format_cell(val, v, o);
				}
				if (v != null) {
					isempty = false;
				}
			}
		}
	}
	return { row, isempty };
}

/** Convert a worksheet to an array of JSON objects */
export function sheet_to_json<T = any>(sheet: WorkSheet, opts?: Sheet2JSONOpts): T[] {
	if (sheet == null || sheet["!ref"] == null) {
		return [];
	}
	let header = 0,
		offset = 1;
	const hdr: any[] = [];
	const o: any = opts || {};
	const range = o.range != null ? o.range : sheet["!ref"];

	if (o.header === 1) {
		header = 1;
	} else if (o.header === "A") {
		header = 2;
	} else if (Array.isArray(o.header)) {
		header = 3;
	} else if (o.header == null) {
		header = 0;
	}

	let r: Range;
	switch (typeof range) {
		case "string":
			r = safe_decode_range(range);
			break;
		case "number":
			r = safe_decode_range(sheet["!ref"]);
			r.s.r = range;
			break;
		default:
			r = range;
	}
	if (header > 0) {
		offset = 0;
	}

	const rr = encode_row(r.s.r);
	const cols: string[] = [];
	const out: any[] = [];
	let outi = 0;
	const dense = (sheet as any)["!data"] != null;
	let R = r.s.r;
	const header_cnt: Record<string, number> = {};
	if (dense && !(sheet as any)["!data"][R]) {
		(sheet as any)["!data"][R] = [];
	}
	const colinfo: any[] = (o.skipHidden && sheet["!cols"]) || [];
	const rowinfo: any[] = (o.skipHidden && sheet["!rows"]) || [];

	for (let C = r.s.c; C <= r.e.c; ++C) {
		if ((colinfo[C] || {}).hidden) {
			continue;
		}
		cols[C] = encode_col(C);
		const val: CellObject | undefined = dense ? (sheet as any)["!data"][R][C] : (sheet as any)[cols[C] + rr];
		let v: any, vv: any;
		switch (header) {
			case 1:
				hdr[C] = C - r.s.c;
				break;
			case 2:
				hdr[C] = cols[C];
				break;
			case 3:
				hdr[C] = (o.header as string[])[C - r.s.c];
				break;
			default: {
				const _val = val == null ? { w: "__EMPTY", t: "s" } : val;
				vv = v = format_cell(_val as CellObject, null, o);
				let counter = header_cnt[v] || 0;
				if (!counter) {
					header_cnt[v] = 1;
				} else {
					do {
						vv = v + "_" + counter++;
					} while (header_cnt[vv]);
					header_cnt[v] = counter;
					header_cnt[vv] = 1;
				}
				hdr[C] = vv;
			}
		}
	}

	for (R = r.s.r + offset; R <= r.e.r; ++R) {
		if ((rowinfo[R] || {}).hidden) {
			continue;
		}
		const row = make_json_row(sheet, r, R, cols, header, hdr, o);
		if (!row.isempty || (header === 1 ? o.blankrows !== false : !!o.blankrows)) {
			out[outi++] = row.row;
		}
	}
	out.length = outi;
	return out;
}

/** Add JSON data to a worksheet */
export function sheet_add_json(_ws: WorkSheet | null, js: any[], opts?: JSON2SheetOpts): WorkSheet {
	const o: any = opts || {};
	const dense = _ws ? (_ws as any)["!data"] != null : !!o.dense;
	const offset = +!o.skipHeader;
	const ws: any = _ws || {};
	if (!_ws && dense) {
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

	const range: Range = { s: { c: 0, r: 0 }, e: { c: _C, r: _R + js.length - 1 + offset } };
	if (ws["!ref"]) {
		const _range = safe_decode_range(ws["!ref"]);
		range.e.c = Math.max(range.e.c, _range.e.c);
		range.e.r = Math.max(range.e.r, _range.e.r);
		if (_R === -1) {
			_R = _range.e.r + 1;
			range.e.r = _R + js.length - 1 + offset;
		}
	} else {
		if (_R === -1) {
			_R = 0;
			range.e.r = js.length - 1 + offset;
		}
	}

	const hdr: string[] = o.header || [];
	let C = 0;
	js.forEach((JS, R) => {
		if (dense && !ws["!data"][_R + R + offset]) {
			ws["!data"][_R + R + offset] = [];
		}
		const ROW = dense ? ws["!data"][_R + R + offset] : null;
		keys(JS).forEach((k: string) => {
			if ((C = hdr.indexOf(k)) === -1) {
				hdr[(C = hdr.length)] = k;
			}
			let v = JS[k];
			let t = "z";
			let z = "";
			const ref = dense ? "" : encode_col(_C + C) + encode_row(_R + R + offset);
			const cell: any = dense ? ROW[_C + C] : ws[ref];

			if (v && typeof v === "object" && !(v instanceof Date)) {
				if (dense) {
					ROW[_C + C] = v;
				} else {
					ws[ref] = v;
				}
			} else {
				if (typeof v === "number") {
					t = "n";
				} else if (typeof v === "boolean") {
					t = "b";
				} else if (typeof v === "string") {
					t = "s";
				} else if (v instanceof Date) {
					t = "d";
					if (!o.UTC) {
						v = local_to_utc(v);
					}
					if (!o.cellDates) {
						t = "n";
						v = datenum(v);
					}
					z =
						cell != null && cell.z && fmt_is_date(String(cell.z))
							? String(cell.z)
							: o.dateNF || table_fmt[14];
				} else if (v === null && o.nullError) {
					t = "e";
					v = 0;
				}

				if (!cell) {
					const newCell: any = { t, v };
					if (z) {
						newCell.z = z;
					}
					if (dense) {
						ROW[_C + C] = newCell;
					} else {
						ws[ref] = newCell;
					}
				} else {
					cell.t = t;
					cell.v = v;
					delete cell.w;
					if (z) {
						cell.z = z;
					}
				}
			}
		});
	});

	range.e.c = Math.max(range.e.c, _C + hdr.length - 1);
	const __R = encode_row(_R);
	if (dense && !ws["!data"][_R]) {
		ws["!data"][_R] = [];
	}
	if (offset) {
		for (C = 0; C < hdr.length; ++C) {
			if (dense) {
				ws["!data"][_R][C + _C] = { t: "s", v: hdr[C] };
			} else {
				ws[encode_col(C + _C) + __R] = { t: "s", v: hdr[C] };
			}
		}
	}
	ws["!ref"] = encode_range(range);
	return ws as WorkSheet;
}

/** Create a new worksheet from JSON data */
export function json_to_sheet(js: any[], opts?: JSON2SheetOpts): WorkSheet {
	return sheet_add_json(null, js, opts);
}
