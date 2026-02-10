import type { WorkSheet, Sheet2JSONOpts, JSON2SheetOpts, CellObject, Range } from "../types.js";
import { decodeCell, encodeCol, encodeRow, encodeRange, safeDecodeRange } from "../utils/cell.js";
import { dateToSerialNumber, serialNumberToDate, utcToLocal, localToUtc } from "../utils/date.js";
import { isDateFormat } from "../ssf/format.js";
import { formatTable } from "../ssf/table.js";
import { objectKeys } from "../utils/helpers.js";
import { formatCell } from "./format.js";

function buildJsonRow(
	sheet: WorkSheet,
	r: Range,
	R: number,
	cols: string[],
	header: number,
	headers: any[],
	o: any,
): { row: any; isempty: boolean } {
	const encodedRow = encodeRow(R);
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
				: (sheet as any)[cols[C] + encodedRow];
			if (val == null || val.t === undefined) {
				if (defval === undefined) {
					continue;
				}
				if (headers[C] != null) {
					row[headers[C]] = defval;
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
					if (!val.z || !isDateFormat(String(val.z))) {
						break;
					}
					v = serialNumberToDate(v as number);
					if (typeof v === "number") {
						break;
					}
				/* falls through */
				case "d":
					if (!(o && (o.UTC || o.raw === false))) {
						v = utcToLocal(new Date(v));
					}
					break;
				default:
					throw new Error("unrecognized type " + val.t);
			}
			if (headers[C] != null) {
				if (v == null) {
					if (val.t === "e" && v === null) {
						row[headers[C]] = null;
					} else if (defval !== undefined) {
						row[headers[C]] = defval;
					} else if (raw && v === null) {
						row[headers[C]] = null;
					} else {
						continue;
					}
				} else {
					row[headers[C]] = (val.t === "n" && typeof o.rawNumbers === "boolean" ? o.rawNumbers : raw)
						? v
						: formatCell(val, v, o);
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
export function sheetToJson<T = any>(sheet: WorkSheet, opts?: Sheet2JSONOpts): T[] {
	if (sheet == null || sheet["!ref"] == null) {
		return [];
	}
	let header = 0,
		offset = 1;
	const headers: any[] = [];
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
			r = safeDecodeRange(range);
			break;
		case "number":
			r = safeDecodeRange(sheet["!ref"]);
			r.s.r = range;
			break;
		default:
			r = range;
	}
	if (header > 0) {
		offset = 0;
	}

	const encodedRow = encodeRow(r.s.r);
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
		cols[C] = encodeCol(C);
		const val: CellObject | undefined = dense ? (sheet as any)["!data"][R][C] : (sheet as any)[cols[C] + encodedRow];
		let v: any, vv: any;
		switch (header) {
			case 1:
				headers[C] = C - r.s.c;
				break;
			case 2:
				headers[C] = cols[C];
				break;
			case 3:
				headers[C] = (o.header as string[])[C - r.s.c];
				break;
			default: {
				const _val = val == null ? { w: "__EMPTY", t: "s" } : val;
				vv = v = formatCell(_val as CellObject, null, o);
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
				headers[C] = vv;
			}
		}
	}

	for (R = r.s.r + offset; R <= r.e.r; ++R) {
		if ((rowinfo[R] || {}).hidden) {
			continue;
		}
		const row = buildJsonRow(sheet, r, R, cols, header, headers, o);
		if (!row.isempty || (header === 1 ? o.blankrows !== false : !!o.blankrows)) {
			out[outi++] = row.row;
		}
	}
	out.length = outi;
	return out;
}

/** Add JSON data to a worksheet */
export function addJsonToSheet(_ws: WorkSheet | null, js: any[], opts?: JSON2SheetOpts): WorkSheet {
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
			const _origin = typeof o.origin === "string" ? decodeCell(o.origin) : o.origin;
			_R = _origin.r;
			_C = _origin.c;
		}
	}

	const range: Range = { s: { c: 0, r: 0 }, e: { c: _C, r: _R + js.length - 1 + offset } };
	if (ws["!ref"]) {
		const _range = safeDecodeRange(ws["!ref"]);
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

	const headers: string[] = o.header || [];
	let C = 0;
	js.forEach((JS, R) => {
		if (dense && !ws["!data"][_R + R + offset]) {
			ws["!data"][_R + R + offset] = [];
		}
		const ROW = dense ? ws["!data"][_R + R + offset] : null;
		objectKeys(JS).forEach((k: string) => {
			if ((C = headers.indexOf(k)) === -1) {
				headers[(C = headers.length)] = k;
			}
			let v = JS[k];
			let t = "z";
			let z = "";
			const ref = dense ? "" : encodeCol(_C + C) + encodeRow(_R + R + offset);
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
						v = localToUtc(v);
					}
					if (!o.cellDates) {
						t = "n";
						v = dateToSerialNumber(v);
					}
					z =
						cell != null && cell.z && isDateFormat(String(cell.z))
							? String(cell.z)
							: o.dateNF || formatTable[14];
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

	range.e.c = Math.max(range.e.c, _C + headers.length - 1);
	const __R = encodeRow(_R);
	if (dense && !ws["!data"][_R]) {
		ws["!data"][_R] = [];
	}
	if (offset) {
		for (C = 0; C < headers.length; ++C) {
			if (dense) {
				ws["!data"][_R][C + _C] = { t: "s", v: headers[C] };
			} else {
				ws[encodeCol(C + _C) + __R] = { t: "s", v: headers[C] };
			}
		}
	}
	ws["!ref"] = encodeRange(range);
	return ws as WorkSheet;
}

/** Create a new worksheet from JSON data */
export function jsonToSheet(js: any[], opts?: JSON2SheetOpts): WorkSheet {
	return addJsonToSheet(null, js, opts);
}
