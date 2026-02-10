import type { WorkSheet, Sheet2JSONOpts, JSON2SheetOpts, CellObject, Range } from "../types.js";
import { decodeCell, encodeCol, encodeRow, encodeRange, safeDecodeRange } from "../utils/cell.js";
import { dateToSerialNumber, serialNumberToDate, utcToLocal, localToUtc } from "../utils/date.js";
import { isDateFormat } from "../ssf/format.js";
import { formatTable } from "../ssf/table.js";
import { formatCell } from "./format.js";

function buildJsonRow(
	sheet: WorkSheet,
	range: Range,
	rowIndex: number,
	cols: string[],
	header: number,
	headers: any[],
	options: any,
): { row: any; isempty: boolean } {
	const encodedRow = encodeRow(rowIndex);
	const defval = options.defval;
	const raw = options.raw || !Object.hasOwn(options, "raw");
	let isempty = true;
	const dense = (sheet as any)["!data"] != null;
	const row: any = header === 1 ? [] : {};

	if (header !== 1) {
		try {
			Object.defineProperty(row, "__rowNum__", { value: rowIndex, enumerable: false });
		} catch {
			row.__rowNum__ = rowIndex;
		}
	}

	if (!dense || (sheet as any)["!data"][rowIndex]) {
		for (let colIdx = range.s.c; colIdx <= range.e.c; ++colIdx) {
			const val: CellObject | undefined = dense
				? ((sheet as any)["!data"][rowIndex] || [])[colIdx]
				: (sheet as any)[cols[colIdx] + encodedRow];
			if (val == null || val.t === undefined) {
				if (defval === undefined) {
					continue;
				}
				if (headers[colIdx] != null) {
					row[headers[colIdx]] = defval;
				}
				continue;
			}
			let cellValue: any = val.v;
			switch (val.t) {
				case "z":
					if (cellValue == null) {
						break;
					}
					continue;
				case "e":
					cellValue = cellValue === 0 ? null : undefined;
					break;
				case "s":
				case "b":
					break;
				case "n":
					if (!val.z || !isDateFormat(String(val.z))) {
						break;
					}
					cellValue = serialNumberToDate(cellValue as number);
					if (typeof cellValue === "number") {
						break;
					}
				/* falls through */
				case "d":
					if (!(options && (options.UTC || options.raw === false))) {
						cellValue = utcToLocal(new Date(cellValue));
					}
					break;
				default:
					throw new Error("unrecognized type " + val.t);
			}
			if (headers[colIdx] != null) {
				if (cellValue == null) {
					if (val.t === "e" && cellValue === null) {
						row[headers[colIdx]] = null;
					} else if (defval !== undefined) {
						row[headers[colIdx]] = defval;
					} else if (raw && cellValue === null) {
						row[headers[colIdx]] = null;
					} else {
						continue;
					}
				} else {
					row[headers[colIdx]] = (val.t === "n" && typeof options.rawNumbers === "boolean" ? options.rawNumbers : raw)
						? cellValue
						: formatCell(val, cellValue, options);
				}
				if (cellValue != null) {
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
	const options: any = opts || {};
	const range = options.range != null ? options.range : sheet["!ref"];

	if (options.header === 1) {
		header = 1;
	} else if (options.header === "A") {
		header = 2;
	} else if (Array.isArray(options.header)) {
		header = 3;
	} else if (options.header == null) {
		header = 0;
	}

	let decodedRange: Range;
	switch (typeof range) {
		case "string":
			decodedRange = safeDecodeRange(range);
			break;
		case "number":
			decodedRange = safeDecodeRange(sheet["!ref"]);
			decodedRange.s.r = range;
			break;
		default:
			decodedRange = range;
	}
	if (header > 0) {
		offset = 0;
	}

	const encodedRow = encodeRow(decodedRange.s.r);
	const cols: string[] = [];
	const out: any[] = [];
	let outputIndex = 0;
	const dense = (sheet as any)["!data"] != null;
	let rowIdx = decodedRange.s.r;
	const header_cnt: Record<string, number> = {};
	if (dense && !(sheet as any)["!data"][rowIdx]) {
		(sheet as any)["!data"][rowIdx] = [];
	}
	const colinfo: any[] = (options.skipHidden && sheet["!cols"]) || [];
	const rowinfo: any[] = (options.skipHidden && sheet["!rows"]) || [];

	for (let colIdx = decodedRange.s.c; colIdx <= decodedRange.e.c; ++colIdx) {
		if ((colinfo[colIdx] || {}).hidden) {
			continue;
		}
		cols[colIdx] = encodeCol(colIdx);
		const val: CellObject | undefined = dense ? (sheet as any)["!data"][rowIdx][colIdx] : (sheet as any)[cols[colIdx] + encodedRow];
		let cellValue: any, headerLabel: any;
		switch (header) {
			case 1:
				headers[colIdx] = colIdx - decodedRange.s.c;
				break;
			case 2:
				headers[colIdx] = cols[colIdx];
				break;
			case 3:
				headers[colIdx] = (options.header as string[])[colIdx - decodedRange.s.c];
				break;
			default: {
				const _val = val == null ? { w: "__EMPTY", t: "s" } : val;
				headerLabel = cellValue = formatCell(_val as CellObject, null, options);
				let counter = header_cnt[cellValue] || 0;
				if (!counter) {
					header_cnt[cellValue] = 1;
				} else {
					do {
						headerLabel = cellValue + "_" + counter++;
					} while (header_cnt[headerLabel]);
					header_cnt[cellValue] = counter;
					header_cnt[headerLabel] = 1;
				}
				headers[colIdx] = headerLabel;
			}
		}
	}

	for (rowIdx = decodedRange.s.r + offset; rowIdx <= decodedRange.e.r; ++rowIdx) {
		if ((rowinfo[rowIdx] || {}).hidden) {
			continue;
		}
		const row = buildJsonRow(sheet, decodedRange, rowIdx, cols, header, headers, options);
		if (!row.isempty || (header === 1 ? options.blankrows !== false : !!options.blankrows)) {
			out[outputIndex++] = row.row;
		}
	}
	out.length = outputIndex;
	return out;
}

/** Add JSON data to a worksheet */
export function addJsonToSheet(existingSheet: WorkSheet | null, jsonData: any[], opts?: JSON2SheetOpts): WorkSheet {
	const options: any = opts || {};
	const dense = existingSheet ? (existingSheet as any)["!data"] != null : !!options.dense;
	const offset = +!options.skipHeader;
	const worksheet: any = existingSheet || {};
	if (!existingSheet && dense) {
		worksheet["!data"] = [];
	}

	let originRow = 0,
		originCol = 0;
	if (worksheet && options.origin != null) {
		if (typeof options.origin === "number") {
			originRow = options.origin;
		} else {
			const parsedOrigin = typeof options.origin === "string" ? decodeCell(options.origin) : options.origin;
			originRow = parsedOrigin.r;
			originCol = parsedOrigin.c;
		}
	}

	const range: Range = { s: { c: 0, r: 0 }, e: { c: originCol, r: originRow + jsonData.length - 1 + offset } };
	if (worksheet["!ref"]) {
		const existingRange = safeDecodeRange(worksheet["!ref"]);
		range.e.c = Math.max(range.e.c, existingRange.e.c);
		range.e.r = Math.max(range.e.r, existingRange.e.r);
		if (originRow === -1) {
			originRow = existingRange.e.r + 1;
			range.e.r = originRow + jsonData.length - 1 + offset;
		}
	} else {
		if (originRow === -1) {
			originRow = 0;
			range.e.r = jsonData.length - 1 + offset;
		}
	}

	const headers: string[] = options.header || [];
	let colIdx = 0;
	jsonData.forEach((rowObj, rowIdx) => {
		if (dense && !worksheet["!data"][originRow + rowIdx + offset]) {
			worksheet["!data"][originRow + rowIdx + offset] = [];
		}
		const denseRow = dense ? worksheet["!data"][originRow + rowIdx + offset] : null;
		Object.keys(rowObj).forEach((key: string) => {
			if ((colIdx = headers.indexOf(key)) === -1) {
				headers[(colIdx = headers.length)] = key;
			}
			let value = rowObj[key];
			let cellType = "z";
			let dateFormat = "";
			const ref = dense ? "" : encodeCol(originCol + colIdx) + encodeRow(originRow + rowIdx + offset);
			const cell: any = dense ? denseRow[originCol + colIdx] : worksheet[ref];

			if (value && typeof value === "object" && !(value instanceof Date)) {
				if (dense) {
					denseRow[originCol + colIdx] = value;
				} else {
					worksheet[ref] = value;
				}
			} else {
				if (typeof value === "number") {
					cellType = "n";
				} else if (typeof value === "boolean") {
					cellType = "b";
				} else if (typeof value === "string") {
					cellType = "s";
				} else if (value instanceof Date) {
					cellType = "d";
					if (!options.UTC) {
						value = localToUtc(value);
					}
					if (!options.cellDates) {
						cellType = "n";
						value = dateToSerialNumber(value);
					}
					dateFormat =
						cell != null && cell.z && isDateFormat(String(cell.z))
							? String(cell.z)
							: options.dateNF || formatTable[14];
				} else if (value === null && options.nullError) {
					cellType = "e";
					value = 0;
				}

				if (!cell) {
					const newCell: any = { t: cellType, v: value };
					if (dateFormat) {
						newCell.z = dateFormat;
					}
					if (dense) {
						denseRow[originCol + colIdx] = newCell;
					} else {
						worksheet[ref] = newCell;
					}
				} else {
					cell.t = cellType;
					cell.v = value;
					delete cell.w;
					if (dateFormat) {
						cell.z = dateFormat;
					}
				}
			}
		});
	});

	range.e.c = Math.max(range.e.c, originCol + headers.length - 1);
	const encodedOriginRow = encodeRow(originRow);
	if (dense && !worksheet["!data"][originRow]) {
		worksheet["!data"][originRow] = [];
	}
	if (offset) {
		for (colIdx = 0; colIdx < headers.length; ++colIdx) {
			if (dense) {
				worksheet["!data"][originRow][colIdx + originCol] = { t: "s", v: headers[colIdx] };
			} else {
				worksheet[encodeCol(colIdx + originCol) + encodedOriginRow] = { t: "s", v: headers[colIdx] };
			}
		}
	}
	worksheet["!ref"] = encodeRange(range);
	return worksheet as WorkSheet;
}

/** Create a new worksheet from JSON data */
export function jsonToSheet(js: any[], opts?: JSON2SheetOpts): WorkSheet {
	return addJsonToSheet(null, js, opts);
}
