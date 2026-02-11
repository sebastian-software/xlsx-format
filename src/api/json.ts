import type { WorkSheet, Sheet2JSONOpts, JSON2SheetOpts, CellObject, Range } from "../types.js";
import { decodeCell, encodeCol, encodeRow, encodeRange, safeDecodeRange, getCell } from "../utils/cell.js";
import { dateToSerialNumber, serialNumberToDate, utcToLocal, localToUtc } from "../utils/date.js";
import { isDateFormat } from "../ssf/format.js";
import { formatTable } from "../ssf/table.js";
import { formatCell } from "./format.js";

/**
 * Build a single JSON row object (or array) from a worksheet row.
 *
 * Reads each cell in the row, converts its value based on type (handling dates,
 * errors, booleans, etc.), and populates the output row keyed by column headers.
 *
 * @returns An object containing the built `row` and an `isempty` flag
 */
function buildJsonRow(
	sheet: WorkSheet,
	range: Range,
	rowIndex: number,
	header: number,
	headers: any[],
	options: any,
): { row: any; isempty: boolean } {
	const defval = options.defval;
	const raw = options.raw || !Object.hasOwn(options, "raw");
	let isempty = true;
	// header===1 means output as arrays (numeric indices), otherwise as objects
	const row: any = header === 1 ? [] : {};

	// Attach a non-enumerable __rowNum__ property for object rows so callers
	// can identify the source row without it polluting serialisation
	if (header !== 1) {
		try {
			Object.defineProperty(row, "__rowNum__", { value: rowIndex, enumerable: false });
		} catch {
			row.__rowNum__ = rowIndex;
		}
	}

	for (let colIdx = range.s.c; colIdx <= range.e.c; ++colIdx) {
		const val = getCell(sheet, rowIndex, colIdx);
		if (val == null || val.t === undefined) {
			// No cell: use default value if provided, otherwise skip
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
			case "z": // stub/blank cell
				if (cellValue == null) {
					break;
				}
				continue;
			case "e": // error cell
				// Error code 0 (#NULL!) maps to null; other codes to undefined
				cellValue = cellValue === 0 ? null : undefined;
				break;
			case "s": // string
			case "b": // boolean
				break;
			case "n": // number — may actually represent a date if the format is date-like
				if (!val.z || !isDateFormat(String(val.z))) {
					break;
				}
				// Number format indicates a date; convert serial number to JS Date
				cellValue = serialNumberToDate(cellValue as number);
				if (typeof cellValue === "number") {
					break;
				}
			/* falls through */
			case "d": // date
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
				// Use raw value when rawNumbers/raw is set, otherwise format for display
				row[headers[colIdx]] = (
					val.t === "n" && typeof options.rawNumbers === "boolean" ? options.rawNumbers : raw
				)
					? cellValue
					: formatCell(val, cellValue, options);
			}
			if (cellValue != null) {
				isempty = false;
			}
		}
	}
	return { row, isempty };
}

/**
 * Convert a worksheet to an array of JSON objects (or arrays).
 *
 * The first row is used as header keys by default. Supports multiple header
 * modes (raw arrays, column-letter keys, custom headers), range overrides,
 * hidden row/column skipping, blank-row handling, and date conversion.
 *
 * @param sheet - The worksheet to convert
 * @param opts - Optional conversion options (header, range, raw, rawNumbers, defval, blankrows, skipHidden, dateNF, UTC)
 * @returns An array of row objects (or arrays when `header: 1`)
 */
export function sheetToJson<T = any>(sheet: WorkSheet, opts?: Sheet2JSONOpts): T[] {
	if (sheet == null || sheet["!ref"] == null) {
		return [];
	}
	let header = 0,
		offset = 1;
	const headers: any[] = [];
	const options: any = opts || {};
	const range = options.range != null ? options.range : sheet["!ref"];

	// Determine header mode:
	//   0 = use first row values as keys (default)
	//   1 = raw array output (numeric indices)
	//   2 = column letters as keys ("A", "B", ...)
	//   3 = caller-supplied header array
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
			// Numeric range means "start from this row"
			decodedRange = safeDecodeRange(sheet["!ref"]);
			decodedRange.s.r = range;
			break;
		default:
			decodedRange = range;
	}
	// When headers are explicitly provided, data starts at the first row (no offset)
	if (header > 0) {
		offset = 0;
	}

	const out: any[] = [];
	let outputIndex = 0;
	let rowIdx = decodedRange.s.r;
	const header_cnt: Record<string, number> = {};
	const colinfo: any[] = (options.skipHidden && sheet["!cols"]) || [];
	const rowinfo: any[] = (options.skipHidden && sheet["!rows"]) || [];

	// Build header labels from the first row (or from explicit options)
	for (let colIdx = decodedRange.s.c; colIdx <= decodedRange.e.c; ++colIdx) {
		if ((colinfo[colIdx] || {}).hidden) {
			continue;
		}
		const val = getCell(sheet, rowIdx, colIdx);
		let cellValue: any, headerLabel: any;
		switch (header) {
			case 1: // Raw array: use zero-based column offset as index
				headers[colIdx] = colIdx - decodedRange.s.c;
				break;
			case 2: // Column-letter keys
				headers[colIdx] = encodeCol(colIdx);
				break;
			case 3: // Caller-supplied headers
				headers[colIdx] = (options.header as string[])[colIdx - decodedRange.s.c];
				break;
			default: {
				// Derive key from the header row cell value; use "__EMPTY" for blank headers
				const _val = val == null ? { w: "__EMPTY", t: "s" } : val;
				headerLabel = cellValue = formatCell(_val as CellObject, null, options);
				// Deduplicate by appending "_1", "_2", ... for repeated header labels
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

	// Build data rows starting after the header row (unless offset is 0)
	for (rowIdx = decodedRange.s.r + offset; rowIdx <= decodedRange.e.r; ++rowIdx) {
		if ((rowinfo[rowIdx] || {}).hidden) {
			continue;
		}
		const row = buildJsonRow(sheet, decodedRange, rowIdx, header, headers, options);
		if (!row.isempty || (header === 1 ? options.blankrows !== false : !!options.blankrows)) {
			out[outputIndex++] = row.row;
		}
	}
	out.length = outputIndex;
	return out;
}

/**
 * Add an array of JSON objects to an existing worksheet, or create a new one.
 *
 * Object keys become column headers (written in the first row unless
 * `skipHeader` is set). Supports dense and sparse storage, origin offsets,
 * date handling, and automatic type detection.
 *
 * @param existingSheet - An existing worksheet to append to, or `null` to create a new one
 * @param jsonData - Array of plain objects whose keys map to column headers
 * @param opts - Optional settings (header, origin, dense, skipHeader, cellDates, UTC, dateNF, nullError)
 * @returns The updated or newly created worksheet
 */
export function addJsonToSheet(existingSheet: WorkSheet | null, jsonData: any[], opts?: JSON2SheetOpts): WorkSheet {
	const options: any = opts || {};
	const dense = existingSheet ? (existingSheet as any)["!data"] != null : !!options.dense;
	// offset is 1 to reserve the first row for headers, 0 when skipHeader is set
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
		// origin of -1 means "append after the last row"
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

	// Collect headers from options or discover them from object keys
	const headers: string[] = options.header || [];
	let colIdx = 0;
	jsonData.forEach((rowObj, rowIdx) => {
		if (dense && !worksheet["!data"][originRow + rowIdx + offset]) {
			worksheet["!data"][originRow + rowIdx + offset] = [];
		}
		const denseRow = dense ? worksheet["!data"][originRow + rowIdx + offset] : null;
		Object.keys(rowObj).forEach((key: string) => {
			// Assign a column index for each key; new keys get appended to the headers array
			if ((colIdx = headers.indexOf(key)) === -1) {
				headers[(colIdx = headers.length)] = key;
			}
			let value = rowObj[key];
			let cellType = "z";
			let dateFormat = "";
			const ref = dense ? "" : encodeCol(originCol + colIdx) + encodeRow(originRow + rowIdx + offset);
			const cell: any = dense ? denseRow[originCol + colIdx] : worksheet[ref];

			// If the value is a pre-built cell object (non-Date object), store it directly
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
					// Preserve an existing date format on the cell if valid, else use option or default
					dateFormat =
						cell != null && cell.z && isDateFormat(String(cell.z))
							? String(cell.z)
							: options.dateNF || formatTable[14];
				} else if (value === null && options.nullError) {
					cellType = "e";
					value = 0;
				}

				if (!cell) {
					// Create a new cell object
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
					// Update the existing cell in place
					cell.t = cellType;
					cell.v = value;
					delete cell.w; // invalidate cached formatted string
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
	// Write header row unless skipHeader was set (offset === 0)
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

/**
 * Create a new worksheet from an array of JSON objects.
 *
 * This is a convenience wrapper around `addJsonToSheet` that always creates
 * a fresh worksheet.
 *
 * @param js - Array of plain objects whose keys map to column headers
 * @param opts - Optional settings (same as `addJsonToSheet`)
 * @returns A new worksheet populated with the given data
 */
export function jsonToSheet(js: any[], opts?: JSON2SheetOpts): WorkSheet {
	return addJsonToSheet(null, js, opts);
}
