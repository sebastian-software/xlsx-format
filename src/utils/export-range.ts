import type { Range, WorkSheet } from "../types.js";
import { XlsxError } from "../errors.js";
import { decodeCell } from "./cell.js";
import { DEFAULT_MAX_EXPORT_CELLS, WorksheetCellBudget, XLSX_MAX_COLUMNS, XLSX_MAX_ROWS } from "./worksheet-budget.js";

const CELL_REF_RE = /^[A-Z]+[1-9]\d*$/;
export function rangeCellCount(range: Range, kind = "worksheet range"): number {
	for (const [name, value] of [
		["start row", range.s.r],
		["start column", range.s.c],
		["end row", range.e.r],
		["end column", range.e.c],
	] as const) {
		if (!Number.isSafeInteger(value) || value < 0) {
			throw new XlsxError("INVALID_ARGUMENT", `Invalid ${kind}: ${name} must be a non-negative safe integer`);
		}
	}
	if (range.e.r < range.s.r || range.e.c < range.s.c) {
		throw new XlsxError("INVALID_ARGUMENT", `Invalid ${kind}: end must not precede start`);
	}
	if (range.e.r >= XLSX_MAX_ROWS || range.e.c >= XLSX_MAX_COLUMNS) {
		throw new XlsxError("INVALID_ARGUMENT", `Invalid ${kind}: range exceeds XLSX worksheet bounds`);
	}
	const rows = range.e.r - range.s.r + 1;
	const cols = range.e.c - range.s.c + 1;
	const count = rows * cols;
	if (!Number.isSafeInteger(count)) {
		throw new XlsxError("INVALID_ARGUMENT", `Invalid ${kind}: cell count is not a safe integer`);
	}
	return count;
}

function occupiedRangeEnd(sheet: WorkSheet, range: Range): { r: number; c: number } | null {
	let maxRow = -1;
	let maxCol = -1;
	const data = (sheet as any)["!data"];
	if (data != null) {
		for (const rowKey of Object.keys(data)) {
			const rowIdx = Number(rowKey);
			if (!Number.isInteger(rowIdx) || rowIdx < range.s.r || rowIdx > range.e.r) {
				continue;
			}
			const row = data[rowIdx];
			if (!row) {
				continue;
			}
			for (const colKey of Object.keys(row)) {
				const colIdx = Number(colKey);
				if (!Number.isInteger(colIdx) || colIdx < range.s.c || colIdx > range.e.c || row[colIdx] == null) {
					continue;
				}
				if (rowIdx > maxRow) {
					maxRow = rowIdx;
				}
				if (colIdx > maxCol) {
					maxCol = colIdx;
				}
			}
		}
	} else {
		for (const ref of Object.keys(sheet)) {
			if (!CELL_REF_RE.test(ref) || (sheet as any)[ref] == null) {
				continue;
			}
			const cell = decodeCell(ref);
			if (cell.r < range.s.r || cell.r > range.e.r || cell.c < range.s.c || cell.c > range.e.c) {
				continue;
			}
			if (cell.r > maxRow) {
				maxRow = cell.r;
			}
			if (cell.c > maxCol) {
				maxCol = cell.c;
			}
		}
	}
	return maxRow === -1 ? null : { r: maxRow, c: maxCol };
}

export function clampLargeExportRange(
	sheet: WorkSheet,
	range: Range,
	budget = new WorksheetCellBudget(undefined, DEFAULT_MAX_EXPORT_CELLS),
	clampToOccupied = true,
): Range | null {
	let count = rangeCellCount(range, "worksheet export range");
	if (count <= budget.limit) {
		budget.charge("worksheet export cell", count);
		return range;
	}
	if (!clampToOccupied) {
		budget.charge("worksheet export cell", count);
		return range;
	}
	const end = occupiedRangeEnd(sheet, range);
	if (!end) {
		return null;
	}
	const clamped = {
		s: { r: range.s.r, c: range.s.c },
		e: {
			r: Math.max(range.s.r, Math.min(range.e.r, end.r)),
			c: Math.max(range.s.c, Math.min(range.e.c, end.c)),
		},
	};
	count = rangeCellCount(clamped, "worksheet export range");
	budget.charge("worksheet export cell", count);
	return clamped;
}
