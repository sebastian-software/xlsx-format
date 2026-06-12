import type { Range, WorkSheet } from "../types.js";
import { decodeCell } from "./cell.js";

const CELL_REF_RE = /^[A-Z]+[1-9][0-9]*$/;
const MAX_EXPORT_CELLS = 1000000;

function rangeCellCount(range: Range): number {
	const rows = range.e.r - range.s.r + 1;
	const cols = range.e.c - range.s.c + 1;
	return rows > 0 && cols > 0 ? rows * cols : 0;
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

export function clampLargeExportRange(sheet: WorkSheet, range: Range): Range | null {
	if (rangeCellCount(range) <= MAX_EXPORT_CELLS) {
		return range;
	}
	const end = occupiedRangeEnd(sheet, range);
	if (!end) {
		return null;
	}
	return {
		s: { r: range.s.r, c: range.s.c },
		e: {
			r: Math.max(range.s.r, Math.min(range.e.r, end.r)),
			c: Math.max(range.s.c, Math.min(range.e.c, end.c)),
		},
	};
}
