import { XlsxError } from "../errors.js";

export const XLSX_MAX_ROWS = 1_048_576;
export const XLSX_MAX_COLUMNS = 16_384;
export const DEFAULT_MAX_EXPORT_CELLS = 1_000_000;
export const DEFAULT_MAX_IMPORT_CELLS = 10_000_000;

export function worksheetOptionLimit(value: number | undefined, fallback: number, name: string): number {
	if (value == null) {
		return fallback;
	}
	if (!Number.isSafeInteger(value) || value < 0) {
		throw new XlsxError("INVALID_ARGUMENT", `${name} must be a non-negative safe integer`);
	}
	return value;
}

/** Cumulative per-sheet budget for explicit and derived worksheet cell work. */
export class WorksheetCellBudget {
	readonly limit: number;
	private used = 0;

	constructor(value: number | undefined, fallback: number) {
		this.limit = worksheetOptionLimit(value, fallback, "maxWorksheetCells");
	}

	charge(kind: string, count: number): void {
		if (!Number.isSafeInteger(count) || count < 0) {
			throw new XlsxError("MALFORMED", `Invalid ${kind} count ${count}`);
		}
		if (count > this.limit - this.used) {
			throw new XlsxError("LIMIT_EXCEEDED", `${kind} count ${this.used + count} exceeds limit ${this.limit}`);
		}
		this.used += count;
	}
}
