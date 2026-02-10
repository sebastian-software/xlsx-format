import type { WorkBook, WorkSheet, CellObject, Hyperlink } from "../types.js";
import { validateSheetName } from "../xlsx/workbook.js";
import { encodeCol, encodeRow, encodeRange, decodeRange, safeDecodeRange } from "../utils/cell.js";

/**
 * Create a new blank workbook, optionally containing an initial worksheet.
 *
 * @param ws - Optional worksheet to include as the first sheet
 * @param wsname - Name for the initial sheet (defaults to "Sheet1")
 * @returns A new workbook object
 */
export function createWorkbook(ws?: WorkSheet, wsname?: string): WorkBook {
	const wb: WorkBook = { SheetNames: [], Sheets: {} };
	if (ws) {
		appendSheet(wb, ws, wsname || "Sheet1");
	}
	return wb;
}

/**
 * Append a worksheet to the end of a workbook's sheet list.
 *
 * If no name is provided, generates one automatically ("Sheet1", "Sheet2", ...).
 * When `roll` is true and the name already exists, appends an incrementing
 * numeric suffix to make it unique (e.g. "Sheet1" -> "Sheet2").
 *
 * @param wb - The workbook to add the sheet to
 * @param ws - The worksheet to append
 * @param name - Optional sheet name; auto-generated if omitted
 * @param roll - If true, auto-increment the name suffix on collision instead of throwing
 * @returns The final sheet name that was used
 */
export function appendSheet(wb: WorkBook, ws: WorkSheet, name?: string, roll?: boolean): string {
	let i = 1;
	if (!name) {
		// Auto-generate "Sheet1", "Sheet2", ... until a unique name is found
		for (; i <= 0xffff; ++i, name = undefined) {
			if (wb.SheetNames.indexOf((name = "Sheet" + i)) === -1) {
				break;
			}
		}
	}
	if (!name || wb.SheetNames.length >= 0xffff) {
		throw new Error("Too many worksheets");
	}
	// When rolling, strip any trailing digits and increment to find a unique name
	if (roll && wb.SheetNames.indexOf(name) >= 0 && name.length < 32) {
		const m = name.match(/\d+$/);
		i = (m && +m[0]) || 0;
		const root = (m && name.slice(0, m.index)) || name;
		for (++i; i <= 0xffff; ++i) {
			if (wb.SheetNames.indexOf((name = root + i)) === -1) {
				break;
			}
		}
	}
	validateSheetName(name);
	if (wb.SheetNames.indexOf(name) >= 0) {
		throw new Error("Worksheet with name |" + name + "| already exists!");
	}

	wb.SheetNames.push(name);
	wb.Sheets[name] = ws;
	return name;
}

/**
 * Create a new empty worksheet.
 *
 * @param opts - Optional settings; set `dense: true` for dense storage mode (array-of-arrays backing)
 * @returns A new empty worksheet object
 */
export function createSheet(opts?: { dense?: boolean }): WorkSheet {
	const out: any = {};
	if (opts?.dense) {
		out["!data"] = [];
	}
	return out as WorkSheet;
}

/**
 * Resolve a sheet name or numeric index to a validated sheet index.
 *
 * @param wb - The workbook to search
 * @param sh - Sheet name (string) or zero-based sheet index (number)
 * @returns The zero-based sheet index
 * @throws If the sheet name or index is not found in the workbook
 */
export function getSheetIndex(wb: WorkBook, sh: number | string): number {
	if (typeof sh === "number") {
		if (sh >= 0 && wb.SheetNames.length > sh) {
			return sh;
		}
		throw new Error("Cannot find sheet # " + sh);
	} else if (typeof sh === "string") {
		const idx = wb.SheetNames.indexOf(sh);
		if (idx > -1) {
			return idx;
		}
		throw new Error("Cannot find sheet name |" + sh + "|");
	}
	throw new Error("Cannot find sheet |" + sh + "|");
}

/**
 * Set the visibility state of a worksheet in the workbook.
 *
 * Initialises the `Workbook.Sheets` metadata array if it does not yet exist.
 *
 * @param wb - The workbook containing the sheet
 * @param sh - Sheet name or zero-based index
 * @param vis - Visibility level: 0 = visible, 1 = hidden, 2 = very hidden
 */
export function setSheetVisibility(wb: WorkBook, sh: number | string, vis: 0 | 1 | 2): void {
	if (!wb.Workbook) {
		wb.Workbook = {};
	}
	if (!wb.Workbook.Sheets) {
		wb.Workbook.Sheets = [];
	}

	const idx = getSheetIndex(wb, sh);
	if (!wb.Workbook.Sheets[idx]) {
		wb.Workbook.Sheets[idx] = {};
	}

	switch (vis) {
		case 0:
		case 1:
		case 2:
			break;
		default:
			throw new Error("Bad sheet visibility setting " + vis);
	}
	wb.Workbook.Sheets[idx].Hidden = vis;
}

/**
 * Set the number format string on a cell.
 *
 * @param cell - The cell object to modify
 * @param fmt - A number format string (e.g. "0.00%") or built-in format ID
 * @returns The same cell object, for chaining
 */
export function setCellNumberFormat(cell: CellObject, fmt: string | number): CellObject {
	cell.z = fmt;
	return cell;
}

/**
 * Set or remove a hyperlink on a cell.
 *
 * Pass `undefined` or an empty string for `target` to remove an existing link.
 *
 * @param cell - The cell object to modify
 * @param target - The hyperlink URL or path; falsy to remove
 * @param tooltip - Optional tooltip text shown on hover
 * @returns The same cell object, for chaining
 */
export function setCellHyperlink(cell: CellObject, target?: string, tooltip?: string): CellObject {
	if (!target) {
		delete cell.l;
	} else {
		cell.l = { Target: target } as Hyperlink;
		if (tooltip) {
			cell.l.Tooltip = tooltip;
		}
	}
	return cell;
}

/**
 * Set an internal (within-workbook) link on a cell.
 *
 * Internal links are prefixed with "#" to distinguish them from external URLs.
 *
 * @param cell - The cell object to modify
 * @param range - The target cell reference or range string (e.g. "Sheet2!A1")
 * @param tooltip - Optional tooltip text shown on hover
 * @returns The same cell object, for chaining
 */
export function setCellInternalLink(cell: CellObject, range: string, tooltip?: string): CellObject {
	return setCellHyperlink(cell, "#" + range, tooltip);
}

/**
 * Add a comment (note) to a cell.
 *
 * Initialises the cell's comment array if it does not yet exist, then appends
 * a new comment entry.
 *
 * @param cell - The cell object to modify
 * @param text - The comment text content
 * @param author - Optional author name (defaults to "SheetJS")
 */
export function addCellComment(cell: CellObject, text: string, author?: string): void {
	if (!cell.c) {
		cell.c = [] as any;
	}
	cell.c!.push({ t: text, a: author || "SheetJS" });
}

/**
 * Set an array formula across a rectangular range of cells.
 *
 * The formula is stored on the top-left cell of the range (`cell.f`), and every
 * cell in the range receives the `cell.F` property indicating the array formula
 * extent. Optionally marks the formula as a dynamic array formula.
 *
 * @param ws - The worksheet to modify
 * @param range - The target range as a string (e.g. "A1:C3") or range object
 * @param formula - The array formula expression (without surrounding braces)
 * @param dynamic - If true, mark as a dynamic array formula (spill)
 * @returns The modified worksheet
 */
export function setArrayFormula(
	ws: WorkSheet,
	range: string | { s: { r: number; c: number }; e: { r: number; c: number } },
	formula: string,
	dynamic?: boolean,
): WorkSheet {
	const rng = typeof range !== "string" ? range : safeDecodeRange(range);
	const rngstr = typeof range === "string" ? range : encodeRange(range);

	for (let R = rng.s.r; R <= rng.e.r; ++R) {
		for (let C = rng.s.c; C <= rng.e.c; ++C) {
			const ref = encodeCol(C) + encodeRow(R);
			const dense = (ws as any)["!data"] != null;
			let cell: any;
			if (dense) {
				if (!(ws as any)["!data"][R]) {
					(ws as any)["!data"][R] = [];
				}
				cell = (ws as any)["!data"][R][C] || ((ws as any)["!data"][R][C] = { t: "z" });
			} else {
				cell = (ws as any)[ref] || ((ws as any)[ref] = { t: "z" });
			}
			cell.t = "n";
			cell.F = rngstr; // array formula range shared by all cells in the group
			delete cell.v;
			// Only the top-left cell carries the actual formula text
			if (R === rng.s.r && C === rng.s.c) {
				cell.f = formula;
				if (dynamic) {
					cell.D = true; // dynamic array (spill) flag
				}
			}
		}
	}

	// Expand the worksheet's !ref to encompass the array formula range
	if (ws["!ref"]) {
		const wsr = decodeRange(ws["!ref"]);
		if (wsr.s.r > rng.s.r) {
			wsr.s.r = rng.s.r;
		}
		if (wsr.s.c > rng.s.c) {
			wsr.s.c = rng.s.c;
		}
		if (wsr.e.r < rng.e.r) {
			wsr.e.r = rng.e.r;
		}
		if (wsr.e.c < rng.e.c) {
			wsr.e.c = rng.e.c;
		}
		ws["!ref"] = encodeRange(wsr);
	}
	return ws;
}

/**
 * Convert a worksheet to an array of formula strings.
 *
 * Each entry has the format "CellRef=Value" (e.g. "A1=42", "B2='Hello").
 * For array formulas, the ref is the full range (e.g. "A1:C3={formula}").
 * String values are prefixed with a single quote; booleans become TRUE/FALSE.
 *
 * @param ws - The worksheet to extract formulas from
 * @returns An array of "ref=value" strings representing every non-empty cell
 */
export function sheetToFormulae(ws: WorkSheet): string[] {
	if (ws == null || ws["!ref"] == null) {
		return [];
	}
	const r = safeDecodeRange(ws["!ref"]);
	const cols: string[] = [];
	const cmds: string[] = [];
	const dense = (ws as any)["!data"] != null;

	// Pre-compute column letters for the range
	for (let C = r.s.c; C <= r.e.c; ++C) {
		cols[C] = encodeCol(C);
	}

	for (let R = r.s.r; R <= r.e.r; ++R) {
		const rr = encodeRow(R);
		for (let C = r.s.c; C <= r.e.c; ++C) {
			const y = cols[C] + rr;
			const x: any = dense ? ((ws as any)["!data"][R] || [])[C] : (ws as any)[y];
			if (x === undefined) {
				continue;
			}
			let val = "";
			let ref = y;
			// For array formulas, use the range as the ref; skip non-origin cells
			if (x.F != null) {
				ref = x.F;
				if (!x.f) {
					continue;
				}
				val = x.f;
				// Normalise single-cell array formula refs to "A1:A1" format
				if (ref.indexOf(":") === -1) {
					ref = ref + ":" + ref;
				}
			}
			if (x.f != null) {
				val = x.f;
			} else if (x.t === "z") {
				continue;
			} else if (x.t === "n" && x.v != null) {
				val = "" + x.v;
			} else if (x.t === "b") {
				val = x.v ? "TRUE" : "FALSE";
			} else if (x.w !== undefined) {
				val = "'" + x.w;
			} else if (x.v === undefined) {
				continue;
			} else if (x.t === "s") {
				val = "'" + x.v; // prefix string values with a single quote
			} else {
				val = "" + x.v;
			}
			cmds.push(ref + "=" + val);
		}
	}
	return cmds;
}
