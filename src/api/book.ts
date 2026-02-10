import type { WorkBook, WorkSheet, CellObject, Hyperlink } from "../types.js";
import { validateSheetName } from "../xlsx/workbook.js";
import { encodeCol, encodeRow, encodeRange, decodeRange, safeDecodeRange } from "../utils/cell.js";

/** Create a new blank workbook, optionally with a first sheet */
export function createWorkbook(ws?: WorkSheet, wsname?: string): WorkBook {
	const wb: WorkBook = { SheetNames: [], Sheets: {} };
	if (ws) {
		appendSheet(wb, ws, wsname || "Sheet1");
	}
	return wb;
}

/** Add a worksheet to the end of a workbook */
export function appendSheet(wb: WorkBook, ws: WorkSheet, name?: string, roll?: boolean): string {
	let i = 1;
	if (!name) {
		for (; i <= 0xffff; ++i, name = undefined) {
			if (wb.SheetNames.indexOf((name = "Sheet" + i)) === -1) {
				break;
			}
		}
	}
	if (!name || wb.SheetNames.length >= 0xffff) {
		throw new Error("Too many worksheets");
	}
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

/** Create a new empty worksheet */
export function createSheet(opts?: { dense?: boolean }): WorkSheet {
	const out: any = {};
	if (opts?.dense) {
		out["!data"] = [];
	}
	return out as WorkSheet;
}

/** Find sheet index for given name or validate index */
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

/** Set sheet visibility (0=visible, 1=hidden, 2=veryHidden) */
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

/** Set a cell's number format */
export function setCellNumberFormat(cell: CellObject, fmt: string | number): CellObject {
	cell.z = fmt;
	return cell;
}

/** Set a cell's hyperlink */
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

/** Set an internal link (starts with #) on a cell */
export function setCellInternalLink(cell: CellObject, range: string, tooltip?: string): CellObject {
	return setCellHyperlink(cell, "#" + range, tooltip);
}

/** Add a comment to a cell */
export function addCellComment(cell: CellObject, text: string, author?: string): void {
	if (!cell.c) {
		cell.c = [] as any;
	}
	cell.c!.push({ t: text, a: author || "SheetJS" });
}

/** Set an array formula on a range of cells */
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
			cell.F = rngstr;
			delete cell.v;
			if (R === rng.s.r && C === rng.s.c) {
				cell.f = formula;
				if (dynamic) {
					cell.D = true;
				}
			}
		}
	}

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

/** Convert a worksheet to an array of formula strings */
export function sheetToFormulae(ws: WorkSheet): string[] {
	if (ws == null || ws["!ref"] == null) {
		return [];
	}
	const r = safeDecodeRange(ws["!ref"]);
	const cols: string[] = [];
	const cmds: string[] = [];
	const dense = (ws as any)["!data"] != null;

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
			if (x.F != null) {
				ref = x.F;
				if (!x.f) {
					continue;
				}
				val = x.f;
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
				val = "'" + x.v;
			} else {
				val = "" + x.v;
			}
			cmds.push(ref + "=" + val);
		}
	}
	return cmds;
}
