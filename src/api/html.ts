import type { WorkSheet, Sheet2HTMLOpts, Range } from "../types.js";
import { BErr } from "../types.js";
import { encodeCol, encodeRow, decodeRange, getCell } from "../utils/cell.js";
import { escapeHtml } from "../xml/escape.js";
import { writeXmlElement } from "../xml/writer.js";
import { formatCell } from "./format.js";

const HTML_BEGIN = '<html><head><meta charset="utf-8"/><title>SheetJS Table Export</title></head><body>';
const HTML_END = "</body></html>";

function buildHtmlRow(ws: WorkSheet, range: Range, rowIndex: number, options: Sheet2HTMLOpts): string {
	const merges = ws["!merges"] || [];
	const cells: string[] = [];

	for (let colIdx = range.s.c; colIdx <= range.e.c; ++colIdx) {
		let rowSpan = 0,
			colSpan = 0;
		for (let j = 0; j < merges.length; ++j) {
			if (merges[j].s.r > rowIndex || merges[j].s.c > colIdx) {
				continue;
			}
			if (merges[j].e.r < rowIndex || merges[j].e.c < colIdx) {
				continue;
			}
			if (merges[j].s.r < rowIndex || merges[j].s.c < colIdx) {
				rowSpan = -1;
				break;
			}
			rowSpan = merges[j].e.r - merges[j].s.r + 1;
			colSpan = merges[j].e.c - merges[j].s.c + 1;
			break;
		}
		if (rowSpan < 0) {
			continue;
		}

		const coord = encodeCol(colIdx) + encodeRow(rowIndex);
		let cell: any = getCell(ws, rowIndex, colIdx);

		if (cell && cell.t === "n" && cell.v != null && !isFinite(cell.v)) {
			if (isNaN(cell.v)) {
				cell = { t: "e", v: 0x24, w: BErr[0x24] };
			} else {
				cell = { t: "e", v: 0x07, w: BErr[0x07] };
			}
		}

		let cellContent = (cell && cell.v != null && (cell.h || escapeHtml(cell.w || (formatCell(cell), cell.w) || ""))) || "";

		const cellAttrs: Record<string, any> = {};
		if (rowSpan > 1) {
			cellAttrs.rowspan = String(rowSpan);
		}
		if (colSpan > 1) {
			cellAttrs.colspan = String(colSpan);
		}

		if (options.editable) {
			cellContent = '<span contenteditable="true">' + cellContent + "</span>";
		} else if (cell) {
			cellAttrs["data-t"] = (cell && cell.t) || "z";
			if (cell.v != null) {
				cellAttrs["data-v"] = escapeHtml(cell.v instanceof Date ? cell.v.toISOString() : String(cell.v));
			}
			if (cell.z != null) {
				cellAttrs["data-z"] = String(cell.z);
			}
			if (cell.f != null) {
				cellAttrs["data-f"] = escapeHtml(cell.f);
			}
			if (
				cell.l &&
				(cell.l.Target || "#").charAt(0) !== "#" &&
				(!options.sanitizeLinks || (cell.l.Target || "").slice(0, 11).toLowerCase() !== "javascript:")
			) {
				cellContent = '<a href="' + escapeHtml(cell.l.Target) + '">' + cellContent + "</a>";
			}
		}
		cellAttrs.id = (options.id || "sjs") + "-" + coord;
		cells.push(writeXmlElement("td", cellContent, cellAttrs));
	}

	return "<tr>" + cells.join("") + "</tr>";
}

/** Convert a worksheet to an HTML table string */
export function sheetToHtml(ws: WorkSheet, opts?: Sheet2HTMLOpts): string {
	const options: Sheet2HTMLOpts = opts || {};
	const header = options.header != null ? options.header : HTML_BEGIN;
	const footer = options.footer != null ? options.footer : HTML_END;
	const out: string[] = [header];
	const range = decodeRange(ws["!ref"] || "A1");
	out.push("<table" + (options.id ? ' id="' + options.id + '"' : "") + ">");
	if (ws["!ref"]) {
		for (let rowIdx = range.s.r; rowIdx <= range.e.r; ++rowIdx) {
			out.push(buildHtmlRow(ws, range, rowIdx, options));
		}
	}
	out.push("</table>" + footer);
	return out.join("");
}
