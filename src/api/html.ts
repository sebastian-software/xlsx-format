import type { WorkSheet, Sheet2HTMLOpts, Range } from "../types.js";
import { BErr } from "../types.js";
import { encodeCol, encodeRow, decodeRange, getCell } from "../utils/cell.js";
import { escapeHtml } from "../xml/escape.js";
import { writeXmlElement } from "../xml/writer.js";
import { formatCell } from "./format.js";

/** Default HTML document prefix wrapping the table in a minimal page structure */
const HTML_BEGIN = '<html><head><meta charset="utf-8"/><title>SheetJS Table Export</title></head><body>';
/** Default HTML document suffix closing the body and html tags */
const HTML_END = "</body></html>";

/**
 * Build a single HTML `<tr>` row from a worksheet row, handling merged cells,
 * error coercion, hyperlinks, editable mode, and data attributes.
 */
function buildHtmlRow(ws: WorkSheet, range: Range, rowIndex: number, options: Sheet2HTMLOpts): string {
	const merges = ws["!merges"] || [];
	const cells: string[] = [];

	for (let colIdx = range.s.c; colIdx <= range.e.c; ++colIdx) {
		let rowSpan = 0,
			colSpan = 0;

		// Determine if this cell is part of a merged region
		for (let j = 0; j < merges.length; ++j) {
			if (merges[j].s.r > rowIndex || merges[j].s.c > colIdx) {
				continue;
			}
			if (merges[j].e.r < rowIndex || merges[j].e.c < colIdx) {
				continue;
			}
			// Cell is inside the merge but is not the top-left origin cell
			if (merges[j].s.r < rowIndex || merges[j].s.c < colIdx) {
				rowSpan = -1;
				break;
			}
			// Cell is the top-left origin of the merge region
			rowSpan = merges[j].e.r - merges[j].s.r + 1;
			colSpan = merges[j].e.c - merges[j].s.c + 1;
			break;
		}
		// rowSpan === -1 means this cell is swallowed by a merge; skip it
		if (rowSpan < 0) {
			continue;
		}

		const coord = encodeCol(colIdx) + encodeRow(rowIndex);
		let cell: any = getCell(ws, rowIndex, colIdx);

		// Coerce non-finite numeric cells into Excel error representations:
		// NaN -> #VALUE! (0x24), Infinity -> #DIV/0! (0x07)
		if (cell && cell.t === "n" && cell.v != null && !isFinite(cell.v)) {
			if (isNaN(cell.v)) {
				cell = { t: "e", v: 0x24, w: BErr[0x24] };
			} else {
				cell = { t: "e", v: 0x07, w: BErr[0x07] };
			}
		}

		// Resolve cell content: prefer cached HTML (cell.h), then formatted text, then empty
		let cellContent = (cell && cell.v != null && (cell.h || escapeHtml(cell.w || (formatCell(cell), cell.w) || ""))) || "";

		const cellAttrs: Record<string, any> = {};
		if (rowSpan > 1) {
			cellAttrs.rowspan = String(rowSpan);
		}
		if (colSpan > 1) {
			cellAttrs.colspan = String(colSpan);
		}

		if (options.editable) {
			// In editable mode, wrap content in a contenteditable span for inline editing
			cellContent = '<span contenteditable="true">' + cellContent + "</span>";
		} else if (cell) {
			// In non-editable mode, attach data attributes for round-tripping
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
			// Wrap in an anchor tag if the cell has a non-internal hyperlink,
			// filtering out javascript: URIs when sanitizeLinks is enabled
			if (
				cell.l &&
				(cell.l.Target || "#").charAt(0) !== "#" &&
				(!options.sanitizeLinks || (cell.l.Target || "").slice(0, 11).toLowerCase() !== "javascript:")
			) {
				cellContent = '<a href="' + escapeHtml(cell.l.Target) + '">' + cellContent + "</a>";
			}
		}
		// Each cell gets a unique id: "{tableId}-{cellRef}" (e.g. "sjs-A1")
		cellAttrs.id = (options.id || "sjs") + "-" + coord;
		cells.push(writeXmlElement("td", cellContent, cellAttrs));
	}

	return "<tr>" + cells.join("") + "</tr>";
}

/**
 * Convert a worksheet to an HTML table string.
 *
 * Generates a full HTML document (or fragment) containing a `<table>` with
 * one `<tr>` per row. Supports merged cells, hyperlinks, editable mode,
 * and data attributes for round-tripping.
 *
 * @param ws - The worksheet to convert
 * @param opts - Optional HTML generation options (header, footer, id, editable, sanitizeLinks)
 * @returns The HTML string representation of the worksheet
 */
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
