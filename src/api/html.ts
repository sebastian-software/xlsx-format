import type { WorkSheet, Sheet2HTMLOpts, Range, CellObject } from "../types.js";
import { BErr } from "../types.js";
import { encodeCol, encodeRow, decodeRange, getCell } from "../utils/cell.js";
import { clampLargeExportRange } from "../utils/export-range.js";
import { escapeHtml } from "../xml/escape.js";
import { escapeHtmlAttribute, writeHtmlElement } from "../xml/writer.js";
import { formatCell } from "./format.js";
import { arrayToSheet } from "./aoa.js";
import { HTML_ENTITY_VALUES } from "./html-entities.js";

/** Default HTML document prefix wrapping the table in a minimal page structure */
const HTML_BEGIN = '<html><head><meta charset="utf-8"/><title>SheetJS Table Export</title></head><body>';
/** Default HTML document suffix closing the body and html tags */
const HTML_END = "</body></html>";
const UNSAFE_LINK_TARGET_RE = /^(?:javascript|vbscript|data):/;

function isIgnorableLinkTargetCode(code: number): boolean {
	return (
		code <= 0x20 ||
		code === 0x7f ||
		code === 0xad ||
		code === 0x61c ||
		code === 0x180e ||
		(code >= 0x200b && code <= 0x200f) ||
		(code >= 0x202a && code <= 0x202e) ||
		(code >= 0x2060 && code <= 0x206f) ||
		code === 0xfeff
	);
}

function isSanitizedLinkTarget(target: string): boolean {
	let normalized = "";
	for (let i = 0; i < target.length; ++i) {
		if (!isIgnorableLinkTargetCode(target.charCodeAt(i))) {
			normalized += target[i];
		}
	}
	normalized = normalized.toLowerCase();
	return !UNSAFE_LINK_TARGET_RE.test(normalized);
}

function sanitizeRichTextStyle(styleText: string): string {
	const safeDeclarations: string[] = [];
	for (const declaration of styleText.split(";")) {
		const separator = declaration.indexOf(":");
		if (separator === -1) {
			continue;
		}
		const property = declaration.slice(0, separator).trim().toLowerCase();
		const value = declaration
			.slice(separator + 1)
			.trim()
			.toLowerCase();
		if (property === "text-decoration" && value === "underline") {
			safeDeclarations.push("text-decoration: underline;");
		} else if (
			property === "text-underline-style" &&
			(value === "double" || value === "single-accounting" || value === "double-accounting")
		) {
			safeDeclarations.push("text-underline-style:" + value + ";");
		} else if (property === "font-size" && /^(?:0|[1-9]\d*)(?:\.\d+)?pt$/.test(value)) {
			safeDeclarations.push("font-size:" + value + ";");
		} else if (property === "text-effect" && value === "outline") {
			safeDeclarations.push("text-effect: outline;");
		} else if (property === "text-shadow" && value === "auto") {
			safeDeclarations.push("text-shadow: auto;");
		}
	}
	return safeDeclarations.join("");
}

/**
 * CellObject.h is an HTML-shaped cache, including for direct worksheet input.
 * Keep the presentation-only subset emitted by the XLSX rich-text reader and
 * render every other tag as inert text.
 */
function sanitizeCellHtml(html: string): string {
	const output: string[] = [];
	const openTags: string[] = [];
	let offset = 0;

	while (offset < html.length) {
		const tagStart = html.indexOf("<", offset);
		if (tagStart === -1) {
			output.push(html.slice(offset));
			break;
		}
		output.push(html.slice(offset, tagStart));
		const tagEnd = html.indexOf(">", tagStart + 1);
		if (tagEnd === -1) {
			output.push(escapeHtml(html.slice(tagStart)));
			break;
		}

		const token = html.slice(tagStart, tagEnd + 1);
		const simpleOpen = token.match(/^<([bis]|sup|sub)>$/i);
		const simpleClose = token.match(/^<\/([bis]|sup|sub)>$/i);
		const spanOpen = token.match(/^<span\s+style\s*=\s*(["'])([\s\S]*)\1\s*>$/i);

		if (/^<br\s*\/?>$/i.test(token)) {
			output.push("<br/>");
		} else if (simpleOpen) {
			const name = simpleOpen[1].toLowerCase();
			openTags.push(name);
			output.push("<" + name + ">");
		} else if (spanOpen) {
			openTags.push("span");
			output.push('<span style="' + escapeHtmlAttribute(sanitizeRichTextStyle(spanOpen[2])) + '">');
		} else if (simpleClose) {
			const name = simpleClose[1].toLowerCase();
			if (openTags[openTags.length - 1] === name) {
				openTags.pop();
				output.push("</" + name + ">");
			} else {
				output.push(escapeHtml(token));
			}
		} else if (/^<\/span>$/i.test(token) && openTags[openTags.length - 1] === "span") {
			openTags.pop();
			output.push("</span>");
		} else {
			output.push(escapeHtml(token));
		}
		offset = tagEnd + 1;
	}

	while (openTags.length) {
		output.push("</" + openTags.pop() + ">");
	}
	return output.join("");
}

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

		// Resolve cell content: preserve the safe rich-text subset in cell.h,
		// then fall back to escaped formatted text.
		let cellContent = "";
		if (cell && cell.v != null) {
			cellContent = cell.h ? sanitizeCellHtml(cell.h) : escapeHtml(cell.w || formatCell(cell) || "");
		}

		const cellAttrs: Record<string, string> = {};
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
				cellAttrs["data-v"] = cell.v instanceof Date ? cell.v.toISOString() : String(cell.v);
			}
			if (cell.z != null) {
				cellAttrs["data-z"] = String(cell.z);
			}
			if (cell.f != null) {
				cellAttrs["data-f"] = cell.f;
			}
			// Wrap in an anchor tag if the cell has a non-internal hyperlink,
			// filtering out unsafe URI schemes unless sanitization is disabled.
			if (
				cell.l &&
				(cell.l.Target || "#").charAt(0) !== "#" &&
				(options.sanitizeLinks === false || isSanitizedLinkTarget(cell.l.Target || ""))
			) {
				cellContent = writeHtmlElement("a", cellContent, { href: cell.l.Target });
			}
		}
		// Each cell gets a unique id: "{tableId}-{cellRef}" (e.g. "sjs-A1")
		cellAttrs.id = (options.id || "sjs") + "-" + coord;
		cells.push(writeHtmlElement("td", cellContent, cellAttrs));
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
	const range = clampLargeExportRange(ws, decodeRange(ws["!ref"] || "A1"));
	const rows: string[] = [];
	if (ws["!ref"] && range) {
		for (let rowIdx = range.s.r; rowIdx <= range.e.r; ++rowIdx) {
			rows.push(buildHtmlRow(ws, range, rowIdx, options));
		}
	}
	out.push(writeHtmlElement("table", rows.join(""), options.id ? { id: options.id } : null));
	out.push(footer);
	return out.join("");
}

const NUMERIC_REFERENCE_REPLACEMENTS: Readonly<Record<number, number>> = {
	0x80: 0x20ac,
	0x82: 0x201a,
	0x83: 0x0192,
	0x84: 0x201e,
	0x85: 0x2026,
	0x86: 0x2020,
	0x87: 0x2021,
	0x88: 0x02c6,
	0x89: 0x2030,
	0x8a: 0x0160,
	0x8b: 0x2039,
	0x8c: 0x0152,
	0x8e: 0x017d,
	0x91: 0x2018,
	0x92: 0x2019,
	0x93: 0x201c,
	0x94: 0x201d,
	0x95: 0x2022,
	0x96: 0x2013,
	0x97: 0x2014,
	0x98: 0x02dc,
	0x99: 0x2122,
	0x9a: 0x0161,
	0x9b: 0x203a,
	0x9c: 0x0153,
	0x9e: 0x017e,
	0x9f: 0x0178,
};

function decodeNumericReference(codePoint: number): string {
	if (codePoint === 0 || codePoint > 0x10ffff || (codePoint >= 0xd800 && codePoint <= 0xdfff)) {
		return "\ufffd";
	}
	return String.fromCodePoint(NUMERIC_REFERENCE_REPLACEMENTS[codePoint] ?? codePoint);
}

/** Decode semicolon-terminated HTML named and numeric character references exactly once. */
function decodeHtmlEntities(s: string): string {
	return s.replace(/&(?:#(\d+)|#[xX]([\dA-Fa-f]+)|([\dA-Za-z]+));/g, (entity, decimal, hex, name) => {
		if (name) {
			const normalizedName = String(name);
			return Object.hasOwn(HTML_ENTITY_VALUES, normalizedName) ? HTML_ENTITY_VALUES[normalizedName] : entity;
		}
		const codePoint = parseInt(decimal || hex, decimal ? 10 : 16);
		return decodeNumericReference(codePoint);
	});
}

/** Strip HTML tags from a string, returning only text content */
function stripTags(s: string): string {
	return s.replace(/<[^>]*>/g, "");
}

/** Extract an attribute value from a tag string */
function getAttr(tag: string, name: string): string | null {
	const re = new RegExp("(?:^|\\s)" + name + "\\s*=\\s*(?:\"([^\"]*)\"|'([^']*)')", "i");
	const m = tag.match(re);
	return m ? (m[1] ?? m[2]) : null;
}

/** Construct a typed cell from an explicit data-t/data-v pair. */
function typedCell(type: string, rawValue: string | null): CellObject | null {
	if (rawValue == null) {
		switch (type) {
			case "n":
				return { t: "n" };
			case "b":
				return { t: "b" };
			case "d":
				return { t: "d" };
			case "e":
				return { t: "e" };
			case "s":
				return { t: "s" };
			case "z":
				return { t: "z" };
		}
		return null;
	}
	const decodedValue = decodeHtmlEntities(rawValue);
	switch (type) {
		case "n": {
			const value = Number(decodedValue);
			return Number.isFinite(value) ? { t: "n", v: value } : null;
		}
		case "b":
			return { t: "b", v: decodedValue === "true" || decodedValue === "1" };
		case "d": {
			const value = new Date(decodedValue);
			return Number.isNaN(value.getTime()) ? null : { t: "d", v: value };
		}
		case "e": {
			const value = Number(decodedValue);
			return Number.isFinite(value) ? { t: "e", v: value } : null;
		}
		case "s":
			return { t: "s", v: decodedValue };
		case "z":
			return { t: "z" };
	}
	return null;
}

/** Try to coerce a plain text value to number or boolean */
function coerceTextValue(text: string): string | number | boolean {
	if (text === "") {
		return text;
	}
	if (text === "TRUE" || text === "true") {
		return true;
	}
	if (text === "FALSE" || text === "false") {
		return false;
	}
	const num = Number(text);
	if (text.length > 0 && !isNaN(num) && isFinite(num)) {
		return num;
	}
	return text;
}

function textFromHtml(innerHtml: string): string {
	return decodeHtmlEntities(stripTags(innerHtml.replace(/<\s*br\s*\/?>/gi, "\n"))).trim();
}

/**
 * Parse an HTML string containing a `<table>` into a WorkSheet.
 *
 * Handles `rowspan`/`colspan` attributes and uses `data-t`/`data-v`
 * attributes (when present) for round-trip fidelity.
 * This is a lightweight table parser: it recognizes quoted attributes,
 * table cells, ordinary `<br>` line breaks, and semicolon-terminated HTML
 * character references without implementing full browser DOM parsing.
 *
 * @param html - HTML string containing a table
 * @returns A WorkSheet with the parsed table data
 */
export function htmlToSheet(html: string): WorkSheet {
	// Find the first <table>...</table> block
	const tableMatch = html.match(/<table[^>]*>([\s\S]*?)<\/table>/i);
	if (!tableMatch) {
		return arrayToSheet([]);
	}

	const tableBody = tableMatch[1];
	const rowMatches = tableBody.match(/<tr[^>]*>([\s\S]*?)<\/tr>/gi) || [];

	const data: any[][] = [];
	// Track cells occupied by rowspan from previous rows: occupied[row][col] = true
	const occupied: Record<number, Record<number, boolean>> = {};

	for (let r = 0; r < rowMatches.length; r++) {
		if (!data[r]) {
			data[r] = [];
		}
		if (!occupied[r]) {
			occupied[r] = {};
		}

		const rowHtml = rowMatches[r];
		const cellMatches = rowHtml.match(/<t[dh][^>]*>[\s\S]*?<\/t[dh]>/gi) || [];

		let col = 0;
		for (let ci = 0; ci < cellMatches.length; ci++) {
			// Skip columns occupied by rowspan from prior rows
			while (occupied[r][col]) {
				col++;
			}

			const cellHtml = cellMatches[ci];
			const tagEnd = cellHtml.indexOf(">");
			const tag = cellHtml.slice(0, tagEnd + 1);

			const rowspanStr = getAttr(tag, "rowspan");
			const colspanStr = getAttr(tag, "colspan");
			const rs = rowspanStr ? parseInt(rowspanStr, 10) : 1;
			const cs = colspanStr ? parseInt(colspanStr, 10) : 1;
			const dataT = getAttr(tag, "data-t");
			const dataV = getAttr(tag, "data-v");
			const dataF = getAttr(tag, "data-f");
			const dataZ = getAttr(tag, "data-z");

			// Extract cell value
			const innerHtml = cellHtml.slice(tagEnd + 1, cellHtml.lastIndexOf("</"));
			const explicitCell = dataT ? typedCell(dataT, dataV) : null;
			let value: CellObject | string | number | boolean;
			if (explicitCell) {
				if (dataF != null) {
					explicitCell.f = decodeHtmlEntities(dataF);
				}
				if (dataZ != null) {
					explicitCell.z = decodeHtmlEntities(dataZ);
				}
				value = explicitCell;
			} else {
				value = coerceTextValue(textFromHtml(innerHtml));
			}

			data[r][col] = value;

			// Mark cells occupied by rowspan/colspan
			for (let dr = 0; dr < rs; dr++) {
				for (let dc = 0; dc < cs; dc++) {
					if (dr === 0 && dc === 0) {
						continue;
					}
					const tr = r + dr;
					const tc = col + dc;
					if (!occupied[tr]) {
						occupied[tr] = {};
					}
					occupied[tr][tc] = true;
				}
			}

			// Fill colspan cells in the current row with empty strings
			for (let dc = 1; dc < cs; dc++) {
				data[r][col + dc] = "";
			}

			col += cs;
		}
	}

	return arrayToSheet(data);
}
