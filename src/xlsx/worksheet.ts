import {
	BErr,
	RBErr,
	type WorkSheet,
	type CellObject,
	type Range,
	type ColInfo,
	type MarginInfo,
	type SheetView,
} from "../types.js";
import { parseXmlTag, XML_HEADER } from "../xml/parser.js";
import { unescapeXml, escapeXml } from "../xml/escape.js";
import { writeXmlElement } from "../xml/writer.js";
import { XMLNS_main, RELS as RELTYPE } from "../xml/namespaces.js";
import {
	assertXmlCountWithinLimit,
	assertXmlPartLimits,
	DEFAULT_MAX_WORKSHEET_CELLS,
	DEFAULT_MAX_WORKSHEET_ROWS,
	xmlOptionLimit,
} from "../xml/limits.js";
import { safeDecodeRange, encodeRange, encodeCell, decodeCell } from "../utils/cell.js";
import { formatNumber, getDateTimeFormatKind } from "../ssf/format.js";
import { formatTable } from "../ssf/table.js";
import { dateToSerialNumber, serialNumberToDate } from "../utils/date.js";
import { normalizeRichTextXml, parseStringItem, type SST, type XLString } from "./shared-strings.js";
import type { StylesData } from "./styles.js";
import { getCellStyleIndex, getStyleFromXf } from "./styles.js";
import type { Relationships } from "../opc/relationships.js";
import { shiftFormulaStr } from "./formula.js";

/** Regex patterns for extracting various worksheet XML elements */
const mergecregex = /<(?:\w+:)?mergeCell ref=["'][A-Z0-9:]+['"]\s*[/]?>/g;
const hlinkregex = /<(?:\w+:)?hyperlink [^<>]*>/gm;
const dimregex = /"(\w*:\w*)"/;
const colregex = /<(?:\w+:)?col\b[^<>]*[/]?>/g;
const afregex = /<(?:\w:)?autoFilter[^>]*([/]|>([\s\S]*)<\/(?:\w:)?autoFilter)>/g;
const marginregex = /<(?:\w+:)?pageMargins[^<>]*\/>/g;
const sheetViewRegex = /<(?:\w+:)?sheetView\b[^>]*(?:\/>|>[\s\S]*?<\/(?:\w+:)?sheetView>)/;

/** Parse the <dimension> element to set the sheet reference range */
function parseWorksheetXml_dim(ws: WorkSheet, s: string): void {
	const d = safeDecodeRange(s);
	if (d.s.r <= d.e.r && d.s.c <= d.e.c && d.s.r >= 0 && d.s.c >= 0) {
		ws["!ref"] = encodeRange(d);
	}
}

/** Parse <pageMargins> attributes with defaults matching Excel's standard margins */
function parseWorksheetXml_margins(tag: Record<string, any>): MarginInfo {
	const margin = (value: unknown, fallback: number): number => {
		const parsed = typeof value === "string" ? parseFloat(value) : Number(value);
		return Number.isFinite(parsed) ? parsed : fallback;
	};
	return {
		left: margin(tag.left, 0.7),
		right: margin(tag.right, 0.7),
		top: margin(tag.top, 0.75),
		bottom: margin(tag.bottom, 0.75),
		header: margin(tag.header, 0.3),
		footer: margin(tag.footer, 0.3),
	};
}

function parseWorksheetXml_error(value: string | null): string | number {
	if (value === null) {
		return "";
	}
	const token = unescapeXml(value);
	return Object.hasOwn(RBErr, token) ? RBErr[token] : token;
}

function serializeWorksheetXml_error(value: unknown): string {
	if (typeof value === "number") {
		return BErr[value] ?? String(value);
	}
	return typeof value === "string" ? value : "";
}

/** Parse <autoFilter> element extracting the filter reference range */
function parseWorksheetXml_autofilter(data: string, opts?: any): { ref: string } {
	const tag = parseXmlTag(data.match(/<[^>]*>/)?.[0] || "", undefined, undefined, opts);
	return { ref: tag.ref || "" };
}

/** Parse <col> elements to populate column width and hidden state */
function parseWorksheetXml_cols(columns: ColInfo[], cols: string[], opts?: any): void {
	for (let i = 0; i < cols.length; ++i) {
		const tag = parseXmlTag(cols[i], undefined, undefined, opts);
		if (!tag.min || !tag.max) {
			continue;
		}
		// min/max are 1-based column indices in the XML
		const min = parseInt(tag.min, 10) - 1;
		const max = parseInt(tag.max, 10) - 1;
		const width = tag.width ? parseFloat(tag.width) : undefined;
		const hidden = tag.hidden === "1";
		for (let j = min; j <= max; ++j) {
			if (!columns[j]) {
				columns[j] = {};
			}
			if (width !== undefined) {
				columns[j].width = width;
			}
			if (hidden) {
				columns[j].hidden = true;
			}
		}
	}
}

function parseWorksheetXml_views(data: string, opts?: any): SheetView[] | undefined {
	const match = data.match(sheetViewRegex);
	if (!match) {
		return undefined;
	}
	const pane = match[0].match(/<(?:\w+:)?pane\b[^>]*\/>/);
	if (!pane) {
		return undefined;
	}
	const tag = parseXmlTag(pane[0], undefined, undefined, opts);
	if (tag.state !== "frozen") {
		return undefined;
	}
	const view: SheetView = { state: "frozen" };
	if (tag.xSplit) {
		view.xSplit = parseFloat(tag.xSplit);
	}
	if (tag.ySplit) {
		view.ySplit = parseFloat(tag.ySplit);
	}
	if (tag.topLeftCell) {
		view.topLeftCell = tag.topLeftCell;
	}
	if (tag.activePane === "topRight" || tag.activePane === "bottomLeft" || tag.activePane === "bottomRight") {
		view.activePane = tag.activePane;
	}
	return [view];
}

/** Parse <hyperlink> elements and attach link objects to the corresponding cells */
function parseWorksheetXml_hlinks(s: WorkSheet, hlinks: string[], rels: Relationships, opts?: any): void {
	for (let i = 0; i < hlinks.length; ++i) {
		const tag = parseXmlTag(hlinks[i], undefined, undefined, opts);
		if (!tag.ref) {
			continue;
		}
		// Hyperlinks can span a range of cells
		const rng = safeDecodeRange(tag.ref);
		for (let R = rng.s.r; R <= rng.e.r; ++R) {
			for (let C = rng.s.c; C <= rng.e.c; ++C) {
				const addr = encodeCell({ r: R, c: C });
				const dense = s["!data"] != null;
				let cell: CellObject | undefined;
				if (dense) {
					if (!s["!data"]![R]) {
						s["!data"]![R] = [];
					}
					cell = s["!data"]![R]![C];
				} else {
					cell = s[addr] as CellObject | undefined;
				}
				if (!cell) {
					cell = { t: "z", v: undefined } as any;
					if (dense) {
						s["!data"]![R]![C] = cell;
					} else {
						s[addr] = cell;
					}
				}
				// Resolve the hyperlink target from the relationship by r:id
				let target = "";
				if (tag.id) {
					const rel = rels["!id"]?.[tag.id];
					if (rel) {
						target = rel.Target;
					}
				}
				if (tag.location) {
					target += "#" + tag.location;
				}
				cell!.l = { Target: target };
				if (tag.tooltip) {
					cell!.l.Tooltip = tag.tooltip;
				}
			}
		}
	}
}

/** Regex to match <c> (cell) elements, capturing inner content */
const cellregex = /<(?:\w+:)?c\b[^>]*?(?:\/>|>([\s\S]*?)<\/(?:\w+:)?c>)/g;

/**
 * Parse the <sheetData> XML into cell objects within the worksheet.
 * Processes rows and cells, handling all cell types (string, number, boolean, etc.).
 */
function parseSheetData(
	sdata: string,
	s: WorkSheet,
	opts: any,
	refguess: Range,
	_themes: any,
	styles: StylesData | undefined,
	wb: any,
): void {
	const dense = s["!data"] != null;
	const date1904 = wb?.WBProps?.date1904;
	const maxWorksheetRows = xmlOptionLimit(opts.maxWorksheetRows, DEFAULT_MAX_WORKSHEET_ROWS, "maxWorksheetRows");
	const maxWorksheetCells = xmlOptionLimit(opts.maxWorksheetCells, DEFAULT_MAX_WORKSHEET_CELLS, "maxWorksheetCells");
	let rowCount = 0;
	let cellCount = 0;
	const sharedFormulas = new Map<string, { formula: string; origin: string }>();
	const pendingSharedFormulas: { cell: CellObject; ref: string; si: string }[] = [];
	const unresolvedSharedFormulaIds = new Set<string>();

	// Split by </row> boundaries to isolate each row's content
	const rows = sdata.split(/<\/(?:\w+:)?row>/);

	for (let ri = 0; ri < rows.length; ++ri) {
		const rowStr = rows[ri];
		if (!rowStr) {
			continue;
		}

		// Find the <row> tag to get the row number
		const rowTagMatch = rowStr.match(/<(?:\w+:)?row\b[^>]*>/);
		if (!rowTagMatch) {
			continue;
		}
		assertXmlCountWithinLimit("worksheet row", ++rowCount, maxWorksheetRows);
		const rowTag = parseXmlTag(rowTagMatch[0], undefined, undefined, opts);
		// Row numbers in XML are 1-based
		const R = parseInt(rowTag.r, 10) - 1;
		if (isNaN(R)) {
			continue;
		}

		const rowIsWithinLimit = !opts.sheetRows || R < opts.sheetRows;
		// Extract row properties (height, hidden) only for retained rows.
		if (rowIsWithinLimit && (rowTag.ht !== undefined || rowTag.hidden !== undefined)) {
			if (!s["!rows"]) {
				s["!rows"] = [];
			}
			if (!s["!rows"][R]) {
				s["!rows"][R] = {};
			}
			if (rowTag.ht !== undefined) {
				const hpt = parseFloat(rowTag.ht);
				if (Number.isFinite(hpt)) {
					s["!rows"][R].hpt = hpt;
				}
			}
			if (rowTag.hidden === "1") {
				s["!rows"][R].hidden = true;
			}
		}

		// Skipped rows are not materialized. If a retained cell references a
		// shared-formula master later in the sheet, scan only until those masters
		// are found and count every inspected cell against the safety limit.
		if (!rowIsWithinLimit) {
			if (opts.cellFormula !== false && unresolvedSharedFormulaIds.size > 0) {
				cellregex.lastIndex = 0;
				let skippedCellMatch;
				while ((skippedCellMatch = cellregex.exec(rowStr))) {
					assertXmlCountWithinLimit("worksheet cell", ++cellCount, maxWorksheetCells);
					const skippedCellTag = parseXmlTag(
						skippedCellMatch[0].match(/<(?:\w+:)?c\b[^>]*/)?.[0] + ">" || "",
						undefined,
						undefined,
						opts,
					);
					const skippedRef = skippedCellTag.r;
					const skippedCellValue = skippedCellMatch[1] || "";
					const skippedFormulaTagMatch = skippedCellValue.match(/<(?:\w+:)?f\b[^>]*>/);
					const skippedFormulaMatch = skippedCellValue.match(/<(?:\w+:)?f\b[^>]*>([\s\S]*?)<\/(?:\w+:)?f>/);
					if (!skippedRef || !skippedFormulaTagMatch || !skippedFormulaMatch?.[1]) {
						continue;
					}
					const skippedFormulaTag = parseXmlTag(skippedFormulaTagMatch[0], undefined, undefined, opts);
					if (skippedFormulaTag.t !== "shared" || skippedFormulaTag.si == null) {
						continue;
					}
					const si = String(skippedFormulaTag.si);
					if (!unresolvedSharedFormulaIds.has(si)) {
						continue;
					}
					sharedFormulas.set(si, {
						formula: unescapeXml(skippedFormulaMatch[1]),
						origin: skippedRef,
					});
					unresolvedSharedFormulaIds.delete(si);
					if (unresolvedSharedFormulaIds.size === 0) {
						break;
					}
				}
			}
			continue;
		}

		// Parse cells in this row
		cellregex.lastIndex = 0;
		let cellMatch;
		while ((cellMatch = cellregex.exec(rowStr))) {
			assertXmlCountWithinLimit("worksheet cell", ++cellCount, maxWorksheetCells);
			const cellTag = parseXmlTag(
				cellMatch[0].match(/<(?:\w+:)?c\b[^>]*/)?.[0] + ">" || "",
				undefined,
				undefined,
				opts,
			);
			const ref = cellTag.r;
			if (!ref) {
				continue;
			}

			// Decode column letter(s) from the cell reference (e.g. "AB12" -> column index)
			let C = 0;
			for (let ci = 0; ci < ref.length; ++ci) {
				const cc = ref.charCodeAt(ci);
				// A-Z: accumulate column index (base-26)
				if (cc >= 65 && cc <= 90) {
					C = 26 * C + (cc - 64);
				} else {
					break;
				}
			}
			C -= 1; // Convert to 0-based

			// Expand the guessed range to include this cell
			if (R < refguess.s.r) {
				refguess.s.r = R;
			}
			if (R > refguess.e.r) {
				refguess.e.r = R;
			}
			if (C < refguess.s.c) {
				refguess.s.c = C;
			}
			if (C > refguess.e.c) {
				refguess.e.c = C;
			}

			// t = cell type attribute: s(shared string), str(formula string), inlineStr, b(boolean), e(error), d(date), n(number/default)
			const cellType = cellTag.t || "n";
			// s = style index
			const cellStyle = cellTag.s ? parseInt(cellTag.s, 10) : 0;
			const cellValue = cellMatch[1] || "";

			let cell: CellObject;

			// Extract <v> (value), <f> (formula), and <is> (inline string) sub-elements
			const vMatch = cellValue.match(/<(?:\w+:)?v>([\s\S]*?)<\/(?:\w+:)?v>/);
			const fTagMatch = cellValue.match(/<(?:\w+:)?f\b[^>]*>/);
			const fMatch = cellValue.match(/<(?:\w+:)?f\b[^>]*>([\s\S]*?)<\/(?:\w+:)?f>/);
			const isMatch = cellValue.match(/<(?:\w+:)?is\b[^>]*>([\s\S]*?)<\/(?:\w+:)?is\s*>/);
			const preserveFormula = fTagMatch != null && opts.cellFormula !== false;

			const v = vMatch ? vMatch[1] : null;

			switch (cellType) {
				case "s": // shared string - value is an index into the SST
					if (v !== null) {
						const idx = parseInt(v, 10);
						cell = { t: "s", v: "" };
						// Store SST index for later resolution via resolveSharedStrings()
						(cell as any)._sstIdx = idx;
					} else {
						cell = { t: "z" };
					}
					break;
				case "str": // formula result that is a string
					cell = { t: "s", v: v ? unescapeXml(v) : "" };
					break;
				case "inlineStr":
					if (isMatch) {
						const inlineString = parseStringItem(isMatch[1], opts);
						cell = { t: "s", v: inlineString.t };
						if (inlineString.r !== undefined) {
							cell.r = inlineString.r;
						}
						if (inlineString.h !== undefined) {
							cell.h = inlineString.h;
						}
					} else {
						cell = { t: "s", v: "" };
					}
					break;
				case "b": // boolean
					cell = { t: "b", v: v === "1" };
					break;
				case "e": // error
					cell = { t: "e", v: parseWorksheetXml_error(v) };
					(cell as any).w = v === null ? "" : unescapeXml(v);
					break;
				case "d": // ISO 8601 date string
					if (v) {
						cell = { t: "d", v: new Date(v) };
					} else {
						cell = { t: "z" };
					}
					break;
				default: // "n" (number) or unrecognized
					if (v !== null) {
						cell = { t: "n", v: parseFloat(v) };
					} else {
						if (!preserveFormula && !opts.sheetStubs) {
							continue;
						}
						cell = preserveFormula ? { t: "n" } : { t: "z" };
					}
					break;
			}

			// Apply style reference from the cellXf table
			if (styles) {
				const xf = styles.CellXf[cellStyle];
				if (xf) {
					const numberFormat = styles.NumberFmt[xf.numFmtId] || formatTable[xf.numFmtId];
					if (cellStyle > 0 || xf.numFmtId !== 0) {
						cell.XF = { numFmtId: xf.numFmtId, numFmt: numberFormat };
					}
					if (opts.cellStyles) {
						const style = getStyleFromXf(styles, cellStyle);
						if (style) {
							cell.s = style;
						}
					}
					if (opts.cellNF) {
						if (numberFormat) {
							cell.z = numberFormat;
						}
					}
				}
			}

			// Extract formula
			if (fTagMatch && opts.cellFormula !== false) {
				const fTag = parseXmlTag(fTagMatch[0], undefined, undefined, opts);
				const formula = fMatch ? unescapeXml(fMatch[1]) : undefined;
				if (fTag.t === "shared" && fTag.si != null) {
					const si = String(fTag.si);
					if (formula) {
						cell.f = formula;
						sharedFormulas.set(si, { formula, origin: ref });
						unresolvedSharedFormulaIds.delete(si);
					} else {
						const master = sharedFormulas.get(si);
						if (master) {
							const origin = decodeCell(master.origin);
							const target = decodeCell(ref);
							cell.f = shiftFormulaStr(master.formula, {
								r: target.r - origin.r,
								c: target.c - origin.c,
							});
						} else {
							pendingSharedFormulas.push({ cell, ref, si });
							unresolvedSharedFormulaIds.add(si);
						}
					}
				} else if (formula !== undefined) {
					cell.f = formula;
				}
				if (fTag.t === "array" && fTag.ref) {
					cell.F = fTag.ref; // Array formula range
					const cellMetadata = cellTag.cm ? opts.xlmeta?.Cell?.[Number(cellTag.cm) - 1] : undefined;
					cell.D = fTag.dt === "1" || cellMetadata?.type === "XLDAPR";
				}
			}

			// Format the cell value as display text
			if (cell.t === "n") {
				const numFmtId = cell.XF?.numFmtId ?? 0;
				const nfmt = cell.z || cell.XF?.numFmt || styles?.NumberFmt[numFmtId] || formatTable[numFmtId];
				if (opts.cellText !== false && typeof cell.v === "number") {
					if (nfmt) {
						try {
							cell.w = formatNumber(nfmt, cell.v, { date1904, dateNF: opts.dateNF });
						} catch {}
					}
				}
				// Date typing is independent from display-text generation.
				if (opts.cellDates && cell.XF) {
					const fmtKind = typeof nfmt === "string" ? getDateTimeFormatKind(nfmt) : "none";
					if (fmtKind !== "none" && fmtKind !== "time" && typeof cell.v === "number") {
						cell.t = "d";
						cell.v = serialNumberToDate(cell.v, date1904);
					}
				}
			}

			// Store cell in dense or sparse mode
			if (dense) {
				if (!s["!data"]![R]) {
					s["!data"]![R] = [];
				}
				s["!data"]![R][C] = cell;
			} else {
				s[ref] = cell;
			}
		}
	}

	for (const pending of pendingSharedFormulas) {
		const master = sharedFormulas.get(pending.si);
		if (!master) {
			continue;
		}
		const origin = decodeCell(master.origin);
		const target = decodeCell(pending.ref);
		pending.cell.f = shiftFormulaStr(master.formula, {
			r: target.r - origin.r,
			c: target.c - origin.c,
		});
	}
}

/**
 * Resolve shared string references in a worksheet by replacing SST index
 * placeholders with actual string values from the shared string table.
 *
 * @param s - Worksheet whose cells may contain _sstIdx placeholders
 * @param sst - Parsed Shared String Table
 * @param opts - Options controlling HTML output (cellHTML)
 */
export function resolveSharedStrings(s: WorkSheet, sst: SST, opts: any): void {
	const dense = s["!data"] != null;
	if (dense) {
		const data = s["!data"]!;
		for (let R = 0; R < data.length; ++R) {
			if (!data[R]) {
				continue;
			}
			for (let C = 0; C < data[R]!.length; ++C) {
				const cell = data[R]![C];
				if (!cell || (cell as any)._sstIdx === undefined) {
					continue;
				}
				const idx = (cell as any)._sstIdx;
				delete (cell as any)._sstIdx;
				if (sst[idx]) {
					cell.v = sst[idx].t;
					if (opts.cellHTML !== false && sst[idx].h) {
						cell.h = sst[idx].h;
					}
					if (sst[idx].r) {
						cell.r = sst[idx].r;
					}
				}
			}
		}
	} else {
		for (const ref of Object.keys(s)) {
			// Skip special keys (e.g. !ref, !merges, !data)
			if (ref.charAt(0) === "!") {
				continue;
			}
			const cell = s[ref] as CellObject;
			if (!cell || (cell as any)._sstIdx === undefined) {
				continue;
			}
			const idx = (cell as any)._sstIdx;
			delete (cell as any)._sstIdx;
			if (sst[idx]) {
				cell.v = sst[idx].t;
				if (opts.cellHTML !== false && sst[idx].h) {
					cell.h = sst[idx].h;
				}
				if (sst[idx].r) {
					cell.r = sst[idx].r;
				}
			}
		}
	}
}

/**
 * Parse a worksheet XML file into a WorkSheet object.
 *
 * Extracts dimensions, columns, cell data, merges, hyperlinks, autofilter,
 * margins, and legacy drawing references from the sheet XML.
 *
 * @param data - Raw XML string of the sheet file (e.g. sheet1.xml)
 * @param opts - Parsing options (dense, sheetRows, cellHTML, cellDates, etc.)
 * @param _idx - Sheet index (unused, reserved)
 * @param rels - Relationships for resolving hyperlink targets
 * @param wb - Parsed workbook properties (for date1904 flag)
 * @param _themes - Theme data (reserved for theme color resolution)
 * @param styles - Parsed styles data for number format resolution
 * @returns Parsed WorkSheet object
 */
export function parseWorksheetXml(
	data: string,
	opts?: any,
	_idx?: number,
	rels?: Relationships,
	wb?: any,
	_themes?: any,
	styles?: StylesData,
): WorkSheet {
	if (!data) {
		return {};
	}
	if (!opts) {
		opts = {};
	}
	assertXmlPartLimits("worksheet.xml", data, opts);
	if (!rels) {
		rels = { "!id": {} };
	}

	const s: WorkSheet = opts.dense ? { "!data": [] } : {};
	// Start with an inverted range that will be narrowed as cells are found
	const refguess: Range = { s: { r: 2000000, c: 2000000 }, e: { r: 0, c: 0 } };

	// Split the XML at the <sheetData> section for efficient parsing
	let data1 = "";
	let data2 = "";
	const sdMatch = data.match(/<(?:\w+:)?sheetData[^>]*>([\s\S]*?)<\/(?:\w+:)?sheetData>/);
	if (sdMatch) {
		data1 = data.slice(0, sdMatch.index); // Content before sheetData
		data2 = data.slice(sdMatch.index! + sdMatch[0].length); // Content after sheetData
	} else {
		data1 = data2 = data;
	}

	// Dimension
	const ridx = (data1.match(/<(?:\w*:)?dimension/) || ({ index: -1 } as any)).index;
	if (ridx > 0) {
		const ref = data1.slice(ridx, ridx + 50).match(dimregex);
		if (ref && !opts.nodim) {
			parseWorksheetXml_dim(s, ref[1]);
		}
	}

	// Columns
	const columns: ColInfo[] = [];
	if (opts.cellStyles) {
		const cols = data1.match(colregex);
		if (cols) {
			parseWorksheetXml_cols(columns, cols, opts);
		}
		const views = parseWorksheetXml_views(data1, opts);
		if (views) {
			s["!views"] = views;
		}
	}

	// SheetData (cells)
	if (sdMatch) {
		parseSheetData(sdMatch[1], s, opts, refguess, _themes, styles, wb);
	}

	// AutoFilter
	const afilter = data2.match(afregex);
	if (afilter) {
		s["!autofilter"] = parseWorksheetXml_autofilter(afilter[0], opts);
	}

	// Merged cells
	const merges: Range[] = [];
	const _merge = data2.match(mergecregex);
	if (_merge) {
		for (let i = 0; i < _merge.length; ++i) {
			// Extract the ref attribute value after the '=' and opening quote
			merges[i] = safeDecodeRange(_merge[i].slice(_merge[i].indexOf("=") + 2));
		}
	}

	// Hyperlinks
	const hlink = data2.match(hlinkregex);
	if (hlink) {
		parseWorksheetXml_hlinks(s, hlink, rels, opts);
	}

	// Page margins
	const margins = data2.match(marginregex);
	if (margins) {
		s["!margins"] = parseWorksheetXml_margins(parseXmlTag(margins[0], undefined, undefined, opts));
	}

	// Legacy drawing reference (for VML comment shapes)
	const legm = data2.match(/legacyDrawing r:id="(.*?)"/);
	if (legm) {
		(s as any)["!legrel"] = legm[1];
	}

	// If nodim mode, start range from (0,0)
	if (opts.nodim) {
		refguess.s.c = refguess.s.r = 0;
	}
	// Set !ref from the guessed range if not already set by <dimension>
	if (!s["!ref"] && refguess.e.c >= refguess.s.c && refguess.e.r >= refguess.s.r) {
		s["!ref"] = encodeRange(refguess);
	}
	// Clamp to sheetRows limit
	if (opts.sheetRows > 0 && s["!ref"]) {
		const tmpref = safeDecodeRange(s["!ref"]);
		if (opts.sheetRows <= tmpref.e.r) {
			tmpref.e.r = opts.sheetRows - 1;
			if (tmpref.e.r > refguess.e.r) {
				tmpref.e.r = refguess.e.r;
			}
			if (tmpref.e.r < tmpref.s.r) {
				tmpref.s.r = tmpref.e.r;
			}
			if (tmpref.e.c > refguess.e.c) {
				tmpref.e.c = refguess.e.c;
			}
			if (tmpref.e.c < tmpref.s.c) {
				tmpref.s.c = tmpref.e.c;
			}
			// Save original full range before clamping
			(s as any)["!fullref"] = s["!ref"];
			s["!ref"] = encodeRange(tmpref);
		}
	}
	if (columns.length > 0) {
		s["!cols"] = columns;
	}
	if (merges.length > 0) {
		s["!merges"] = merges;
	}

	return s;
}

/** Generate <mergeCells> XML from an array of merge ranges */
function writeWorksheetXml_merges(merges: Range[]): string {
	if (merges.length === 0) {
		return "";
	}
	const lines = ['<mergeCells count="' + merges.length + '">'];
	for (let i = 0; i < merges.length; ++i) {
		lines.push('<mergeCell ref="' + encodeRange(merges[i]) + '"/>');
	}
	lines.push("</mergeCells>");
	return lines.join("");
}

function getFrozenPaneTopLeftCell(view: SheetView): string | undefined {
	if (view.topLeftCell) {
		return view.topLeftCell;
	}
	const xSplit = view.xSplit || 0;
	const ySplit = view.ySplit || 0;
	if (xSplit <= 0 && ySplit <= 0) {
		return undefined;
	}
	return encodeCell({ r: ySplit, c: xSplit });
}

function getFrozenPaneActivePane(view: SheetView): "topRight" | "bottomLeft" | "bottomRight" {
	if (view.activePane) {
		return view.activePane;
	}
	if (view.xSplit && view.ySplit) {
		return "bottomRight";
	}
	return view.xSplit ? "topRight" : "bottomLeft";
}

function writeWorksheetXml_sheetViews(ws: WorkSheet, idx: number): string {
	const view = ws["!views"]?.find(
		(item) => item?.state === "frozen" && ((item.xSplit || 0) > 0 || (item.ySplit || 0) > 0),
	);
	const attrs = idx === 0 ? ' tabSelected="1"' : "";
	if (!view) {
		return '<sheetViews><sheetView workbookViewId="0"' + attrs + "/></sheetViews>";
	}
	const paneAttrs: Record<string, string> = { state: "frozen", activePane: getFrozenPaneActivePane(view) };
	if (view.xSplit) {
		paneAttrs.xSplit = String(view.xSplit);
	}
	if (view.ySplit) {
		paneAttrs.ySplit = String(view.ySplit);
	}
	const topLeftCell = getFrozenPaneTopLeftCell(view);
	if (topLeftCell) {
		paneAttrs.topLeftCell = topLeftCell;
	}
	const pane = writeXmlElement("pane", null, paneAttrs);
	const selection = writeXmlElement("selection", null, { pane: paneAttrs.activePane });
	return '<sheetViews><sheetView workbookViewId="0"' + attrs + ">" + pane + selection + "</sheetView></sheetViews>";
}

/** Return reusable rich-text XML only when it still represents the cell's current value. */
function getCellRichText(cell: CellObject): string | undefined {
	const text = String(cell.v);
	if (typeof cell.r !== "string" || cell.r.length === 0) {
		return undefined;
	}
	try {
		if (parseStringItem(cell.r, { cellHTML: false }).t === text) {
			return normalizeRichTextXml(cell.r);
		}
	} catch {}
	return undefined;
}

/** Add a string cell to the workbook-wide shared string table and return its index. */
function addSharedString(
	cell: CellObject,
	richText: string | undefined,
	opts: { Strings: SST; revStrings: Map<string, number> },
): number {
	const text = String(cell.v);
	const key = richText === undefined ? "t\0" + text : "r\0" + richText;
	const existing = opts.revStrings.get(key);
	opts.Strings.Count = (opts.Strings.Count || 0) + 1;
	if (existing !== undefined) {
		return existing;
	}

	const entry: XLString = { t: text };
	if (richText !== undefined) {
		entry.r = richText;
	}
	const index = opts.Strings.length;
	opts.Strings.push(entry);
	opts.Strings.Unique = (opts.Strings.Unique || 0) + 1;
	opts.revStrings.set(key, index);
	return index;
}

/**
 * Write a worksheet as XML.
 *
 * Serializes cell data, row properties, column definitions, merged cells,
 * autofilter, and page margins into a complete sheet XML document.
 *
 * @param ws - WorkSheet to serialize
 * @param opts - Write options (cellDates, etc.)
 * @param _idx - Sheet index (0-based), used to set tabSelected on the first sheet
 * @param _rels - Relationships object (reserved for hyperlink writing)
 * @param _wb - WorkBook reference (reserved)
 * @returns Complete worksheet XML string
 */
export function writeWorksheetXml(ws: WorkSheet, opts: any, _idx: number, _rels: Relationships, _wb: any): string {
	const lines: string[] = [XML_HEADER];
	lines.push(
		writeXmlElement("worksheet", null, {
			xmlns: XMLNS_main[0],
			"xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
		}),
	);

	const ref = ws["!ref"] || "A1";
	lines.push('<dimension ref="' + ref + '"/>');

	lines.push(writeWorksheetXml_sheetViews(ws, _idx));

	lines.push('<sheetFormatPr defaultRowHeight="15"/>');

	// Column definitions
	if (ws["!cols"]) {
		lines.push("<cols>");
		for (let i = 0; i < ws["!cols"].length; ++i) {
			if (!ws["!cols"][i]) {
				continue;
			}
			const col = ws["!cols"][i];
			const attrs: Record<string, string> = {
				min: String(i + 1), // 1-based
				max: String(i + 1),
			};
			if (col.width) {
				attrs.width = String(col.width);
			} else {
				attrs.width = "9.140625"; // Excel default column width
			}
			if (col.hidden) {
				attrs.hidden = "1";
			}
			attrs.customWidth = "1";
			lines.push(writeXmlElement("col", null, attrs));
		}
		lines.push("</cols>");
	}

	lines.push("<sheetData>");

	const dense = ws["!data"] != null;
	const range = safeDecodeRange(ref);
	const appendRow = (rowIdx: number, rowCells: string[]): void => {
		const rowInfo = ws["!rows"]?.[rowIdx];
		const hasFiniteHeight = rowInfo != null && Number.isFinite(rowInfo.hpt);
		const hasMetadata = hasFiniteHeight || rowInfo?.hidden === true;
		if (rowCells.length === 0 && !hasMetadata) {
			return;
		}
		let rowTag = '<row r="' + (rowIdx + 1) + '"'; // 1-based row number
		if (rowInfo) {
			if (hasFiniteHeight) {
				rowTag += ' ht="' + rowInfo.hpt + '" customHeight="1"';
			}
			if (rowInfo.hidden === true) {
				rowTag += ' hidden="1"';
			}
		}
		rowTag += ">";
		lines.push(rowTag);
		lines.push(rowCells.join(""));
		lines.push("</row>");
	};
	const extraRows = Object.keys(ws["!rows"] ?? {})
		.map(Number)
		.filter((rowIdx) => Number.isSafeInteger(rowIdx) && rowIdx >= 0 && (rowIdx < range.s.r || rowIdx > range.e.r))
		.sort((a, b) => a - b);
	for (const rowIdx of extraRows) {
		if (rowIdx < range.s.r) {
			appendRow(rowIdx, []);
		}
	}

	for (let rowIdx = range.s.r; rowIdx <= range.e.r; ++rowIdx) {
		const row_cells: string[] = [];
		for (let colIdx = range.s.c; colIdx <= range.e.c; ++colIdx) {
			let cell: CellObject | undefined;
			if (dense) {
				cell = ws["!data"]?.[rowIdx]?.[colIdx];
			} else {
				const addr = encodeCell({ r: rowIdx, c: colIdx });
				cell = ws[addr] as CellObject | undefined;
			}
			const styleIndex = cell ? getCellStyleIndex(opts, cell) : undefined;
			// Skip empty and unstyled "z" (stub) cells
			if (!cell || (cell.t === "z" && styleIndex == null)) {
				continue;
			}

			const addr = encodeCell({ r: rowIdx, c: colIdx });
			let cellValueStr = "";
			let cellTypeAttr = "";
			let cellValueTag = "v";
			const richText = cell.t === "s" ? getCellRichText(cell) : undefined;

			switch (cell.t) {
				case "b":
					cellValueStr = cell.v ? "1" : "0";
					cellTypeAttr = "b";
					break;
				case "n":
					if (cell.v != null) {
						cellValueStr = String(cell.v);
					}
					break;
				case "e":
					cellValueStr = escapeXml(serializeWorksheetXml_error(cell.v));
					cellTypeAttr = "e";
					break;
				case "d":
					if (opts.cellDates) {
						cellValueStr = (cell.v as Date).toISOString();
						cellTypeAttr = "d";
					} else {
						// Convert date to serial number for non-cellDates mode
						cellValueStr = String(dateToSerialNumber(cell.v as Date, _wb?.Workbook?.WBProps?.date1904));
					}
					break;
				case "s":
					if (!cell.f && opts.bookSST) {
						cellValueStr = String(addSharedString(cell, richText, opts));
						cellTypeAttr = "s";
					} else if (!cell.f && richText !== undefined) {
						cellValueStr = richText;
						cellTypeAttr = "inlineStr";
						cellValueTag = "is";
					} else {
						cellValueStr = escapeXml(String(cell.v));
						cellTypeAttr = "str"; // Formula result or plain inline value
					}
					break;
			}

			let cellXml = '<c r="' + addr + '"';
			if (cellTypeAttr) {
				cellXml += ' t="' + cellTypeAttr + '"';
			}
			if (styleIndex != null && styleIndex > 0) {
				cellXml += ' s="' + styleIndex + '"';
			}
			if (cell.D) {
				cellXml += ' cm="1"';
			}
			if (!cell.f && cellValueStr === "") {
				cellXml += "/>";
				row_cells.push(cellXml);
				continue;
			}
			cellXml += ">";
			if (cell.f) {
				cellXml += "<f";
				if (cell.F) {
					// Array formula with reference range
					cellXml += ' ref="' + cell.F + '" t="array"';
				}
				cellXml += ">" + escapeXml(cell.f) + "</f>";
			}
			if (cellValueStr !== "") {
				cellXml += "<" + cellValueTag + ">" + cellValueStr + "</" + cellValueTag + ">";
			}
			cellXml += "</c>";
			row_cells.push(cellXml);
		}
		appendRow(rowIdx, row_cells);
	}

	// A worksheet's data range does not necessarily include rows with only layout metadata.
	// Iterate the existing sparse metadata keys so a distant row cannot cause a huge scan.
	for (const rowIdx of extraRows) {
		if (rowIdx > range.e.r) {
			appendRow(rowIdx, []);
		}
	}

	lines.push("</sheetData>");

	// Merged cells
	if (ws["!merges"] && ws["!merges"].length > 0) {
		lines.push(writeWorksheetXml_merges(ws["!merges"]));
	}

	// AutoFilter
	if (ws["!autofilter"]) {
		lines.push('<autoFilter ref="' + ws["!autofilter"].ref + '"/>');
	}

	// Page margins
	if (ws["!margins"]) {
		const margins = ws["!margins"];
		lines.push(
			writeXmlElement("pageMargins", null, {
				left: String(margins.left ?? 0.7),
				right: String(margins.right ?? 0.7),
				top: String(margins.top ?? 0.75),
				bottom: String(margins.bottom ?? 0.75),
				header: String(margins.header ?? 0.3),
				footer: String(margins.footer ?? 0.3),
			}),
		);
	}

	const legacyDrawing = Object.values(_rels?.["!id"] || {}).find((relationship) => relationship.Type === RELTYPE.VML);
	if (legacyDrawing) {
		lines.push(writeXmlElement("legacyDrawing", null, { "r:id": legacyDrawing.Id }));
	}

	lines.push("</worksheet>");
	// Convert self-closing <worksheet .../> to opening tag <worksheet ...>
	lines[1] = lines[1].replace("/>", ">");
	return lines.join("");
}
