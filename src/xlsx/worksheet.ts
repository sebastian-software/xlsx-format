import type { WorkSheet, CellObject, Range, ColInfo, RowInfo, MarginInfo } from "../types.js";
import { parseXmlTag, XML_HEADER } from "../xml/parser.js";
import { unescapeXml, escapeXml } from "../xml/escape.js";
import { writeXmlElement } from "../xml/writer.js";
import { XMLNS_main } from "../xml/namespaces.js";
import { safeDecodeRange, encodeRange, encodeCell } from "../utils/cell.js";
import { formatNumber, isDateFormat } from "../ssf/format.js";
import { formatTable } from "../ssf/table.js";
import { dateToSerialNumber, serialNumberToDate } from "../utils/date.js";
import type { SST } from "./shared-strings.js";
import type { StylesData } from "./styles.js";
import type { Relationships } from "../opc/relationships.js";

/** Regex patterns for extracting various worksheet XML elements */
const mergecregex = /<(?:\w+:)?mergeCell ref=["'][A-Z0-9:]+['"]\s*[/]?>/g;
const hlinkregex = /<(?:\w+:)?hyperlink [^<>]*>/gm;
const dimregex = /"(\w*:\w*)"/;
const colregex = /<(?:\w+:)?col\b[^<>]*[/]?>/g;
const afregex = /<(?:\w:)?autoFilter[^>]*([/]|>([\s\S]*)<\/(?:\w:)?autoFilter)>/g;
const marginregex = /<(?:\w+:)?pageMargins[^<>]*\/>/g;

/** Parse the <dimension> element to set the sheet reference range */
function parseWorksheetXml_dim(ws: WorkSheet, s: string): void {
	const d = safeDecodeRange(s);
	if (d.s.r <= d.e.r && d.s.c <= d.e.c && d.s.r >= 0 && d.s.c >= 0) {
		ws["!ref"] = encodeRange(d);
	}
}

/** Parse <pageMargins> attributes with defaults matching Excel's standard margins */
function parseWorksheetXml_margins(tag: Record<string, any>): MarginInfo {
	return {
		left: parseFloat(tag.left) || 0.7,
		right: parseFloat(tag.right) || 0.7,
		top: parseFloat(tag.top) || 0.75,
		bottom: parseFloat(tag.bottom) || 0.75,
		header: parseFloat(tag.header) || 0.3,
		footer: parseFloat(tag.footer) || 0.3,
	};
}

/** Parse <autoFilter> element extracting the filter reference range */
function parseWorksheetXml_autofilter(data: string): { ref: string } {
	const tag = parseXmlTag(data.match(/<[^>]*>/)?.[0] || "");
	return { ref: tag.ref || "" };
}

/** Parse <col> elements to populate column width and hidden state */
function parseWorksheetXml_cols(columns: ColInfo[], cols: string[]): void {
	for (let i = 0; i < cols.length; ++i) {
		const tag = parseXmlTag(cols[i]);
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
				columns[j] = {} as ColInfo;
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

/** Parse <hyperlink> elements and attach link objects to the corresponding cells */
function parseWorksheetXml_hlinks(s: WorkSheet, hlinks: string[], rels: Relationships): void {
	for (let i = 0; i < hlinks.length; ++i) {
		const tag = parseXmlTag(hlinks[i]);
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
const cellregex = /<(?:\w+:)?c\b[^>]*(?:\/>|>([\s\S]*?)<\/(?:\w+:)?c>)/g;

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
		const rowTag = parseXmlTag(rowTagMatch[0]);
		// Row numbers in XML are 1-based
		const R = parseInt(rowTag.r, 10) - 1;
		if (isNaN(R)) {
			continue;
		}

		// Extract row properties (height, hidden)
		if (rowTag.ht || rowTag.hidden) {
			if (!s["!rows"]) {
				s["!rows"] = [];
			}
			if (!s["!rows"][R]) {
				s["!rows"][R] = {} as RowInfo;
			}
			if (rowTag.ht) {
				s["!rows"][R].hpt = parseFloat(rowTag.ht);
			}
			if (rowTag.hidden === "1") {
				s["!rows"][R].hidden = true;
			}
		}

		// Skip rows beyond the sheetRows limit
		if (opts.sheetRows && R >= opts.sheetRows) {
			continue;
		}

		// Parse cells in this row
		cellregex.lastIndex = 0;
		let cellMatch;
		while ((cellMatch = cellregex.exec(rowStr))) {
			const cellTag = parseXmlTag(cellMatch[0].match(/<(?:\w+:)?c\b[^>]*/)?.[0] + ">" || "");
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
			const fMatch = cellValue.match(/<(?:\w+:)?f[^>]*>([\s\S]*?)<\/(?:\w+:)?f>/);
			const isMatch = cellValue.match(/<(?:\w+:)?is>([\s\S]*?)<\/(?:\w+:)?is>/);

			const v = vMatch ? vMatch[1] : null;

			switch (cellType) {
				case "s": // shared string - value is an index into the SST
					if (v !== null) {
						const idx = parseInt(v, 10);
						cell = { t: "s", v: "" } as CellObject;
						// Store SST index for later resolution via resolveSharedStrings()
						(cell as any)._sstIdx = idx;
					} else {
						cell = { t: "z" } as CellObject;
					}
					break;
				case "str": // formula result that is a string
					cell = { t: "s", v: v ? unescapeXml(v) : "" } as CellObject;
					break;
				case "inlineStr":
					if (isMatch) {
						const tMatch = isMatch[1].match(/<(?:\w+:)?t[^>]*>([\s\S]*?)<\/(?:\w+:)?t>/);
						cell = { t: "s", v: tMatch ? unescapeXml(tMatch[1]) : "" } as CellObject;
					} else {
						cell = { t: "s", v: "" } as CellObject;
					}
					break;
				case "b": // boolean
					cell = { t: "b", v: v === "1" } as CellObject;
					break;
				case "e": // error
					cell = { t: "e", v: v ? parseInt(v, 10) || 0 : 0 } as CellObject;
					(cell as any).w = v || "";
					break;
				case "d": // ISO 8601 date string
					if (v) {
						cell = { t: "d", v: new Date(v) } as CellObject;
					} else {
						cell = { t: "z" } as CellObject;
					}
					break;
				default: // "n" (number) or unrecognized
					if (v !== null) {
						cell = { t: "n", v: parseFloat(v) } as CellObject;
					} else {
						if (!opts.sheetStubs) {
							continue;
						}
						cell = { t: "z" } as CellObject;
					}
					break;
			}

			// Apply style reference from the cellXf table
			if (cellStyle > 0 && styles) {
				const xf = styles.CellXf[cellStyle];
				if (xf) {
					cell.XF = { numFmtId: xf.numFmtId };
					if (opts.cellNF) {
						const nf = styles.NumberFmt[xf.numFmtId] || formatTable[xf.numFmtId];
						if (nf) {
							cell.z = nf;
						}
					}
				}
			}

			// Extract formula
			if (fMatch && opts.cellFormula !== false) {
				cell.f = unescapeXml(fMatch[1]);
				const fTag = parseXmlTag(cellValue.match(/<(?:\w+:)?f[^>]*/)?.[0] + ">" || "");
				if (fTag.t === "shared" && fTag.si != null) {
					// Shared formula (master or reference)
				}
				if (fTag.t === "array" && fTag.ref) {
					cell.F = fTag.ref; // Array formula range
					cell.D = fTag.dt === "1"; // Dynamic array flag
				}
			}

			// Format the cell value as display text
			if (opts.cellText !== false) {
				if (cell.t === "n") {
					const nfmt =
						cell.z ||
						(cell.XF && cell.XF.numFmtId != null && styles?.NumberFmt[cell.XF.numFmtId]) ||
						formatTable[(cell.XF && cell.XF.numFmtId) || 0];
					if (nfmt) {
						try {
							cell.w = formatNumber(nfmt, cell.v as number, { date1904 });
						} catch {}
					}
					// Convert numeric cells with date formats to Date objects if cellDates is enabled
					if (opts.cellDates && cell.XF) {
						const fmtStr = nfmt || formatTable[cell.XF.numFmtId || 0] || "";
						if (typeof fmtStr === "string" && isDateFormat(fmtStr) && typeof cell.v === "number") {
							cell.t = "d";
							cell.v = serialNumberToDate(cell.v, date1904);
						}
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
		return {} as WorkSheet;
	}
	if (!opts) {
		opts = {};
	}
	if (!rels) {
		rels = { "!id": {} } as any;
	}

	const s: WorkSheet = opts.dense ? { "!data": [] } : ({} as any);
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
			parseWorksheetXml_cols(columns, cols);
		}
	}

	// SheetData (cells)
	if (sdMatch) {
		parseSheetData(sdMatch[1], s, opts, refguess, _themes, styles, wb);
	}

	// AutoFilter
	const afilter = data2.match(afregex);
	if (afilter) {
		s["!autofilter"] = parseWorksheetXml_autofilter(afilter[0]);
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
		parseWorksheetXml_hlinks(s, hlink, rels!);
	}

	// Page margins
	const margins = data2.match(marginregex);
	if (margins) {
		s["!margins"] = parseWorksheetXml_margins(parseXmlTag(margins[0]));
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

	lines.push('<sheetViews><sheetView workbookViewId="0"');
	// Mark the first sheet as the active/selected tab
	if (_idx === 0) {
		lines.push(' tabSelected="1"');
	}
	lines.push("/></sheetViews>");

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
			// Skip empty and "z" (stub) cells
			if (!cell || cell.t === "z") {
				continue;
			}

			const addr = encodeCell({ r: rowIdx, c: colIdx });
			let cellValueStr = "";
			let cellTypeAttr = "";

			switch (cell.t) {
				case "b":
					cellValueStr = cell.v ? "1" : "0";
					cellTypeAttr = "b";
					break;
				case "n":
					cellValueStr = String(cell.v);
					break;
				case "e":
					cellValueStr = String(cell.v);
					cellTypeAttr = "e";
					break;
				case "d":
					if (opts.cellDates) {
						cellValueStr = (cell.v as Date).toISOString();
						cellTypeAttr = "d";
					} else {
						// Convert date to serial number for non-cellDates mode
						cellValueStr = String(dateToSerialNumber(cell.v as Date));
					}
					break;
				case "s":
					cellValueStr = escapeXml(String(cell.v));
					cellTypeAttr = "str"; // Inline string (not shared)
					break;
			}

			let cellXml = '<c r="' + addr + '"';
			if (cellTypeAttr) {
				cellXml += ' t="' + cellTypeAttr + '"';
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
				cellXml += "<v>" + cellValueStr + "</v>";
			}
			cellXml += "</c>";
			row_cells.push(cellXml);
		}
		// Only emit rows that contain at least one cell
		if (row_cells.length > 0) {
			let rowTag = '<row r="' + (rowIdx + 1) + '"'; // 1-based row number
			if (ws["!rows"]?.[rowIdx]) {
				if (ws["!rows"][rowIdx].hpt) {
					rowTag += ' ht="' + ws["!rows"][rowIdx].hpt + '" customHeight="1"';
				}
				if (ws["!rows"][rowIdx].hidden) {
					rowTag += ' hidden="1"';
				}
			}
			rowTag += ">";
			lines.push(rowTag);
			lines.push(row_cells.join(""));
			lines.push("</row>");
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
				left: String(margins.left || 0.7),
				right: String(margins.right || 0.7),
				top: String(margins.top || 0.75),
				bottom: String(margins.bottom || 0.75),
				header: String(margins.header || 0.3),
				footer: String(margins.footer || 0.3),
			}),
		);
	}

	lines.push("</worksheet>");
	// Convert self-closing <worksheet .../> to opening tag <worksheet ...>
	lines[1] = lines[1].replace("/>", ">");
	return lines.join("");
}
