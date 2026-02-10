import type { WorkSheet, CellObject, Range, ColInfo, RowInfo, MarginInfo } from "../types.js";
import { BErr } from "../types.js";
import { parsexmltag, tagregex, XML_HEADER, parsexmlbool } from "../xml/parser.js";
import { unescapexml, escapexml } from "../xml/escape.js";
import { writextag } from "../xml/writer.js";
import { XMLNS_main } from "../xml/namespaces.js";
import { safe_decode_range, encode_range, encode_cell, encode_col } from "../utils/cell.js";
import { utf8read } from "../utils/buffer.js";
import { SSF_format, fmt_is_date } from "../ssf/format.js";
import { table_fmt } from "../ssf/table.js";
import { datenum, numdate } from "../utils/date.js";
import type { SST } from "./shared-strings.js";
import type { StylesData } from "./styles.js";
import type { Relationships } from "../opc/relationships.js";

const mergecregex = /<(?:\w+:)?mergeCell ref=["'][A-Z0-9:]+['"]\s*[/]?>/g;
const hlinkregex = /<(?:\w+:)?hyperlink [^<>]*>/gm;
const dimregex = /"(\w*:\w*)"/;
const colregex = /<(?:\w+:)?col\b[^<>]*[/]?>/g;
const afregex = /<(?:\w:)?autoFilter[^>]*([/]|>([\s\S]*)<\/(?:\w:)?autoFilter)>/g;
const marginregex = /<(?:\w+:)?pageMargins[^<>]*\/>/g;

function parse_ws_xml_dim(ws: WorkSheet, s: string): void {
	const d = safe_decode_range(s);
	if (d.s.r <= d.e.r && d.s.c <= d.e.c && d.s.r >= 0 && d.s.c >= 0) {
		ws["!ref"] = encode_range(d);
	}
}

function parse_ws_xml_margins(tag: Record<string, any>): MarginInfo {
	return {
		left: parseFloat(tag.left) || 0.7,
		right: parseFloat(tag.right) || 0.7,
		top: parseFloat(tag.top) || 0.75,
		bottom: parseFloat(tag.bottom) || 0.75,
		header: parseFloat(tag.header) || 0.3,
		footer: parseFloat(tag.footer) || 0.3,
	};
}

function parse_ws_xml_autofilter(data: string): { ref: string } {
	const tag = parsexmltag(data.match(/<[^>]*>/)?.[0] || "");
	return { ref: tag.ref || "" };
}

function parse_ws_xml_cols(columns: ColInfo[], cols: string[]): void {
	for (let i = 0; i < cols.length; ++i) {
		const tag = parsexmltag(cols[i]);
		if (!tag.min || !tag.max) {
			continue;
		}
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

function parse_ws_xml_hlinks(s: WorkSheet, hlinks: string[], rels: Relationships): void {
	for (let i = 0; i < hlinks.length; ++i) {
		const tag = parsexmltag(hlinks[i]);
		if (!tag.ref) {
			continue;
		}
		const rng = safe_decode_range(tag.ref);
		for (let R = rng.s.r; R <= rng.e.r; ++R) {
			for (let C = rng.s.c; C <= rng.e.c; ++C) {
				const addr = encode_cell({ r: R, c: C });
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

// Row tag regex
const rowregex = /<(?:\w+:)?row\b[^>]*>/g;
const cellregex = /<(?:\w+:)?c\b[^>]*(?:\/>|>([\s\S]*?)<\/(?:\w+:)?c>)/g;

/** Parse sheetData XML into worksheet */
function parse_ws_xml_data(
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

	// Parse row by row
	const rowMatches = sdata.match(rowregex) || [];
	// Split by row boundaries
	const rows = sdata.split(/<\/(?:\w+:)?row>/);

	for (let ri = 0; ri < rows.length; ++ri) {
		const rowStr = rows[ri];
		if (!rowStr) {
			continue;
		}

		// Find row tag
		const rowTagMatch = rowStr.match(/<(?:\w+:)?row\b[^>]*>/);
		if (!rowTagMatch) {
			continue;
		}
		const rowTag = parsexmltag(rowTagMatch[0]);
		const R = parseInt(rowTag.r, 10) - 1;
		if (isNaN(R)) {
			continue;
		}

		// Row properties
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

		if (opts.sheetRows && R >= opts.sheetRows) {
			continue;
		}

		// Parse cells in this row
		cellregex.lastIndex = 0;
		let cellMatch;
		while ((cellMatch = cellregex.exec(rowStr))) {
			const cellTag = parsexmltag(cellMatch[0].match(/<(?:\w+:)?c\b[^>]*/)?.[0] + ">" || "");
			const ref = cellTag.r;
			if (!ref) {
				continue;
			}

			// Decode cell address
			let C = 0;
			for (let ci = 0; ci < ref.length; ++ci) {
				const cc = ref.charCodeAt(ci);
				if (cc >= 65 && cc <= 90) {
					C = 26 * C + (cc - 64);
				} else {
					break;
				}
			}
			C -= 1;

			// Update refguess
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

			const cellType = cellTag.t || "n";
			const cellStyle = cellTag.s ? parseInt(cellTag.s, 10) : 0;
			const cellValue = cellMatch[1] || "";

			let cell: CellObject;

			// Extract value from <v> tag
			const vMatch = cellValue.match(/<(?:\w+:)?v>([\s\S]*?)<\/(?:\w+:)?v>/);
			const fMatch = cellValue.match(/<(?:\w+:)?f[^>]*>([\s\S]*?)<\/(?:\w+:)?f>/);
			const isMatch = cellValue.match(/<(?:\w+:)?is>([\s\S]*?)<\/(?:\w+:)?is>/);

			const v = vMatch ? vMatch[1] : null;

			switch (cellType) {
				case "s": // shared string
					if (v !== null) {
						const idx = parseInt(v, 10);
						cell = { t: "s", v: "" } as CellObject;
						// Will be resolved later with SST
						(cell as any)._sstIdx = idx;
					} else {
						cell = { t: "z" } as CellObject;
					}
					break;
				case "str": // formula string
					cell = { t: "s", v: v ? unescapexml(v) : "" } as CellObject;
					break;
				case "inlineStr":
					if (isMatch) {
						const tMatch = isMatch[1].match(/<(?:\w+:)?t[^>]*>([\s\S]*?)<\/(?:\w+:)?t>/);
						cell = { t: "s", v: tMatch ? unescapexml(tMatch[1]) : "" } as CellObject;
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
				case "d": // date
					if (v) {
						cell = { t: "d", v: new Date(v) } as CellObject;
					} else {
						cell = { t: "z" } as CellObject;
					}
					break;
				default: // 'n' number
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

			// Style reference
			if (cellStyle > 0 && styles) {
				const xf = styles.CellXf[cellStyle];
				if (xf) {
					cell.XF = { numFmtId: xf.numFmtId };
					if (opts.cellNF) {
						const nf = styles.NumberFmt[xf.numFmtId] || table_fmt[xf.numFmtId];
						if (nf) {
							cell.z = nf;
						}
					}
				}
			}

			// Formula
			if (fMatch && opts.cellFormula !== false) {
				cell.f = unescapexml(fMatch[1]);
				const fTag = parsexmltag(cellValue.match(/<(?:\w+:)?f[^>]*/)?.[0] + ">" || "");
				if (fTag.t === "shared" && fTag.si != null) {
					// Shared formula
				}
				if (fTag.t === "array" && fTag.ref) {
					cell.F = fTag.ref;
					cell.D = fTag.dt === "1";
				}
			}

			// Number formatting / cell text
			if (opts.cellText !== false) {
				if (cell.t === "n") {
					const nfmt =
						cell.z ||
						(cell.XF && cell.XF.numFmtId != null && styles?.NumberFmt[cell.XF.numFmtId]) ||
						table_fmt[(cell.XF && cell.XF.numFmtId) || 0];
					if (nfmt) {
						try {
							cell.w = SSF_format(nfmt, cell.v as number, { date1904 });
						} catch {}
					}
					// Handle dates
					if (opts.cellDates && cell.XF) {
						const fmtStr = nfmt || table_fmt[cell.XF.numFmtId || 0] || "";
						if (typeof fmtStr === "string" && fmt_is_date(fmtStr) && typeof cell.v === "number") {
							cell.t = "d";
							cell.v = numdate(cell.v, date1904);
						}
					}
				}
			}

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

/** Resolve SST references in a worksheet */
export function resolve_sst(s: WorkSheet, sst: SST, opts: any): void {
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

/** Parse a worksheet XML */
export function parse_ws_xml(
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
	const refguess: Range = { s: { r: 2000000, c: 2000000 }, e: { r: 0, c: 0 } };

	// Split at sheetData
	let data1 = "";
	let data2 = "";
	const sdMatch = data.match(/<(?:\w+:)?sheetData[^>]*>([\s\S]*?)<\/(?:\w+:)?sheetData>/);
	if (sdMatch) {
		data1 = data.slice(0, sdMatch.index);
		data2 = data.slice(sdMatch.index! + sdMatch[0].length);
	} else {
		data1 = data2 = data;
	}

	// Dimension
	const ridx = (data1.match(/<(?:\w*:)?dimension/) || ({ index: -1 } as any)).index;
	if (ridx > 0) {
		const ref = data1.slice(ridx, ridx + 50).match(dimregex);
		if (ref && !opts.nodim) {
			parse_ws_xml_dim(s, ref[1]);
		}
	}

	// Columns
	const columns: ColInfo[] = [];
	if (opts.cellStyles) {
		const cols = data1.match(colregex);
		if (cols) {
			parse_ws_xml_cols(columns, cols);
		}
	}

	// SheetData
	if (sdMatch) {
		parse_ws_xml_data(sdMatch[1], s, opts, refguess, _themes, styles, wb);
	}

	// AutoFilter
	const afilter = data2.match(afregex);
	if (afilter) {
		s["!autofilter"] = parse_ws_xml_autofilter(afilter[0]);
	}

	// Merges
	const merges: Range[] = [];
	const _merge = data2.match(mergecregex);
	if (_merge) {
		for (let i = 0; i < _merge.length; ++i) {
			merges[i] = safe_decode_range(_merge[i].slice(_merge[i].indexOf("=") + 2));
		}
	}

	// Hyperlinks
	const hlink = data2.match(hlinkregex);
	if (hlink) {
		parse_ws_xml_hlinks(s, hlink, rels!);
	}

	// Margins
	const margins = data2.match(marginregex);
	if (margins) {
		s["!margins"] = parse_ws_xml_margins(parsexmltag(margins[0]));
	}

	// Legacy drawing
	const legm = data2.match(/legacyDrawing r:id="(.*?)"/);
	if (legm) {
		(s as any)["!legrel"] = legm[1];
	}

	if (opts.nodim) {
		refguess.s.c = refguess.s.r = 0;
	}
	if (!s["!ref"] && refguess.e.c >= refguess.s.c && refguess.e.r >= refguess.s.r) {
		s["!ref"] = encode_range(refguess);
	}
	if (opts.sheetRows > 0 && s["!ref"]) {
		const tmpref = safe_decode_range(s["!ref"]);
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
			(s as any)["!fullref"] = s["!ref"];
			s["!ref"] = encode_range(tmpref);
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

/** Write merges XML */
function write_ws_xml_merges(merges: Range[]): string {
	if (merges.length === 0) {
		return "";
	}
	const o = ['<mergeCells count="' + merges.length + '">'];
	for (let i = 0; i < merges.length; ++i) {
		o.push('<mergeCell ref="' + encode_range(merges[i]) + '"/>');
	}
	o.push("</mergeCells>");
	return o.join("");
}

/** Write a worksheet XML */
export function write_ws_xml(ws: WorkSheet, opts: any, _idx: number, _rels: Relationships, _wb: any): string {
	const o: string[] = [XML_HEADER];
	o.push(
		writextag("worksheet", null, {
			xmlns: XMLNS_main[0],
			"xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
		}),
	);

	const ref = ws["!ref"] || "A1";
	o.push('<dimension ref="' + ref + '"/>');

	o.push('<sheetViews><sheetView workbookViewId="0"');
	// Only add tabSelected for first sheet
	if (_idx === 0) {
		o.push(' tabSelected="1"');
	}
	o.push("/></sheetViews>");

	o.push('<sheetFormatPr defaultRowHeight="15"/>');

	// Columns
	if (ws["!cols"]) {
		o.push("<cols>");
		for (let i = 0; i < ws["!cols"].length; ++i) {
			if (!ws["!cols"][i]) {
				continue;
			}
			const col = ws["!cols"][i];
			const attrs: Record<string, string> = {
				min: String(i + 1),
				max: String(i + 1),
			};
			if (col.width) {
				attrs.width = String(col.width);
			} else {
				attrs.width = "9.140625";
			}
			if (col.hidden) {
				attrs.hidden = "1";
			}
			attrs.customWidth = "1";
			o.push(writextag("col", null, attrs));
		}
		o.push("</cols>");
	}

	o.push("<sheetData>");

	const dense = ws["!data"] != null;
	const range = safe_decode_range(ref);

	for (let R = range.s.r; R <= range.e.r; ++R) {
		const row_cells: string[] = [];
		for (let C = range.s.c; C <= range.e.c; ++C) {
			let cell: CellObject | undefined;
			if (dense) {
				cell = ws["!data"]?.[R]?.[C];
			} else {
				const addr = encode_cell({ r: R, c: C });
				cell = ws[addr] as CellObject | undefined;
			}
			if (!cell || cell.t === "z") {
				continue;
			}

			const addr = encode_cell({ r: R, c: C });
			let v = "";
			let t = "";

			switch (cell.t) {
				case "b":
					v = cell.v ? "1" : "0";
					t = "b";
					break;
				case "n":
					v = String(cell.v);
					break;
				case "e":
					v = String(cell.v);
					t = "e";
					break;
				case "d":
					if (opts.cellDates) {
						v = (cell.v as Date).toISOString();
						t = "d";
					} else {
						v = String(datenum(cell.v as Date));
					}
					break;
				case "s":
					v = escapexml(String(cell.v));
					t = "str";
					break;
			}

			let cellXml = '<c r="' + addr + '"';
			if (t) {
				cellXml += ' t="' + t + '"';
			}
			cellXml += ">";
			if (cell.f) {
				cellXml += "<f";
				if (cell.F) {
					cellXml += ' ref="' + cell.F + '" t="array"';
				}
				cellXml += ">" + escapexml(cell.f) + "</f>";
			}
			if (v !== "") {
				cellXml += "<v>" + v + "</v>";
			}
			cellXml += "</c>";
			row_cells.push(cellXml);
		}
		if (row_cells.length > 0) {
			let rowTag = '<row r="' + (R + 1) + '"';
			if (ws["!rows"]?.[R]) {
				if (ws["!rows"][R].hpt) {
					rowTag += ' ht="' + ws["!rows"][R].hpt + '" customHeight="1"';
				}
				if (ws["!rows"][R].hidden) {
					rowTag += ' hidden="1"';
				}
			}
			rowTag += ">";
			o.push(rowTag);
			o.push(row_cells.join(""));
			o.push("</row>");
		}
	}

	o.push("</sheetData>");

	// Merges
	if (ws["!merges"] && ws["!merges"].length > 0) {
		o.push(write_ws_xml_merges(ws["!merges"]));
	}

	// AutoFilter
	if (ws["!autofilter"]) {
		o.push('<autoFilter ref="' + ws["!autofilter"].ref + '"/>');
	}

	// Margins
	if (ws["!margins"]) {
		const m = ws["!margins"];
		o.push(
			writextag("pageMargins", null, {
				left: String(m.left || 0.7),
				right: String(m.right || 0.7),
				top: String(m.top || 0.75),
				bottom: String(m.bottom || 0.75),
				header: String(m.header || 0.3),
				footer: String(m.footer || 0.3),
			}),
		);
	}

	o.push("</worksheet>");
	o[1] = o[1].replace("/>", ">");
	return o.join("");
}
