import type {
	WorkBook,
	WorkSheet,
	CellObject,
	CellStyle,
	StyleColor,
	CellFont,
	CellFill,
	CellBorder,
	CellBorderSide,
	CellAlignment,
} from "../types.js";
import { XlsxError } from "../errors.js";
import { parseXmlTag, XML_TAG_REGEX, XML_HEADER } from "../xml/parser.js";
import { escapeXml, unescapeXml } from "../xml/escape.js";
import { writeXmlElement } from "../xml/writer.js";
import { XMLNS_main } from "../xml/namespaces.js";
import { formatTable, loadFormat } from "../ssf/table.js";

/** Parsed style information from styles.xml */
export interface StylesData {
	NumberFmt: Record<number, string>;
	CellXf: CellXfEntry[];
	Fonts: CellFont[];
	Fills: CellFill[];
	Borders: CellBorder[];
}

/** A single cell format (xf) entry from the cellXfs collection */
export interface CellXfEntry {
	numFmtId: number;
	fontId?: number;
	fillId?: number;
	borderId?: number;
	xfId?: number;
	applyNumberFormat?: boolean;
	applyFont?: boolean;
	applyFill?: boolean;
	applyBorder?: boolean;
	applyAlignment?: boolean;
	alignment?: CellAlignment;
}

interface NormalizedCellStyle {
	font?: CellFont;
	fill?: CellFill;
	border?: CellBorder;
	alignment?: CellAlignment;
	numFmt?: string | number;
}

interface StyleXf {
	numFmtId: number;
	fontId: number;
	fillId: number;
	borderId: number;
	alignment?: CellAlignment;
	applyNumberFormat?: boolean;
	applyFont?: boolean;
	applyFill?: boolean;
	applyBorder?: boolean;
	applyAlignment?: boolean;
}

export interface StyleRegistry {
	cellStyleIds: WeakMap<CellObject, number>;
	numFmts: Map<number, string>;
	fonts: CellFont[];
	fills: CellFill[];
	borders: CellBorder[];
	cellXfs: StyleXf[];
	hasStyles: boolean;
}

const DEFAULT_FONT: CellFont = { name: "Calibri", size: 11 };
const DEFAULT_FILL: CellFill = {};
const GRAY125_FILL: CellFill = { patternType: "solid" };
const DEFAULT_BORDER: CellBorder = {};

function styleKey(value: unknown): string {
	if (value == null) {
		return "";
	}
	if (typeof value !== "object") {
		return JSON.stringify(value);
	}
	if (Array.isArray(value)) {
		return "[" + value.map((item) => styleKey(item)).join(",") + "]";
	}
	const record = value as Record<string, unknown>;
	return (
		"{" +
		Object.keys(record)
			.sort()
			.filter((key) => record[key] !== undefined)
			.map((key) => JSON.stringify(key) + ":" + styleKey(record[key]))
			.join(",") +
		"}"
	);
}

function normalizeColor(color?: StyleColor, opts?: any): StyleColor | undefined {
	const raw = color?.argb || color?.rgb;
	if (!raw) {
		return undefined;
	}
	const normalized = raw.replace(/^#/, "").toUpperCase();
	let rgb = "";
	if (/^[0-9A-F]{6}$/.test(normalized)) {
		rgb = "FF" + normalized;
	} else if (/^[0-9A-F]{8}$/.test(normalized)) {
		rgb = normalized;
	} else {
		if (opts?.WTF) {
			throw new XlsxError("UNSUPPORTED", "Unsupported style color: " + raw);
		}
		return undefined;
	}
	return { rgb };
}

function normalizeFont(font?: CellFont, opts?: any): CellFont | undefined {
	if (!font) {
		return undefined;
	}
	const out: CellFont = {};
	if (font.name) {
		out.name = font.name;
	}
	if (typeof font.size === "number") {
		out.size = font.size;
	}
	if (font.bold) {
		out.bold = true;
	}
	const color = normalizeColor(font.color, opts);
	if (color) {
		out.color = color;
	}
	return Object.keys(out).length > 0 ? out : undefined;
}

function normalizeFill(fill?: CellFill, opts?: any): CellFill | undefined {
	if (!fill) {
		return undefined;
	}
	if (fill.patternType && fill.patternType !== "solid") {
		if (opts?.WTF) {
			throw new XlsxError("UNSUPPORTED", "Unsupported fill pattern: " + fill.patternType);
		}
		return undefined;
	}
	const color = normalizeColor(fill.fgColor, opts);
	if (!color) {
		return undefined;
	}
	return { patternType: "solid", fgColor: color };
}

function normalizeBorder(border?: CellBorder, opts?: any): CellBorder | undefined {
	if (!border) {
		return undefined;
	}
	const out: CellBorder = {};
	for (const side of ["top", "right", "bottom", "left"] as const) {
		const input = border[side];
		if (!input) {
			continue;
		}
		if (input.style !== "thin" && input.style !== "medium") {
			if (opts?.WTF) {
				throw new XlsxError("UNSUPPORTED", "Unsupported border style: " + input.style);
			}
			continue;
		}
		const color = normalizeColor(input.color, opts);
		out[side] = color ? { style: input.style, color } : { style: input.style };
	}
	return Object.keys(out).length > 0 ? out : undefined;
}

function normalizeAlignment(alignment?: CellAlignment, opts?: any): CellAlignment | undefined {
	if (!alignment) {
		return undefined;
	}
	const out: CellAlignment = {};
	if (alignment.horizontal === "left" || alignment.horizontal === "center" || alignment.horizontal === "right") {
		out.horizontal = alignment.horizontal;
	} else if (alignment.horizontal && opts?.WTF) {
		throw new XlsxError("UNSUPPORTED", "Unsupported horizontal alignment: " + alignment.horizontal);
	}
	if (alignment.vertical === "top" || alignment.vertical === "middle" || alignment.vertical === "bottom") {
		out.vertical = alignment.vertical;
	} else if (alignment.vertical && opts?.WTF) {
		throw new XlsxError("UNSUPPORTED", "Unsupported vertical alignment: " + alignment.vertical);
	}
	if (alignment.wrapText) {
		out.wrapText = true;
	}
	return Object.keys(out).length > 0 ? out : undefined;
}

function normalizeCellStyle(cell: CellObject, opts?: any): NormalizedCellStyle | undefined {
	const style = cell.s;
	const out: NormalizedCellStyle = {};
	if (style) {
		out.font = normalizeFont(style.font, opts);
		out.fill = normalizeFill(style.fill, opts);
		out.border = normalizeBorder(style.border, opts);
		out.alignment = normalizeAlignment(style.alignment, opts);
		if (style.numFmt != null) {
			out.numFmt = style.numFmt;
		}
	}
	if (out.numFmt == null && cell.z != null) {
		out.numFmt = cell.z;
	}
	return Object.keys(out).some((key) => (out as any)[key] !== undefined) ? out : undefined;
}

function getOrAdd<T>(items: T[], seen: Map<string, number>, value: T): number {
	const key = styleKey(value);
	const existing = seen.get(key);
	if (existing != null) {
		return existing;
	}
	const id = items.length;
	items.push(value);
	seen.set(key, id);
	return id;
}

function getBuiltinNumFmtId(format: string): number | undefined {
	for (const key of Object.keys(formatTable)) {
		if (formatTable[+key] === format) {
			return +key;
		}
	}
	return undefined;
}

function getNumFmtId(
	numFmt: string | number | undefined,
	customNumFmts: Map<string, number>,
	numFmts: Map<number, string>,
): number {
	if (numFmt == null) {
		return 0;
	}
	if (typeof numFmt === "number") {
		return numFmt;
	}
	const builtin = getBuiltinNumFmtId(numFmt);
	if (builtin != null) {
		return builtin;
	}
	const existing = customNumFmts.get(numFmt);
	if (existing != null) {
		return existing;
	}
	let id = 164;
	while (numFmts.has(id) || formatTable[id]) {
		id++;
	}
	customNumFmts.set(numFmt, id);
	numFmts.set(id, numFmt);
	return id;
}

function eachWorksheetCell(ws: WorkSheet, callback: (cell: CellObject) => void): void {
	if (ws["!data"]) {
		for (const row of ws["!data"]) {
			if (!row) {
				continue;
			}
			for (const cell of row) {
				if (cell) {
					callback(cell);
				}
			}
		}
		return;
	}
	for (const key of Object.keys(ws)) {
		if (key.charAt(0) === "!") {
			continue;
		}
		const cell = ws[key] as CellObject | undefined;
		if (cell) {
			callback(cell);
		}
	}
}

export function buildStyleRegistry(wb: WorkBook, opts?: any): StyleRegistry {
	const registry: StyleRegistry = {
		cellStyleIds: new WeakMap(),
		numFmts: new Map(),
		fonts: [DEFAULT_FONT],
		fills: [DEFAULT_FILL, GRAY125_FILL],
		borders: [DEFAULT_BORDER],
		cellXfs: [{ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 }],
		hasStyles: false,
	};
	const fontIds = new Map<string, number>([[styleKey(DEFAULT_FONT), 0]]);
	const fillIds = new Map<string, number>([
		[styleKey(DEFAULT_FILL), 0],
		[styleKey(GRAY125_FILL), 1],
	]);
	const borderIds = new Map<string, number>([[styleKey(DEFAULT_BORDER), 0]]);
	const xfIds = new Map<string, number>([[styleKey(registry.cellXfs[0]), 0]]);
	const customNumFmts = new Map<string, number>();

	for (const sheetName of wb.SheetNames) {
		const ws = wb.Sheets[sheetName];
		if (!ws) {
			continue;
		}
		eachWorksheetCell(ws, (cell) => {
			const normalized = normalizeCellStyle(cell, opts);
			if (!normalized) {
				return;
			}
			const fontId = normalized.font ? getOrAdd(registry.fonts, fontIds, normalized.font) : 0;
			const fillId = normalized.fill ? getOrAdd(registry.fills, fillIds, normalized.fill) : 0;
			const borderId = normalized.border ? getOrAdd(registry.borders, borderIds, normalized.border) : 0;
			const numFmtId = getNumFmtId(normalized.numFmt, customNumFmts, registry.numFmts);
			const xf: StyleXf = { numFmtId, fontId, fillId, borderId };
			if (numFmtId !== 0) {
				xf.applyNumberFormat = true;
			}
			if (fontId !== 0) {
				xf.applyFont = true;
			}
			if (fillId !== 0) {
				xf.applyFill = true;
			}
			if (borderId !== 0) {
				xf.applyBorder = true;
			}
			if (normalized.alignment) {
				xf.alignment = normalized.alignment;
				xf.applyAlignment = true;
			}
			const styleId = getOrAdd(registry.cellXfs, xfIds, xf);
			registry.cellStyleIds.set(cell, styleId);
			if (styleId !== 0) {
				registry.hasStyles = true;
			}
		});
	}
	return registry;
}

export function getCellStyleIndex(opts: any, cell: CellObject): number | undefined {
	return opts?.styleRegistry?.cellStyleIds?.get(cell);
}

/**
 * Parse <numFmts> section, registering custom number formats into the styles object
 * and the global SSF format table.
 */
function parseNumberFormats(t: string, styles: StylesData, _opts: any): void {
	const matches = t.match(XML_TAG_REGEX);
	if (!matches) {
		return;
	}
	for (let i = 0; i < matches.length; ++i) {
		const parsedTag = parseXmlTag(matches[i]);
		switch (stripTagNamespace(parsedTag[0])) {
			case "<numFmt": {
				const formatCode = unescapeXml(parsedTag.formatCode);
				const fmtId = parseInt(parsedTag.numFmtId, 10);
				styles.NumberFmt[fmtId] = formatCode;
				if (fmtId > 0) {
					if (fmtId > 0x188) {
						// Format IDs above 0x188 (392) are beyond the standard reserved range
					}
					loadFormat(formatCode, fmtId);
				}
				break;
			}
		}
	}
}

function parseColor(tag: Record<string, any>): StyleColor | undefined {
	if (tag.rgb) {
		return { rgb: String(tag.rgb).toUpperCase() };
	}
	return undefined;
}

function parseFonts(t: string, styles: StylesData): void {
	const fonts = t.match(/<(?:\w+:)?font\b[^>]*>[\s\S]*?<\/(?:\w+:)?font>|<(?:\w+:)?font\b[^>]*\/>/g);
	if (!fonts) {
		return;
	}
	for (const fontXml of fonts) {
		const font: CellFont = {};
		const name = fontXml.match(/<(?:\w+:)?name\b[^>]*\/>/);
		if (name) {
			const tag = parseXmlTag(name[0]);
			if (tag.val) {
				font.name = tag.val;
			}
		}
		const size = fontXml.match(/<(?:\w+:)?sz\b[^>]*\/>/);
		if (size) {
			const tag = parseXmlTag(size[0]);
			if (tag.val) {
				font.size = parseFloat(tag.val);
			}
		}
		if (/<(?:\w+:)?b\b[^>]*\/>/.test(fontXml)) {
			font.bold = true;
		}
		const color = fontXml.match(/<(?:\w+:)?color\b[^>]*\/>/);
		if (color) {
			const parsed = parseColor(parseXmlTag(color[0]));
			if (parsed) {
				font.color = parsed;
			}
		}
		styles.Fonts.push(font);
	}
}

function parseFills(t: string, styles: StylesData): void {
	const fills = t.match(/<(?:\w+:)?fill\b[^>]*>[\s\S]*?<\/(?:\w+:)?fill>|<(?:\w+:)?fill\b[^>]*\/>/g);
	if (!fills) {
		return;
	}
	for (const fillXml of fills) {
		const pattern = fillXml.match(/<(?:\w+:)?patternFill\b[^>]*>/);
		const fill: CellFill = {};
		if (pattern) {
			const patternTag = parseXmlTag(pattern[0]);
			if (patternTag.patternType === "solid") {
				fill.patternType = "solid";
				const fg = fillXml.match(/<(?:\w+:)?fgColor\b[^>]*\/>/);
				if (fg) {
					const parsed = parseColor(parseXmlTag(fg[0]));
					if (parsed) {
						fill.fgColor = parsed;
					}
				}
			}
		}
		styles.Fills.push(fill);
	}
}

function parseBorderSide(borderXml: string, side: "top" | "right" | "bottom" | "left"): CellBorderSide | undefined {
	const match = borderXml.match(
		new RegExp("<(?:\\w+:)?" + side + "\\b[^>]*(?:/>|>[\\s\\S]*?</(?:\\w+:)?" + side + ">)"),
	);
	if (!match) {
		return undefined;
	}
	const open = match[0].match(/<[^>]*>/);
	if (!open) {
		return undefined;
	}
	const tag = parseXmlTag(open[0]);
	if (tag.style !== "thin" && tag.style !== "medium") {
		return undefined;
	}
	const color = match[0].match(/<(?:\w+:)?color\b[^>]*\/>/);
	const parsedColor = color ? parseColor(parseXmlTag(color[0])) : undefined;
	return parsedColor ? { style: tag.style, color: parsedColor } : { style: tag.style };
}

function parseBorders(t: string, styles: StylesData): void {
	const borders = t.match(/<(?:\w+:)?border\b[^>]*>[\s\S]*?<\/(?:\w+:)?border>|<(?:\w+:)?border\b[^>]*\/>/g);
	if (!borders) {
		return;
	}
	for (const borderXml of borders) {
		const border: CellBorder = {};
		for (const side of ["top", "right", "bottom", "left"] as const) {
			const parsed = parseBorderSide(borderXml, side);
			if (parsed) {
				border[side] = parsed;
			}
		}
		styles.Borders.push(border);
	}
}

/**
 * Parse <cellXfs> section, extracting cell format entries that map style indices
 * to number format, font, fill, and border IDs.
 */
function parseCellFormats(t: string, styles: StylesData): void {
	const matches = t.match(XML_TAG_REGEX);
	if (!matches) {
		return;
	}
	let xf: CellXfEntry | null = null;
	for (let i = 0; i < matches.length; ++i) {
		const parsedTag = parseXmlTag(matches[i]);
		switch (stripTagNamespace(parsedTag[0])) {
			case "<xf":
				xf = {
					numFmtId: parseInt(parsedTag.numFmtId, 10) || 0,
					fontId: parseInt(parsedTag.fontId, 10) || 0,
					fillId: parseInt(parsedTag.fillId, 10) || 0,
					borderId: parseInt(parsedTag.borderId, 10) || 0,
					xfId: parseInt(parsedTag.xfId, 10) || 0,
				};
				if (parsedTag.applyNumberFormat) {
					xf.applyNumberFormat = parsedTag.applyNumberFormat === "1";
				}
				if (parsedTag.applyFont) {
					xf.applyFont = parsedTag.applyFont === "1";
				}
				if (parsedTag.applyFill) {
					xf.applyFill = parsedTag.applyFill === "1";
				}
				if (parsedTag.applyBorder) {
					xf.applyBorder = parsedTag.applyBorder === "1";
				}
				if (parsedTag.applyAlignment) {
					xf.applyAlignment = parsedTag.applyAlignment === "1";
				}
				styles.CellXf.push(xf);
				break;
			case "<alignment":
				if (xf) {
					const alignment: CellAlignment = {};
					if (
						parsedTag.horizontal === "left" ||
						parsedTag.horizontal === "center" ||
						parsedTag.horizontal === "right"
					) {
						alignment.horizontal = parsedTag.horizontal;
					}
					if (
						parsedTag.vertical === "top" ||
						parsedTag.vertical === "center" ||
						parsedTag.vertical === "bottom"
					) {
						alignment.vertical = parsedTag.vertical === "center" ? "middle" : parsedTag.vertical;
					}
					if (parsedTag.wrapText === "1") {
						alignment.wrapText = true;
					}
					if (Object.keys(alignment).length > 0) {
						xf.alignment = alignment;
					}
				}
				break;
		}
	}
}

export function getStyleFromXf(styles: StylesData, styleIndex: number): CellStyle | undefined {
	const xf = styles.CellXf[styleIndex];
	if (!xf) {
		return undefined;
	}
	const style: CellStyle = {};
	const font = styles.Fonts[xf.fontId || 0];
	if (font && Object.keys(font).length > 0) {
		style.font = font;
	}
	const fill = styles.Fills[xf.fillId || 0];
	if (fill && fill.patternType === "solid" && fill.fgColor) {
		style.fill = fill;
	}
	const border = styles.Borders[xf.borderId || 0];
	if (border && Object.keys(border).length > 0) {
		style.border = border;
	}
	if (xf.alignment) {
		style.alignment = xf.alignment;
	}
	if (xf.numFmtId) {
		style.numFmt = styles.NumberFmt[xf.numFmtId] || formatTable[xf.numFmtId] || xf.numFmtId;
	}
	return Object.keys(style).length > 0 ? style : undefined;
}

/** Strip XML namespace prefix from a tag name (e.g. "<x:numFmt" -> "<numFmt") */
function stripTagNamespace(tag: string): string {
	return tag.replace(/<\w+:/, "<");
}

/**
 * Parse a styles.xml file into a StylesData structure.
 *
 * Extracts custom number formats and cell format (xf) entries. Fonts, fills,
 * and borders arrays are initialized but not fully parsed in this implementation.
 *
 * @param data - Raw XML string of the styles.xml file
 * @param _themes - Parsed theme data (reserved for theme-based color resolution)
 * @param opts - Parsing options
 * @returns Parsed style data containing number formats and cell format entries
 */
export function parseStylesXml(data: string, _themes?: any, opts?: any): StylesData {
	const styles: StylesData = {
		NumberFmt: {},
		CellXf: [],
		Fonts: [],
		Fills: [],
		Borders: [],
	};

	if (!data) {
		return styles;
	}

	/* numFmts - custom number format definitions */
	const numFmts = data.match(/<(?:\w+:)?numFmts[^>]*>([\s\S]*?)<\/(?:\w+:)?numFmts>/);
	if (numFmts) {
		parseNumberFormats(numFmts[1], styles, opts);
	}

	const fonts = data.match(/<(?:\w+:)?fonts[^>]*>([\s\S]*?)<\/(?:\w+:)?fonts>/);
	if (fonts) {
		parseFonts(fonts[1], styles);
	}

	const fills = data.match(/<(?:\w+:)?fills[^>]*>([\s\S]*?)<\/(?:\w+:)?fills>/);
	if (fills) {
		parseFills(fills[1], styles);
	}

	const borders = data.match(/<(?:\w+:)?borders[^>]*>([\s\S]*?)<\/(?:\w+:)?borders>/);
	if (borders) {
		parseBorders(borders[1], styles);
	}

	/* cellXfs - cell format entries (style index -> format/font/fill/border IDs) */
	const cellXfs = data.match(/<(?:\w+:)?cellXfs[^>]*>([\s\S]*?)<\/(?:\w+:)?cellXfs>/);
	if (cellXfs) {
		parseCellFormats(cellXfs[1], styles);
	}

	return styles;
}

/**
 * Write a minimal styles.xml with default formatting.
 *
 * Produces a stylesheet with one "General" number format, one Calibri font,
 * two standard fills (none + gray125), one empty border, and two cell formats.
 * This is the minimum required for a valid XLSX file.
 *
 * @param _wb - WorkBook (reserved for future style extraction)
 * @param _opts - Write options
 * @returns Complete styles.xml string
 */
export function writeStylesXml(_wb: any, _opts: any): string {
	const lines: string[] = [XML_HEADER];
	const registry = _opts?.styleRegistry as StyleRegistry | undefined;
	lines.push(
		writeXmlElement("styleSheet", null, {
			xmlns: XMLNS_main[0],
			"xmlns:vt": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
		}),
	);

	const numFmts = registry?.numFmts || new Map<number, string>();
	if (numFmts.size > 0) {
		lines.push(
			'<numFmts count="' +
				numFmts.size +
				'">' +
				[...numFmts.entries()]
					.sort((a, b) => a[0] - b[0])
					.map(([id, code]) => '<numFmt numFmtId="' + id + '" formatCode="' + escapeXml(code) + '"/>')
					.join("") +
				"</numFmts>",
		);
	}

	const fonts = registry?.fonts || [DEFAULT_FONT];
	lines.push('<fonts count="' + fonts.length + '">' + fonts.map((font) => writeFont(font)).join("") + "</fonts>");

	const fills = registry?.fills || [DEFAULT_FILL, GRAY125_FILL];
	lines.push('<fills count="' + fills.length + '">' + fills.map((fill) => writeFill(fill)).join("") + "</fills>");

	const borders = registry?.borders || [DEFAULT_BORDER];
	lines.push(
		'<borders count="' +
			borders.length +
			'">' +
			borders.map((border) => writeBorder(border)).join("") +
			"</borders>",
	);

	// Cell Style Xfs - base format for the "Normal" style
	lines.push('<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>');

	const cellXfs = registry?.cellXfs || [
		{ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 },
		{ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 },
	];
	lines.push(
		'<cellXfs count="' +
			cellXfs.length +
			'">' +
			cellXfs.map((cellXf) => writeCellXf(cellXf)).join("") +
			"</cellXfs>",
	);

	// Cell Styles - the built-in "Normal" style
	lines.push('<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>');

	lines.push("</styleSheet>");
	// Convert self-closing <styleSheet .../> to opening tag <styleSheet ...>
	lines[1] = lines[1].replace("/>", ">");
	return lines.join("");
}

function writeFont(font: CellFont): string {
	const parts: string[] = [];
	parts.push('<sz val="' + (font.size || 11) + '"/>');
	if (font.color?.rgb) {
		parts.push('<color rgb="' + font.color.rgb + '"/>');
	} else {
		parts.push('<color theme="1"/>');
	}
	parts.push('<name val="' + escapeXml(font.name || "Calibri") + '"/>');
	parts.push('<family val="2"/>');
	if (font.bold) {
		parts.push("<b/>");
	}
	return "<font>" + parts.join("") + "</font>";
}

function writeFill(fill: CellFill): string {
	if (fill.patternType === "solid" && fill.fgColor?.rgb) {
		return (
			'<fill><patternFill patternType="solid"><fgColor rgb="' +
			fill.fgColor.rgb +
			'"/><bgColor indexed="64"/></patternFill></fill>'
		);
	}
	if (fill.patternType === "solid") {
		return '<fill><patternFill patternType="gray125"/></fill>';
	}
	return '<fill><patternFill patternType="none"/></fill>';
}

function writeBorder(border: CellBorder): string {
	const sides = ["left", "right", "top", "bottom"] as const;
	const out = sides.map((side) => writeBorderSide(side, border[side])).join("");
	return "<border>" + out + "<diagonal/></border>";
}

function writeBorderSide(side: string, border?: { style: "thin" | "medium"; color?: StyleColor }): string {
	if (!border) {
		return "<" + side + "/>";
	}
	const color = border.color?.rgb ? '<color rgb="' + border.color.rgb + '"/>' : "";
	return "<" + side + ' style="' + border.style + '">' + color + "</" + side + ">";
}

function writeCellXf(xf: StyleXf): string {
	const attrs: Record<string, string> = {
		numFmtId: String(xf.numFmtId),
		fontId: String(xf.fontId),
		fillId: String(xf.fillId),
		borderId: String(xf.borderId),
		xfId: "0",
	};
	if (xf.applyNumberFormat) {
		attrs.applyNumberFormat = "1";
	}
	if (xf.applyFont) {
		attrs.applyFont = "1";
	}
	if (xf.applyFill) {
		attrs.applyFill = "1";
	}
	if (xf.applyBorder) {
		attrs.applyBorder = "1";
	}
	if (xf.applyAlignment) {
		attrs.applyAlignment = "1";
	}
	if (!xf.alignment) {
		return writeXmlElement("xf", null, attrs);
	}
	const alignmentAttrs: Record<string, string> = {};
	if (xf.alignment.horizontal) {
		alignmentAttrs.horizontal = xf.alignment.horizontal;
	}
	if (xf.alignment.vertical) {
		alignmentAttrs.vertical = xf.alignment.vertical === "middle" ? "center" : xf.alignment.vertical;
	}
	if (xf.alignment.wrapText) {
		alignmentAttrs.wrapText = "1";
	}
	return writeXmlElement("xf", writeXmlElement("alignment", null, alignmentAttrs), attrs);
}
