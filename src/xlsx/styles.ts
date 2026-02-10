import { parseXmlTag, XML_TAG_REGEX, XML_HEADER } from "../xml/parser.js";
import { unescapeXml, escapeXml } from "../xml/escape.js";
import { writeXmlElement } from "../xml/writer.js";
import { XMLNS_main } from "../xml/namespaces.js";
import { loadFormat } from "../ssf/table.js";

/** Parsed style information from styles.xml */
export interface StylesData {
	NumberFmt: Record<number, string>;
	CellXf: CellXfEntry[];
	Fonts: any[];
	Fills: any[];
	Borders: any[];
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
}

/**
 * Parse <numFmts> section, registering custom number formats into the styles object
 * and the global SSF format table.
 */
function parseNumberFormats(t: string, styles: StylesData, opts: any): void {
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
				styles.CellXf.push(xf);
				break;
		}
	}
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
	lines.push(
		writeXmlElement("styleSheet", null, {
			xmlns: XMLNS_main[0],
			"xmlns:vt": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
		}),
	);

	// Number Formats - single "General" format with ID 164 (first custom ID slot)
	lines.push('<numFmts count="1"><numFmt numFmtId="164" formatCode="General"/></numFmts>');

	// Fonts - default Calibri 11pt
	lines.push(
		'<fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>',
	);

	// Fills - "none" and "gray125" are required by the spec as the first two fills
	lines.push(
		'<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>',
	);

	// Borders - single empty border
	lines.push('<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>');

	// Cell Style Xfs - base format for the "Normal" style
	lines.push('<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>');

	// Cell Xfs - two identical default formats (index 0 and 1)
	lines.push(
		'<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>',
	);

	// Cell Styles - the built-in "Normal" style
	lines.push('<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>');

	lines.push("</styleSheet>");
	// Convert self-closing <styleSheet .../> to opening tag <styleSheet ...>
	lines[1] = lines[1].replace("/>", ">");
	return lines.join("");
}
