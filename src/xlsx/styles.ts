import { parseXmlTag, XML_TAG_REGEX, XML_HEADER } from "../xml/parser.js";
import { unescapeXml, escapeXml } from "../xml/escape.js";
import { writeXmlElement } from "../xml/writer.js";
import { XMLNS_main } from "../xml/namespaces.js";
import { loadFormat } from "../ssf/table.js";

export interface StylesData {
	NumberFmt: Record<number, string>;
	CellXf: CellXfEntry[];
	Fonts: any[];
	Fills: any[];
	Borders: any[];
}

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
						// high-numbered format
					}
					loadFormat(formatCode, fmtId);
				}
				break;
			}
		}
	}
}

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

function stripTagNamespace(tag: string): string {
	return tag.replace(/<\w+:/, "<");
}

/** Parse a styles XML file */
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

	/* numFmts */
	const numFmts = data.match(/<(?:\w+:)?numFmts[^>]*>([\s\S]*?)<\/(?:\w+:)?numFmts>/);
	if (numFmts) {
		parseNumberFormats(numFmts[1], styles, opts);
	}

	/* cellXfs */
	const cellXfs = data.match(/<(?:\w+:)?cellXfs[^>]*>([\s\S]*?)<\/(?:\w+:)?cellXfs>/);
	if (cellXfs) {
		parseCellFormats(cellXfs[1], styles);
	}

	return styles;
}

/** Write a minimal styles XML */
export function writeStylesXml(_wb: any, _opts: any): string {
	const lines: string[] = [XML_HEADER];
	lines.push(
		writeXmlElement("styleSheet", null, {
			xmlns: XMLNS_main[0],
			"xmlns:vt": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
		}),
	);

	// Number Formats
	lines.push('<numFmts count="1"><numFmt numFmtId="164" formatCode="General"/></numFmts>');

	// Fonts
	lines.push(
		'<fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>',
	);

	// Fills
	lines.push(
		'<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>',
	);

	// Borders
	lines.push('<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>');

	// Cell Style Xfs
	lines.push('<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>');

	// Cell Xfs
	lines.push(
		'<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>',
	);

	// Cell Styles
	lines.push('<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>');

	lines.push("</styleSheet>");
	lines[1] = lines[1].replace("/>", ">");
	return lines.join("");
}
