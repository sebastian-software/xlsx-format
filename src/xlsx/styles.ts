import { parsexmltag, tagregex, XML_HEADER } from "../xml/parser.js";
import { unescapexml, escapexml } from "../xml/escape.js";
import { writextag } from "../xml/writer.js";
import { XMLNS_main } from "../xml/namespaces.js";
import { SSF_load } from "../ssf/table.js";

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

function parse_numFmts(t: string, styles: StylesData, opts: any): void {
	const m = t.match(tagregex);
	if (!m) {
		return;
	}
	for (let i = 0; i < m.length; ++i) {
		const y = parsexmltag(m[i]);
		switch (strip_tag(y[0])) {
			case "<numFmt": {
				const f = unescapexml(y.formatCode);
				const j = parseInt(y.numFmtId, 10);
				styles.NumberFmt[j] = f;
				if (j > 0) {
					if (j > 0x188) {
						// high-numbered format
					}
					SSF_load(f, j);
				}
				break;
			}
		}
	}
}

function parse_cellXfs(t: string, styles: StylesData): void {
	const m = t.match(tagregex);
	if (!m) {
		return;
	}
	let xf: CellXfEntry | null = null;
	for (let i = 0; i < m.length; ++i) {
		const y = parsexmltag(m[i]);
		switch (strip_tag(y[0])) {
			case "<xf":
				xf = {
					numFmtId: parseInt(y.numFmtId, 10) || 0,
					fontId: parseInt(y.fontId, 10) || 0,
					fillId: parseInt(y.fillId, 10) || 0,
					borderId: parseInt(y.borderId, 10) || 0,
					xfId: parseInt(y.xfId, 10) || 0,
				};
				if (y.applyNumberFormat) {
					xf.applyNumberFormat = y.applyNumberFormat === "1";
				}
				styles.CellXf.push(xf);
				break;
		}
	}
}

function strip_tag(tag: string): string {
	return tag.replace(/<\w+:/, "<");
}

/** Parse a styles XML file */
export function parse_sty_xml(data: string, _themes?: any, opts?: any): StylesData {
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
		parse_numFmts(numFmts[1], styles, opts);
	}

	/* cellXfs */
	const cellXfs = data.match(/<(?:\w+:)?cellXfs[^>]*>([\s\S]*?)<\/(?:\w+:)?cellXfs>/);
	if (cellXfs) {
		parse_cellXfs(cellXfs[1], styles);
	}

	return styles;
}

/** Write a minimal styles XML */
export function write_sty_xml(_wb: any, _opts: any): string {
	const o: string[] = [XML_HEADER];
	o.push(
		writextag("styleSheet", null, {
			xmlns: XMLNS_main[0],
			"xmlns:vt": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
		}),
	);

	// Number Formats
	o.push('<numFmts count="1"><numFmt numFmtId="164" formatCode="General"/></numFmts>');

	// Fonts
	o.push(
		'<fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>',
	);

	// Fills
	o.push(
		'<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>',
	);

	// Borders
	o.push('<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>');

	// Cell Style Xfs
	o.push('<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>');

	// Cell Xfs
	o.push(
		'<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>',
	);

	// Cell Styles
	o.push('<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>');

	o.push("</styleSheet>");
	o[1] = o[1].replace("/>", ">");
	return o.join("");
}
