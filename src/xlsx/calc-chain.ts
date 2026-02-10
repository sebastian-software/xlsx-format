import { parseXmlTag, XML_TAG_REGEX } from "../xml/parser.js";

/** A single entry in the calculation chain */
export interface CalcChainEntry {
	/** Cell reference (e.g. "A1") */
	r?: string;
	/** Sheet index */
	i?: any;
	/** "new dependency level" flag */
	l?: string;
	/** Array formula flag */
	a?: string;
	[key: string]: any;
}

/**
 * Parse the calculation chain XML (ECMA-376 18.6).
 *
 * The calculation chain records the order in which cells with formulas should
 * be recalculated. Each <c> entry references a cell and its sheet index.
 * The sheet index (i) is sticky: if omitted, it inherits from the previous entry.
 *
 * @param data - Raw XML string of calcChain.xml
 * @returns Array of calculation chain entries
 */
export function parseCalcChainXml(data: string): CalcChainEntry[] {
	const d: CalcChainEntry[] = [];
	if (!data) {
		return d;
	}
	// Sticky sheet index: persists across entries when not explicitly set
	let i: any = 1;
	(data.match(XML_TAG_REGEX) || []).forEach((x) => {
		const y: any = parseXmlTag(x);
		switch (y[0]) {
			case "<?xml":
				break;
			case "<calcChain":
			case "<calcChain>":
			case "</calcChain>":
				break;
			case "<c":
				delete y[0];
				if (y.i) {
					i = y.i; // Update sticky sheet index
				} else {
					y.i = i; // Inherit from previous entry
				}
				d.push(y);
				break;
		}
	});
	return d;
}
