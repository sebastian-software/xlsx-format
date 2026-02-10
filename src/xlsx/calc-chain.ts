import { parseXmlTag, XML_TAG_REGEX } from "../xml/parser.js";

export interface CalcChainEntry {
	r?: string;
	i?: any;
	l?: string;
	a?: string;
	[key: string]: any;
}

/** Parse calculation chain XML (18.6) */
export function parseCalcChainXml(data: string): CalcChainEntry[] {
	const d: CalcChainEntry[] = [];
	if (!data) {
		return d;
	}
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
					i = y.i;
				} else {
					y.i = i;
				}
				d.push(y);
				break;
		}
	});
	return d;
}
