import { parsexmltag, tagregex } from "../xml/parser.js";

export interface CalcChainEntry {
	r?: string;
	i?: any;
	l?: string;
	a?: string;
	[key: string]: any;
}

/** Parse calculation chain XML (18.6) */
export function parse_cc_xml(data: string): CalcChainEntry[] {
	const d: CalcChainEntry[] = [];
	if (!data) {
		return d;
	}
	let i: any = 1;
	(data.match(tagregex) || []).forEach((x) => {
		const y: any = parsexmltag(x);
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
