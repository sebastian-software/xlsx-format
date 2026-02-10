/** Match XML namespace-agnostic tags globally */
export function matchXmlTagGlobal(xmlString: string, tag: string): string[] | null {
	const re = new RegExp("<(?:\\w+:)?" + tag + "[\\s>][\\s\\S]*?<\\/(?:\\w+:)?" + tag + ">", "g");
	return xmlString.match(re);
}

/** Match XML namespace-agnostic tag (first) */
export function matchXmlTagFirst(xmlString: string, tag: string): string | null {
	const m = matchXmlTagGlobal(xmlString, tag);
	return m ? m[0] : null;
}
