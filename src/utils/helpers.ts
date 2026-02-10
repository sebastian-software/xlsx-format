/** Fill string with repeated character */
export function repeatChar(char: string, targetLength: number): string {
	let result = "";
	while (result.length < targetLength) {
		result += char;
	}
	return result;
}

/** Object keys helper */
export function objectKeys(obj: object): string[] {
	return Object.keys(obj);
}

/** Reverse map: values become keys */
export function invertMapping(obj: Record<string, string>): Record<string, string> {
	const inverted: Record<string, string> = {};
	for (const key of Object.keys(obj)) {
		inverted[obj[key]] = key;
	}
	return inverted;
}

/** Shallow clone */
export function shallowClone<T>(source: T): T {
	if (typeof source === "object" && source !== null) {
		if (Array.isArray(source)) {
			return source.slice() as T;
		}
		const result = {} as any;
		for (const key of Object.keys(source as any)) {
			result[key] = (source as any)[key];
		}
		return result as T;
	}
	return source;
}

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
