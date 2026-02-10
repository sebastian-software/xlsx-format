/** Fill string with repeated character */
export function fill(c: string, l: number): string {
	let o = "";
	while (o.length < l) {
		o += c;
	}
	return o;
}

/** Object keys helper */
export function keys(o: object): string[] {
	return Object.keys(o);
}

/** Reverse map: values become keys */
export function evert(obj: Record<string, string>): Record<string, string> {
	const o: Record<string, string> = {};
	for (const k of Object.keys(obj)) {
		o[obj[k]] = k;
	}
	return o;
}

/** Shallow clone */
export function dup<T>(o: T): T {
	if (typeof o === "object" && o !== null) {
		if (Array.isArray(o)) {
			return o.slice() as T;
		}
		const out = {} as any;
		for (const k of Object.keys(o as any)) {
			out[k] = (o as any)[k];
		}
		return out as T;
	}
	return o;
}

/** Match XML namespace-agnostic tags globally */
export function str_match_xml_ns_g(str: string, tag: string): string[] | null {
	const re = new RegExp("<(?:\\w+:)?" + tag + "[\\s>][\\s\\S]*?<\\/(?:\\w+:)?" + tag + ">", "g");
	return str.match(re);
}

/** Match XML namespace-agnostic tag (first) */
export function str_match_xml_ns(str: string, tag: string): string | null {
	const m = str_match_xml_ns_g(str, tag);
	return m ? m[0] : null;
}
