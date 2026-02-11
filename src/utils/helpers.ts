/**
 * Find all occurrences of an XML tag in a string, ignoring namespace prefixes.
 *
 * Builds a regex that matches `<tag>...</tag>` or `<ns:tag>...</ns:tag>`
 * and returns all matches. Handles attributes and nested content via [\s\S]*?.
 *
 * @param xmlString - The XML string to search
 * @param tag - The local tag name (without namespace prefix)
 * @returns Array of matched tag strings, or null if no matches
 */
export function matchXmlTagGlobal(xmlString: string, tag: string): string[] | null {
	const re = new RegExp("<(?:\\w+:)?" + tag + "[\\s>][\\s\\S]*?<\\/(?:\\w+:)?" + tag + ">", "g");
	return xmlString.match(re);
}

/**
 * Find the first occurrence of an XML tag in a string, ignoring namespace prefixes.
 *
 * Convenience wrapper around {@link matchXmlTagGlobal} that returns only the first match.
 *
 * @param xmlString - The XML string to search
 * @param tag - The local tag name (without namespace prefix)
 * @returns The first matched tag string, or null if not found
 */
export function matchXmlTagFirst(xmlString: string, tag: string): string | null {
	const m = matchXmlTagGlobal(xmlString, tag);
	return m ? m[0] : null;
}
