// Matches attribute key="value" pairs within an XML tag. Captures the attribute
// name, then handles double-quoted, single-quoted, or unquoted values.
const attregexg = /\s([^"\s?>/]+)\s*=\s*((?:")([^"]*)(?:")|(?:')([^']*)(?:')|([^'">\s]+))/g;

// Strict XML tag regex: matches well-formed tags with optional attributes
const tagregex1 = /<[/?]?[a-zA-Z0-9:_-]+(?:\s+[^"\s?<>/]+\s*=\s*(?:"[^"]*"|'[^']*'|[^'"<>\s=]+))*\s*[/?]?>/gm;
// Lenient fallback: matches anything between < and >
const tagregex2 = /<[^<>]*>/g;

/** Standard XML declaration header with UTF-8 encoding and Windows line ending */
export const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

/**
 * The regex used to tokenize XML tags throughout the codebase.
 * Uses the strict pattern if it can match the XML header, otherwise falls back to the lenient one.
 */
export const XML_TAG_REGEX: RegExp = XML_HEADER.match(tagregex1) ? tagregex1 : tagregex2;

// Matches namespace-prefixed opening or closing tags like <ns:tag or </ns:tag
const nsregex2 = /<(\/?)\w+:/;

/**
 * Parse XML attributes from a raw tag string into a key-value record.
 * Handles namespace prefixes by stripping them (e.g., "r:id" becomes "id").
 * Stores the tag name itself under key `0` unless skip_root is true.
 * @param tag - the raw XML tag string (e.g., `<Relationship Id="rId1" Target="..."/>`)
 * @param skip_root - if true, do not store the tag name under key `0`
 * @param skip_LC - if true, do not store lowercase copies of attribute names
 * @returns a record mapping attribute names to their values
 */
export function parseXmlTag(tag: string, skip_root?: boolean, skip_LC?: boolean): Record<string, any> {
	const attrs: Record<string, any> = {};
	let scanPos = 0;
	let charCode = 0;
	// Scan forward to find the end of the tag name (terminated by space, LF, or CR)
	for (; scanPos !== tag.length; ++scanPos) {
		if ((charCode = tag.charCodeAt(scanPos)) === 32 || charCode === 10 || charCode === 13) {
			break; // 32=space, 10=LF, 13=CR
		}
	}
	if (!skip_root) {
		attrs[0] = tag.slice(0, scanPos);
	}
	if (scanPos === tag.length) {
		return attrs;
	}
	const matches = tag.match(attregexg);
	if (matches) {
		for (let i = 0; i < matches.length; ++i) {
			// Strip the leading whitespace character captured by the regex
			const attrStr = matches[i].slice(1);
			let eqPos = 0;
			// Find the '=' separator (charCode 61)
			for (eqPos = 0; eqPos < attrStr.length; ++eqPos) {
				if (attrStr.charCodeAt(eqPos) === 61) {
					break;
				}
			}
			let attrName = attrStr.slice(0, eqPos).trim();
			// Skip any spaces between '=' and the value
			while (attrStr.charCodeAt(eqPos + 1) === 32) {
				++eqPos;
			}
			// Detect and skip surrounding quotes (charCode 34=" or 39=')
			const quoteOffset = (scanPos = attrStr.charCodeAt(eqPos + 1)) === 34 || scanPos === 39 ? 1 : 0;
			const attrValue = attrStr.slice(eqPos + 1 + quoteOffset, attrStr.length - quoteOffset);

			// Check for namespace prefix by finding ':' (charCode 58)
			let colonPos = 0;
			for (colonPos = 0; colonPos < attrName.length; ++colonPos) {
				if (attrName.charCodeAt(colonPos) === 58) {
					break;
				}
			}
			if (colonPos === attrName.length) {
				// No namespace prefix -- truncate at underscore if present (handles extended names)
				if (attrName.indexOf("_") > 0) {
					attrName = attrName.slice(0, attrName.indexOf("_"));
				}
				attrs[attrName] = attrValue;
				if (!skip_LC) {
					attrs[attrName.toLowerCase()] = attrValue;
				}
			} else {
				// Has namespace prefix -- extract the local name.
				// Special case: "xmlns:foo" keeps "xmlns" prefix for the local name.
				const localName = (colonPos === 5 && attrName.slice(0, 5) === "xmlns" ? "xmlns" : "") + attrName.slice(colonPos + 1);
				// Skip "ext" namespace attributes if a non-ext value already exists
				if (attrs[localName] && attrName.slice(colonPos - 3, colonPos) === "ext") {
					continue;
				}
				attrs[localName] = attrValue;
				if (!skip_LC) {
					attrs[localName.toLowerCase()] = attrValue;
				}
			}
		}
	}
	return attrs;
}

/**
 * Strip namespace prefixes from XML tag names (e.g., `<a:foo>` becomes `<foo>`).
 * @param x - XML string with possible namespace prefixes
 * @returns the XML string with namespace prefixes removed from tag names
 */
export function stripNamespace(x: string): string {
	return x.replace(nsregex2, "<$1");
}

/**
 * Parse xsd:boolean-compatible values to a native boolean.
 * Accepts 1, true, "1", "true" as truthy; 0, false, "0", "false" as falsy.
 * @param value - the value to interpret as boolean
 * @returns true or false
 */
export function parseXmlBoolean(value: any): boolean {
	switch (value) {
		case 1:
		case true:
		case "1":
		case "true":
			return true;
		case 0:
		case false:
		case "0":
		case "false":
			return false;
	}
	return false;
}
