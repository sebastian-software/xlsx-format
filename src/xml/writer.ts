import { escapeXml } from "./escape.js";

// Detects leading/trailing whitespace or embedded newlines that require xml:space="preserve"
const wtregex = /(^\s|\s$|\n)/;

/**
 * Write a simple XML tag wrapping text content.
 * Automatically adds xml:space="preserve" when the content has
 * leading/trailing whitespace or embedded newlines.
 * @param tagName - the XML element name
 * @param content - the text content to wrap
 * @returns an XML string like `<tag>content</tag>`
 */
export function writeXmlTag(tagName: string, content: string): string {
	return "<" + tagName + (content.match(wtregex) ? ' xml:space="preserve"' : "") + ">" + content + "</" + tagName + ">";
}

/**
 * Format a key-value record as XML attribute pairs.
 * @param attributes - attribute name/value pairs
 * @returns a string of ` key="value"` segments (with leading spaces)
 */
function formatXmlAttributes(attributes: Record<string, string>): string {
	return Object.keys(attributes)
		.map((key) => " " + key + '="' + attributes[key] + '"')
		.join("");
}

/**
 * Write an XML element with optional attributes and optional content.
 * When content is null/undefined, emits a self-closing tag (`<tag .../>`).
 * When content is provided, adds xml:space="preserve" if needed.
 * @param tagName - the XML element name
 * @param content - text content, or null for self-closing tag
 * @param attributes - optional attribute key-value pairs
 * @returns the complete XML element string
 */
export function writeXmlElement(tagName: string, content?: string | null, attributes?: Record<string, string> | null): string {
	return (
		"<" +
		tagName +
		(attributes != null ? formatXmlAttributes(attributes) : "") +
		// If content is non-null, emit open+close tags; otherwise self-close with "/>"
		(content != null ? (content.match(wtregex) ? ' xml:space="preserve"' : "") + ">" + content + "</" + tagName : "/") +
		">"
	);
}

/**
 * Write a W3C datetime string (ISO 8601 without fractional seconds) from a Date object.
 * Used for dcterms:created/modified and vt:filetime elements.
 * @param date - the Date to serialize
 * @param throwOnError - when true, rethrow invalid Date errors instead of returning ""
 * @returns an ISO datetime string like "2024-01-15T10:30:00Z", or "" on error
 */
export function writeW3cDatetime(date: Date, throwOnError?: boolean): string {
	try {
		// Strip fractional seconds (.000) from ISO string
		return date.toISOString().replace(/\.\d*/, "");
	} catch (error) {
		if (throwOnError) {
			throw error;
		}
	}
	return "";
}

/**
 * Write an OPC variant-type (vt:) XML element for a given JavaScript value.
 * Dispatches on the runtime type to choose the appropriate vt: element:
 * - string  -> vt:lpwstr
 * - number  -> vt:i4 (integer) or vt:r8 (float)
 * - boolean -> vt:bool
 * - Date    -> vt:filetime
 * @param value - the value to serialize
 * @param xlsx - when true, escape double-quotes as _x0022_ (OOXML convention)
 * @returns an XML string containing the appropriate vt: element
 * @throws if value is an unsupported type
 */
export function writeVariantType(value: any, xlsx?: boolean): string {
	switch (typeof value) {
		case "string": {
			let output = writeXmlElement("vt:lpwstr", escapeXml(value));
			if (xlsx) {
				// OOXML uses _x0022_ instead of &quot; for double quotes in custom properties
				output = output.replace(/&quot;/g, "_x0022_");
			}
			return output;
		}
		case "number":
			// Use vt:i4 for integers (bitwise OR test), vt:r8 for floating-point
			return writeXmlElement((value | 0) === value ? "vt:i4" : "vt:r8", escapeXml(String(value)));
		case "boolean":
			return writeXmlElement("vt:bool", value ? "true" : "false");
	}
	if (value instanceof Date) {
		return writeXmlElement("vt:filetime", writeW3cDatetime(value));
	}
	throw new Error("Unable to serialize " + value);
}
