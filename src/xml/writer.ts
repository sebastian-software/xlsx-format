import { escapeXml } from "./escape.js";

const wtregex = /(^\s|\s$|\n)/;

/** Write an XML tag with text content */
export function writeXmlTag(tagName: string, content: string): string {
	return "<" + tagName + (content.match(wtregex) ? ' xml:space="preserve"' : "") + ">" + content + "</" + tagName + ">";
}

function formatXmlAttributes(attributes: Record<string, string>): string {
	return Object.keys(attributes)
		.map((key) => " " + key + '="' + attributes[key] + '"')
		.join("");
}

/** Write an XML tag with optional attributes and content */
export function writeXmlElement(tagName: string, content?: string | null, attributes?: Record<string, string> | null): string {
	return (
		"<" +
		tagName +
		(attributes != null ? formatXmlAttributes(attributes) : "") +
		(content != null ? (content.match(wtregex) ? ' xml:space="preserve"' : "") + ">" + content + "</" + tagName : "/") +
		">"
	);
}

/** Write a W3C datetime string from a Date */
export function writeW3cDatetime(date: Date, throwOnError?: boolean): string {
	try {
		return date.toISOString().replace(/\.\d*/, "");
	} catch (error) {
		if (throwOnError) {
			throw error;
		}
	}
	return "";
}

/** Write a vt: (variant type) XML element */
export function writeVariantType(value: any, xlsx?: boolean): string {
	switch (typeof value) {
		case "string": {
			let output = writeXmlElement("vt:lpwstr", escapeXml(value));
			if (xlsx) {
				output = output.replace(/&quot;/g, "_x0022_");
			}
			return output;
		}
		case "number":
			return writeXmlElement((value | 0) === value ? "vt:i4" : "vt:r8", escapeXml(String(value)));
		case "boolean":
			return writeXmlElement("vt:bool", value ? "true" : "false");
	}
	if (value instanceof Date) {
		return writeXmlElement("vt:filetime", writeW3cDatetime(value));
	}
	throw new Error("Unable to serialize " + value);
}
