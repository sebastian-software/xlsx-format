import { XML_HEADER, parseXmlTag } from "../xml/parser.js";
import { unescapeXml, escapeXml } from "../xml/escape.js";
import { parseXmlBoolean, stripNamespace } from "../xml/parser.js";
import { writeXmlElement, writeVariantType } from "../xml/writer.js";
import { XMLNS } from "../xml/namespaces.js";

// Matches sequences of: an XML tag followed by optional non-tag text content
const custregex = /<[^<>]+>[^<]*/g;

/**
 * Parse OPC custom properties XML into a key-value record.
 * Custom properties are user-defined metadata stored as typed vt: variant elements
 * (strings, booleans, integers, floats, dates, etc.).
 * @param data - raw XML string of the custom properties part
 * @param opts - optional settings: WTF enables console warnings for unknown types
 * @returns a record mapping property names to their parsed values
 */
export function parseCustomProperties(data: string, opts?: { WTF?: boolean }): Record<string, any> {
	const p: Record<string, any> = {};
	let name = "";
	const matches = data.match(custregex);
	if (matches) {
		for (let i = 0; i < matches.length; ++i) {
			const tagStr = matches[i];
			const parsedTag = parseXmlTag(tagStr);
			switch (stripNamespace(parsedTag[0])) {
				case "<?xml":
					break;
				case "<Properties":
					break;
				case "<property":
					name = unescapeXml(parsedTag.name);
					break;
				case "</property>":
					name = "";
					break;
				default:
					if (tagStr.indexOf("<vt:") === 0) {
						// Parse the vt: (variant type) element: split on ">" to get type and text
						const tokens = tagStr.split(">");
						const type = tokens[0].slice(4); // strip "<vt:" prefix
						const text = tokens[1];
						switch (type) {
							// String types
							case "lpstr":
							case "bstr":
							case "lpwstr":
								p[name] = unescapeXml(text);
								break;
							case "bool":
								p[name] = parseXmlBoolean(text);
								break;
							// Integer types (various widths)
							case "i1":
							case "i2":
							case "i4":
							case "i8":
							case "int":
							case "uint":
								p[name] = parseInt(text, 10);
								break;
							// Floating-point types
							case "r4":
							case "r8":
							case "decimal":
								p[name] = parseFloat(text);
								break;
							// Date/time types
							case "filetime":
							case "date":
								p[name] = new Date(text);
								break;
							// Passthrough types (kept as strings)
							case "cy":
							case "error":
								p[name] = unescapeXml(text);
								break;
							default:
								// Self-closing tags (e.g., <vt:empty/>) end with "/"
								if (type.slice(-1) === "/") {
									break;
								}
								if (opts?.WTF && typeof console !== "undefined") {
									console.warn("Unexpected", tagStr, type, tokens);
								}
						}
					}
			}
		}
	}
	return p;
}

/**
 * Serialize custom properties to OPC custom properties XML.
 * Each property is wrapped in a `<property>` element with a unique pid (property ID)
 * and the well-known fmtid GUID for custom properties.
 * @param cp - record of property name-value pairs (may be undefined)
 * @returns the complete XML string for the custom properties part
 */
export function writeCustomProperties(cp: Record<string, any> | undefined): string {
	const lines: string[] = [
		XML_HEADER,
		writeXmlElement("Properties", null, {
			xmlns: XMLNS.CUST_PROPS,
			"xmlns:vt": XMLNS.vt,
		}),
	];
	if (!cp) {
		return lines.join("");
	}
	// Property IDs (pid) start at 2 per the OPC specification (pid=1 is reserved)
	let pid = 1;
	for (const propName of Object.keys(cp)) {
		++pid;
		lines.push(
			writeXmlElement("property", writeVariantType(cp[propName], true), {
				fmtid: "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}", // Standard GUID for custom property sets
				pid: String(pid),
				name: escapeXml(propName),
			}),
		);
	}
	// Only close the root element if child elements were added
	if (lines.length > 2) {
		lines.push("</Properties>");
		// Convert the self-closing root tag to an opening tag
		lines[1] = lines[1].replace("/>", ">");
	}
	return lines.join("");
}
