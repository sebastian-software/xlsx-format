import { XML_HEADER, parseXmlTag } from "../xml/parser.js";
import { unescapeXml, escapeXml } from "../xml/escape.js";
import { parseXmlBoolean, stripNamespace } from "../xml/parser.js";
import { writeXmlElement, writeVariantType } from "../xml/writer.js";
import { XMLNS } from "../xml/namespaces.js";

const custregex = /<[^<>]+>[^<]*/g;

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
						const tokens = tagStr.split(">");
						const type = tokens[0].slice(4);
						const text = tokens[1];
						switch (type) {
							case "lpstr":
							case "bstr":
							case "lpwstr":
								p[name] = unescapeXml(text);
								break;
							case "bool":
								p[name] = parseXmlBoolean(text);
								break;
							case "i1":
							case "i2":
							case "i4":
							case "i8":
							case "int":
							case "uint":
								p[name] = parseInt(text, 10);
								break;
							case "r4":
							case "r8":
							case "decimal":
								p[name] = parseFloat(text);
								break;
							case "filetime":
							case "date":
								p[name] = new Date(text);
								break;
							case "cy":
							case "error":
								p[name] = unescapeXml(text);
								break;
							default:
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
	let pid = 1;
	for (const propName of Object.keys(cp)) {
		++pid;
		lines.push(
			writeXmlElement("property", writeVariantType(cp[propName], true), {
				fmtid: "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}",
				pid: String(pid),
				name: escapeXml(propName),
			}),
		);
	}
	if (lines.length > 2) {
		lines.push("</Properties>");
		lines[1] = lines[1].replace("/>", ">");
	}
	return lines.join("");
}
