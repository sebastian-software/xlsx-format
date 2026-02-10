import { XML_HEADER, parsexmltag } from "../xml/parser.js";
import { unescapexml, escapexml } from "../xml/escape.js";
import { parsexmlbool, strip_ns } from "../xml/parser.js";
import { writextag, write_vt } from "../xml/writer.js";
import { XMLNS } from "../xml/namespaces.js";

const custregex = /<[^<>]+>[^<]*/g;

export function parse_cust_props(data: string, opts?: { WTF?: boolean }): Record<string, any> {
	const p: Record<string, any> = {};
	let name = "";
	const m = data.match(custregex);
	if (m) {
		for (let i = 0; i < m.length; ++i) {
			const x = m[i];
			const y = parsexmltag(x);
			switch (strip_ns(y[0])) {
				case "<?xml":
					break;
				case "<Properties":
					break;
				case "<property":
					name = unescapexml(y.name);
					break;
				case "</property>":
					name = "";
					break;
				default:
					if (x.indexOf("<vt:") === 0) {
						const toks = x.split(">");
						const type = toks[0].slice(4);
						const text = toks[1];
						switch (type) {
							case "lpstr":
							case "bstr":
							case "lpwstr":
								p[name] = unescapexml(text);
								break;
							case "bool":
								p[name] = parsexmlbool(text);
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
								p[name] = unescapexml(text);
								break;
							default:
								if (type.slice(-1) === "/") {
									break;
								}
								if (opts?.WTF && typeof console !== "undefined") {
									console.warn("Unexpected", x, type, toks);
								}
						}
					}
			}
		}
	}
	return p;
}

export function write_cust_props(cp: Record<string, any> | undefined): string {
	const o: string[] = [
		XML_HEADER,
		writextag("Properties", null, {
			xmlns: XMLNS.CUST_PROPS,
			"xmlns:vt": XMLNS.vt,
		}),
	];
	if (!cp) {
		return o.join("");
	}
	let pid = 1;
	for (const k of Object.keys(cp)) {
		++pid;
		o.push(
			writextag("property", write_vt(cp[k], true), {
				fmtid: "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}",
				pid: String(pid),
				name: escapexml(k),
			}),
		);
	}
	if (o.length > 2) {
		o.push("</Properties>");
		o[1] = o[1].replace("/>", ">");
	}
	return o.join("");
}
