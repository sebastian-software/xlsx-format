const attregexg = /\s([^"\s?>/]+)\s*=\s*((?:")([^"]*)(?:")|(?:')([^']*)(?:')|([^'">\s]+))/g;
const tagregex1 = /<[/?]?[a-zA-Z0-9:_-]+(?:\s+[^"\s?<>/]+\s*=\s*(?:"[^"]*"|'[^']*'|[^'"<>\s=]+))*\s*[/?]?>/gm;
const tagregex2 = /<[^<>]*>/g;

export const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

export const XML_TAG_REGEX: RegExp = XML_HEADER.match(tagregex1) ? tagregex1 : tagregex2;

const nsregex2 = /<(\/?)\w+:/;

/** Parse XML attributes from a tag string */
export function parseXmlTag(tag: string, skip_root?: boolean, skip_LC?: boolean): Record<string, any> {
	const z: Record<string, any> = {};
	let eq = 0;
	let c = 0;
	for (; eq !== tag.length; ++eq) {
		if ((c = tag.charCodeAt(eq)) === 32 || c === 10 || c === 13) {
			break;
		}
	}
	if (!skip_root) {
		z[0] = tag.slice(0, eq);
	}
	if (eq === tag.length) {
		return z;
	}
	const m = tag.match(attregexg);
	if (m) {
		for (let i = 0; i < m.length; ++i) {
			const cc = m[i].slice(1);
			let c2 = 0;
			for (c2 = 0; c2 < cc.length; ++c2) {
				if (cc.charCodeAt(c2) === 61) {
					break;
				}
			}
			let q = cc.slice(0, c2).trim();
			while (cc.charCodeAt(c2 + 1) === 32) {
				++c2;
			}
			const quot = (eq = cc.charCodeAt(c2 + 1)) === 34 || eq === 39 ? 1 : 0;
			const v = cc.slice(c2 + 1 + quot, cc.length - quot);
			let j = 0;
			for (j = 0; j < q.length; ++j) {
				if (q.charCodeAt(j) === 58) {
					break;
				}
			}
			if (j === q.length) {
				if (q.indexOf("_") > 0) {
					q = q.slice(0, q.indexOf("_"));
				}
				z[q] = v;
				if (!skip_LC) {
					z[q.toLowerCase()] = v;
				}
			} else {
				const k = (j === 5 && q.slice(0, 5) === "xmlns" ? "xmlns" : "") + q.slice(j + 1);
				if (z[k] && q.slice(j - 3, j) === "ext") {
					continue;
				}
				z[k] = v;
				if (!skip_LC) {
					z[k.toLowerCase()] = v;
				}
			}
		}
	}
	return z;
}

/** Strip namespace prefixes from XML */
export function stripNamespace(x: string): string {
	return x.replace(nsregex2, "<$1");
}

/** Parse xsd:boolean values */
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
