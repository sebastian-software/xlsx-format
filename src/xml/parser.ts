const attregexg = /\s([^"\s?>/]+)\s*=\s*((?:")([^"]*)(?:")|(?:')([^']*)(?:')|([^'">\s]+))/g;
const tagregex1 = /<[/?]?[a-zA-Z0-9:_-]+(?:\s+[^"\s?<>/]+\s*=\s*(?:"[^"]*"|'[^']*'|[^'"<>\s=]+))*\s*[/?]?>/gm;
const tagregex2 = /<[^<>]*>/g;

export const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

export const XML_TAG_REGEX: RegExp = XML_HEADER.match(tagregex1) ? tagregex1 : tagregex2;

const nsregex2 = /<(\/?)\w+:/;

/** Parse XML attributes from a tag string */
export function parseXmlTag(tag: string, skip_root?: boolean, skip_LC?: boolean): Record<string, any> {
	const attrs: Record<string, any> = {};
	let scanPos = 0;
	let charCode = 0;
	for (; scanPos !== tag.length; ++scanPos) {
		if ((charCode = tag.charCodeAt(scanPos)) === 32 || charCode === 10 || charCode === 13) {
			break;
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
			const attrStr = matches[i].slice(1);
			let eqPos = 0;
			for (eqPos = 0; eqPos < attrStr.length; ++eqPos) {
				if (attrStr.charCodeAt(eqPos) === 61) {
					break;
				}
			}
			let attrName = attrStr.slice(0, eqPos).trim();
			while (attrStr.charCodeAt(eqPos + 1) === 32) {
				++eqPos;
			}
			const quoteOffset = (scanPos = attrStr.charCodeAt(eqPos + 1)) === 34 || scanPos === 39 ? 1 : 0;
			const attrValue = attrStr.slice(eqPos + 1 + quoteOffset, attrStr.length - quoteOffset);
			let colonPos = 0;
			for (colonPos = 0; colonPos < attrName.length; ++colonPos) {
				if (attrName.charCodeAt(colonPos) === 58) {
					break;
				}
			}
			if (colonPos === attrName.length) {
				if (attrName.indexOf("_") > 0) {
					attrName = attrName.slice(0, attrName.indexOf("_"));
				}
				attrs[attrName] = attrValue;
				if (!skip_LC) {
					attrs[attrName.toLowerCase()] = attrValue;
				}
			} else {
				const localName = (colonPos === 5 && attrName.slice(0, 5) === "xmlns" ? "xmlns" : "") + attrName.slice(colonPos + 1);
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
