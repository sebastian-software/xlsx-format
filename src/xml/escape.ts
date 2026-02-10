const encodings: Record<string, string> = {
	"&quot;": '"',
	"&apos;": "'",
	"&gt;": ">",
	"&lt;": "<",
	"&amp;": "&",
};

const XML_ESCAPE_MAP: Record<string, string> = {
	'"': "&quot;",
	"'": "&apos;",
	">": "&gt;",
	"<": "&lt;",
	"&": "&amp;",
};

const encregex = /&(?:quot|apos|gt|lt|amp|#x?([\da-fA-F]+));/gi;
const coderegex = /_x([\da-fA-F]{4})_/g;
const decregex = /[&<>'"]/g;
// eslint-disable-next-line no-control-regex
const charegex = /[\u0000-\u0008\u000b-\u001f\uFFFE-\uFFFF]/g;

function rawUnescapeXml(text: string): string {
	const str = text;
	const i = str.indexOf("<![CDATA[");
	if (i === -1) {
		return str
			.replace(encregex, ($$, $1) => {
				return encodings[$$] || String.fromCharCode(parseInt($1, $$.indexOf("x") > -1 ? 16 : 10)) || $$;
			})
			.replace(coderegex, (_m, c) => {
				return String.fromCharCode(parseInt(c, 16));
			});
	}
	const cdataEndIdx = str.indexOf("]]>");
	return rawUnescapeXml(str.slice(0, i)) + str.slice(i + 9, cdataEndIdx) + rawUnescapeXml(str.slice(cdataEndIdx + 3));
}

/** Unescape XML entities. When xlsx=true, normalize \r\n -> \n */
export function unescapeXml(text: string, xlsx?: boolean): string {
	const out = rawUnescapeXml(text);
	return xlsx ? out.replace(/\r\n/g, "\n") : out;
}

/** Escape a string for XML text content */
export function escapeXml(text: string): string {
	const str = text;
	return str
		.replace(decregex, (char) => XML_ESCAPE_MAP[char])
		.replace(charegex, (char) => "_x" + ("000" + char.charCodeAt(0).toString(16)).slice(-4) + "_");
}

/** Escape a string for XML tag names (spaces -> _x0020_) */
export function escapeXmlTag(text: string): string {
	return escapeXml(text).replace(/ /g, "_x0020_");
}

// eslint-disable-next-line no-control-regex
const htmlcharegex = /[\u0000-\u001f]/g;

/** Escape a string for HTML output */
export function escapeHtml(text: string): string {
	const str = text;
	return str
		.replace(decregex, (char) => XML_ESCAPE_MAP[char])
		.replace(/\n/g, "<br/>")
		.replace(htmlcharegex, (char) => "&#x" + ("000" + char.charCodeAt(0).toString(16)).slice(-4) + ";");
}

const entities: [RegExp, string][] = [
	["nbsp", " "],
	["middot", "\u00B7"],
	["quot", '"'],
	["apos", "'"],
	["gt", ">"],
	["lt", "<"],
	["amp", "&"],
].map(([name, ch]) => [new RegExp("&" + name + ";", "gi"), ch]);

/** Decode HTML entities */
export function htmlDecode(str: string): string {
	let result = str
		.replace(/^[\t\n\r ]+/, "")
		.replace(/(^|[^\t\n\r ])[\t\n\r ]+$/, "$1")
		.replace(/>\s+/g, ">")
		.replace(/\b\s+</g, "<")
		.replace(/[\t\n\r ]+/g, " ")
		.replace(/<\s*[bB][rR]\s*\/?>/g, "\n")
		.replace(/<[^<>]*>/g, "");
	for (let i = 0; i < entities.length; ++i) {
		result = result.replace(entities[i][0], entities[i][1]);
	}
	return result;
}
