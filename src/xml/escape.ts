const encodings: Record<string, string> = {
	"&quot;": '"',
	"&apos;": "'",
	"&gt;": ">",
	"&lt;": "<",
	"&amp;": "&",
};

const rencoding: Record<string, string> = {
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

function raw_unescapexml(text: string): string {
	const s = text;
	const i = s.indexOf("<![CDATA[");
	if (i === -1) {
		return s
			.replace(encregex, ($$, $1) => {
				return encodings[$$] || String.fromCharCode(parseInt($1, $$.indexOf("x") > -1 ? 16 : 10)) || $$;
			})
			.replace(coderegex, (_m, c) => {
				return String.fromCharCode(parseInt(c, 16));
			});
	}
	const j = s.indexOf("]]>");
	return raw_unescapexml(s.slice(0, i)) + s.slice(i + 9, j) + raw_unescapexml(s.slice(j + 3));
}

/** Unescape XML entities. When xlsx=true, normalize \r\n -> \n */
export function unescapexml(text: string, xlsx?: boolean): string {
	const out = raw_unescapexml(text);
	return xlsx ? out.replace(/\r\n/g, "\n") : out;
}

/** Escape a string for XML text content */
export function escapexml(text: string): string {
	const s = text;
	return s
		.replace(decregex, (y) => rencoding[y])
		.replace(charegex, (s) => "_x" + ("000" + s.charCodeAt(0).toString(16)).slice(-4) + "_");
}

/** Escape a string for XML tag names (spaces -> _x0020_) */
export function escapexmltag(text: string): string {
	return escapexml(text).replace(/ /g, "_x0020_");
}

// eslint-disable-next-line no-control-regex
const htmlcharegex = /[\u0000-\u001f]/g;

/** Escape a string for HTML output */
export function escapehtml(text: string): string {
	const s = text;
	return s
		.replace(decregex, (y) => rencoding[y])
		.replace(/\n/g, "<br/>")
		.replace(htmlcharegex, (s) => "&#x" + ("000" + s.charCodeAt(0).toString(16)).slice(-4) + ";");
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
export function htmldecode(str: string): string {
	let o = str
		.replace(/^[\t\n\r ]+/, "")
		.replace(/(^|[^\t\n\r ])[\t\n\r ]+$/, "$1")
		.replace(/>\s+/g, ">")
		.replace(/\b\s+</g, "<")
		.replace(/[\t\n\r ]+/g, " ")
		.replace(/<\s*[bB][rR]\s*\/?>/g, "\n")
		.replace(/<[^<>]*>/g, "");
	for (let i = 0; i < entities.length; ++i) {
		o = o.replace(entities[i][0], entities[i][1]);
	}
	return o;
}
