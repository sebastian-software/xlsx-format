/** Lookup table mapping XML named entities to their decoded characters */
const encodings: Record<string, string> = {
	"&quot;": '"',
	"&apos;": "'",
	"&gt;": ">",
	"&lt;": "<",
	"&amp;": "&",
};

/** Reverse lookup: special characters to their XML entity representations */
const XML_ESCAPE_MAP: Record<string, string> = {
	'"': "&quot;",
	"'": "&apos;",
	">": "&gt;",
	"<": "&lt;",
	"&": "&amp;",
};

// Matches XML named entities (&quot; etc.) and numeric character references (&#x1A; or &#26;)
const encregex = /&(?:quot|apos|gt|lt|amp|#(?:x([\da-f]+)|(\d+)));/gi;
// Matches OOXML-style escaped Unicode characters like _x0020_ (underscore-hex encoding)
const coderegex = /_x([\da-fA-F]{4})_/g;
// Matches the leading underscore of literal text that would otherwise be interpreted as an OOXML escape on read.
const literalCoderegex = /_(?=x[\da-fA-F]{4}_)/g;
// Matches characters that must be escaped in XML: & < > ' "
const decregex = /[&<>'"]/g;
// Matches XML-illegal control characters (U+0000-U+0008, U+000B-U+001F, U+FFFE-U+FFFF)
// eslint-disable-next-line no-control-regex
const charegex = /[\u0000-\u0008\u000b-\u001f\uFFFE\uFFFF]/g;

/**
 * Check whether a number is a character allowed by the XML 1.0 Char production.
 * XML excludes surrogate code points and values outside the listed ranges.
 */
function isValidXmlCodePoint(codePoint: number): boolean {
	return (
		codePoint === 0x9 ||
		codePoint === 0xa ||
		codePoint === 0xd ||
		(codePoint >= 0x20 && codePoint <= 0xd7ff) ||
		(codePoint >= 0xe000 && codePoint <= 0xfffd) ||
		(codePoint >= 0x10000 && codePoint <= 0x10ffff)
	);
}

/**
 * Recursively unescape XML entities, handling CDATA sections.
 * @param text - raw XML text potentially containing entities and CDATA blocks
 * @returns the fully unescaped string
 */
function rawUnescapeXml(text: string): string {
	const str = text;
	const i = str.indexOf("<![CDATA[");
	if (i === -1) {
		// No CDATA: replace named/numeric entities then OOXML _xHHHH_ escapes
		return str
			.replace(encregex, ($$, $1, $2) => {
				// Try named entity first, then parse as hex (&#xHH;) or decimal (&#DD;)
				if (encodings[$$]) {
					return encodings[$$];
				}
				const codePoint = Number.parseInt($1 || $2, $1 ? 16 : 10);
				return isValidXmlCodePoint(codePoint) ? String.fromCodePoint(codePoint) : $$;
			})
			.replace(coderegex, (_m, c) => {
				return String.fromCharCode(parseInt(c, 16));
			});
	}
	// Split around CDATA: unescape text before and after, keep CDATA content verbatim
	const cdataEndIdx = str.indexOf("]]>");
	return rawUnescapeXml(str.slice(0, i)) + str.slice(i + 9, cdataEndIdx) + rawUnescapeXml(str.slice(cdataEndIdx + 3));
}

/**
 * Unescape XML entities in a string. Optionally normalize line endings for XLSX.
 * @param text - XML-encoded text to unescape
 * @param xlsx - when true, normalize \r\n line endings to \n
 * @returns the unescaped string
 */
export function unescapeXml(text: string, xlsx?: boolean): string {
	const out = rawUnescapeXml(text);
	return xlsx ? out.replace(/\r\n/g, "\n") : out;
}

/**
 * Escape a string for safe inclusion as XML text content.
 * Replaces &, <, >, ', " with named entities, and encodes illegal
 * control characters using OOXML _xHHHH_ notation.
 * @param text - plain text to escape
 * @returns XML-safe string
 */
export function escapeXml(text: string): string {
	const str = text;
	return (
		str
			.replace(decregex, (char) => XML_ESCAPE_MAP[char])
			// A literal _xHHHH_ would be decoded as an OOXML escape during a roundtrip.
			// Escaping its leading underscore as _x005F_ makes the reader decode it once.
			.replace(literalCoderegex, () => "_x005F_")
			.replace(charegex, (char) => "_x" + ("000" + char.charCodeAt(0).toString(16)).slice(-4) + "_")
	);
}

/**
 * Escape a string for use in XML tag names.
 * In addition to standard XML escaping, spaces are encoded as _x0020_.
 * @param text - raw tag name text
 * @returns escaped tag name suitable for XML
 */
export function escapeXmlTag(text: string): string {
	return escapeXml(text).replace(/ /g, "_x0020_");
}

// Matches control characters U+0000-U+001F (all ASCII control chars including tab, newline, etc.)
// eslint-disable-next-line no-control-regex
const htmlcharegex = /[\u0000-\u001f]/g;

/**
 * Escape a string for HTML output.
 * Replaces XML-special characters with entities, converts newlines to <br/>,
 * and encodes remaining control characters as hex character references.
 * @param text - plain text to escape for HTML
 * @returns HTML-safe string
 */
export function escapeHtml(text: string): string {
	const str = text;
	return str
		.replace(decregex, (char) => XML_ESCAPE_MAP[char])
		.replace(/\n/g, "<br/>")
		.replace(htmlcharegex, (char) => "&#x" + ("000" + char.charCodeAt(0).toString(16)).slice(-4) + ";");
}

/** Pre-compiled entity patterns paired with their replacement characters for HTML decoding */
const entities: [RegExp, string][] = [
	["nbsp", " "],
	["middot", "\u00B7"],
	["quot", '"'],
	["apos", "'"],
	["gt", ">"],
	["lt", "<"],
	["amp", "&"], // &amp; must be last so earlier replacements don't get double-decoded
].map(([name, ch]) => [new RegExp("&" + name + ";", "gi"), ch]);

/**
 * Decode an HTML string to plain text.
 * Strips tags, collapses whitespace, converts <br> to newlines,
 * and decodes named HTML entities.
 * @param str - HTML string to decode
 * @returns decoded plain text
 */
export function htmlDecode(str: string): string {
	let result = str
		.replace(/^[\t\n\r ]+/, "") // trim leading whitespace
		.replace(/(^|[^\t\n\r ])[\t\n\r ]+$/, "$1") // trim trailing whitespace (preserve last non-ws char)
		.replace(/>\s+/g, ">") // collapse whitespace after closing brackets
		.replace(/\b\s+</g, "<") // collapse whitespace before opening brackets
		.replace(/[\t\n\r ]+/g, " ") // normalize internal whitespace to single spaces
		.replace(/<\s*br\s*\/?>/gi, "\n") // convert <br> / <BR/> variants to newlines
		.replace(/<[^<>]*>/g, ""); // strip all remaining HTML tags
	for (let i = 0; i < entities.length; ++i) {
		result = result.replace(entities[i][0], entities[i][1]);
	}
	return result;
}
