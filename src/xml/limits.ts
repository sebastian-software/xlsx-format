import { XlsxError } from "../errors.js";

/** Limits applied while reading XML-based workbook parts. */
export interface XmlLimitOptions {
	/** Maximum decoded XML part size. */
	maxXmlPartBytes?: number;
	/** Maximum number of raw XML tags in a single part. */
	maxXmlTags?: number;
	/** Maximum nesting depth for XML elements in a single part. */
	maxXmlNestingDepth?: number;
	/** Maximum number of characters in a single XML tag. */
	maxXmlTagLength?: number;
	/** Maximum number of attributes in a single XML tag. */
	maxXmlAttributesPerTag?: number;
	/** Maximum number of shared string items to parse. */
	maxSharedStringItems?: number;
	/** Maximum number of worksheet row elements to scan. */
	maxWorksheetRows?: number;
	/** Maximum number of worksheet cell elements to scan. */
	maxWorksheetCells?: number;
}

export const DEFAULT_MAX_XML_PART_BYTES = 128 * 1024 * 1024;
export const DEFAULT_MAX_XML_TAGS = 5_000_000;
export const DEFAULT_MAX_XML_NESTING_DEPTH = 256;
export const DEFAULT_MAX_XML_TAG_LENGTH = 1024 * 1024;
export const DEFAULT_MAX_XML_ATTRIBUTES_PER_TAG = 10_000;
export const DEFAULT_MAX_SHARED_STRING_ITEMS = 1_000_000;
export const DEFAULT_MAX_WORKSHEET_ROWS = 1_048_576;
export const DEFAULT_MAX_WORKSHEET_CELLS = 10_000_000;

export function xmlOptionLimit(value: number | undefined, fallback: number, name: string): number {
	if (value == null) {
		return fallback;
	}
	if (!Number.isFinite(value) || value < 0) {
		throw new XlsxError("INVALID_ARGUMENT", `Invalid XML option: ${name} must be a non-negative finite number`);
	}
	return value;
}

export function assertXmlCountWithinLimit(kind: string, count: number, limit: number): void {
	if (count > limit) {
		throw new XlsxError("LIMIT_EXCEEDED", `Invalid XML: ${kind} count ${count} exceeds limit ${limit}`);
	}
}

export function assertXmlPartLimits(partName: string, data: string, opts?: XmlLimitOptions): void {
	const maxXmlPartBytes = xmlOptionLimit(opts?.maxXmlPartBytes, DEFAULT_MAX_XML_PART_BYTES, "maxXmlPartBytes");
	if (data.length > maxXmlPartBytes) {
		throw new XlsxError(
			"LIMIT_EXCEEDED",
			`Invalid XML: ${partName} size ${data.length} exceeds limit ${maxXmlPartBytes}`,
		);
	}

	const maxXmlTags = xmlOptionLimit(opts?.maxXmlTags, DEFAULT_MAX_XML_TAGS, "maxXmlTags");
	const maxXmlNestingDepth = xmlOptionLimit(
		opts?.maxXmlNestingDepth,
		DEFAULT_MAX_XML_NESTING_DEPTH,
		"maxXmlNestingDepth",
	);
	const maxXmlTagLength = xmlOptionLimit(opts?.maxXmlTagLength, DEFAULT_MAX_XML_TAG_LENGTH, "maxXmlTagLength");
	let tagCount = 0;
	let depth = 0;
	let offset = -1;
	while ((offset = data.indexOf("<", offset + 1)) !== -1) {
		assertXmlCountWithinLimit(`${partName} tag`, ++tagCount, maxXmlTags);
		const closeOffset = data.indexOf(">", offset + 1);
		const tagLength = closeOffset === -1 ? data.length - offset : closeOffset - offset + 1;
		if (tagLength > maxXmlTagLength) {
			throw new XlsxError(
				"LIMIT_EXCEEDED",
				`Invalid XML: ${partName} tag length ${tagLength} exceeds limit ${maxXmlTagLength}`,
			);
		}
		if (closeOffset === -1) {
			break;
		}
		let tagStart = offset + 1;
		while (tagStart < closeOffset && data.charCodeAt(tagStart) <= 32) {
			++tagStart;
		}
		const first = data.charCodeAt(tagStart);
		const isClosingTag = first === 47;
		const isProcessingOrDeclaration = first === 33 || first === 63;
		const isSelfClosing = data.charCodeAt(closeOffset - 1) === 47;
		if (isClosingTag) {
			depth = Math.max(0, depth - 1);
		} else if (!isProcessingOrDeclaration && !isSelfClosing && ++depth > maxXmlNestingDepth) {
			throw new XlsxError(
				"LIMIT_EXCEEDED",
				`Invalid XML: ${partName} nesting depth ${depth} exceeds limit ${maxXmlNestingDepth}`,
			);
		}
		offset = closeOffset;
	}
}
