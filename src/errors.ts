/**
 * Stable failure categories exposed by {@link XlsxError}.
 *
 * Codes are intentionally broader than individual messages so callers can
 * handle failures without depending on human-readable text.
 */
export type XlsxErrorCode =
	| "INVALID_ARGUMENT"
	| "MALFORMED"
	| "LIMIT_EXCEEDED"
	| "UNSUPPORTED"
	| "CRC_MISMATCH"
	| "NOT_FOUND"
	| "DUPLICATE";

/** Error subclass thrown by xlsx-format for deterministic failure handling. */
export class XlsxError extends Error {
	readonly code: XlsxErrorCode;

	constructor(code: XlsxErrorCode, message: string, options?: ErrorOptions) {
		super(message, options);
		this.name = "XlsxError";
		this.code = code;
		Object.setPrototypeOf(this, new.target.prototype);
	}
}
