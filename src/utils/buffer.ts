const encoder = new TextEncoder();
const decoder = new TextDecoder();

/**
 * Encode a string as a UTF-8 Uint8Array using the platform TextEncoder.
 * @param s - String to encode
 * @returns UTF-8 encoded byte array
 */
export function utf8encode(s: string): Uint8Array {
	return encoder.encode(s);
}

/**
 * Decode a UTF-8 Uint8Array to a string using the platform TextDecoder.
 * @param data - UTF-8 encoded byte array
 * @returns Decoded string
 */
export function utf8decode(data: Uint8Array): string {
	return decoder.decode(data);
}

/** Regex matching NUL (U+0000) characters globally */

export const NULL_CHAR_REGEX = /\0/g;

/** Regex matching control characters U+0001 through U+0006 globally */
// eslint-disable-next-line no-control-regex
export const CONTROL_CHAR_REGEX = /[\u0001-\u0006]/g;
