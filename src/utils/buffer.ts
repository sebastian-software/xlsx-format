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

/**
 * Decode a UTF-8 binary string (where each character's charCode is a byte value)
 * into a proper JavaScript string.
 *
 * This manually implements UTF-8 decoding for strings that store raw byte values
 * as character codes (common in legacy binary formats).
 *
 * UTF-8 byte patterns:
 *   - 0xxxxxxx (0-127): single-byte ASCII
 *   - 110xxxxx 10xxxxxx (192-223): two-byte sequence
 *   - 1110xxxx 10xxxxxx 10xxxxxx (224-239): three-byte sequence
 *   - 11110xxx 10xxxxxx 10xxxxxx 10xxxxxx (240-247): four-byte sequence (produces surrogate pair)
 *
 * @param orig - Binary string with byte values as char codes
 * @returns Properly decoded JavaScript string
 */
export function utf8read(orig: string): string {
	let out = "";
	let i = 0;
	let byte1 = 0;
	let byte2 = 0;
	let byte3 = 0;
	let byte4 = 0;
	let codePoint = 0;
	while (i < orig.length) {
		byte1 = orig.charCodeAt(i++);
		if (byte1 < 128) {
			// Single-byte ASCII character
			out += String.fromCharCode(byte1);
			continue;
		}
		byte2 = orig.charCodeAt(i++);
		if (byte1 > 191 && byte1 < 224) {
			// Two-byte sequence: 110xxxxx 10xxxxxx
			codePoint = ((byte1 & 31) << 6) | (byte2 & 63);
			out += String.fromCharCode(codePoint);
			continue;
		}
		byte3 = orig.charCodeAt(i++);
		if (byte1 < 240) {
			// Three-byte sequence: 1110xxxx 10xxxxxx 10xxxxxx
			out += String.fromCharCode(((byte1 & 15) << 12) | ((byte2 & 63) << 6) | (byte3 & 63));
			continue;
		}
		// Four-byte sequence: 11110xxx 10xxxxxx 10xxxxxx 10xxxxxx
		// Produces a code point above U+FFFF, requiring a UTF-16 surrogate pair
		byte4 = orig.charCodeAt(i++);
		codePoint = (((byte1 & 7) << 18) | ((byte2 & 63) << 12) | ((byte3 & 63) << 6) | (byte4 & 63)) - 65536;
		// High surrogate: 0xD800 + upper 10 bits
		out += String.fromCharCode(0xd800 + ((codePoint >>> 10) & 1023));
		// Low surrogate: 0xDC00 + lower 10 bits
		out += String.fromCharCode(0xdc00 + (codePoint & 1023));
	}
	return out;
}

/** Regex matching NUL (U+0000) characters globally */
// eslint-disable-next-line no-control-regex
export const NULL_CHAR_REGEX = /\u0000/g;

/** Regex matching control characters U+0001 through U+0006 globally */
// eslint-disable-next-line no-control-regex
export const CONTROL_CHAR_REGEX = /[\u0001-\u0006]/g;
