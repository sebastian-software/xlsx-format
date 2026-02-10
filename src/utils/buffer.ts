const encoder = new TextEncoder();
const decoder = new TextDecoder();

/** Encode a string as UTF-8 Uint8Array */
export function utf8encode(s: string): Uint8Array {
	return encoder.encode(s);
}

/** Decode UTF-8 Uint8Array to string */
export function utf8decode(data: Uint8Array): string {
	return decoder.decode(data);
}

/** Read a UTF-8 encoded binary string (charCode-based) into a proper string */
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
			out += String.fromCharCode(byte1);
			continue;
		}
		byte2 = orig.charCodeAt(i++);
		if (byte1 > 191 && byte1 < 224) {
			codePoint = ((byte1 & 31) << 6) | (byte2 & 63);
			out += String.fromCharCode(codePoint);
			continue;
		}
		byte3 = orig.charCodeAt(i++);
		if (byte1 < 240) {
			out += String.fromCharCode(((byte1 & 15) << 12) | ((byte2 & 63) << 6) | (byte3 & 63));
			continue;
		}
		byte4 = orig.charCodeAt(i++);
		codePoint = (((byte1 & 7) << 18) | ((byte2 & 63) << 12) | ((byte3 & 63) << 6) | (byte4 & 63)) - 65536;
		out += String.fromCharCode(0xd800 + ((codePoint >>> 10) & 1023));
		out += String.fromCharCode(0xdc00 + (codePoint & 1023));
	}
	return out;
}

// eslint-disable-next-line no-control-regex
export const NULL_CHAR_REGEX = /\u0000/g;
// eslint-disable-next-line no-control-regex
export const CONTROL_CHAR_REGEX = /[\u0001-\u0006]/g;
