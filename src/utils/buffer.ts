const encoder = new TextEncoder();
const decoder = new TextDecoder();

/** Create a new zero-filled Uint8Array */
export function createZeroBuffer(len: number): Uint8Array {
	return new Uint8Array(len);
}

/** Create a new uninitialized Uint8Array */
export function createBuffer(len: number): Uint8Array {
	return new Uint8Array(len);
}

/** Convert a binary string to Uint8Array */
export function binaryStringToUint8Array(str: string): Uint8Array {
	const buf = new Uint8Array(str.length);
	for (let i = 0; i < str.length; ++i) {
		buf[i] = str.charCodeAt(i) & 0xff;
	}
	return buf;
}

/** Convert Uint8Array to binary string */
export function uint8ArrayToBinaryString(data: Uint8Array | number[]): string {
	const result: string[] = [];
	for (let i = 0; i < data.length; ++i) {
		result[i] = String.fromCharCode(data[i]);
	}
	return result.join("");
}

/** Concatenate multiple Uint8Arrays */
export function concatUint8Arrays(bufs: Uint8Array[]): Uint8Array {
	let maxlen = 0;
	for (let i = 0; i < bufs.length; ++i) {
		maxlen += bufs[i].length;
	}
	const result = new Uint8Array(maxlen);
	let offset = 0;
	for (let i = 0; i < bufs.length; ++i) {
		result.set(bufs[i], offset);
		offset += bufs[i].length;
	}
	return result;
}

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

/** Write a string to UTF-8 encoded binary string */
export function utf8write(orig: string): string {
	const out: string[] = [];
	let i = 0;
	let charCode = 0;
	let lowSurrogate = 0;
	while (i < orig.length) {
		charCode = orig.charCodeAt(i++);
		if (charCode < 128) {
			out.push(String.fromCharCode(charCode));
		} else if (charCode < 2048) {
			out.push(String.fromCharCode(192 + (charCode >> 6)));
			out.push(String.fromCharCode(128 + (charCode & 63)));
		} else if (charCode >= 55296 && charCode < 57344) {
			charCode -= 55296;
			lowSurrogate = orig.charCodeAt(i++) - 56320 + (charCode << 10);
			out.push(String.fromCharCode(240 + ((lowSurrogate >> 18) & 7)));
			out.push(String.fromCharCode(144 + ((lowSurrogate >> 12) & 63)));
			out.push(String.fromCharCode(128 + ((lowSurrogate >> 6) & 63)));
			out.push(String.fromCharCode(128 + (lowSurrogate & 63)));
		} else {
			out.push(String.fromCharCode(224 + (charCode >> 12)));
			out.push(String.fromCharCode(128 + ((charCode >> 6) & 63)));
			out.push(String.fromCharCode(128 + (charCode & 63)));
		}
	}
	return out.join("");
}

// eslint-disable-next-line no-control-regex
export const NULL_CHAR_REGEX = /\u0000/g;
// eslint-disable-next-line no-control-regex
export const CONTROL_CHAR_REGEX = /[\u0001-\u0006]/g;
