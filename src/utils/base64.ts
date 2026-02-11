/**
 * Decode a base64 string to a Uint8Array.
 *
 * Automatically strips a data-URI prefix (e.g. "data:application/octet-stream;base64,...")
 * if present before decoding.
 *
 * @param input - Base64-encoded string, optionally prefixed with a data URI scheme
 * @returns Decoded byte array
 */
export function base64decode(input: string): Uint8Array {
	let str = input;
	// Strip data-URI prefix if present (e.g. "data:image/png;base64,...")
	if (str.slice(0, 5) === "data:") {
		const i = str.slice(0, 1024).indexOf(";base64,");
		if (i > -1) {
			str = str.slice(i + 8);
		}
	}
	const binaryStr = atob(str);
	const len = binaryStr.length;
	const bytes = new Uint8Array(len);
	for (let i = 0; i < len; i++) {
		bytes[i] = binaryStr.charCodeAt(i);
	}
	return bytes;
}

/**
 * Encode a Uint8Array to a base64 string.
 *
 * Converts each byte to a character and uses the built-in btoa() for encoding.
 *
 * @param data - Byte array to encode
 * @returns Base64-encoded string
 */
export function base64encode(data: Uint8Array): string {
	let binaryStr = "";
	for (let i = 0; i < data.length; i++) {
		binaryStr += String.fromCharCode(data[i]);
	}
	return btoa(binaryStr);
}
