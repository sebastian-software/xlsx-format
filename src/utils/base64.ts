/** Decode a base64 string to Uint8Array */
export function base64decode(input: string): Uint8Array {
	let str = input;
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

/** Encode a Uint8Array to base64 string */
export function base64encode(data: Uint8Array): string {
	let binaryStr = "";
	for (let i = 0; i < data.length; i++) {
		binaryStr += String.fromCharCode(data[i]);
	}
	return btoa(binaryStr);
}
