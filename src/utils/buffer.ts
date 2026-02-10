const encoder = new TextEncoder();
const decoder = new TextDecoder();

/** Create a new zero-filled Uint8Array */
export function new_buf(len: number): Uint8Array {
	return new Uint8Array(len);
}

/** Create a new uninitialized Uint8Array */
export function new_unsafe_buf(len: number): Uint8Array {
	return new Uint8Array(len);
}

/** Convert a binary string to Uint8Array */
export function s2a(s: string): Uint8Array {
	const buf = new Uint8Array(s.length);
	for (let i = 0; i < s.length; ++i) {
		buf[i] = s.charCodeAt(i) & 0xff;
	}
	return buf;
}

/** Convert Uint8Array to binary string */
export function a2s(data: Uint8Array | number[]): string {
	const o: string[] = [];
	for (let i = 0; i < data.length; ++i) {
		o[i] = String.fromCharCode(data[i]);
	}
	return o.join("");
}

/** Concatenate multiple Uint8Arrays */
export function bconcat(bufs: Uint8Array[]): Uint8Array {
	let maxlen = 0;
	for (let i = 0; i < bufs.length; ++i) {
		maxlen += bufs[i].length;
	}
	const o = new Uint8Array(maxlen);
	let offset = 0;
	for (let i = 0; i < bufs.length; ++i) {
		o.set(bufs[i], offset);
		offset += bufs[i].length;
	}
	return o;
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
	let c = 0;
	let d = 0;
	let e = 0;
	let f = 0;
	let w = 0;
	while (i < orig.length) {
		c = orig.charCodeAt(i++);
		if (c < 128) {
			out += String.fromCharCode(c);
			continue;
		}
		d = orig.charCodeAt(i++);
		if (c > 191 && c < 224) {
			f = ((c & 31) << 6) | (d & 63);
			out += String.fromCharCode(f);
			continue;
		}
		e = orig.charCodeAt(i++);
		if (c < 240) {
			out += String.fromCharCode(((c & 15) << 12) | ((d & 63) << 6) | (e & 63));
			continue;
		}
		f = orig.charCodeAt(i++);
		w = (((c & 7) << 18) | ((d & 63) << 12) | ((e & 63) << 6) | (f & 63)) - 65536;
		out += String.fromCharCode(0xd800 + ((w >>> 10) & 1023));
		out += String.fromCharCode(0xdc00 + (w & 1023));
	}
	return out;
}

/** Write a string to UTF-8 encoded binary string */
export function utf8write(orig: string): string {
	const out: string[] = [];
	let i = 0;
	let c = 0;
	let d = 0;
	while (i < orig.length) {
		c = orig.charCodeAt(i++);
		if (c < 128) {
			out.push(String.fromCharCode(c));
		} else if (c < 2048) {
			out.push(String.fromCharCode(192 + (c >> 6)));
			out.push(String.fromCharCode(128 + (c & 63)));
		} else if (c >= 55296 && c < 57344) {
			c -= 55296;
			d = orig.charCodeAt(i++) - 56320 + (c << 10);
			out.push(String.fromCharCode(240 + ((d >> 18) & 7)));
			out.push(String.fromCharCode(144 + ((d >> 12) & 63)));
			out.push(String.fromCharCode(128 + ((d >> 6) & 63)));
			out.push(String.fromCharCode(128 + (d & 63)));
		} else {
			out.push(String.fromCharCode(224 + (c >> 12)));
			out.push(String.fromCharCode(128 + ((c >> 6) & 63)));
			out.push(String.fromCharCode(128 + (c & 63)));
		}
	}
	return out.join("");
}

// eslint-disable-next-line no-control-regex
export const chr0 = /\u0000/g;
// eslint-disable-next-line no-control-regex
export const chr1 = /[\u0001-\u0006]/g;
