const T0 = new Int32Array(256);
for (let i = 0; i < 256; ++i) {
	let c = i;
	for (let j = 0; j < 8; ++j) {
		c = c & 1 ? -306674912 ^ ((c >>> 1) & 0x7fffffff) : (c >>> 1) & 0x7fffffff;
	}
	T0[i] = c;
}

/** Compute CRC32 of a Uint8Array */
export function crc32(buf: Uint8Array, seed?: number): number {
	let C = seed !== undefined ? ~seed : -1;
	for (let i = 0; i < buf.length; ++i) {
		C = T0[(C ^ buf[i]) & 0xff] ^ ((C >>> 8) & 0x00ffffff);
	}
	return ~C;
}
