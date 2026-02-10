/** Pre-computed CRC32 lookup table (ISO 3309 / ITU-T V.42 polynomial). */
const TABLE = new Uint32Array(256);
for (let i = 0; i < 256; i++) {
	let c = i;
	for (let j = 0; j < 8; j++) {
		c = c & 1 ? (c >>> 1) ^ 0xedb88320 : c >>> 1;
	}
	TABLE[i] = c;
}

/**
 * Compute the CRC32 checksum of a byte buffer.
 *
 * @param buf - Input bytes
 * @returns Unsigned 32-bit CRC32 value
 */
export function crc32(buf: Uint8Array): number {
	let crc = 0xffffffff;
	for (let i = 0; i < buf.length; i++) {
		crc = (crc >>> 8) ^ TABLE[(crc ^ buf[i]) & 0xff];
	}
	return (crc ^ 0xffffffff) >>> 0;
}
