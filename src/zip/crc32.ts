/**
 * Pre-computed CRC32 lookup table using the standard polynomial 0xEDB88320
 * (ISO 3309 / ITU-T V.42, bit-reversed form of 0x04C11DB7).
 *
 * Each entry TABLE[i] is the CRC32 of the single byte i, computed by
 * shifting through all 8 bits and XOR-ing with the polynomial when the LSB is set.
 */
const TABLE = new Uint32Array(256);
for (let i = 0; i < 256; i++) {
	let c = i;
	for (let j = 0; j < 8; j++) {
		// If LSB is set, shift right and XOR with polynomial; otherwise just shift
		c = c & 1 ? (c >>> 1) ^ 0xedb88320 : c >>> 1;
	}
	TABLE[i] = c;
}

/**
 * Compute the CRC32 checksum of a byte buffer.
 *
 * Uses the standard table-driven algorithm: start with all bits set (0xFFFFFFFF),
 * fold each byte through the lookup table, and invert at the end.
 *
 * @param buf - Input bytes
 * @returns Unsigned 32-bit CRC32 value
 */
export function crc32(buf: Uint8Array): number {
	let crc = 0xffffffff;
	for (let i = 0; i < buf.length; i++) {
		// XOR bottom byte with input byte to get table index, then shift and XOR
		crc = (crc >>> 8) ^ TABLE[(crc ^ buf[i]) & 0xff];
	}
	// Final inversion; >>> 0 ensures unsigned 32-bit result
	return (crc ^ 0xffffffff) >>> 0;
}
