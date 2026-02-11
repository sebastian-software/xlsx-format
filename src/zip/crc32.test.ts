import { describe, it, expect } from "vitest";
import { crc32 } from "./crc32.js";

describe("crc32", () => {
	it("should return 0 for empty input", () => {
		expect(crc32(new Uint8Array(0))).toBe(0);
	});

	it("should compute correct CRC for known values", () => {
		// CRC32 of "123456789" is 0xCBF43926
		const input = new TextEncoder().encode("123456789");
		expect(crc32(input)).toBe(0xcbf43926);
	});

	it("should compute correct CRC for single byte", () => {
		// CRC32 of a single null byte
		const input = new Uint8Array([0]);
		expect(crc32(input)).toBe(0xd202ef8d);
	});

	it("should produce different results for different inputs", () => {
		const a = new TextEncoder().encode("hello");
		const b = new TextEncoder().encode("world");
		expect(crc32(a)).not.toBe(crc32(b));
	});

	it("should be deterministic", () => {
		const input = new TextEncoder().encode("test data");
		expect(crc32(input)).toBe(crc32(input));
	});
});
