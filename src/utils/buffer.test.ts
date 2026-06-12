import { describe, it, expect } from "vitest";
import { utf8read, utf8encode, utf8decode, NULL_CHAR_REGEX, CONTROL_CHAR_REGEX } from "./buffer.js";

describe("utils/buffer", () => {
	it("utf8encode and utf8decode should roundtrip", () => {
		const s = "Hello, Wörld! 日本語";
		expect(utf8decode(utf8encode(s))).toBe(s);
	});

	it("utf8read should decode ASCII", () => {
		expect(utf8read("Hello")).toBe("Hello");
	});

	it("utf8read should decode 2-byte UTF-8", () => {
		// ö = U+00F6 = 0xC3 0xB6 in UTF-8
		const binary = String.fromCharCode(0xc3, 0xb6);
		expect(utf8read(binary)).toBe("ö");
	});

	it("utf8read should decode 3-byte UTF-8", () => {
		// 日 = U+65E5 = 0xE6 0x97 0xA5
		const binary = String.fromCharCode(0xe6, 0x97, 0xa5);
		expect(utf8read(binary)).toBe("日");
	});

	it("utf8read should decode 4-byte UTF-8 (surrogate pair)", () => {
		// 𝄞 (musical symbol G clef) = U+1D11E = 0xF0 0x9D 0x84 0x9E
		const binary = String.fromCharCode(0xf0, 0x9d, 0x84, 0x9e);
		expect(utf8read(binary)).toBe("𝄞");
	});

	it("should export regex patterns", () => {
		expect("abc\u0000def".replace(NULL_CHAR_REGEX, "")).toBe("abcdef");
		expect("a\u0001b\u0003c".replace(CONTROL_CHAR_REGEX, "")).toBe("abc");
	});
});
