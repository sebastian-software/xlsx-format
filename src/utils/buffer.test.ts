import { describe, it, expect } from "vitest";
import { utf8encode, utf8decode, NULL_CHAR_REGEX, CONTROL_CHAR_REGEX } from "./buffer.js";

describe("utils/buffer", () => {
	it("utf8encode and utf8decode should roundtrip", () => {
		const s = "Hello, Wörld! 日本語";
		expect(utf8decode(utf8encode(s))).toBe(s);
	});

	it("should export regex patterns", () => {
		expect("abc\u0000def".replace(NULL_CHAR_REGEX, "")).toBe("abcdef");
		expect("a\u0001b\u0003c".replace(CONTROL_CHAR_REGEX, "")).toBe("abc");
	});
});
