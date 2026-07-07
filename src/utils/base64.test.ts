import { describe, it, expect } from "vitest";
import { base64decode, base64encode } from "./base64.js";

describe("utils/base64", () => {
	it("should roundtrip encode/decode", () => {
		const data = new Uint8Array([72, 101, 108, 108, 111]);
		const encoded = base64encode(data);
		const decoded = base64decode(encoded);
		expect(decoded).toStrictEqual(data);
	});

	it("should strip data URI prefix", () => {
		const data = new Uint8Array([1, 2, 3]);
		const b64 = base64encode(data);
		const withPrefix = "data:application/octet-stream;base64," + b64;
		expect(base64decode(withPrefix)).toStrictEqual(data);
	});
});
