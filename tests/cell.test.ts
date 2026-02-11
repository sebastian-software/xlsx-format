import { describe, it, expect } from "vitest";
import {
	decodeCell,
	encodeCell,
	decodeRange,
	encodeRange,
	decodeCol,
	encodeCol,
	decodeRow,
	encodeRow,
} from "../src/utils/cell.js";

describe("decodeCol / encodeCol", () => {
	it("should decode single-letter columns", () => {
		expect(decodeCol("A")).toBe(0);
		expect(decodeCol("B")).toBe(1);
		expect(decodeCol("Z")).toBe(25);
	});

	it("should decode multi-letter columns", () => {
		expect(decodeCol("AA")).toBe(26);
		expect(decodeCol("AB")).toBe(27);
		expect(decodeCol("AZ")).toBe(51);
		expect(decodeCol("BA")).toBe(52);
	});

	it("should strip $ absolute marker", () => {
		expect(decodeCol("$A")).toBe(0);
		expect(decodeCol("$Z")).toBe(25);
	});

	it("should encode single-letter columns", () => {
		expect(encodeCol(0)).toBe("A");
		expect(encodeCol(1)).toBe("B");
		expect(encodeCol(25)).toBe("Z");
	});

	it("should encode multi-letter columns", () => {
		expect(encodeCol(26)).toBe("AA");
		expect(encodeCol(27)).toBe("AB");
		expect(encodeCol(51)).toBe("AZ");
		expect(encodeCol(52)).toBe("BA");
		expect(encodeCol(701)).toBe("ZZ");
		expect(encodeCol(702)).toBe("AAA");
	});

	it("should throw for negative column index", () => {
		expect(() => encodeCol(-1)).toThrow("invalid column");
	});

	it("should roundtrip encode/decode", () => {
		for (let i = 0; i < 1000; i++) {
			expect(decodeCol(encodeCol(i))).toBe(i);
		}
	});
});

describe("decodeRow / encodeRow", () => {
	it("should decode row strings to zero-based index", () => {
		expect(decodeRow("1")).toBe(0);
		expect(decodeRow("2")).toBe(1);
		expect(decodeRow("100")).toBe(99);
	});

	it("should strip $ absolute marker", () => {
		expect(decodeRow("$5")).toBe(4);
	});

	it("should encode zero-based index to row string", () => {
		expect(encodeRow(0)).toBe("1");
		expect(encodeRow(1)).toBe("2");
		expect(encodeRow(99)).toBe("100");
	});
});

describe("decodeCell / encodeCell", () => {
	it("should decode A1 to {c:0, r:0}", () => {
		expect(decodeCell("A1")).toEqual({ c: 0, r: 0 });
	});

	it("should decode B3 to {c:1, r:2}", () => {
		expect(decodeCell("B3")).toEqual({ c: 1, r: 2 });
	});

	it("should decode multi-letter cell references", () => {
		expect(decodeCell("AA1")).toEqual({ c: 26, r: 0 });
		expect(decodeCell("AB12")).toEqual({ c: 27, r: 11 });
	});

	it("should encode {c:0, r:0} to A1", () => {
		expect(encodeCell({ c: 0, r: 0 })).toBe("A1");
	});

	it("should encode {c:1, r:2} to B3", () => {
		expect(encodeCell({ c: 1, r: 2 })).toBe("B3");
	});

	it("should encode multi-letter columns", () => {
		expect(encodeCell({ c: 26, r: 0 })).toBe("AA1");
	});

	it("should roundtrip cell addresses", () => {
		const refs = ["A1", "Z1", "AA1", "AZ100", "BA50"];
		for (const ref of refs) {
			expect(encodeCell(decodeCell(ref))).toBe(ref);
		}
	});
});

describe("decodeRange / encodeRange", () => {
	it("should decode a normal range", () => {
		const r = decodeRange("A1:C5");
		expect(r.s).toEqual({ c: 0, r: 0 });
		expect(r.e).toEqual({ c: 2, r: 4 });
	});

	it("should decode a single-cell range", () => {
		const r = decodeRange("B2");
		expect(r.s).toEqual({ c: 1, r: 1 });
		expect(r.e).toEqual({ c: 1, r: 1 });
	});

	it("should encode a range object", () => {
		expect(encodeRange({ s: { c: 0, r: 0 }, e: { c: 2, r: 4 } })).toBe("A1:C5");
	});

	it("should encode a single-cell range without colon", () => {
		expect(encodeRange({ s: { c: 0, r: 0 }, e: { c: 0, r: 0 } })).toBe("A1");
	});

	it("should encode from two cell addresses", () => {
		expect(encodeRange({ c: 0, r: 0 }, { c: 2, r: 4 })).toBe("A1:C5");
	});

	it("should roundtrip ranges", () => {
		const ranges = ["A1:C5", "A1:Z100", "AA1:AZ50"];
		for (const range of ranges) {
			expect(encodeRange(decodeRange(range))).toBe(range);
		}
	});
});
