import { describe, it, expect } from "vitest";
import { formatCell } from "../index.js";

describe("api/format", () => {
	it("formatCell should return empty for null/stub cells", () => {
		expect(formatCell(null as any)).toBe("");
		expect(formatCell({ t: "z" } as any)).toBe("");
	});

	it("formatCell should return cached w value", () => {
		expect(formatCell({ t: "s", v: "x", w: "cached" } as any)).toBe("cached");
	});

	it("formatCell should format error cells", () => {
		const result = formatCell({ t: "e", v: 0x07 } as any);
		expect(result).toBe("#DIV/0!");
	});

	it("formatCell should format with explicit value", () => {
		const cell: any = { t: "n", v: 0 };
		const result = formatCell(cell, 42);
		expect(result).toBeDefined();
	});

	it("formatCell should use dateNF option", () => {
		const cell: any = { t: "d", v: new Date("2024-01-15") };
		formatCell(cell, null, { dateNF: "yyyy-mm-dd" });
		expect(cell.z).toBe("yyyy-mm-dd");
	});
});
