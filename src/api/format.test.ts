import { describe, it, expect } from "vitest";
import { formatCell } from "../index.js";
import { formatCellForOutput, getCellDateTimeFormatKind } from "./format.js";

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

	it("classifies cell formats from XF records and custom format tables", () => {
		expect(getCellDateTimeFormatKind({ t: "n", v: 45292, XF: { numFmtId: 14 } })).toBe("date");
		expect(getCellDateTimeFormatKind({ t: "n", v: 45292, XF: { numFmtId: 27 } })).toBe("date");
		expect(
			getCellDateTimeFormatKind({ t: "n", v: 45292.5, XF: { numFmtId: 200 } }, { table: { 200: "m/d/yy h:mm" } }),
		).toBe("datetime");
		expect(getCellDateTimeFormatKind({ t: "n", v: 42, XF: { numFmtId: 63 } })).toBe("none");
		expect(getCellDateTimeFormatKind({ t: "n", v: 42 })).toBe("none");
	});

	it("formats numeric date cells as ISO from XF records", () => {
		expect(
			formatCellForOutput({ t: "n", v: 45292, XF: { numFmtId: 14 } }, null, {
				dateOutput: "iso",
				dateNF: "yyyy-mm-dd",
				UTC: true,
			}),
		).toBe("2024-01-01");
		expect(
			formatCellForOutput({ t: "n", v: 45292.5, XF: { numFmtId: 200 } }, null, {
				dateOutput: "iso",
				table: { 200: "m/d/yy h:mm" },
				UTC: true,
			}),
		).toBe("2024-01-01T12:00:00");
	});

	it("formats local ISO strings once when UTC is false", () => {
		const originalTimezone = process.env.TZ;
		process.env.TZ = "Etc/GMT-5";
		try {
			expect(
				formatCellForOutput({ t: "n", v: 45292.5, z: "m/d/yy h:mm" }, null, {
					dateOutput: "iso",
					UTC: false,
				}),
			).toBe("2024-01-01T17:00:00");
		} finally {
			process.env.TZ = originalTimezone;
		}
	});

	it("infers ISO output shape for Date objects without explicit formats", () => {
		expect(
			formatCellForOutput({ t: "d", v: new Date("2024-01-01T00:00:00Z") }, null, {
				dateOutput: "iso",
				UTC: true,
			}),
		).toBe("2024-01-01");
		expect(
			formatCellForOutput({ t: "d", v: new Date("2024-01-01T12:34:56Z") }, null, {
				dateOutput: "iso",
				UTC: true,
			}),
		).toBe("2024-01-01T12:34:56");
	});

	it("pads ISO years before 1000 to four digits", () => {
		expect(
			formatCellForOutput({ t: "d", v: new Date("0100-03-25T00:00:00.000Z"), z: "yyyy-mm-dd" }, null, {
				dateOutput: "iso",
				UTC: true,
			}),
		).toBe("0100-03-25");
	});
});
