import { describe, it, expect } from "vitest";
import { formatNumber } from "./format.js";

describe("formatNumber", () => {
	describe("General format", () => {
		it("should format integers", () => {
			expect(formatNumber("General", 0)).toBe("0");
			expect(formatNumber("General", 1)).toBe("1");
			expect(formatNumber("General", -1)).toBe("-1");
			expect(formatNumber("General", 42)).toBe("42");
		});

		it("should format decimals", () => {
			expect(formatNumber("General", 1.5)).toBe("1.5");
			expect(formatNumber("General", 0.1)).toBe("0.1");
		});

		it("should format booleans", () => {
			expect(formatNumber("General", true)).toBe("TRUE");
			expect(formatNumber("General", false)).toBe("FALSE");
		});
	});

	describe("number format codes", () => {
		it("should format with fixed decimals (0.00)", () => {
			expect(formatNumber("0.00", 1)).toBe("1.00");
			expect(formatNumber("0.00", 1.5)).toBe("1.50");
			expect(formatNumber("0.00", 1.555)).toBe("1.56");
		});

		it("should format with thousands separator (#,##0)", () => {
			expect(formatNumber("#,##0", 1000)).toBe("1,000");
			expect(formatNumber("#,##0", 0)).toBe("0");
		});

		it("should format with thousands and decimals (#,##0.00)", () => {
			expect(formatNumber("#,##0.00", 1234.5)).toBe("1,234.50");
		});

		it("should format scientific notation (0.00E+00)", () => {
			expect(formatNumber("0.00E+00", 12345)).toBe("1.23E+04");
		});

		it("should format text (@)", () => {
			expect(formatNumber("@", "hello")).toBe("hello");
		});
	});

	describe("built-in format IDs", () => {
		it("should format with ID 0 (General)", () => {
			expect(formatNumber(0, 42)).toBe("42");
		});

		it("should format with ID 1 (0)", () => {
			expect(formatNumber(1, 1.5)).toBe("2");
		});

		it("should format with ID 2 (0.00)", () => {
			expect(formatNumber(2, 1.5)).toBe("1.50");
		});

		it("should format with ID 3 (#,##0)", () => {
			expect(formatNumber(3, 1234)).toBe("1,234");
		});

		it("should format with ID 4 (#,##0.00)", () => {
			expect(formatNumber(4, 1234.5)).toBe("1,234.50");
		});

		it("should format with ID 49 (@)", () => {
			expect(formatNumber(49, "text")).toBe("text");
		});
	});

	describe("date formats", () => {
		it("should format m/d/yy (ID 14)", () => {
			const result = formatNumber(14, 44928);
			expect(result).toContain("1");
		});

		it("should format h:mm (24h)", () => {
			expect(formatNumber("h:mm", 0.75)).toBe("18:00");
			expect(formatNumber("h:mm", 0.5)).toBe("12:00");
		});

		it("should format h:mm:ss (24h)", () => {
			expect(formatNumber("h:mm:ss", 0.5)).toBe("12:00:00");
			expect(formatNumber("h:mm:ss", 0.75)).toBe("18:00:00");
		});

		it("should format mm:ss (ID 45)", () => {
			// 0.5 = 12h = 720 minutes
			const result = formatNumber("mm:ss", 0.5);
			expect(result).toBe("00:00");
		});
	});

	describe("negative numbers", () => {
		it("should handle negative with positive;negative format", () => {
			expect(formatNumber("#,##0;(#,##0)", -1234)).toBe("(1,234)");
			expect(formatNumber("#,##0;(#,##0)", 1234)).toBe("1,234");
		});
	});

	describe("edge cases", () => {
		it("should format zero", () => {
			expect(formatNumber("0.00", 0)).toBe("0.00");
			expect(formatNumber("#,##0", 0)).toBe("0");
		});

		it("should format strings as text", () => {
			expect(formatNumber("General", "hello")).toBe("hello");
		});
	});
});
