import { describe, it, expect } from "vitest";
import { dateToSerialNumber, serialNumberToDate, localToUtc, utcToLocal } from "./date.js";

describe("dateToSerialNumber", () => {
	it("should convert 1900-01-01 to serial 2", () => {
		// Epoch is 1899-12-30, so 1900-01-01 is 2 days later
		const d = new Date(Date.UTC(1900, 0, 1));
		expect(dateToSerialNumber(d)).toBe(2);
	});

	it("should add 1 for dates at or after serial 60 (Lotus bug)", () => {
		// 1900-02-28 is 60 days from epoch; >= 60 triggers the +1 adjustment
		const d = new Date(Date.UTC(1900, 1, 28));
		expect(dateToSerialNumber(d)).toBe(61);
	});

	it("should convert 1900-03-01 to serial 62", () => {
		const d = new Date(Date.UTC(1900, 2, 1));
		expect(dateToSerialNumber(d)).toBe(62);
	});

	it("should convert 2023-01-01 correctly", () => {
		const d = new Date(Date.UTC(2023, 0, 1));
		expect(dateToSerialNumber(d)).toBe(44928);
	});

	it("should handle 1904 date system", () => {
		const d = new Date(Date.UTC(2023, 0, 1));
		const serial1900 = dateToSerialNumber(d);
		const serial1904 = dateToSerialNumber(d, true);
		// 1904 system is 1462 days less
		expect(serial1900 - serial1904).toBe(1462);
	});
});

describe("serialNumberToDate", () => {
	it("should convert serial 2 to 1900-01-01", () => {
		// Serial 1 maps to 1899-12-31 due to epoch at 1899-12-30
		const d = serialNumberToDate(2);
		expect(d.getUTCFullYear()).toBe(1900);
		expect(d.getUTCMonth()).toBe(0);
		expect(d.getUTCDate()).toBe(1);
	});

	it("should convert serial 44928 to 2023-01-01", () => {
		const d = serialNumberToDate(44928);
		expect(d.getUTCFullYear()).toBe(2023);
		expect(d.getUTCMonth()).toBe(0);
		expect(d.getUTCDate()).toBe(1);
	});

	it("should roundtrip with dateToSerialNumber", () => {
		const original = new Date(Date.UTC(2024, 5, 15));
		const serial = dateToSerialNumber(original);
		const result = serialNumberToDate(serial);
		expect(result.getUTCFullYear()).toBe(2024);
		expect(result.getUTCMonth()).toBe(5);
		expect(result.getUTCDate()).toBe(15);
	});

	it("should handle 1904 date system", () => {
		const d = serialNumberToDate(44928, true);
		// 1904 system adds 1462 days
		const d1900 = serialNumberToDate(44928);
		const diff = d.getTime() - d1900.getTime();
		expect(diff).toBe(1462 * 24 * 60 * 60 * 1000);
	});
});

describe("localToUtc / utcToLocal", () => {
	it("should be inverse operations", () => {
		const d = new Date(2024, 0, 15, 12, 30, 0);
		const utc = localToUtc(d);
		const back = utcToLocal(utc);
		expect(back.getTime()).toBe(d.getTime());
	});
});
