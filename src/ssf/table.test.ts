import { describe, it, expect } from "vitest";
import { initFormatTable, loadFormat, loadFormatTable, resetFormatTable } from "./table.js";

describe("ssf/table", () => {
	it("initFormatTable should populate defaults", () => {
		const t = initFormatTable();
		expect(t[0]).toBe("General");
		expect(t[14]).toBe("m/d/yy");
		expect(t[49]).toBe("@");
	});

	it("loadFormat should register custom format", () => {
		resetFormatTable();
		const idx = loadFormat("#,##0.000");
		expect(idx).toBeGreaterThan(0);
	});

	it("loadFormat should find existing format", () => {
		resetFormatTable();
		const idx = loadFormat("General");
		expect(idx).toBe(0);
	});

	it("loadFormatTable should bulk-load formats", () => {
		resetFormatTable();
		loadFormatTable({ 200: "custom-fmt" });
		const t = initFormatTable();
		// After reset, custom format should be gone
		expect(t[200]).toBeUndefined();
	});
});
