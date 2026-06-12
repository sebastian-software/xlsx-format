import { describe, expect, it } from "vitest";
import { addArrayToSheet, arrayToSheet, sheetToArray } from "./aoa.js";

describe("sheetToArray", () => {
	it("returns worksheet values as an array of arrays", () => {
		const ws = arrayToSheet([
			["Name", "Age"],
			["Alice", 30],
			["Bob", 25],
		]);

		expect(sheetToArray(ws)).toEqual([
			["Name", "Age"],
			["Alice", 30],
			["Bob", 25],
		]);
	});
});

describe("aoa.ts — addArrayToSheet edge cases", () => {
	it("should create dense worksheet", () => {
		const ws = addArrayToSheet(
			null,
			[
				["a", 1],
				["b", 2],
			],
			{ dense: true },
		);
		expect((ws as any)["!data"]).toBeDefined();
		expect((ws as any)["!data"][0][0].v).toBe("a");
		expect((ws as any)["!data"][1][1].v).toBe(2);
	});

	it("should handle numeric origin", () => {
		const ws = addArrayToSheet(null, [["A"]], { origin: 3 } as any);
		// When no prior ref, range starts at 0,0 and data at row 3 expands e.r to 3
		expect(ws["!ref"]).toBe("A1:A4");
		expect((ws as any)["A4"]).toBeDefined();
	});

	it("should handle origin as cell ref string", () => {
		const ws = addArrayToSheet(null, [["X"]], { origin: "C5" } as any);
		expect((ws as any)["C5"].v).toBe("X");
	});

	it("should handle origin -1 (append)", () => {
		let ws = arrayToSheet([["Row1"]]);
		ws = addArrayToSheet(ws, [["Row2"]], { origin: -1 } as any);
		expect((ws as any)["A2"].v).toBe("Row2");
	});

	it("should handle null values with nullError", () => {
		const ws = arrayToSheet([[null, "ok"]], { nullError: true } as any);
		expect((ws as any)["A1"].t).toBe("e");
		expect((ws as any)["A1"].v).toBe(0);
	});

	it("should handle null values with sheetStubs", () => {
		const ws = arrayToSheet([[null, "ok"]], { sheetStubs: true } as any);
		expect((ws as any)["A1"].t).toBe("z");
	});

	it("should handle NaN and Infinity", () => {
		const ws = arrayToSheet([[NaN, Infinity]]);
		expect((ws as any)["A1"].t).toBe("e");
		expect((ws as any)["A1"].v).toBe(0x0f); // #VALUE!
		expect((ws as any)["B1"].t).toBe("e");
		expect((ws as any)["B1"].v).toBe(0x07); // #DIV/0!
	});

	it("should handle array values [value, formula]", () => {
		const ws = arrayToSheet([[["result", "=SUM(A2:A10)"]]]);
		expect((ws as any)["A1"].v).toBe("result");
		expect((ws as any)["A1"].f).toBe("=SUM(A2:A10)");
	});

	it("should handle pre-built cell objects", () => {
		const ws = arrayToSheet([[{ t: "n", v: 42, z: "#,##0" }]]);
		expect((ws as any)["A1"].v).toBe(42);
		expect((ws as any)["A1"].z).toBe("#,##0");
	});

	it("should handle Date values", () => {
		const date = new Date("2024-06-15T00:00:00Z");
		const ws = arrayToSheet([[date]], { UTC: true });
		expect((ws as any)["A1"].t).toBe("n");
		expect((ws as any)["A1"].v).toBeGreaterThan(40000);
	});

	it("should handle Date with cellDates", () => {
		const date = new Date("2024-06-15T00:00:00Z");
		const ws = arrayToSheet([[date]], { cellDates: true, UTC: true });
		expect((ws as any)["A1"].t).toBe("d");
	});

	it("should skip undefined values", () => {
		const ws = arrayToSheet([[undefined, "b"]]);
		expect((ws as any)["A1"]).toBeUndefined();
		expect((ws as any)["B1"].v).toBe("b");
	});

	it("should skip null rows", () => {
		const data: any[][] = [["a"], null as any, ["c"]];
		const ws = arrayToSheet(data);
		expect((ws as any)["A1"].v).toBe("a");
		expect((ws as any)["A3"].v).toBe("c");
	});

	it("should throw for non-array rows", () => {
		expect(() => arrayToSheet(["not an array" as any])).toThrow("array of arrays");
	});

	it("should handle null with formula", () => {
		const ws = arrayToSheet([[[null, "=NOW()"]]]);
		expect((ws as any)["A1"].t).toBe("n");
		expect((ws as any)["A1"].f).toBe("=NOW()");
	});

	it("preserves existing number formats in sparse worksheets", () => {
		const ws = arrayToSheet([[1]]);
		ws["A1"].z = "0.00";

		addArrayToSheet(ws, [[2]]);

		expect(ws["A1"].v).toBe(2);
		expect(ws["A1"].z).toBe("0.00");
	});

	it("preserves existing number formats in dense worksheets", () => {
		const ws = arrayToSheet([[1]], { dense: true });
		ws["!data"]![0]![0]!.z = "0.00";

		addArrayToSheet(ws, [[2]]);

		expect(ws["!data"]![0]![0]!.v).toBe(2);
		expect(ws["!data"]![0]![0]!.z).toBe("0.00");
	});
});
