import { describe, it, expect } from "vitest";
import { rcToA1, a1ToRc, shiftFormulaStr, shiftFormulaXlsx, isFuzzyFormula, stripXlFunctionPrefix } from "./formula.js";

describe("xlsx/formula", () => {
	it("rcToA1 should convert absolute references", () => {
		expect(rcToA1("R1C1", { r: 0, c: 0 })).toBe("$A$1");
		expect(rcToA1("R2C3", { r: 0, c: 0 })).toBe("$C$2");
	});

	it("rcToA1 should convert relative references", () => {
		expect(rcToA1("R[1]C[1]", { r: 0, c: 0 })).toBe("B2");
		expect(rcToA1("R[0]C[0]", { r: 2, c: 3 })).toBe("D3");
		expect(rcToA1("RC", { r: 5, c: 2 })).toBe("C6");
	});

	it("a1ToRc should convert A1 to R1C1", () => {
		expect(a1ToRc("$A$1", { r: 0, c: 0 })).toBe("R1C1");
		expect(a1ToRc("B2", { r: 1, c: 1 })).toBe("RC");
		expect(a1ToRc("C3", { r: 0, c: 0 })).toBe("R[2]C[2]");
	});

	it("shiftFormulaStr should shift relative references", () => {
		const result = shiftFormulaStr("A1+B2", { r: 1, c: 1 });
		expect(result).toBe("B2+C3");
	});

	it("shiftFormulaStr should not shift absolute references", () => {
		const result = shiftFormulaStr("$A$1+B2", { r: 1, c: 1 });
		expect(result).toBe("$A$1+C3");
	});

	it("shiftFormulaXlsx should shift based on range and cell", () => {
		const result = shiftFormulaXlsx("A1*2", "A1:A10", "A3");
		expect(result).toBe("A3*2");
	});

	it("isFuzzyFormula should reject single chars", () => {
		expect(isFuzzyFormula("=")).toBe(false);
		expect(isFuzzyFormula("=SUM(A1)")).toBe(true);
	});

	it("stripXlFunctionPrefix should remove _xlfn.", () => {
		expect(stripXlFunctionPrefix("_xlfn.CONCAT(A1,B1)")).toBe("CONCAT(A1,B1)");
		expect(stripXlFunctionPrefix("SUM(A1)")).toBe("SUM(A1)");
	});
});
