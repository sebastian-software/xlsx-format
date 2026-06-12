import { describe, it, expect } from "vitest";
import { parseCalcChainXml } from "./calc-chain.js";

describe("xlsx/calc-chain", () => {
	it("parseCalcChainXml should parse chain entries", () => {
		const xml = `<?xml version="1.0"?><calcChain><c r="A1" i="1"/><c r="B2"/><c r="C3" i="2"/></calcChain>`;
		const chain = parseCalcChainXml(xml);
		expect(chain).toHaveLength(3);
		expect(chain[0].r).toBe("A1");
		expect(chain[0].i).toBe("1");
		expect(chain[1].r).toBe("B2");
		expect(chain[1].i).toBe("1"); // sticky
		expect(chain[2].i).toBe("2");
	});

	it("parseCalcChainXml should handle empty input", () => {
		expect(parseCalcChainXml("")).toEqual([]);
	});
});
