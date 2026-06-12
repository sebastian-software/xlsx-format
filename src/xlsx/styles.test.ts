import { describe, expect, it } from "vitest";
import { createWorkbook } from "../api/book.js";
import type { CellStyle, WorkSheet } from "../types.js";
import { buildStyleRegistry, parseStylesXml, writeStylesXml } from "./styles.js";

describe("xlsx styles", () => {
	it("deduplicates equivalent style parts and cell formats", () => {
		const headerStyle: CellStyle = {
			font: { name: "Calibri", size: 10, bold: true, color: { argb: "FFFFFFFF" } },
			fill: { patternType: "solid", fgColor: { rgb: "2E75B6" } },
			border: {
				top: { style: "thin", color: { argb: "FFB4B4B4" } },
				right: { style: "thin", color: { argb: "FFB4B4B4" } },
				bottom: { style: "thin", color: { argb: "FFB4B4B4" } },
				left: { style: "thin", color: { argb: "FFB4B4B4" } },
			},
			alignment: { horizontal: "center", vertical: "middle", wrapText: true },
		};
		const ws: WorkSheet = {
			"!ref": "A1:B1",
			A1: { t: "s", v: "A", s: headerStyle },
			B1: { t: "s", v: "B", s: { ...headerStyle } },
		};

		const registry = buildStyleRegistry(createWorkbook(ws, "S"), {});

		expect(registry.fonts).toHaveLength(2);
		expect(registry.fills).toHaveLength(3);
		expect(registry.borders).toHaveLength(2);
		expect(registry.cellXfs).toHaveLength(2);
		expect(registry.cellStyleIds.get(ws.A1)).toBe(registry.cellStyleIds.get(ws.B1));
	});

	it("writes and parses supported style metadata", () => {
		const ws: WorkSheet = {
			"!ref": "A1",
			A1: {
				t: "n",
				v: 42,
				s: {
					font: { bold: true, color: { argb: "FFFF0000" } },
					fill: { patternType: "solid", fgColor: { argb: "FFEDF2F7" } },
					alignment: { horizontal: "right", vertical: "bottom" },
					numFmt: "0.0000%",
				},
			},
		};
		const registry = buildStyleRegistry(createWorkbook(ws, "S"), {});
		const xml = writeStylesXml(null, { styleRegistry: registry });
		const parsed = parseStylesXml(xml);

		expect(xml).toContain('formatCode="0.0000%"');
		expect(parsed.Fonts[1].bold).toBe(true);
		expect(parsed.Fills[2].fgColor?.rgb).toBe("FFEDF2F7");
		expect(parsed.CellXf[1].alignment?.horizontal).toBe("right");
		expect(parsed.CellXf[1].alignment?.vertical).toBe("bottom");
	});
});
