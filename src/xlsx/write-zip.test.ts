import { describe, it, expect } from "vitest";
import { read, write, createWorkbook, appendSheet, arrayToSheet } from "../index.js";

describe("write-zip.ts: comments in roundtrip", () => {
	it("legacy comments survive write → read", async () => {
		const ws = arrayToSheet([["Value"]]);
		// Prepare comments in the format expected by the writer
		(ws as any)["!comments"] = [["A1", [{ a: "Alice", t: "A note", T: false }]]];
		(ws as any)["!legacy"] = true;

		const wb = createWorkbook(ws, "Sheet1");
		const bytes = await write(wb);
		const wb2 = await read(bytes);
		const ws2 = wb2.Sheets[wb2.SheetNames[0]];
		// Check that the comment was read back
		const a1 = (ws2 as any)["A1"];
		expect(a1).toBeDefined();
		if (a1.c) {
			expect(a1.c.length).toBeGreaterThanOrEqual(1);
			expect(a1.c[0].t).toContain("A note");
		}
	});

	it("threaded comments survive write → read", async () => {
		const ws = arrayToSheet([["Value"]]);
		(ws as any)["!comments"] = [
			[
				"A1",
				[
					{ a: "Alice", t: "First comment", T: true, ID: undefined },
					{ a: "Bob", t: "Reply", T: true, ID: undefined },
				],
			],
		];
		(ws as any)["!legacy"] = true;

		const wb = createWorkbook(ws, "Sheet1");
		const bytes = await write(wb);
		const wb2 = await read(bytes);
		const ws2 = wb2.Sheets[wb2.SheetNames[0]];
		const a1 = (ws2 as any)["A1"];
		expect(a1).toBeDefined();
		if (a1.c) {
			expect(a1.c.length).toBeGreaterThanOrEqual(1);
		}
	});

	it("workbook with veryHidden sheets filters app props", async () => {
		const ws1 = arrayToSheet([["A"]]);
		const ws2 = arrayToSheet([["B"]]);
		const wb = createWorkbook(ws1, "Visible");
		appendSheet(wb, ws2, "VeryHidden");
		wb.Workbook = {
			Sheets: [{ Hidden: 0 }, { Hidden: 2 }],
		} as any;

		const bytes = await write(wb);
		const wb2 = await read(bytes);
		expect(wb2.SheetNames).toContain("Visible");
	});
});
