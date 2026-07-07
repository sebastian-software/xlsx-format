import { describe, it, expect } from "vitest";
import { arrayToSheet, sheetToJson, sheetToCsv, csvToSheet } from "../index.js";
import { sheetToTxt } from "./csv.js";

describe("csv.ts: advanced CSV features", () => {
	it("sheetToCsv with skipHidden rows", () => {
		const ws = arrayToSheet([["A"], ["B"], ["C"]]);
		ws["!rows"] = [undefined as any, { hidden: true }, undefined as any];
		const csv = sheetToCsv(ws, { skipHidden: true });
		expect(csv).toContain("A");
		expect(csv).not.toContain("B");
		expect(csv).toContain("C");
	});

	it("sheetToCsv with skipHidden cols", () => {
		const ws = arrayToSheet([["A", "B", "C"]]);
		ws["!cols"] = [undefined as any, { hidden: true }, undefined as any];
		const csv = sheetToCsv(ws, { skipHidden: true });
		expect(csv).toContain("A");
		expect(csv).not.toContain("B");
		expect(csv).toContain("C");
	});

	it("sheetToCsv with strip option", () => {
		const ws = arrayToSheet([["A", "", ""]]);
		const csv = sheetToCsv(ws, { strip: true });
		expect(csv).toBe("A");
	});

	it("sheetToCsv with blankrows=false", () => {
		const ws = arrayToSheet([["A"], [], ["C"]]);
		const csv = sheetToCsv(ws, { blankrows: false });
		expect(csv).toBe("A\nC");
	});

	it("sheetToCsv with forceQuotes", () => {
		const ws = arrayToSheet([["Simple"]]);
		const csv = sheetToCsv(ws, { forceQuotes: true });
		expect(csv).toBe('"Simple"');
	});

	it("sheetToCsv with custom RS", () => {
		const ws = arrayToSheet([["A"], ["B"]]);
		const csv = sheetToCsv(ws, { RS: "\r\n" });
		expect(csv).toBe("A\r\nB");
	});

	it("sheetToCsv with rawNumbers", () => {
		const ws = arrayToSheet([[1.23456789]]);
		const csv = sheetToCsv(ws, { rawNumbers: true });
		expect(csv).toBe("1.23456789");
	});

	it("sheetToCsv with ISO date output", () => {
		const ws = arrayToSheet([
			["Date", "DateTime", "Time"],
			[
				{ t: "n", v: 45292, z: "m/d/yy" },
				{ t: "n", v: 45292.5, z: "m/d/yy h:mm" },
				{ t: "n", v: 0.5, z: "h:mm" },
			],
		]);

		expect(sheetToCsv(ws, { dateOutput: "iso", UTC: true })).toBe(
			"Date,DateTime,Time\n2024-01-01,2024-01-01T12:00:00,12:00",
		);
		expect(sheetToCsv(ws, { dateOutput: "iso", rawNumbers: true, UTC: true })).toBe(
			"Date,DateTime,Time\n2024-01-01,2024-01-01T12:00:00,0.5",
		);
	});

	it("sheetToCsv keeps default ISO datetime output timezone-stable", () => {
		const originalTimezone = process.env.TZ;
		process.env.TZ = "Europe/Berlin";
		try {
			const ws = arrayToSheet([
				["Date", "Text"],
				[{ t: "n", v: 45292.5, z: "m/d/yy h:mm" }, "hello"],
			]);

			expect(sheetToCsv(ws, { dateOutput: "iso", strip: true, blankrows: false })).toBe(
				"Date,Text\n2024-01-01T12:00:00,hello",
			);
			expect(sheetToCsv(ws, { dateOutput: "iso", UTC: false, strip: true, blankrows: false })).toBe(
				"Date,Text\n2024-01-01T13:00:00,hello",
			);
		} finally {
			if (originalTimezone === undefined) {
				delete process.env.TZ;
			} else {
				process.env.TZ = originalTimezone;
			}
		}
	});

	it("sheetToCsv quotes commas", () => {
		const ws = arrayToSheet([["A,B"]]);
		const csv = sheetToCsv(ws);
		expect(csv).toBe('"A,B"');
	});

	it("sheetToCsv quotes newlines", () => {
		const ws = arrayToSheet([["A\nB"]]);
		const csv = sheetToCsv(ws);
		expect(csv).toBe('"A\nB"');
	});

	it("sheetToCsv quotes double-quotes", () => {
		const ws = arrayToSheet([['A"B']]);
		const csv = sheetToCsv(ws);
		expect(csv).toBe('"A""B"');
	});

	it("sheetToCsv quotes bare ID", () => {
		const ws = arrayToSheet([["ID", "Name"]]);
		const csv = sheetToCsv(ws);
		expect(csv).toContain('"ID"');
	});

	it("sheetToCsv with formula-only cell", () => {
		const ws: any = { A1: { t: "z", f: "SUM(B1:B10)" }, "!ref": "A1:A1" };
		const csv = sheetToCsv(ws);
		expect(csv).toContain("=SUM(B1:B10)");
	});

	it("sheetToCsv formula with comma gets quoted", () => {
		const ws: any = { A1: { t: "z", f: "IF(A2,B2,C2)" }, "!ref": "A1:A1" };
		const csv = sheetToCsv(ws);
		expect(csv).toContain('"\'=IF(A2,B2,C2)"');
		expect(sheetToCsv(ws, { escapeFormulae: false })).toContain('"=IF(A2,B2,C2)"');
	});

	it("sheetToTxt produces tab-separated output", () => {
		const ws = arrayToSheet([
			["A", "B"],
			[1, 2],
		]);
		const tsv = sheetToTxt(ws);
		expect(tsv).toContain("A\tB");
		expect(tsv).toContain("1\t2");
	});

	it("csvToSheet with tab separator", () => {
		const ws = csvToSheet("A\tB\n1\t2", { FS: "\t" });
		const rows = sheetToJson(ws, { header: 1 });
		expect(rows[0]).toContain("A");
		expect(rows[0]).toContain("B");
	});

	it("csvToSheet with CRLF line endings", () => {
		const ws = csvToSheet("A,B\r\n1,2\r\n3,4");
		const rows = sheetToJson(ws, { header: 1 });
		expect(rows.length).toBeGreaterThanOrEqual(3);
	});

	it("csvToSheet with quoted fields containing newlines", () => {
		const ws = csvToSheet('"Line1\nLine2",B\n1,2');
		const rows = sheetToJson(ws, { header: 1 });
		expect(rows[0][0]).toBe("Line1\nLine2");
	});

	it("csvToSheet with escaped double-quotes", () => {
		const ws = csvToSheet('"He said ""hi""",B\n1,2');
		const rows = sheetToJson(ws, { header: 1 });
		expect(rows[0][0]).toBe('He said "hi"');
	});

	it("csvToSheet coerces booleans", () => {
		const ws = csvToSheet("TRUE,FALSE,true,false");
		const rows = sheetToJson(ws, { header: 1 });
		expect(rows[0][0]).toBe(true);
		expect(rows[0][1]).toBe(false);
		expect(rows[0][2]).toBe(true);
		expect(rows[0][3]).toBe(false);
	});

	it("null sheet returns empty string", () => {
		expect(sheetToCsv(null as any)).toBe("");
		expect(sheetToCsv({} as any)).toBe("");
	});
});
