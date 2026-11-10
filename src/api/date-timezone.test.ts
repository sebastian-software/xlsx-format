import { afterEach, describe, expect, it } from "vitest";
import { arrayToSheet, createWorkbook, jsonToSheet, sheetToJson, write } from "../index.js";
import { localToUtc, utcToLocal } from "../utils/date.js";
import { zipRead, zipReadString } from "../zip/index.js";

const originalTimezone = process.env.TZ;
afterEach(() => {
	if (originalTimezone === undefined) {
		delete process.env.TZ;
	} else {
		process.env.TZ = originalTimezone;
	}
});

for (const timezone of ["Europe/Berlin", "America/New_York"]) {
	describe(`date conversions in ${timezone}`, () => {
		// Expected serials are fixed Excel values, independent of the date helpers.
		for (const [month, day, hour, serial] of [
			[0, 1, 0, 45292],
			[6, 1, 12, 45474.5],
			[2, 10, 0, 45361],
			[2, 10, 4, 45361 + 4 / 24],
			[2, 31, 0, 45382],
			[2, 31, 4, 45382 + 4 / 24],
			[9, 27, 4, 45592 + 4 / 24],
			[10, 3, 4, 45599 + 4 / 24],
		]) {
			it(`preserves local components for ${month + 1}/${day} at ${hour}:00`, () => {
				process.env.TZ = timezone;
				const local = new Date(2024, month, day, hour);
				const utc = new Date(Date.UTC(2024, month, day, hour));
				expect(localToUtc(local).getTime()).toBe(utc.getTime());
				expect(utcToLocal(utc).getTime()).toBe(local.getTime());
				for (const UTC of [undefined, false, true]) {
					const input = UTC ? utc : local;
					for (const cellDates of [false, true]) {
						const sheets = [
							arrayToSheet([[input]], { UTC, cellDates }),
							jsonToSheet([{ date: input }], { UTC, cellDates, skipHeader: true }),
						];
						for (const sheet of sheets) {
							expect(sheet.A1.v).toStrictEqual(cellDates ? utc : serial);
						}
					}
					const sheet = { "!ref": "A1", A1: { t: "n" as const, v: serial, z: "yyyy-mm-dd hh:mm:ss" } };
					const output = sheetToJson<Date[]>(sheet, { header: 1, UTC });
					expect(output[0][0].getTime()).toBe(input.getTime());
				}
			});
		}

		it("keeps ISO output's UTC default and explicit local formatting", () => {
			process.env.TZ = timezone;
			const sheet = { "!ref": "A1", A1: { t: "n" as const, v: 45292, z: "yyyy-mm-dd hh:mm:ss" } };
			for (const UTC of [undefined, true]) {
				expect(sheetToJson(sheet, { header: 1, dateOutput: "iso", UTC })).toStrictEqual([
					["2024-01-01T00:00:00"],
				]);
			}
			expect(sheetToJson(sheet, { header: 1, dateOutput: "iso", UTC: false })).toStrictEqual([
				[timezone === "Europe/Berlin" ? "2024-01-01T01:00:00" : "2023-12-31T19:00:00"],
			]);
		});

		it("writes the independent Excel serial into XLSX XML", async () => {
			process.env.TZ = timezone;
			for (const cellDates of [false, true]) {
				const sheet = arrayToSheet([[new Date(2024, 0, 1)]], { cellDates });
				const zip = await zipRead(await write(createWorkbook(sheet)));
				expect(zipReadString(zip, "xl/worksheets/sheet1.xml")).toContain("<v>45292</v>");
			}
		});
	});
}
