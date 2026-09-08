import { describe, it, expect } from "vitest";
import { read, write, arrayToSheet, createWorkbook, jsonToSheet, XlsxError } from "../index.js";
import { zipRead, zipWrite } from "../zip/index.js";

const encoder = new TextEncoder();
const decoder = new TextDecoder();

async function workbookWithSharedStrings(
	stringItems = "<si><t>one</t></si><si><t>two</t></si>",
	indexes: [number, number] = [0, 1],
): Promise<Uint8Array> {
	const zip = await zipRead(await write(createWorkbook(arrayToSheet([["one", "two"]]), "S")));
	const sheetPath = "xl/worksheets/sheet1.xml";
	zip.files[sheetPath] = encoder.encode(
		decoder
			.decode(zip.files[sheetPath])
			.replace('t="str"><v>one</v>', `t="s"><v>${indexes[0]}</v>`)
			.replace('t="str"><v>two</v>', `t="s"><v>${indexes[1]}</v>`),
	);
	zip.files["[Content_Types].xml"] = encoder.encode(
		decoder
			.decode(zip.files["[Content_Types].xml"])
			.replace(
				"</Types>",
				'<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>',
			),
	);
	zip.files["xl/sharedStrings.xml"] = encoder.encode(
		`<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${stringItems}</sst>`,
	);
	return zipWrite(zip);
}

async function removeZipPart(data: Uint8Array, path: string): Promise<Uint8Array> {
	const zip = await zipRead(data);
	Reflect.deleteProperty(zip.files, path);
	return zipWrite(zip);
}

async function expectReadError(
	data: Uint8Array,
	options: Parameters<typeof read>[1],
	code: XlsxError["code"],
	message: RegExp,
): Promise<void> {
	const result = read(data, options);
	await expect(result).rejects.toBeInstanceOf(XlsxError);
	await expect(result).rejects.toMatchObject({ code });
	await expect(result).rejects.toThrow(message);
}

describe("parse-zip: bookSheets and bookProps options", () => {
	it("reads with bookSheets to get only sheet names", async () => {
		const ws = jsonToSheet([{ a: 1 }, { a: 2 }]);
		const wb = createWorkbook(ws, "TestSheet");
		const buf = await write(wb);
		const result = await read(buf, { bookSheets: true });
		expect(result.SheetNames).toBeDefined();
		expect(result.SheetNames).toContain("TestSheet");
		// Should not have full sheet data
	});

	it("reads with bookProps to get document properties", async () => {
		const ws = jsonToSheet([{ a: 1 }]);
		const wb = createWorkbook(ws, "S1");
		const buf = await write(wb);
		const result = await read(buf, { bookProps: true });
		expect(result.Props).toBeDefined();
	});

	it("reads with specific sheet by index", async () => {
		const wb = createWorkbook(jsonToSheet([{ a: 1 }]), "S1");
		const ws2 = jsonToSheet([{ b: 2 }]);
		wb.SheetNames.push("S2");
		wb.Sheets.S2 = ws2;
		const buf = await write(wb);
		const result = await read(buf, { sheets: 0 });
		expect(result.SheetNames).toContain("S1");
		// S2 should exist in SheetNames but may have empty data
	});

	it("reads with specific sheet by name", async () => {
		const wb = createWorkbook(jsonToSheet([{ a: 1 }]), "S1");
		const ws2 = jsonToSheet([{ b: 2 }]);
		wb.SheetNames.push("S2");
		wb.Sheets.S2 = ws2;
		const buf = await write(wb);
		const result = await read(buf, { sheets: "S2" });
		expect(result.SheetNames).toBeDefined();
	});

	it("reads with sheets as array of indices and names", async () => {
		const wb = createWorkbook(jsonToSheet([{ a: 1 }]), "S1");
		wb.SheetNames.push("S2");
		wb.Sheets.S2 = jsonToSheet([{ b: 2 }]);
		wb.SheetNames.push("S3");
		wb.Sheets.S3 = jsonToSheet([{ c: 3 }]);
		const buf = await write(wb);
		const result = await read(buf, { sheets: [0, "S3"] });
		expect(result.SheetNames).toBeDefined();
	});
});

describe("parse-zip: dense mode and cellStyles", () => {
	it("reads in dense mode", async () => {
		const ws = jsonToSheet([{ a: 1, b: "text" }]);
		const wb = createWorkbook(ws, "S1");
		const buf = await write(wb);
		const result = await read(buf, { dense: true });
		const sheet = result.Sheets.S1;
		expect(sheet["!data"]).toBeDefined();
	});

	it("reads with cellStyles to populate column info", async () => {
		const ws = jsonToSheet([{ a: 1, b: 2 }]);
		ws["!cols"] = [{ width: 15 }, { width: 20 }];
		const wb = createWorkbook(ws, "S1");
		const buf = await write(wb);
		const result = await read(buf, { cellStyles: true });
		const sheet = result.Sheets.S1;
		expect(sheet["!cols"]).toBeDefined();
	});
});

describe("parse-zip: required parts and configured limits", () => {
	it("enforces worksheet cell limits without requiring WTF mode", async () => {
		const data = await write(createWorkbook(arrayToSheet([[1, 2]]), "S"));

		const workbook = await read(data, { maxWorksheetCells: 2 });
		expect(workbook.Sheets.S.A1?.v).toBe(1);
		expect(workbook.Sheets.S.B1?.v).toBe(2);

		await expectReadError(data, { maxWorksheetCells: 1 }, "LIMIT_EXCEEDED", /worksheet cell count 2/);
		await expectReadError(data, { maxWorksheetCells: -1 }, "INVALID_ARGUMENT", /maxWorksheetCells/);
	});

	it("counts shared strings exactly and propagates their configured limit", async () => {
		const data = await workbookWithSharedStrings();

		const workbook = await read(data, { maxSharedStringItems: 2 });
		expect(workbook.Sheets.S.A1?.v).toBe("one");
		expect(workbook.Sheets.S.B1?.v).toBe("two");

		await expectReadError(data, { maxSharedStringItems: 1 }, "LIMIT_EXCEEDED", /shared string item count 2/);
	});

	it("rejects a missing selected worksheet but does not load unselected sheets", async () => {
		const workbook = createWorkbook(arrayToSheet([[1]]), "Present");
		workbook.SheetNames.push("Missing");
		workbook.Sheets.Missing = arrayToSheet([[2]]);
		const data = await removeZipPart(await write(workbook), "xl/worksheets/sheet2.xml");

		const selected = await read(data, { sheets: "Present" });
		expect(selected.Sheets.Present.A1?.v).toBe(1);
		expect(selected.Sheets.Missing).toBeUndefined();

		await expectReadError(data, { sheets: "Missing" }, "NOT_FOUND", /xl\/worksheets\/sheet2\.xml/);
	});

	it("rejects a missing declared shared strings part for full reads", async () => {
		const data = await removeZipPart(await workbookWithSharedStrings(), "xl/sharedStrings.xml");

		await expectReadError(data, undefined, "NOT_FOUND", /xl\/sharedStrings\.xml/);

		const namesOnly = await read(data, { bookSheets: true });
		expect(namesOnly.SheetNames).toStrictEqual(["S"]);
		const propsOnly = await read(data, { bookProps: true });
		expect(propsOnly.Props).toBeDefined();
	});

	it("recovers from malformed optional worksheet relationships in normal mode", async () => {
		const source = createWorkbook(arrayToSheet([["value"]]), "S");
		const zip = await zipRead(await write(source));
		zip.files["xl/worksheets/_rels/sheet1.xml.rels"] = encoder.encode(
			'<Relationships><Relationship Id="rId1"/></Relationships>',
		);
		const data = await zipWrite(zip);

		const workbook = await read(data);
		expect(workbook.Sheets.S.A1?.v).toBe("value");
		await expect(read(data, { WTF: true })).rejects.toThrow(/undefined/);
	});

	it("does not suppress configured limits for optional worksheet relationships", async () => {
		const source = createWorkbook(arrayToSheet([["value"]]), "S");
		const zip = await zipRead(await write(source));
		const relsPath = "xl/worksheets/_rels/sheet1.xml.rels";
		const largestRequiredPart = Math.max(
			...Object.entries(zip.files)
				.filter(([path]) => path.endsWith(".xml"))
				.map(([, bytes]) => bytes.length),
		);
		zip.files[relsPath] = encoder.encode(
			`<Relationships><!--${"x".repeat(largestRequiredPart)}--><Relationship Id="rId1"/></Relationships>`,
		);
		const data = await zipWrite(zip);

		await expectReadError(
			data,
			{ maxXmlPartBytes: largestRequiredPart },
			"LIMIT_EXCEEDED",
			/sheet1\.xml\.rels size/,
		);
	});

	it("preserves cell references after empty shared string entries", async () => {
		const data = await workbookWithSharedStrings(
			'<si/><si><t>one</t></si><x:si xmlns:x="urn:test"><x:t>two</x:t></x:si>',
			[1, 2],
		);

		const workbook = await read(data, { maxSharedStringItems: 3 });
		expect(workbook.Sheets.S.A1?.v).toBe("one");
		expect(workbook.Sheets.S.B1?.v).toBe("two");
		await expectReadError(data, { maxSharedStringItems: 2 }, "LIMIT_EXCEEDED", /shared string item count 3/);
	});
});
