import { describe, expect, it } from "vitest";
import {
	appendSheet,
	createSheet,
	createWorkbook,
	getSheetIndex,
	read,
	setSheetVisibility,
	sheetToJson,
	XlsxError,
} from "./index.js";
import { parseContentTypes } from "./opc/content-types.js";
import { formatNumber } from "./ssf/format.js";
import type { CellStyle, WorkBook, WorkSheet } from "./types.js";
import { parseZip } from "./xlsx/parse-zip.js";
import { parseSstXml } from "./xlsx/shared-strings.js";
import { buildStyleRegistry } from "./xlsx/styles.js";
import { zipRead, zipWrite } from "./zip/index.js";

const encoder = new TextEncoder();

function bytes(text: string): Uint8Array {
	return encoder.encode(text);
}

function readU16(data: Uint8Array, off: number): number {
	return data[off] | (data[off + 1] << 8);
}

function readU32(data: Uint8Array, off: number): number {
	return (data[off] | (data[off + 1] << 8) | (data[off + 2] << 16) | (data[off + 3] << 24)) >>> 0;
}

function writeU16(data: Uint8Array, off: number, value: number): void {
	data[off] = value & 0xff;
	data[off + 1] = (value >> 8) & 0xff;
}

function writeU32(data: Uint8Array, off: number, value: number): void {
	data[off] = value & 0xff;
	data[off + 1] = (value >> 8) & 0xff;
	data[off + 2] = (value >> 16) & 0xff;
	data[off + 3] = (value >> 24) & 0xff;
}

function findEocd(data: Uint8Array): number {
	for (let i = data.length - 22; i >= 0; --i) {
		if (readU32(data, i) === 0x06054b50) {
			return i;
		}
	}
	throw new Error("EOCD not found in test archive");
}

async function expectXlsxError(promise: Promise<unknown>, code: XlsxError["code"], message: RegExp): Promise<void> {
	await expect(promise).rejects.toBeInstanceOf(XlsxError);
	await expect(promise).rejects.toMatchObject({ code });
	await expect(promise).rejects.toThrow(message);
}

function expectSyncXlsxError(action: () => unknown, code: XlsxError["code"], message: RegExp): void {
	let error: unknown;
	try {
		action();
	} catch (error_) {
		error = error_;
	}

	expect(error).toBeInstanceOf(XlsxError);
	expect(error).toMatchObject({ code });
	expect((error as Error).message).toMatch(message);
}

function workbookWithStyle(style: CellStyle): WorkBook {
	const ws: WorkSheet = {
		"!ref": "A1",
		A1: { t: "s", v: "styled", s: style },
	};
	return createWorkbook(ws, "Styles");
}

describe("XlsxError", () => {
	it("exposes stable code metadata", () => {
		const cause = new Error("source");
		const error = new XlsxError("MALFORMED", "Invalid workbook", { cause });

		expect(error).toBeInstanceOf(Error);
		expect(error).toBeInstanceOf(XlsxError);
		expect(error.name).toBe("XlsxError");
		expect(error.code).toBe("MALFORMED");
		expect(error.message).toBe("Invalid workbook");
		expect(error.cause).toBe(cause);
	});

	it("classifies read() argument and file format failures", async () => {
		await expectXlsxError(read(123), "INVALID_ARGUMENT", /Unsupported data type/);
		await expectXlsxError(read(new Uint8Array([0x25, 0x50, 0x44, 0x46])), "UNSUPPORTED", /PDF/);
	});

	it("classifies malformed ZIP data and resource limits", async () => {
		await expectXlsxError(zipRead(new Uint8Array([0x50, 0x4b])), "MALFORMED", /EOCD not found/);

		const archive = await zipWrite({
			files: {
				"a.txt": bytes("a"),
				"b.txt": bytes("b"),
			},
		});

		await expectXlsxError(zipRead(archive, { maxZipEntries: 1 }), "LIMIT_EXCEEDED", /entry count 2/);
	});

	it("classifies ZIP CRC mismatches", async () => {
		const archive = await zipWrite({ files: { "a.txt": bytes("a") } });
		const corrupted = archive.slice();
		const nameLength = readU16(corrupted, 26);
		const extraLength = readU16(corrupted, 28);
		corrupted[30 + nameLength + extraLength] ^= 0xff;

		await expectXlsxError(zipRead(corrupted), "CRC_MISMATCH", /CRC mismatch/);
	});

	it("classifies additional ZIP structure failures", async () => {
		const archive = await zipWrite({ files: { "a.txt": bytes("abcd") } });
		const eocd = findEocd(archive);
		const centralDirectoryOffset = readU32(archive, eocd + 16);

		const overlappingCentralDirectory = archive.slice();
		writeU32(overlappingCentralDirectory, eocd + 12, 1);
		writeU32(overlappingCentralDirectory, eocd + 16, eocd);
		await expectXlsxError(zipRead(overlappingCentralDirectory), "MALFORMED", /central directory overlaps EOCD/);

		const truncatedCentralDirectoryHeader = archive.slice();
		writeU32(truncatedCentralDirectoryHeader, eocd + 12, 1);
		await expectXlsxError(
			zipRead(truncatedCentralDirectoryHeader),
			"MALFORMED",
			/central directory entry exceeds declared size/,
		);

		const truncatedCentralDirectoryEntry = archive.slice();
		writeU32(truncatedCentralDirectoryEntry, eocd + 12, 46);
		await expectXlsxError(
			zipRead(truncatedCentralDirectoryEntry),
			"MALFORMED",
			/central directory entry exceeds declared size/,
		);

		const badCentralDirectorySignature = archive.slice();
		writeU32(badCentralDirectorySignature, centralDirectoryOffset, 0);
		await expectXlsxError(zipRead(badCentralDirectorySignature), "MALFORMED", /bad central directory entry/);

		const multiDiskCentralDirectoryEntry = archive.slice();
		writeU16(multiDiskCentralDirectoryEntry, centralDirectoryOffset + 34, 1);
		await expectXlsxError(zipRead(multiDiskCentralDirectoryEntry), "UNSUPPORTED", /multi-disk archives/);

		const badLocalHeaderSignature = archive.slice();
		writeU32(badLocalHeaderSignature, 0, 0);
		await expectXlsxError(zipRead(badLocalHeaderSignature), "MALFORMED", /bad local file header/);

		const deflated = await zipWrite({ files: { "a.txt": bytes("abcd") } }, true);
		const deflatedEocd = findEocd(deflated);
		const deflatedCentralDirectoryOffset = readU32(deflated, deflatedEocd + 16);
		const mismatchedInflatedSize = deflated.slice();
		writeU32(mismatchedInflatedSize, deflatedCentralDirectoryOffset + 24, 5);
		await expectXlsxError(zipRead(mismatchedInflatedSize), "MALFORMED", /inflated to 4 bytes, expected 5/);
	});

	it("classifies workbook and sheet API failures", () => {
		const fullWorkbook: WorkBook = {
			SheetNames: Array.from({ length: 0xffff }, (_, index) => "Sheet" + index),
			Sheets: {},
		};
		expectSyncXlsxError(
			() => appendSheet(fullWorkbook, createSheet(), "Overflow"),
			"LIMIT_EXCEEDED",
			/Too many worksheets/,
		);

		const wb = createWorkbook(createSheet(), "A");
		expectSyncXlsxError(() => appendSheet(wb, createSheet(), "A"), "DUPLICATE", /already exists/);
		expectSyncXlsxError(() => getSheetIndex(wb, "missing"), "NOT_FOUND", /Cannot find sheet name \|missing\|/);
		expectSyncXlsxError(() => getSheetIndex(wb, true as any), "NOT_FOUND", /Cannot find sheet \|true\|/);
		expectSyncXlsxError(
			() => {
				setSheetVisibility(wb, "A", 3 as any);
			},
			"INVALID_ARGUMENT",
			/Bad sheet visibility/,
		);
	});

	it("classifies malformed worksheet and package metadata", () => {
		expectSyncXlsxError(
			() => sheetToJson({ "!ref": "A1", A1: { t: "x" as any, v: 1 } }, { header: 1 }),
			"INVALID_ARGUMENT",
			/unrecognized type x/,
		);
		expectSyncXlsxError(
			() => parseContentTypes('<Types xmlns="urn:unknown"></Types>'),
			"UNSUPPORTED",
			/Unknown Namespace/,
		);
		expectSyncXlsxError(
			() =>
				parseSstXml(
					'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><r><rPr><foo/></rPr><t>x</t></r></si></sst>',
				),
			"UNSUPPORTED",
			/Unrecognized rich format/,
		);
	});

	it("classifies SSF format failures", () => {
		expectSyncXlsxError(() => formatNumber("General", (() => 1) as any), "UNSUPPORTED", /unsupported value/);
		expectSyncXlsxError(() => formatNumber("hhh AM/PM", 0.5), "MALFORMED", /bad hour format/);
		expectSyncXlsxError(() => formatNumber("hhh", 0.5), "MALFORMED", /bad hour format/);
		expectSyncXlsxError(() => formatNumber("h:mmm", 0.5), "MALFORMED", /bad minute format/);
		expectSyncXlsxError(() => formatNumber("sss", 0.5), "MALFORMED", /bad second format/);
		expectSyncXlsxError(() => formatNumber("[hhh]", 0.5), "MALFORMED", /bad abstime format/);
		expectSyncXlsxError(() => formatNumber("0,0", 1), "UNSUPPORTED", /unsupported format \|0,0\|/);
		expectSyncXlsxError(() => formatNumber("0,0", 1.5), "UNSUPPORTED", /unsupported format \|0,0\|/);
		expectSyncXlsxError(() => formatNumber("Gx", 1), "MALFORMED", /unrecognized character G/);
		expectSyncXlsxError(() => formatNumber("[Red", 1), "MALFORMED", /unterminated "\[" block/);
		expectSyncXlsxError(() => formatNumber("0Q", 1), "MALFORMED", /unrecognized character Q/);
		expectSyncXlsxError(() => formatNumber("0;0;0;@;0", 1), "MALFORMED", /cannot find right format/);
	});

	it("classifies XLSX package structure failures", () => {
		expectSyncXlsxError(() => parseZip({ files: {} }), "UNSUPPORTED", /Unsupported ZIP file/);

		const emptyContentTypes =
			'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>';
		expectSyncXlsxError(
			() => parseZip({ files: { "[Content_Types].xml": bytes(emptyContentTypes) } }),
			"NOT_FOUND",
			/Could not find workbook/,
		);

		const missingWorkbook =
			'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>';
		expectSyncXlsxError(
			() => parseZip({ files: { "[Content_Types].xml": bytes(missingWorkbook) } }),
			"NOT_FOUND",
			/Could not find xl\/workbook\.xml/,
		);
	});

	it("classifies strict style serialization failures", () => {
		expectSyncXlsxError(
			() => buildStyleRegistry(workbookWithStyle({ font: { color: { argb: "not-a-color" } } }), { WTF: true }),
			"UNSUPPORTED",
			/Unsupported style color/,
		);
		expectSyncXlsxError(
			() =>
				buildStyleRegistry(
					workbookWithStyle({ fill: { patternType: "darkGrid" as any, fgColor: { argb: "FFFFFFFF" } } }),
					{ WTF: true },
				),
			"UNSUPPORTED",
			/Unsupported fill pattern/,
		);
		expectSyncXlsxError(
			() => buildStyleRegistry(workbookWithStyle({ border: { top: { style: "dashed" as any } } }), { WTF: true }),
			"UNSUPPORTED",
			/Unsupported border style/,
		);
		expectSyncXlsxError(
			() => buildStyleRegistry(workbookWithStyle({ alignment: { horizontal: "justify" as any } }), { WTF: true }),
			"UNSUPPORTED",
			/Unsupported horizontal alignment/,
		);
		expectSyncXlsxError(
			() => buildStyleRegistry(workbookWithStyle({ alignment: { vertical: "justify" as any } }), { WTF: true }),
			"UNSUPPORTED",
			/Unsupported vertical alignment/,
		);
	});
});
