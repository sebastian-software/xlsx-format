import type { WorkBook, WriteOptions, WriteResult } from "./types.js";
import { XlsxError } from "./errors.js";
import { zipWrite } from "./zip/index.js";
import { writeZipXlsx } from "./xlsx/write-zip.js";
import { validateWorkbook } from "./xlsx/workbook.js";
import { base64encode } from "./utils/base64.js";
import { resetFormatTable } from "./ssf/table.js";
import { sheetToCsv, sheetToTxt } from "./api/csv.js";
import { sheetToHtml } from "./api/html.js";

/** Encode a text string as a UTF-8 Uint8Array */
function textToUint8Array(text: string): Uint8Array {
	return new TextEncoder().encode(text);
}

/** Convert text output to the requested output type */
function textOutput(text: string, type?: string): WriteResult {
	switch (type) {
		case "string":
			return text;
		case "base64":
			return base64encode(textToUint8Array(text));
		case "buffer":
			if (typeof Buffer !== "undefined") {
				return Buffer.from(text, "utf8");
			}
			return textToUint8Array(text);
		case "array":
			return textToUint8Array(text);
		default:
			return text;
	}
}

type Base64WriteOptions = Omit<WriteOptions, "type"> & { type: "base64" };
type BinaryWriteOptions = Omit<WriteOptions, "type"> & { type: "array" | "buffer" };
type TextWriteOptions = Omit<WriteOptions, "bookType" | "type"> & {
	bookType: "csv" | "tsv" | "html";
	type?: "string";
};
type SpreadsheetWriteOptions = Omit<WriteOptions, "bookType" | "type"> & {
	bookType?: "xlsx" | "xlsm";
	type?: "string";
};

function hasVbaPayload(value: unknown): boolean {
	if (value == null) {
		return false;
	}
	if (typeof value === "string" || Array.isArray(value)) {
		return value.length > 0;
	}
	if (value instanceof ArrayBuffer || ArrayBuffer.isView(value)) {
		return value.byteLength > 0;
	}
	return true;
}

function rejectUnsupportedWriteOptions(wb: WorkBook, options: WriteOptions): void {
	if (options.bookVBA) {
		throw new XlsxError("UNSUPPORTED", 'Write option "bookVBA" is not supported');
	}
	if (options.themeXLSX) {
		throw new XlsxError("UNSUPPORTED", 'Write option "themeXLSX" is not supported');
	}
	if (hasVbaPayload(wb.vbaraw)) {
		throw new XlsxError("UNSUPPORTED", "Workbooks containing VBA data cannot be written");
	}
}

/** Get the first worksheet from a workbook */
function firstSheet(wb: WorkBook) {
	return wb.Sheets[wb.SheetNames[0]];
}

/**
 * Write a WorkBook to an in-memory representation.
 *
 * Supports XLSX (default), CSV, TSV, and HTML output formats via opts.bookType.
 *
 * @param wb - WorkBook object to serialize
 * @param opts - Write options controlling output format and behavior
 * @returns A string for base64 or text output; otherwise a portable Uint8Array
 * @throws XlsxError if a requested compatibility option or non-empty VBA payload is unsupported
 */
export function write(wb: WorkBook, opts?: BinaryWriteOptions | SpreadsheetWriteOptions): Promise<Uint8Array>;
export function write(wb: WorkBook, opts: Base64WriteOptions | TextWriteOptions): Promise<string>;
export function write(wb: WorkBook, opts?: WriteOptions): Promise<WriteResult>;
export async function write(wb: WorkBook, opts?: WriteOptions): Promise<WriteResult> {
	const options: any = { ...opts };
	if (options.password) {
		throw new XlsxError("UNSUPPORTED", "Password-protected workbooks are not supported");
	}
	rejectUnsupportedWriteOptions(wb, options);
	resetFormatTable();
	if (!opts || !(opts as any).unsafe) {
		validateWorkbook(wb);
	}
	// cellStyles implies cellNF (number format) and sheetStubs (empty cell placeholders)
	if (options.cellStyles) {
		options.cellNF = true;
		options.sheetStubs = true;
	}

	const bookType = options.bookType || "xlsx";

	switch (bookType) {
		case "csv": {
			const ws = firstSheet(wb);
			return textOutput(ws ? sheetToCsv(ws, options) : "", options.type);
		}
		case "tsv": {
			const ws = firstSheet(wb);
			return textOutput(ws ? sheetToTxt(ws, options) : "", options.type);
		}
		case "html": {
			const ws = firstSheet(wb);
			return textOutput(ws ? sheetToHtml(ws, options) : "", options.type);
		}
		default: {
			const zip = writeZipXlsx(wb, options);
			const compressed = await zipWrite(zip, !!options.compression);

			switch (options.type) {
				case "base64":
					return base64encode_u8(compressed);
				case "buffer":
					if (typeof Buffer !== "undefined") {
						return Buffer.from(compressed.buffer, compressed.byteOffset, compressed.byteLength);
					}
					return compressed;
				case "array":
					return compressed;
				default:
					return compressed;
			}
		}
	}
}

/** Thin wrapper to encode Uint8Array to base64 */
function base64encode_u8(data: Uint8Array): string {
	return base64encode(data);
}
