import type { WorkBook, WriteOptions } from "./types.js";
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
function textOutput(text: string, type?: string): any {
	switch (type) {
		case "string":
			return text;
		case "base64":
			return base64encode(textToUint8Array(text));
		case "buffer":
			if (typeof Buffer !== "undefined") {
				return Buffer.from(text, "utf-8");
			}
			return textToUint8Array(text);
		case "array":
			return textToUint8Array(text);
		default:
			return text;
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
 * @returns Promise resolving to the serialized data in the requested format
 */
export async function write(wb: WorkBook, opts?: WriteOptions): Promise<any> {
	resetFormatTable();
	if (!opts || !(opts as any).unsafe) {
		validateWorkbook(wb);
	}
	const options: any = { ...(opts || {}) };
	// cellStyles implies cellNF (number format) and sheetStubs (empty cell placeholders)
	if (options.cellStyles) {
		options.cellNF = true;
		options.sheetStubs = true;
	}

	const bookType = options.bookType || "xlsx";

	switch (bookType) {
		case "csv": {
			const ws = firstSheet(wb);
			return textOutput(ws ? sheetToCsv(ws) : "", options.type);
		}
		case "tsv": {
			const ws = firstSheet(wb);
			return textOutput(ws ? sheetToTxt(ws) : "", options.type);
		}
		case "html": {
			const ws = firstSheet(wb);
			return textOutput(ws ? sheetToHtml(ws) : "", options.type);
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
