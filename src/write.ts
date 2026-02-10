import type { WorkBook, WriteOptions } from "./types.js";
import { zipWrite } from "./zip/index.js";
import { writeZipXlsx } from "./xlsx/write-zip.js";
import { validateWorkbook } from "./xlsx/workbook.js";
import { base64encode } from "./utils/base64.js";
import { resetFormatTable } from "./ssf/table.js";
import * as fs from "node:fs";

/**
 * Write a WorkBook to an in-memory representation (XLSX format).
 *
 * Serializes the workbook into a ZIP archive containing the OOXML parts,
 * then returns the result in the format specified by opts.type.
 *
 * @param wb - WorkBook object to serialize
 * @param opts - Write options controlling output format and behavior
 * @returns Promise resolving to Uint8Array (default/"array"), base64 string, or Buffer depending on opts.type
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

/** Thin wrapper to encode Uint8Array to base64 */
function base64encode_u8(data: Uint8Array): string {
	return base64encode(data);
}

/**
 * Write a WorkBook to a file on the local filesystem (Node.js only).
 *
 * Convenience wrapper that calls {@link write} with type "buffer" and
 * writes the result synchronously with fs.writeFileSync.
 *
 * @param wb - WorkBook object to serialize
 * @param filename - Output file path
 * @param opts - Write options controlling output format and behavior
 * @returns Promise that resolves when the file has been written
 */
export async function writeFile(wb: WorkBook, filename: string, opts?: WriteOptions): Promise<void> {
	const options: any = opts ? { ...opts } : {};
	options.type = "buffer";
	const data = await write(wb, options);
	fs.writeFileSync(filename, data instanceof Uint8Array ? Buffer.from(data) : data);
}
