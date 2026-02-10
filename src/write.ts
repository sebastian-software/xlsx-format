import type { WorkBook, WriteOptions } from "./types.js";
import { zipWrite } from "./zip/index.js";
import { writeZipXlsx } from "./xlsx/write-zip.js";
import { validateWorkbook } from "./xlsx/workbook.js";
import { base64encode } from "./utils/base64.js";
import { resetFormatTable } from "./ssf/table.js";
import * as fs from "node:fs";

/**
 * Write a WorkBook to a Uint8Array (XLSX format).
 *
 * @param wb - WorkBook object to write
 * @param opts - Write options
 * @returns Promise resolving to file contents as Uint8Array, base64 string, or Buffer depending on opts.type
 */
export async function write(wb: WorkBook, opts?: WriteOptions): Promise<any> {
	resetFormatTable();
	if (!opts || !(opts as any).unsafe) {
		validateWorkbook(wb);
	}
	const o: any = { ...(opts || {}) };
	if (o.cellStyles) {
		o.cellNF = true;
		o.sheetStubs = true;
	}

	const zip = writeZipXlsx(wb, o);
	const compressed = await zipWrite(zip, !!o.compression);

	switch (o.type) {
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

function base64encode_u8(data: Uint8Array): string {
	return base64encode(data);
}

/**
 * Write a WorkBook to a file (XLSX format).
 *
 * @param wb - WorkBook object to write
 * @param filename - Output file path
 * @param opts - Write options
 * @returns Promise that resolves when the file has been written
 */
export async function writeFile(wb: WorkBook, filename: string, opts?: WriteOptions): Promise<void> {
	const o: any = opts ? { ...opts } : {};
	o.type = "buffer";
	const data = await write(wb, o);
	fs.writeFileSync(filename, data instanceof Uint8Array ? Buffer.from(data) : data);
}
