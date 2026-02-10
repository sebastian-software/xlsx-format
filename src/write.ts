import type { WorkBook, WriteOptions } from "./types.js";
import { zip_write } from "./zip/index.js";
import { write_zip_xlsx } from "./xlsx/write-zip.js";
import { check_wb } from "./xlsx/workbook.js";
import { base64encode } from "./utils/base64.js";
import { make_ssf } from "./ssf/table.js";
import { dup } from "./utils/helpers.js";
import * as fs from "node:fs";

/**
 * Write a WorkBook to a Uint8Array (XLSX format).
 *
 * @param wb - WorkBook object to write
 * @param opts - Write options
 * @returns File contents as Uint8Array, base64 string, or Buffer depending on opts.type
 */
export function write(wb: WorkBook, opts?: WriteOptions): any {
	make_ssf();
	if (!opts || !(opts as any).unsafe) {
		check_wb(wb);
	}
	const o: any = dup(opts || {});
	if (o.cellStyles) {
		o.cellNF = true;
		o.sheetStubs = true;
	}

	const zip = write_zip_xlsx(wb, o);
	const compressed = zip_write(zip, !!o.compression);

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
 */
export function writeFile(wb: WorkBook, filename: string, opts?: WriteOptions): void {
	const o: any = opts ? dup(opts) : {};
	o.type = "buffer";
	const data = write(wb, o);
	fs.writeFileSync(filename, data instanceof Uint8Array ? Buffer.from(data) : data);
}
