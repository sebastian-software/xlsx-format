import type { WorkBook, ReadOptions } from "./types.js";
import { zipRead } from "./zip/index.js";
import { parseZip } from "./xlsx/parse-zip.js";
import { base64decode } from "./utils/base64.js";
import { resetFormatTable } from "./ssf/table.js";
import * as fs from "node:fs";

/**
 * Normalize any supported input type into a Uint8Array for ZIP parsing.
 *
 * Handles Uint8Array, ArrayBuffer, Node Buffer, base64 strings, binary strings, and plain arrays.
 */
function to_uint8array(data: any, opts: ReadOptions): Uint8Array {
	if (data instanceof Uint8Array) {
		return data;
	}
	if (data instanceof ArrayBuffer) {
		return new Uint8Array(data);
	}
	if (typeof Buffer !== "undefined" && Buffer.isBuffer(data)) {
		// Node.js Buffer: create a Uint8Array view over the same memory
		return new Uint8Array(data.buffer, data.byteOffset, data.length);
	}
	if (typeof data === "string") {
		if (opts.type === "base64") {
			return base64decode(data);
		}
		// Treat as a binary string where each character's charCode is a byte value
		const u8 = new Uint8Array(data.length);
		for (let i = 0; i < data.length; ++i) {
			u8[i] = data.charCodeAt(i);
		}
		return u8;
	}
	if (Array.isArray(data)) {
		return new Uint8Array(data);
	}
	throw new Error("Unsupported data type for read()");
}

/**
 * Auto-detect the input data type based on its JavaScript type.
 *
 * Used when the caller does not explicitly set opts.type.
 */
function detect_type(data: any): ReadOptions["type"] {
	if (data instanceof Uint8Array || data instanceof ArrayBuffer) {
		return "array";
	}
	if (typeof Buffer !== "undefined" && Buffer.isBuffer(data)) {
		return "buffer";
	}
	if (typeof data === "string") {
		return "base64";
	}
	return "array";
}

/**
 * Read an XLSX file from an in-memory data source.
 *
 * Accepts multiple input types (Uint8Array, ArrayBuffer, Buffer, base64 string, binary string).
 * Validates the file signature (must be a ZIP/PK file) and delegates parsing to the XLSX module.
 *
 * @param data - File contents as Uint8Array, ArrayBuffer, Buffer, base64 string, or binary string
 * @param opts - Read options controlling parsing behavior
 * @returns Promise resolving to a parsed WorkBook object
 * @throws Error if the input is a PDF, PNG, or other unsupported format
 */
export async function read(data: any, opts?: ReadOptions): Promise<WorkBook> {
	resetFormatTable();
	const options: any = opts ? { ...opts } : {};
	if (!options.type) {
		options.type = detect_type(data);
	}

	const u8 = to_uint8array(data, options);

	// 0x504B = "PK" -- ZIP file magic number (Phil Katz)
	if (u8[0] === 0x50 && u8[1] === 0x4b) {
		const zip = await zipRead(u8);
		return parseZip(zip, options);
	}

	// 0x25504446 = "%PDF" -- PDF file magic number
	if (u8[0] === 0x25 && u8[1] === 0x50 && u8[2] === 0x44 && u8[3] === 0x46) {
		throw new Error("PDF File is not a spreadsheet");
	}

	// 0x89504E47 = "\x89PNG" -- PNG file magic number
	if (u8[0] === 0x89 && u8[1] === 0x50 && u8[2] === 0x4e && u8[3] === 0x47) {
		throw new Error("PNG Image File is not a spreadsheet");
	}

	throw new Error("Unsupported file format. xlsx-format only supports XLSX files.");
}

/**
 * Read an XLSX file from the local filesystem (Node.js only).
 *
 * Convenience wrapper that reads the file synchronously with fs.readFileSync
 * and then delegates to {@link read}.
 *
 * @param filename - Path to the XLSX file
 * @param opts - Read options controlling parsing behavior
 * @returns Promise resolving to a parsed WorkBook object
 */
export async function readFile(filename: string, opts?: ReadOptions): Promise<WorkBook> {
	const data = fs.readFileSync(filename);
	return read(new Uint8Array(data), opts);
}
