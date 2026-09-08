import type {
	BookPropsResult,
	BookSheetsAndPropsResult,
	BookSheetsResult,
	ReadOptions,
	ReadResult,
	WorkBook,
} from "./types.js";
import { XlsxError } from "./errors.js";
import { zipRead } from "./zip/index.js";
import { parseZip } from "./xlsx/parse-zip.js";
import { base64decode } from "./utils/base64.js";
import { resetFormatTable } from "./ssf/table.js";
import { csvToSheet } from "./api/csv.js";
import { htmlToSheet } from "./api/html.js";

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
	throw new XlsxError("INVALID_ARGUMENT", "Unsupported data type for read()");
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

/** Wrap a single worksheet into a WorkBook */
function sheetToWorkBook(ws: any, name?: string): WorkBook {
	const n = name || "Sheet1";
	return {
		SheetNames: [n],
		Sheets: { [n]: ws },
	};
}

type BookSheetsReadOptions = Omit<ReadOptions, "bookSheets" | "bookProps"> & {
	bookSheets: true;
	bookProps?: false;
};

type BookPropsReadOptions = Omit<ReadOptions, "bookSheets" | "bookProps"> & {
	bookSheets?: false;
	bookProps: true;
};

type BookSheetsAndPropsReadOptions = Omit<ReadOptions, "bookSheets" | "bookProps"> & {
	bookSheets: true;
	bookProps: true;
};

type FullReadOptions = Omit<ReadOptions, "bookSheets" | "bookProps"> & {
	bookSheets?: false;
	bookProps?: false;
};

function rejectUnsupportedReadOptions(options: ReadOptions): void {
	for (const [name, requested] of [
		["bookFiles", options.bookFiles],
		["bookVBA", options.bookVBA],
		["bookDeps", options.bookDeps],
		["xlfn", options.xlfn],
	] as const) {
		if (requested) {
			throw new XlsxError("UNSUPPORTED", `Read option "${name}" is not supported`);
		}
	}
}

function selectReadResult(workbook: WorkBook, options: ReadOptions): ReadResult {
	if (options.bookSheets && options.bookProps) {
		return {
			SheetNames: workbook.SheetNames,
			Props: workbook.Props ?? {},
			Custprops: workbook.Custprops ?? {},
		};
	}
	if (options.bookSheets) {
		return { SheetNames: workbook.SheetNames };
	}
	if (options.bookProps) {
		return { Props: workbook.Props ?? {}, Custprops: workbook.Custprops ?? {} };
	}
	return workbook;
}

/**
 * Read a spreadsheet from an in-memory data source.
 *
 * Supports XLSX (ZIP), CSV, and HTML input. For string input with type "string",
 * auto-detects HTML (starts with "<") vs CSV.
 *
 * @param data - File contents as Uint8Array, ArrayBuffer, Buffer, base64 string, binary string, or plain text string
 * @param opts - Read options controlling parsing behavior
 * @returns A full workbook, or the metadata-only shape selected by bookSheets and bookProps
 * @throws XlsxError if the format or an affirmative compatibility option is unsupported
 */
export function read(data: any, opts?: FullReadOptions): Promise<WorkBook>;
export function read(data: any, opts: BookSheetsAndPropsReadOptions): Promise<BookSheetsAndPropsResult>;
export function read(data: any, opts: BookSheetsReadOptions): Promise<BookSheetsResult>;
export function read(data: any, opts: BookPropsReadOptions): Promise<BookPropsResult>;
export function read(data: any, opts?: ReadOptions): Promise<ReadResult>;
export async function read(data: any, opts?: ReadOptions): Promise<ReadResult> {
	const options: ReadOptions = opts ? { ...opts } : {};
	if (options.password) {
		throw new XlsxError("UNSUPPORTED", "Password-protected workbooks are not supported");
	}
	rejectUnsupportedReadOptions(options);
	resetFormatTable();
	if (!options.type) {
		options.type = detect_type(data);
	}

	// Handle plain text string input (CSV or HTML)
	if (options.type === "string" && typeof data === "string") {
		const trimmed = data.trimStart();
		if (trimmed.charAt(0) === "<") {
			return selectReadResult(sheetToWorkBook(htmlToSheet(data)), options);
		}
		return selectReadResult(sheetToWorkBook(csvToSheet(data, options)), options);
	}

	const u8 = to_uint8array(data, options);

	// 0x504B = "PK" -- ZIP file magic number (Phil Katz)
	if (u8[0] === 0x50 && u8[1] === 0x4b) {
		const zip = await zipRead(u8, options);
		return parseZip(zip, options);
	}

	// 0x25504446 = "%PDF" -- PDF file magic number
	if (u8[0] === 0x25 && u8[1] === 0x50 && u8[2] === 0x44 && u8[3] === 0x46) {
		throw new XlsxError("UNSUPPORTED", "PDF File is not a spreadsheet");
	}

	// 0x89504E47 = "\x89PNG" -- PNG file magic number
	if (u8[0] === 0x89 && u8[1] === 0x50 && u8[2] === 0x4e && u8[3] === 0x47) {
		throw new XlsxError("UNSUPPORTED", "PNG Image File is not a spreadsheet");
	}

	throw new XlsxError("UNSUPPORTED", "Unsupported file format. xlsx-format only supports XLSX files.");
}
