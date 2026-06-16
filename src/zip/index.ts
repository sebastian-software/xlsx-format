import { XlsxError } from "../errors.js";
import { crc32 } from "./crc32.js";
import { inflate, deflate } from "./streams.js";

/** In-memory representation of a ZIP archive as a flat filename-to-data map. */
export interface ZipArchive {
	/** Map of file paths to their uncompressed contents. */
	files: Record<string, Uint8Array>;
}

/** Limits applied while reading untrusted ZIP archives. */
export interface ZipReadOptions {
	/** Maximum number of central-directory entries to parse. */
	maxZipEntries?: number;
	/** Maximum total uncompressed size across file entries. */
	maxTotalUncompressedBytes?: number;
	/** Maximum uncompressed size for a single file entry. */
	maxEntryUncompressedBytes?: number;
}

const encoder = new TextEncoder();
const decoder = new TextDecoder();

const DEFAULT_MAX_ZIP_ENTRIES = 10000;
const DEFAULT_MAX_TOTAL_UNCOMPRESSED_BYTES = 512 * 1024 * 1024;
const DEFAULT_MAX_ENTRY_UNCOMPRESSED_BYTES = 256 * 1024 * 1024;

// -- ZIP format signatures (little-endian magic numbers) --
/** Local file header signature: "PK\x03\x04" */
const SIG_LOCAL = 0x04034b50;
/** Central directory file header signature: "PK\x01\x02" */
const SIG_CENTRAL = 0x02014b50;
/** End of Central Directory record signature: "PK\x05\x06" */
const SIG_EOCD = 0x06054b50;

const ZIP64_U16 = 0xffff;
const ZIP64_U32 = 0xffffffff;

function optionLimit(value: number | undefined, fallback: number, name: string): number {
	if (value == null) {
		return fallback;
	}
	if (!Number.isFinite(value) || value < 0) {
		throw new XlsxError("INVALID_ARGUMENT", `Invalid ZIP option: ${name} must be a non-negative finite number`);
	}
	return value;
}

function assertRange(data: Uint8Array, off: number, len: number, what: string): void {
	if (!Number.isInteger(off) || !Number.isInteger(len) || off < 0 || len < 0 || off > data.length - len) {
		throw new XlsxError("MALFORMED", `Invalid ZIP: ${what} out of bounds`);
	}
}

function readU16At(buf: Uint8Array, off: number, what: string): number {
	assertRange(buf, off, 2, what);
	return readU16(buf, off);
}

function readU32At(buf: Uint8Array, off: number, what: string): number {
	assertRange(buf, off, 4, what);
	return readU32(buf, off);
}

function rejectZip64(): never {
	throw new XlsxError("UNSUPPORTED", "Unsupported ZIP: Zip64 archives are not supported");
}

function assertCrc(name: string, data: Uint8Array, expected: number): void {
	const actual = crc32(data);
	if (actual !== expected) {
		throw new XlsxError("CRC_MISMATCH", `Invalid ZIP: CRC mismatch for ${name}`);
	}
}

/** Read an unsigned 16-bit little-endian integer from a buffer */
function readU16(buf: Uint8Array, off: number): number {
	return buf[off] | (buf[off + 1] << 8);
}

/** Read an unsigned 32-bit little-endian integer from a buffer */
function readU32(buf: Uint8Array, off: number): number {
	return (buf[off] | (buf[off + 1] << 8) | (buf[off + 2] << 16) | (buf[off + 3] << 24)) >>> 0;
}

/** Write an unsigned 16-bit little-endian integer to a buffer */
function writeU16(buf: Uint8Array, off: number, val: number): void {
	buf[off] = val & 0xff;
	buf[off + 1] = (val >> 8) & 0xff;
}

/** Write an unsigned 32-bit little-endian integer to a buffer */
function writeU32(buf: Uint8Array, off: number, val: number): void {
	buf[off] = val & 0xff;
	buf[off + 1] = (val >> 8) & 0xff;
	buf[off + 2] = (val >> 16) & 0xff;
	buf[off + 3] = (val >> 24) & 0xff;
}

/**
 * Parse a ZIP archive from raw bytes.
 *
 * Locates the End of Central Directory (EOCD) record by scanning backward,
 * then reads the central directory to find all file entries. Deflated entries
 * are decompressed one at a time to keep peak memory bounded by entry limits.
 *
 * @param data - Raw ZIP file bytes
 * @returns Parsed archive with decompressed file contents
 * @throws XlsxError if the ZIP structure is invalid or uses an unsupported compression method
 */
export async function zipRead(data: Uint8Array, opts?: ZipReadOptions): Promise<ZipArchive> {
	const maxZipEntries = optionLimit(opts?.maxZipEntries, DEFAULT_MAX_ZIP_ENTRIES, "maxZipEntries");
	const maxTotalUncompressedBytes = optionLimit(
		opts?.maxTotalUncompressedBytes,
		DEFAULT_MAX_TOTAL_UNCOMPRESSED_BYTES,
		"maxTotalUncompressedBytes",
	);
	const maxEntryUncompressedBytes = optionLimit(
		opts?.maxEntryUncompressedBytes,
		DEFAULT_MAX_ENTRY_UNCOMPRESSED_BYTES,
		"maxEntryUncompressedBytes",
	);

	// Scan backward for EOCD signature; EOCD is at least 22 bytes, max 65557 with comment
	let eocdOffset = -1;
	for (let i = data.length - 22; i >= 0 && i >= data.length - 65557; i--) {
		if (readU32(data, i) === SIG_EOCD) {
			const commentLen = readU16(data, i + 20);
			if (i + 22 + commentLen === data.length) {
				eocdOffset = i;
				break;
			}
		}
	}
	if (eocdOffset === -1) {
		throw new XlsxError("MALFORMED", "Invalid ZIP: EOCD not found");
	}

	assertRange(data, eocdOffset, 22, "EOCD");
	const eocdCommentLen = readU16At(data, eocdOffset + 20, "EOCD comment length");
	assertRange(data, eocdOffset, 22 + eocdCommentLen, "EOCD comment");

	const diskNumber = readU16At(data, eocdOffset + 4, "EOCD disk number");
	const cdDiskNumber = readU16At(data, eocdOffset + 6, "EOCD central directory disk number");
	const cdEntriesOnDisk = readU16At(data, eocdOffset + 8, "EOCD disk entry count");
	const cdEntries = readU16At(data, eocdOffset + 10, "EOCD entry count");
	const cdSize = readU32At(data, eocdOffset + 12, "EOCD central directory size");
	const cdOffset = readU32At(data, eocdOffset + 16, "EOCD central directory offset");
	if (cdEntriesOnDisk === ZIP64_U16 || cdEntries === ZIP64_U16 || cdSize === ZIP64_U32 || cdOffset === ZIP64_U32) {
		rejectZip64();
	}
	if (diskNumber !== 0 || cdDiskNumber !== 0 || cdEntriesOnDisk !== cdEntries) {
		throw new XlsxError("UNSUPPORTED", "Unsupported ZIP: multi-disk archives are not supported");
	}
	if (cdEntries > maxZipEntries) {
		throw new XlsxError("LIMIT_EXCEEDED", `Invalid ZIP: entry count ${cdEntries} exceeds limit ${maxZipEntries}`);
	}
	assertRange(data, cdOffset, cdSize, "central directory");
	if (cdOffset + cdSize > eocdOffset) {
		throw new XlsxError("MALFORMED", "Invalid ZIP: central directory overlaps EOCD");
	}

	// Parse central directory entries
	interface CdEntry {
		method: number;
		crc: number;
		compSize: number;
		uncompSize: number;
		name: string;
		localOffset: number;
	}
	const entries: CdEntry[] = [];
	let pos = cdOffset;
	const cdEnd = cdOffset + cdSize;
	for (let i = 0; i < cdEntries; i++) {
		assertRange(data, pos, 46, "central directory entry");
		if (pos + 46 > cdEnd) {
			throw new XlsxError("MALFORMED", "Invalid ZIP: central directory entry exceeds declared size");
		}
		if (readU32At(data, pos, "central directory signature") !== SIG_CENTRAL) {
			throw new XlsxError("MALFORMED", "Invalid ZIP: bad central directory entry");
		}
		const method = readU16At(data, pos + 10, "central directory compression method");
		const crcVal = readU32At(data, pos + 16, "central directory CRC");
		const compSize = readU32At(data, pos + 20, "central directory compressed size");
		const uncompSize = readU32At(data, pos + 24, "central directory uncompressed size");
		const nameLen = readU16At(data, pos + 28, "central directory file name length");
		const extraLen = readU16At(data, pos + 30, "central directory extra length");
		const commentLen = readU16At(data, pos + 32, "central directory comment length");
		const diskStart = readU16At(data, pos + 34, "central directory disk start");
		const localOffset = readU32At(data, pos + 42, "central directory local header offset");
		if (
			compSize === ZIP64_U32 ||
			uncompSize === ZIP64_U32 ||
			localOffset === ZIP64_U32 ||
			diskStart === ZIP64_U16
		) {
			rejectZip64();
		}
		if (diskStart !== 0) {
			throw new XlsxError("UNSUPPORTED", "Unsupported ZIP: multi-disk archives are not supported");
		}
		const entryEnd = pos + 46 + nameLen + extraLen + commentLen;
		if (entryEnd > cdEnd) {
			throw new XlsxError("MALFORMED", "Invalid ZIP: central directory entry exceeds declared size");
		}
		const nameBytes = data.subarray(pos + 46, pos + 46 + nameLen);
		const name = decoder.decode(nameBytes);
		entries.push({ method, crc: crcVal, compSize, uncompSize, name, localOffset });
		pos = entryEnd;
	}

	// Read file data from local headers -- collect deflated entries for bounded decompression
	const files: Record<string, Uint8Array> = Object.create(null);
	const seenNames = new Set<string>();
	const inflateJobs: { name: string; compressed: Uint8Array; expectedSize: number; crc: number }[] = [];
	let totalUncompressedBytes = 0;

	for (const entry of entries) {
		const name = entry.name;

		// Skip directory entries (paths ending with "/")
		if (name.endsWith("/")) {
			continue;
		}
		if (seenNames.has(name)) {
			throw new XlsxError("DUPLICATE", `Invalid ZIP: duplicate entry ${name}`);
		}
		seenNames.add(name);

		const loc = entry.localOffset;
		assertRange(data, loc, 30, `local file header for ${entry.name}`);
		if (readU32At(data, loc, "local file header signature") !== SIG_LOCAL) {
			throw new XlsxError("MALFORMED", "Invalid ZIP: bad local file header");
		}
		const localMethod = readU16At(data, loc + 8, "local file header compression method");
		const localNameLen = readU16At(data, loc + 26, "local file header file name length");
		const localExtraLen = readU16At(data, loc + 28, "local file header extra length");
		const dataStart = loc + 30 + localNameLen + localExtraLen;
		assertRange(data, dataStart, entry.compSize, `file data for ${entry.name}`);
		const dataEnd = dataStart + entry.compSize;
		if (localMethod !== entry.method) {
			throw new XlsxError("MALFORMED", `Invalid ZIP: local header method mismatch for ${entry.name}`);
		}
		const localName = decoder.decode(data.subarray(loc + 30, loc + 30 + localNameLen));
		if (localName !== name) {
			throw new XlsxError("MALFORMED", `Invalid ZIP: local header file name mismatch for ${name}`);
		}
		if (entry.uncompSize > maxEntryUncompressedBytes) {
			throw new XlsxError(
				"LIMIT_EXCEEDED",
				`Invalid ZIP: entry ${name} uncompressed size ${entry.uncompSize} exceeds limit ${maxEntryUncompressedBytes}`,
			);
		}
		totalUncompressedBytes += entry.uncompSize;
		if (totalUncompressedBytes > maxTotalUncompressedBytes) {
			throw new XlsxError(
				"LIMIT_EXCEEDED",
				`Invalid ZIP: total uncompressed size ${totalUncompressedBytes} exceeds limit ${maxTotalUncompressedBytes}`,
			);
		}

		if (entry.method === 0) {
			if (entry.compSize !== entry.uncompSize) {
				throw new XlsxError(
					"MALFORMED",
					`Invalid ZIP: stored entry ${name} has mismatched compressed and uncompressed sizes`,
				);
			}
			// Method 0: Stored (no compression)
			const fileData = data.subarray(dataStart, dataEnd);
			assertCrc(name, fileData, entry.crc);
			files[name] = fileData;
		} else if (entry.method === 8) {
			// Method 8: Deflated -- decompress after all structural checks pass
			inflateJobs.push({
				name,
				compressed: data.subarray(dataStart, dataEnd),
				expectedSize: entry.uncompSize,
				crc: entry.crc,
			});
		} else {
			throw new XlsxError("UNSUPPORTED", `Unsupported ZIP compression method: ${entry.method}`);
		}
	}

	// Decompress deflated entries sequentially so a failing archive cannot keep
	// multiple max-sized decompressed buffers alive after the first rejection.
	for (const job of inflateJobs) {
		const result = await inflate(job.compressed, job.expectedSize);
		if (result.length !== job.expectedSize) {
			throw new XlsxError(
				"MALFORMED",
				`Invalid ZIP: entry ${job.name} inflated to ${result.length} bytes, expected ${job.expectedSize}`,
			);
		}
		assertCrc(job.name, result, job.crc);
		files[job.name] = result;
	}

	return { files };
}

/**
 * Serialize a {@link ZipArchive} to raw ZIP file bytes.
 *
 * Writes local file headers, central directory, and EOCD record.
 * When {@link compress} is `true`, entries are deflated in parallel using {@link CompressionStream}.
 *
 * @param archive - Archive to serialize
 * @param compress - Whether to deflate file entries (default: stored uncompressed)
 * @returns Raw ZIP file bytes
 */
export async function zipWrite(archive: ZipArchive, compress?: boolean): Promise<Uint8Array> {
	const names = Object.keys(archive.files);
	const rawEntries = names.map((name) => ({
		name,
		nameBytes: encoder.encode(name),
		data: archive.files[name],
		crc: crc32(archive.files[name]),
	}));

	// Compress all entries in parallel if requested
	let compressedDatas: Uint8Array[];
	let methods: number[];
	if (compress) {
		const results = await Promise.all(rawEntries.map((e) => deflate(e.data)));
		compressedDatas = results;
		methods = rawEntries.map(() => 8); // method 8 = deflate
	} else {
		compressedDatas = rawEntries.map((e) => e.data);
		methods = rawEntries.map(() => 0); // method 0 = stored
	}

	// Pre-calculate total buffer size: local headers + central directory + EOCD
	let totalSize = 0;
	for (let i = 0; i < rawEntries.length; i++) {
		totalSize += 30 + rawEntries[i].nameBytes.length + compressedDatas[i].length; // local header + data
		totalSize += 46 + rawEntries[i].nameBytes.length; // central directory entry
	}
	totalSize += 22; // EOCD record (fixed 22 bytes with no comment)

	const buf = new Uint8Array(totalSize);
	let offset = 0;
	const centralEntries: { offset: number; index: number }[] = [];

	// Write local file headers + file data
	for (let i = 0; i < rawEntries.length; i++) {
		const entry = rawEntries[i];
		const compData = compressedDatas[i];
		centralEntries.push({ offset, index: i });

		writeU32(buf, offset, SIG_LOCAL);
		writeU16(buf, offset + 4, 20); // version needed to extract (2.0)
		writeU16(buf, offset + 6, 0); // general purpose bit flags
		writeU16(buf, offset + 8, methods[i]); // compression method
		writeU16(buf, offset + 10, 0); // last mod file time
		writeU16(buf, offset + 12, 0); // last mod file date
		writeU32(buf, offset + 14, entry.crc); // CRC-32
		writeU32(buf, offset + 18, compData.length); // compressed size
		writeU32(buf, offset + 22, entry.data.length); // uncompressed size
		writeU16(buf, offset + 26, entry.nameBytes.length); // file name length
		writeU16(buf, offset + 28, 0); // extra field length
		buf.set(entry.nameBytes, offset + 30);
		buf.set(compData, offset + 30 + entry.nameBytes.length);
		offset += 30 + entry.nameBytes.length + compData.length;
	}

	// Write central directory headers
	const cdStart = offset;
	for (const ce of centralEntries) {
		const entry = rawEntries[ce.index];
		const compData = compressedDatas[ce.index];

		writeU32(buf, offset, SIG_CENTRAL);
		writeU16(buf, offset + 4, 20); // version made by (2.0)
		writeU16(buf, offset + 6, 20); // version needed to extract (2.0)
		writeU16(buf, offset + 8, 0); // general purpose bit flags
		writeU16(buf, offset + 10, methods[ce.index]); // compression method
		writeU16(buf, offset + 12, 0); // last mod file time
		writeU16(buf, offset + 14, 0); // last mod file date
		writeU32(buf, offset + 16, entry.crc); // CRC-32
		writeU32(buf, offset + 20, compData.length); // compressed size
		writeU32(buf, offset + 24, entry.data.length); // uncompressed size
		writeU16(buf, offset + 28, entry.nameBytes.length); // file name length
		writeU16(buf, offset + 30, 0); // extra field length
		writeU16(buf, offset + 32, 0); // file comment length
		writeU16(buf, offset + 34, 0); // disk number start
		writeU16(buf, offset + 36, 0); // internal file attributes
		writeU32(buf, offset + 38, 0); // external file attributes
		writeU32(buf, offset + 42, ce.offset); // relative offset of local header
		buf.set(entry.nameBytes, offset + 46);
		offset += 46 + entry.nameBytes.length;
	}

	// Write End of Central Directory record
	const cdSize = offset - cdStart;
	writeU32(buf, offset, SIG_EOCD);
	writeU16(buf, offset + 4, 0); // number of this disk
	writeU16(buf, offset + 6, 0); // disk where CD starts
	writeU16(buf, offset + 8, rawEntries.length); // number of CD entries on this disk
	writeU16(buf, offset + 10, rawEntries.length); // total number of CD entries
	writeU32(buf, offset + 12, cdSize); // size of central directory
	writeU32(buf, offset + 16, cdStart); // offset of start of CD
	writeU16(buf, offset + 20, 0); // ZIP file comment length

	return buf;
}

/**
 * Read a file from a ZIP archive as a UTF-8 string.
 *
 * Falls back to trying with/without a leading slash if the exact path is not found.
 *
 * @param archive - ZIP archive to read from
 * @param path - File path within the archive
 * @returns Decoded string, or `null` if the file is not found
 */
export function zipReadString(archive: ZipArchive, path: string): string | null {
	let data = archive.files[path];
	if (!data) {
		// Try with or without leading slash as fallback
		const normalized = path.startsWith("/") ? path.slice(1) : "/" + path;
		data = archive.files[normalized];
	}
	if (!data) {
		return null;
	}
	return decoder.decode(data);
}

/**
 * Add a UTF-8 string as a file entry in the archive.
 *
 * @param archive - Target archive
 * @param path - File path within the archive
 * @param content - String content to encode as UTF-8
 */
export function zipAddString(archive: ZipArchive, path: string, content: string): void {
	archive.files[path] = encoder.encode(content);
}

/** Create a new empty ZIP archive with no file entries. */
export function zipCreate(): ZipArchive {
	return { files: {} };
}

/**
 * Check if a file exists in the archive.
 *
 * Tries an exact match first, then with/without a leading slash,
 * and finally a case-insensitive search as a last resort.
 *
 * @param archive - ZIP archive to search
 * @param path - File path to look for
 * @returns `true` if the file exists
 */
export function zipHas(archive: ZipArchive, path: string): boolean {
	if (archive.files[path]) {
		return true;
	}
	const normalized = path.startsWith("/") ? path.slice(1) : "/" + path;
	if (archive.files[normalized]) {
		return true;
	}
	// Case-insensitive fallback for interoperability with different ZIP tools
	const lpath = path.toLowerCase();
	for (const k of Object.keys(archive.files)) {
		if (k.toLowerCase() === lpath) {
			return true;
		}
	}
	return false;
}
