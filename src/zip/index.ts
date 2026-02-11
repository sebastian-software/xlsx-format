import { crc32 } from "./crc32.js";
import { inflate, deflate } from "./streams.js";

/** In-memory representation of a ZIP archive as a flat filename-to-data map. */
export interface ZipArchive {
	/** Map of file paths to their uncompressed contents. */
	files: Record<string, Uint8Array>;
}

const encoder = new TextEncoder();
const decoder = new TextDecoder();

// -- ZIP format signatures (little-endian magic numbers) --
/** Local file header signature: "PK\x03\x04" */
const SIG_LOCAL = 0x04034b50;
/** Central directory file header signature: "PK\x01\x02" */
const SIG_CENTRAL = 0x02014b50;
/** End of Central Directory record signature: "PK\x05\x06" */
const SIG_EOCD = 0x06054b50;

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
 * are decompressed in parallel using {@link DecompressionStream}.
 *
 * @param data - Raw ZIP file bytes
 * @returns Parsed archive with decompressed file contents
 * @throws Error if the ZIP structure is invalid or uses an unsupported compression method
 */
export async function zipRead(data: Uint8Array): Promise<ZipArchive> {
	// Scan backward for EOCD signature; EOCD is at least 22 bytes, max 65557 with comment
	let eocdOffset = -1;
	for (let i = data.length - 22; i >= 0 && i >= data.length - 65557; i--) {
		if (readU32(data, i) === SIG_EOCD) {
			eocdOffset = i;
			break;
		}
	}
	if (eocdOffset === -1) {
		throw new Error("Invalid ZIP: EOCD not found");
	}

	// EOCD layout: +8 = total entries on disk, +10 = total entries, +12 = CD size, +16 = CD offset
	const cdEntries = readU16(data, eocdOffset + 10);
	const _cdSize = readU32(data, eocdOffset + 12);
	const cdOffset = readU32(data, eocdOffset + 16);

	// Parse central directory entries
	interface CdEntry {
		method: number;
		crc: number;
		compSize: number;
		uncompSize: number;
		nameBytes: Uint8Array;
		localOffset: number;
	}
	const entries: CdEntry[] = [];
	let pos = cdOffset;
	for (let i = 0; i < cdEntries; i++) {
		if (readU32(data, pos) !== SIG_CENTRAL) {
			throw new Error("Invalid ZIP: bad central directory entry");
		}
		const method = readU16(data, pos + 10); // compression method (0=stored, 8=deflate)
		const crcVal = readU32(data, pos + 16);
		const compSize = readU32(data, pos + 20);
		const uncompSize = readU32(data, pos + 24);
		const nameLen = readU16(data, pos + 28);
		const extraLen = readU16(data, pos + 30);
		const commentLen = readU16(data, pos + 32);
		const localOffset = readU32(data, pos + 42); // offset to local file header
		const nameBytes = data.subarray(pos + 46, pos + 46 + nameLen);
		entries.push({ method, crc: crcVal, compSize, uncompSize, nameBytes, localOffset });
		pos += 46 + nameLen + extraLen + commentLen;
	}

	// Read file data from local headers -- collect inflate promises for parallel decompression
	const files: Record<string, Uint8Array> = {};
	const inflateJobs: { name: string; compressed: Uint8Array }[] = [];

	for (const entry of entries) {
		const loc = entry.localOffset;
		if (readU32(data, loc) !== SIG_LOCAL) {
			throw new Error("Invalid ZIP: bad local file header");
		}
		const localNameLen = readU16(data, loc + 26);
		const localExtraLen = readU16(data, loc + 28);
		const dataStart = loc + 30 + localNameLen + localExtraLen;
		const name = decoder.decode(entry.nameBytes);

		// Skip directory entries (paths ending with "/")
		if (name.endsWith("/")) {
			continue;
		}

		if (entry.method === 0) {
			// Method 0: Stored (no compression)
			files[name] = data.subarray(dataStart, dataStart + entry.uncompSize);
		} else if (entry.method === 8) {
			// Method 8: Deflated -- batch for parallel decompression
			inflateJobs.push({ name, compressed: data.subarray(dataStart, dataStart + entry.compSize) });
		} else {
			throw new Error(`Unsupported ZIP compression method: ${entry.method}`);
		}
	}

	// Decompress all deflated entries in parallel
	if (inflateJobs.length > 0) {
		const results = await Promise.all(inflateJobs.map((j) => inflate(j.compressed)));
		for (let i = 0; i < inflateJobs.length; i++) {
			files[inflateJobs[i].name] = results[i];
		}
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
