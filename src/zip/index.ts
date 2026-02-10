import { crc32 } from "./crc32.js";
import { inflate, deflate } from "./streams.js";

/** In-memory representation of a ZIP archive as a flat filename-to-data map. */
export interface ZipArchive {
	/** Map of file paths to their uncompressed contents. */
	files: Record<string, Uint8Array>;
}

const encoder = new TextEncoder();
const decoder = new TextDecoder();

// -- ZIP format constants --
const SIG_LOCAL = 0x04034b50;
const SIG_CENTRAL = 0x02014b50;
const SIG_EOCD = 0x06054b50;

function readU16(buf: Uint8Array, off: number): number {
	return buf[off] | (buf[off + 1] << 8);
}

function readU32(buf: Uint8Array, off: number): number {
	return (buf[off] | (buf[off + 1] << 8) | (buf[off + 2] << 16) | (buf[off + 3] << 24)) >>> 0;
}

function writeU16(buf: Uint8Array, off: number, val: number): void {
	buf[off] = val & 0xff;
	buf[off + 1] = (val >> 8) & 0xff;
}

function writeU32(buf: Uint8Array, off: number, val: number): void {
	buf[off] = val & 0xff;
	buf[off + 1] = (val >> 8) & 0xff;
	buf[off + 2] = (val >> 16) & 0xff;
	buf[off + 3] = (val >> 24) & 0xff;
}

/**
 * Parse a ZIP archive from raw bytes.
 *
 * Deflated entries are decompressed in parallel using {@link DecompressionStream}.
 *
 * @param data - Raw ZIP file bytes
 * @returns Parsed archive with decompressed file contents
 */
export async function zipRead(data: Uint8Array): Promise<ZipArchive> {
	// Find EOCD record by scanning backward for signature
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

	const cdEntries = readU16(data, eocdOffset + 10);
	const cdSize = readU32(data, eocdOffset + 12);
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
		const method = readU16(data, pos + 10);
		const crcVal = readU32(data, pos + 16);
		const compSize = readU32(data, pos + 20);
		const uncompSize = readU32(data, pos + 24);
		const nameLen = readU16(data, pos + 28);
		const extraLen = readU16(data, pos + 30);
		const commentLen = readU16(data, pos + 32);
		const localOffset = readU32(data, pos + 42);
		const nameBytes = data.subarray(pos + 46, pos + 46 + nameLen);
		entries.push({ method, crc: crcVal, compSize, uncompSize, nameBytes, localOffset });
		pos += 46 + nameLen + extraLen + commentLen;
	}

	// Read file data from local headers — collect inflate promises
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

		// Skip directories
		if (name.endsWith("/")) continue;

		if (entry.method === 0) {
			// Stored
			files[name] = data.subarray(dataStart, dataStart + entry.uncompSize);
		} else if (entry.method === 8) {
			// Deflated — batch for parallel inflate
			inflateJobs.push({ name, compressed: data.subarray(dataStart, dataStart + entry.compSize) });
		} else {
			throw new Error(`Unsupported ZIP compression method: ${entry.method}`);
		}
	}

	// Parallel decompression
	if (inflateJobs.length > 0) {
		const results = await Promise.all(inflateJobs.map((j) => inflate(j.compressed)));
		for (let i = 0; i < inflateJobs.length; i++) {
			files[inflateJobs[i].name] = results[i];
		}
	}

	return { files };
}

/**
 * Serialize a {@link ZipArchive} to a ZIP file as raw bytes.
 *
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

	// Compress in parallel if requested
	let compressedDatas: Uint8Array[];
	let methods: number[];
	if (compress) {
		const results = await Promise.all(rawEntries.map((e) => deflate(e.data)));
		compressedDatas = results;
		methods = rawEntries.map(() => 8);
	} else {
		compressedDatas = rawEntries.map((e) => e.data);
		methods = rawEntries.map(() => 0);
	}

	// Calculate total size
	let totalSize = 0;
	for (let i = 0; i < rawEntries.length; i++) {
		totalSize += 30 + rawEntries[i].nameBytes.length + compressedDatas[i].length; // local
		totalSize += 46 + rawEntries[i].nameBytes.length; // central
	}
	totalSize += 22; // EOCD

	const buf = new Uint8Array(totalSize);
	let offset = 0;
	const centralEntries: { offset: number; index: number }[] = [];

	// Write local file headers + data
	for (let i = 0; i < rawEntries.length; i++) {
		const entry = rawEntries[i];
		const compData = compressedDatas[i];
		centralEntries.push({ offset, index: i });

		writeU32(buf, offset, SIG_LOCAL);
		writeU16(buf, offset + 4, 20); // version needed
		writeU16(buf, offset + 6, 0); // flags
		writeU16(buf, offset + 8, methods[i]); // method
		writeU16(buf, offset + 10, 0); // mod time
		writeU16(buf, offset + 12, 0); // mod date
		writeU32(buf, offset + 14, entry.crc);
		writeU32(buf, offset + 18, compData.length); // compressed size
		writeU32(buf, offset + 22, entry.data.length); // uncompressed size
		writeU16(buf, offset + 26, entry.nameBytes.length);
		writeU16(buf, offset + 28, 0); // extra length
		buf.set(entry.nameBytes, offset + 30);
		buf.set(compData, offset + 30 + entry.nameBytes.length);
		offset += 30 + entry.nameBytes.length + compData.length;
	}

	// Write central directory
	const cdStart = offset;
	for (const ce of centralEntries) {
		const entry = rawEntries[ce.index];
		const compData = compressedDatas[ce.index];

		writeU32(buf, offset, SIG_CENTRAL);
		writeU16(buf, offset + 4, 20); // version made by
		writeU16(buf, offset + 6, 20); // version needed
		writeU16(buf, offset + 8, 0); // flags
		writeU16(buf, offset + 10, methods[ce.index]); // method
		writeU16(buf, offset + 12, 0); // mod time
		writeU16(buf, offset + 14, 0); // mod date
		writeU32(buf, offset + 16, entry.crc);
		writeU32(buf, offset + 20, compData.length); // compressed size
		writeU32(buf, offset + 24, entry.data.length); // uncompressed size
		writeU16(buf, offset + 28, entry.nameBytes.length);
		writeU16(buf, offset + 30, 0); // extra length
		writeU16(buf, offset + 32, 0); // comment length
		writeU16(buf, offset + 34, 0); // disk number start
		writeU16(buf, offset + 36, 0); // internal attributes
		writeU32(buf, offset + 38, 0); // external attributes
		writeU32(buf, offset + 42, ce.offset); // local header offset
		buf.set(entry.nameBytes, offset + 46);
		offset += 46 + entry.nameBytes.length;
	}

	// Write EOCD
	const cdSize = offset - cdStart;
	writeU32(buf, offset, SIG_EOCD);
	writeU16(buf, offset + 4, 0); // disk number
	writeU16(buf, offset + 6, 0); // disk with CD
	writeU16(buf, offset + 8, rawEntries.length); // entries on disk
	writeU16(buf, offset + 10, rawEntries.length); // total entries
	writeU32(buf, offset + 12, cdSize);
	writeU32(buf, offset + 16, cdStart);
	writeU16(buf, offset + 20, 0); // comment length

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
 * and finally a case-insensitive search.
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
	const lpath = path.toLowerCase();
	for (const k of Object.keys(archive.files)) {
		if (k.toLowerCase() === lpath) {
			return true;
		}
	}
	return false;
}
