import { unzipSync, zipSync, type Zippable } from "fflate";

export interface ZipEntry {
	data: Uint8Array;
	name: string;
}

export interface ZipArchive {
	files: Record<string, Uint8Array>;
}

const encoder = new TextEncoder();
const decoder = new TextDecoder();

/** Read a ZIP archive from a Uint8Array */
export function zip_read(data: Uint8Array): ZipArchive {
	const unzipped = unzipSync(data);
	return { files: unzipped as Record<string, Uint8Array> };
}

/** Write a ZIP archive to a Uint8Array */
export function zip_write(archive: ZipArchive, compress?: boolean): Uint8Array {
	const zippable: Zippable = {};
	for (const [name, data] of Object.entries(archive.files)) {
		zippable[name] = compress ? data : [data, { level: 0 }];
	}
	return zipSync(zippable);
}

/** Get a file from a ZIP archive as string */
export function zip_read_str(archive: ZipArchive, path: string): string | null {
	// Try exact path first
	let data = archive.files[path];
	if (!data) {
		// Try without leading slash
		const normalized = path.startsWith("/") ? path.slice(1) : "/" + path;
		data = archive.files[normalized];
	}
	if (!data) {
		return null;
	}
	return decoder.decode(data);
}

/** Get a file from a ZIP archive as Uint8Array */
export function zip_read_bin(archive: ZipArchive, path: string): Uint8Array | null {
	let data = archive.files[path];
	if (!data) {
		const normalized = path.startsWith("/") ? path.slice(1) : "/" + path;
		data = archive.files[normalized];
	}
	return data ?? null;
}

/** Add a string file to a ZIP archive */
export function zip_add_str(archive: ZipArchive, path: string, content: string): void {
	archive.files[path] = encoder.encode(content);
}

/** Add a binary file to a ZIP archive */
export function zip_add_bin(archive: ZipArchive, path: string, data: Uint8Array): void {
	archive.files[path] = data;
}

/** Create a new empty ZIP archive */
export function zip_new(): ZipArchive {
	return { files: {} };
}

/** List all file paths in a ZIP archive */
export function zip_list(archive: ZipArchive): string[] {
	return Object.keys(archive.files);
}

/** Check if a file exists in the archive (case-insensitive fallback) */
export function zip_has(archive: ZipArchive, path: string): boolean {
	if (archive.files[path]) {
		return true;
	}
	const normalized = path.startsWith("/") ? path.slice(1) : "/" + path;
	if (archive.files[normalized]) {
		return true;
	}
	// Case-insensitive search
	const lpath = path.toLowerCase();
	for (const k of Object.keys(archive.files)) {
		if (k.toLowerCase() === lpath) {
			return true;
		}
	}
	return false;
}

/** Find a file path in the archive (case-insensitive) */
export function zip_find(archive: ZipArchive, path: string): string | null {
	if (archive.files[path]) {
		return path;
	}
	const normalized = path.startsWith("/") ? path.slice(1) : "/" + path;
	if (archive.files[normalized]) {
		return normalized;
	}
	const lpath = path.toLowerCase();
	for (const k of Object.keys(archive.files)) {
		if (k.toLowerCase() === lpath) {
			return k;
		}
	}
	return null;
}
