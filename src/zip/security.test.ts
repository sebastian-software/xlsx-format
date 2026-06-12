import { describe, expect, it } from "vitest";
import { zipRead, zipWrite } from "./index.js";

const encoder = new TextEncoder();

function bytes(text: string): Uint8Array {
	return encoder.encode(text);
}

function readU16(data: Uint8Array, off: number): number {
	return data[off] | (data[off + 1] << 8);
}

function readU32(data: Uint8Array, off: number): number {
	return (data[off] | (data[off + 1] << 8) | (data[off + 2] << 16) | (data[off + 3] << 24)) >>> 0;
}

function writeU16(data: Uint8Array, off: number, value: number): void {
	data[off] = value & 0xff;
	data[off + 1] = (value >> 8) & 0xff;
}

function writeU32(data: Uint8Array, off: number, value: number): void {
	data[off] = value & 0xff;
	data[off + 1] = (value >> 8) & 0xff;
	data[off + 2] = (value >> 16) & 0xff;
	data[off + 3] = (value >> 24) & 0xff;
}

function findEocd(data: Uint8Array): number {
	for (let i = data.length - 22; i >= 0; --i) {
		if (readU32(data, i) === 0x06054b50) {
			return i;
		}
	}
	throw new Error("EOCD not found in test archive");
}

function centralDirectoryEntryOffsets(data: Uint8Array): number[] {
	const eocd = findEocd(data);
	const entryCount = readU16(data, eocd + 10);
	const offsets: number[] = [];
	let pos = readU32(data, eocd + 16);
	for (let i = 0; i < entryCount; ++i) {
		offsets.push(pos);
		const nameLen = readU16(data, pos + 28);
		const extraLen = readU16(data, pos + 30);
		const commentLen = readU16(data, pos + 32);
		pos += 46 + nameLen + extraLen + commentLen;
	}
	return offsets;
}

describe("zipRead security guards", () => {
	it("rejects archives without EOCD", async () => {
		await expect(zipRead(new Uint8Array([0x50, 0x4b]))).rejects.toThrow(/EOCD not found/);
	});

	it("rejects central directories outside archive bounds", async () => {
		const archive = await zipWrite({ files: { "a.txt": bytes("a") } });
		const corrupted = archive.slice();
		writeU32(corrupted, findEocd(corrupted) + 16, corrupted.length);

		await expect(zipRead(corrupted)).rejects.toThrow(/central directory out of bounds/);
	});

	it("rejects Zip64 markers explicitly", async () => {
		const archive = await zipWrite({ files: { "a.txt": bytes("a") } });
		const corrupted = archive.slice();
		writeU16(corrupted, findEocd(corrupted) + 10, 0xffff);

		await expect(zipRead(corrupted)).rejects.toThrow(/Zip64 archives are not supported/);
	});

	it("enforces entry and uncompressed-size budgets", async () => {
		const archive = await zipWrite({
			files: {
				"a.txt": bytes("abcd"),
				"b.txt": bytes("efgh"),
			},
		});

		await expect(zipRead(archive, { maxZipEntries: 1 })).rejects.toThrow(/entry count 2 exceeds limit 1/);
		await expect(zipRead(archive, { maxEntryUncompressedBytes: 3 })).rejects.toThrow(
			/a.txt uncompressed size 4 exceeds limit 3/,
		);
		await expect(zipRead(archive, { maxTotalUncompressedBytes: 7 })).rejects.toThrow(
			/total uncompressed size 8 exceeds limit 7/,
		);
	});

	it("rejects invalid budget options", async () => {
		const archive = await zipWrite({ files: { "a.txt": bytes("a") } });

		await expect(zipRead(archive, { maxZipEntries: -1 })).rejects.toThrow(/maxZipEntries/);
		await expect(zipRead(archive, { maxTotalUncompressedBytes: Number.POSITIVE_INFINITY })).rejects.toThrow(
			/maxTotalUncompressedBytes/,
		);
	});

	it("rejects multi-disk archives", async () => {
		const archive = await zipWrite({ files: { "a.txt": bytes("a") } });
		const corrupted = archive.slice();
		writeU16(corrupted, findEocd(corrupted) + 4, 1);

		await expect(zipRead(corrupted)).rejects.toThrow(/multi-disk archives are not supported/);
	});

	it("rejects duplicate central-directory names", async () => {
		const archive = await zipWrite({
			files: {
				"a.txt": bytes("a"),
				"b.txt": bytes("b"),
			},
		});
		const corrupted = archive.slice();
		const [, secondEntry] = centralDirectoryEntryOffsets(corrupted);
		corrupted.set(bytes("a.txt"), secondEntry + 46);

		await expect(zipRead(corrupted)).rejects.toThrow(/duplicate entry a\.txt/);
	});

	it("rejects local-header file name mismatches", async () => {
		const archive = await zipWrite({ files: { "a.txt": bytes("a") } });
		const corrupted = archive.slice();
		const [entry] = centralDirectoryEntryOffsets(corrupted);
		corrupted.set(bytes("b.txt"), entry + 46);

		await expect(zipRead(corrupted)).rejects.toThrow(/local header file name mismatch for b\.txt/);
	});

	it("rejects local-header method mismatches", async () => {
		const archive = await zipWrite({ files: { "a.txt": bytes("a") } });
		const corrupted = archive.slice();
		const [entry] = centralDirectoryEntryOffsets(corrupted);
		const localOffset = readU32(corrupted, entry + 42);
		writeU16(corrupted, localOffset + 8, 8);

		await expect(zipRead(corrupted)).rejects.toThrow(/local header method mismatch for a\.txt/);
	});

	it("rejects unsupported compression methods", async () => {
		const archive = await zipWrite({ files: { "a.txt": bytes("a") } });
		const corrupted = archive.slice();
		const [entry] = centralDirectoryEntryOffsets(corrupted);
		const localOffset = readU32(corrupted, entry + 42);
		writeU16(corrupted, entry + 10, 99);
		writeU16(corrupted, localOffset + 8, 99);

		await expect(zipRead(corrupted)).rejects.toThrow(/Unsupported ZIP compression method: 99/);
	});

	it("rejects stored entries with mismatched sizes", async () => {
		const archive = await zipWrite({ files: { "a.txt": bytes("a") } });
		const corrupted = archive.slice();
		const [entry] = centralDirectoryEntryOffsets(corrupted);
		writeU32(corrupted, entry + 20, 0);

		await expect(zipRead(corrupted)).rejects.toThrow(/stored entry a\.txt has mismatched/);
	});

	it("rejects CRC mismatches", async () => {
		const archive = await zipWrite({ files: { "a.txt": bytes("a") } });
		const corrupted = archive.slice();
		const [entry] = centralDirectoryEntryOffsets(corrupted);
		writeU32(corrupted, entry + 16, 0);

		await expect(zipRead(corrupted)).rejects.toThrow(/CRC mismatch for a\.txt/);
	});

	it("caps deflate output at the declared uncompressed size", async () => {
		const archive = await zipWrite({ files: { "a.txt": bytes("hello") } }, true);
		const corrupted = archive.slice();
		const [entry] = centralDirectoryEntryOffsets(corrupted);
		writeU32(corrupted, entry + 24, 1);

		await expect(zipRead(corrupted)).rejects.toThrow(/decompressed data exceeds limit|inflated to/);
	});
});
