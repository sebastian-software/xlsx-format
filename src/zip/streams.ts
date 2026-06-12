/**
 * Read all chunks from a {@link ReadableStream} and concatenate them into a single {@link Uint8Array}.
 *
 * Collects chunks incrementally, then copies them into a contiguous buffer.
 */
async function collectStream(readable: ReadableStream<Uint8Array>, maxBytes = Infinity): Promise<Uint8Array> {
	const reader = readable.getReader();
	const chunks: Uint8Array[] = [];
	let totalLength = 0;
	for (;;) {
		const { done, value } = await reader.read();
		if (done) {
			break;
		}
		chunks.push(value);
		totalLength += value.length;
		if (totalLength > maxBytes) {
			throw new Error(`Invalid ZIP: decompressed data exceeds limit (${maxBytes} bytes)`);
		}
	}
	const result = new Uint8Array(totalLength);
	let offset = 0;
	for (const chunk of chunks) {
		result.set(chunk, offset);
		offset += chunk.length;
	}
	return result;
}

/**
 * Decompress raw DEFLATE data using the built-in {@link DecompressionStream} API.
 *
 * Uses "deflate-raw" format (no zlib header or gzip wrapper), which matches
 * the compression used inside ZIP archives (method 8).
 *
 * @param data - Compressed bytes (raw deflate, no zlib/gzip wrapper)
 * @param maxBytes - Maximum allowed decompressed size
 * @returns Decompressed bytes
 */
export async function inflate(data: Uint8Array, maxBytes = Infinity): Promise<Uint8Array> {
	const ds = new DecompressionStream("deflate-raw");
	const writer = ds.writable.getWriter();
	const writePromise = writer.write(data as Uint8Array<ArrayBuffer>).then(() => writer.close());
	try {
		const [result] = await Promise.all([collectStream(ds.readable, maxBytes), writePromise]);
		return result;
	} catch (err) {
		try {
			await writer.abort(err);
		} catch {
			/* ignore abort errors while preserving the original stream error */
		}
		throw err;
	}
}

/**
 * Compress data using raw DEFLATE via the built-in {@link CompressionStream} API.
 *
 * Uses "deflate-raw" format (no zlib header or gzip wrapper), which matches
 * the compression expected inside ZIP archives (method 8).
 *
 * @param data - Uncompressed bytes
 * @returns Compressed bytes (raw deflate, no zlib/gzip wrapper)
 */
export async function deflate(data: Uint8Array): Promise<Uint8Array> {
	const cs = new CompressionStream("deflate-raw");
	const writer = cs.writable.getWriter();
	const writePromise = writer.write(data as Uint8Array<ArrayBuffer>).then(() => writer.close());
	try {
		const [result] = await Promise.all([collectStream(cs.readable), writePromise]);
		return result;
	} catch (err) {
		try {
			await writer.abort(err);
		} catch {
			/* ignore abort errors while preserving the original stream error */
		}
		throw err;
	}
}
