/**
 * Read all chunks from a {@link ReadableStream} and concatenate them into a single {@link Uint8Array}.
 */
async function collectStream(readable: ReadableStream<Uint8Array>): Promise<Uint8Array> {
	const reader = readable.getReader();
	const chunks: Uint8Array[] = [];
	let totalLength = 0;
	for (;;) {
		const { done, value } = await reader.read();
		if (done) break;
		chunks.push(value);
		totalLength += value.length;
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
 * @param data - Compressed bytes (raw deflate, no zlib/gzip wrapper)
 * @returns Decompressed bytes
 */
export async function inflate(data: Uint8Array): Promise<Uint8Array> {
	const ds = new DecompressionStream("deflate-raw");
	const writer = ds.writable.getWriter();
	void writer.write(data as Uint8Array<ArrayBuffer>);
	void writer.close();
	return collectStream(ds.readable);
}

/**
 * Compress data using raw DEFLATE via the built-in {@link CompressionStream} API.
 *
 * @param data - Uncompressed bytes
 * @returns Compressed bytes (raw deflate, no zlib/gzip wrapper)
 */
export async function deflate(data: Uint8Array): Promise<Uint8Array> {
	const cs = new CompressionStream("deflate-raw");
	const writer = cs.writable.getWriter();
	void writer.write(data as Uint8Array<ArrayBuffer>);
	void writer.close();
	return collectStream(cs.readable);
}
