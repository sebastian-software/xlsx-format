import {
	read,
	write,
	type ReadOptions,
	type ReadResult,
	type WorkBook,
	type WriteOptions,
	type WriteResult,
} from "../index.js";

declare const bytes: Uint8Array;
declare const workbook: WorkBook;
declare const dynamicReadOptions: ReadOptions;
declare const optionalReadOptions: ReadOptions | undefined;
declare const dynamicWriteOptions: WriteOptions;
declare const optionalWriteOptions: WriteOptions | undefined;
declare const dynamicBookSheets: boolean;
declare const dynamicBookProps: boolean;

async function verifyReadContracts(): Promise<void> {
	const full = await read(bytes);
	const fullSheets = full.Sheets;

	const sheetNames = await read(bytes, { bookSheets: true });
	// @ts-expect-error Sheet-only reads intentionally do not expose worksheet data.
	const invalidSheetData = sheetNames.Sheets;

	const properties = await read(bytes, { bookProps: true });
	// @ts-expect-error Properties-only reads intentionally do not expose sheet names.
	const invalidSheetNames = properties.SheetNames;

	const combined = await read(bytes, { bookSheets: true, bookProps: true });
	// @ts-expect-error Combined metadata reads intentionally do not expose worksheet data.
	const invalidCombinedSheets = combined.Sheets;

	const dynamic = await read(bytes, dynamicReadOptions);
	const readResult: ReadResult = dynamic;
	// @ts-expect-error Dynamic options do not guarantee that worksheet data was parsed.
	const invalidDynamicSheets = dynamic.Sheets;

	const partlyDynamic = await read(bytes, { bookSheets: true, bookProps: dynamicBookProps });
	// @ts-expect-error An independently dynamic flag does not guarantee one metadata shape.
	const invalidDynamicProps = partlyDynamic.Props;
	const dynamicSheets = await read(bytes, { bookSheets: dynamicBookSheets });
	// @ts-expect-error A dynamic sheet-only flag does not guarantee worksheet data.
	const invalidMaybeSheets = dynamicSheets.Sheets;
	const optionalDynamic: ReadResult = await read(bytes, optionalReadOptions);

	void [
		fullSheets,
		sheetNames.SheetNames,
		properties.Props,
		combined.SheetNames,
		combined.Props,
		readResult,
		invalidSheetData,
		invalidSheetNames,
		invalidCombinedSheets,
		invalidDynamicSheets,
		invalidDynamicProps,
		invalidMaybeSheets,
		optionalDynamic,
	];
}

async function verifyWriteContracts(): Promise<void> {
	const defaultBytes: Uint8Array = await write(workbook);
	const xlsxStringModeBytes: Uint8Array = await write(workbook, { type: "string" });
	const text: string = await write(workbook, { bookType: "csv", type: "string" });
	const defaultText: string = await write(workbook, { bookType: "html" });
	const base64: string = await write(workbook, { type: "base64" });
	const array: Uint8Array = await write(workbook, { type: "array" });
	const portableBuffer: Uint8Array = await write(workbook, { type: "buffer" });
	const dynamic = await write(workbook, dynamicWriteOptions);
	const writeResult: WriteResult = dynamic;
	const optionalDynamic: WriteResult = await write(workbook, optionalWriteOptions);

	// @ts-expect-error XLSX string mode preserves the runtime byte result.
	const invalidXlsxString: string = await write(workbook, { type: "string" });
	// @ts-expect-error Broad write options can produce text or bytes.
	const invalidDynamicBytes: Uint8Array = await write(workbook, dynamicWriteOptions);
	// @ts-expect-error Buffer output is typed portably because browsers return Uint8Array.
	const invalidPortableBuffer: Buffer = await write(workbook, { type: "buffer" });

	void [
		defaultBytes,
		xlsxStringModeBytes,
		text,
		defaultText,
		base64,
		array,
		portableBuffer,
		writeResult,
		optionalDynamic,
		invalidXlsxString,
		invalidDynamicBytes,
		invalidPortableBuffer,
	];
}

void [verifyReadContracts, verifyWriteContracts];
