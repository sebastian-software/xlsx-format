// xlsx-format - Modern TypeScript XLSX parser and writer
// Public API

// Types
export type {
	WorkBook,
	WorkSheet,
	CellObject,
	CellAddress,
	Range,
	Properties,
	FullProperties,
	ReadOptions,
	WriteOptions,
	ExcelDataType,
	Comment,
	Comments,
	Hyperlink,
	ColInfo,
	RowInfo,
	MarginInfo,
	ProtectInfo,
	AutoFilterInfo,
	DenseSheetData,
	Sheet,
	SheetProps,
	DefinedName,
	WBView,
	WorkbookProperties,
	WBProps,
	Sheet2CSVOpts,
	Sheet2HTMLOpts,
	Sheet2JSONOpts,
	AOA2SheetOpts,
	JSON2SheetOpts,
	NumberFormat,
} from "./types.js";

// Read / Write
export { read, readFile } from "./read.js";
export { write, writeFile } from "./write.js";

// Utilities - workbook/sheet manipulation
export {
	book_new,
	book_append_sheet,
	book_set_sheet_visibility,
	sheet_new,
	wb_sheet_idx,
	cell_set_number_format,
	cell_set_hyperlink,
	cell_set_internal_link,
	cell_add_comment,
	sheet_set_array_formula,
	sheet_to_formulae,
} from "./api/book.js";

// Utilities - format conversions
export { aoa_to_sheet, sheet_add_aoa } from "./api/aoa.js";
export { json_to_sheet, sheet_to_json, sheet_add_json } from "./api/json.js";
export { sheet_to_csv, sheet_to_txt } from "./api/csv.js";
export { sheet_to_html } from "./api/html.js";
export { format_cell } from "./api/format.js";

// Cell utilities
export {
	decode_cell,
	encode_cell,
	decode_range,
	encode_range,
	decode_col,
	encode_col,
	decode_row,
	encode_row,
} from "./utils/cell.js";

// SSF (Number Formatting)
export { SSF_format } from "./ssf/format.js";

// Version
export const version = "1.0.0-alpha.0";
