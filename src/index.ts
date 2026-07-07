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
	StyleColor,
	CellFont,
	CellFill,
	CellBorderStyle,
	CellBorderSide,
	CellBorder,
	CellAlignment,
	CellStyle,
	ColInfo,
	RowInfo,
	SheetView,
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
export type { XlsxErrorCode } from "./errors.js";
export { XlsxError } from "./errors.js";

// Read / Write
export { read } from "./read.js";
export { write } from "./write.js";

// Utilities - workbook/sheet manipulation
export {
	createWorkbook,
	appendSheet,
	setSheetVisibility,
	createSheet,
	getSheetIndex,
	setCellNumberFormat,
	setCellStyle,
	styleRange,
	mergeCells,
	setRowHeight,
	setColumnWidth,
	freezePanes,
	setCellHyperlink,
	setCellInternalLink,
	addCellComment,
	setArrayFormula,
	sheetToFormulae,
} from "./api/book.js";

// Utilities - format conversions
export { arrayToSheet, sheetToArray, addArrayToSheet } from "./api/aoa.js";
export { jsonToSheet, sheetToJson, addJsonToSheet } from "./api/json.js";
export { sheetToCsv, sheetToTxt, csvToSheet } from "./api/csv.js";
export { sheetToHtml, htmlToSheet } from "./api/html.js";
export { formatCell } from "./api/format.js";

// Cell utilities
export {
	decodeCell,
	encodeCell,
	decodeRange,
	encodeRange,
	decodeCol,
	encodeCol,
	decodeRow,
	encodeRow,
} from "./utils/cell.js";

// SSF (Number Formatting)
export { formatNumber } from "./ssf/format.js";

// Version
export const version = "2.4.0"; // x-release-please-version
