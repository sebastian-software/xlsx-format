/** Number Format (either a format string like "#,##0.00" or an index into the format table) */
export type NumberFormat = string | number;

/** Basic file properties from the OPC Core Properties part */
export interface Properties {
	/** Document title */
	Title?: string;
	/** Document subject/description */
	Subject?: string;
	/** Primary author */
	Author?: string;
	/** Manager name */
	Manager?: string;
	/** Company or organization */
	Company?: string;
	/** Category for grouping */
	Category?: string;
	/** Keywords / tags for search */
	Keywords?: string;
	/** Free-form comments or description */
	Comments?: string;
	/** Most recent editor */
	LastAuthor?: string;
	/** Date the document was created */
	CreatedDate?: Date;
}

/** Extended file properties (combines Core Properties with App-specific metadata) */
export interface FullProperties extends Properties {
	/** Date the document was last modified */
	ModifiedDate?: Date;
	/** Application that created the document (e.g. "Microsoft Excel") */
	Application?: string;
	/** Version of the creating application */
	AppVersion?: string;
	/** Document security level (as a string code) */
	DocSecurity?: string;
	/** Whether hyperlinks were changed outside the document */
	HyperlinksChanged?: boolean;
	/** Whether this is a shared document */
	SharedDoc?: boolean;
	/** Whether links are up to date */
	LinksUpToDate?: boolean;
	/** Whether the thumbnail should be cropped to fit */
	ScaleCrop?: boolean;
	/** Number of worksheets in the workbook */
	Worksheets?: number;
	/** List of worksheet names */
	SheetNames?: string[];
	/** Content status (e.g. "Draft", "Final") */
	ContentStatus?: string;
	/** Date the document was last printed */
	LastPrinted?: string;
	/** Revision number */
	Revision?: string | number;
	/** Version string */
	Version?: string;
	/** Unique document identifier */
	Identifier?: string;
	/** Document language (e.g. "en-US") */
	Language?: string;
}

/** Options common to both reading and writing operations */
export interface CommonOptions {
	/** If true, throw errors on unexpected situations instead of silently recovering */
	WTF?: boolean;
	/** If true, expose VBA macro data in workbook.vbaraw */
	bookVBA?: boolean;
	/** If true, store dates as Date objects instead of serial numbers */
	cellDates?: boolean;
	/** If true, include stub cells for empty cells within the used range */
	sheetStubs?: boolean;
	/** If true, include style/theme information on cells */
	cellStyles?: boolean;
	/** Workbook password for encrypted files */
	password?: string;
}

/** Options for reading/parsing workbook files */
export interface ReadOptions extends CommonOptions {
	/** Input data type: "base64" for base64 string, "buffer" for Node Buffer, "array" for Uint8Array, "string" for plain text (CSV/HTML) */
	type?: "base64" | "buffer" | "array" | "string";
	/** If true, parse and store cell formulas */
	cellFormula?: boolean;
	/** If true, generate HTML representation of rich text */
	cellHTML?: boolean;
	/** If true, store the number format string on each cell */
	cellNF?: boolean;
	/** If true, generate formatted text for each cell */
	cellText?: boolean;
	/** Override date format string (replaces default "m/d/yy" for format 14) */
	dateNF?: string;
	/** Maximum number of rows to read per sheet (0 = all rows) */
	sheetRows?: number;
	/** If true, parse inter-sheet dependencies */
	bookDeps?: boolean;
	/** If true, expose raw ZIP file entries */
	bookFiles?: boolean;
	/** If true, only parse workbook properties (skip sheet data) */
	bookProps?: boolean;
	/** If true, only parse sheet names (skip sheet data) */
	bookSheets?: boolean;
	/** Restrict parsing to specific sheets by index or name */
	sheets?: number | string | Array<number | string>;
	/** If true, do not infer sheet dimensions from data */
	nodim?: boolean;
	/** If true, preserve _xlfn. prefixes on formula function names */
	xlfn?: boolean;
	/** If true, use dense (2D array) storage mode instead of sparse (object) mode */
	dense?: boolean;
	/** If true, all dates are interpreted as UTC (no timezone adjustment) */
	UTC?: boolean;
}

/** Options for writing/serializing workbook files */
export interface WriteOptions extends CommonOptions {
	/** Output data type: "base64" for base64 string, "buffer" for Node Buffer, "array" for Uint8Array, "string" for plain text */
	type?: "base64" | "buffer" | "array" | "string";
	/** Output file format (default: "xlsx") */
	bookType?: "xlsx" | "xlsm" | "csv" | "tsv" | "html";
	/** If true, generate a Shared Strings Table for string deduplication */
	bookSST?: boolean;
	/** If true, compress (deflate) ZIP entries */
	compression?: boolean;
	/** Custom theme XML string to embed */
	themeXLSX?: string;
	/** If true, skip error-checking in the output */
	ignoreEC?: boolean;
	/** File properties to embed in the output */
	Props?: Properties;
}

/**
 * Excel data type codes for cell values.
 * - "b": Boolean
 * - "n": Number
 * - "e": Error
 * - "s": String
 * - "d": Date
 * - "z": Empty/stub cell
 */
export type ExcelDataType = "b" | "n" | "e" | "s" | "d" | "z";

/** A single comment entry within a cell's comment thread */
export interface Comment {
	/** Author of the comment */
	a?: string;
	/** Comment text content */
	t: string;
	/** If true, the comment is a threaded reply (Excel 365+) */
	T?: boolean;
}

/** Array of comments attached to a cell, with optional visibility flag */
export interface Comments extends Array<Comment> {
	/** If true, the comment indicator is hidden */
	hidden?: boolean;
}

/** Hyperlink target and optional tooltip */
export interface Hyperlink {
	/** URL or cell reference target */
	Target: string;
	/** Hover tooltip text */
	Tooltip?: string;
}

/** Worksheet Cell Object containing value, format, formula, and metadata */
export interface CellObject {
	/** Raw cell value (string, number, boolean, or Date) */
	v?: string | number | boolean | Date;
	/** Formatted text representation of the cell value */
	w?: string;
	/** Cell data type code */
	t: ExcelDataType;
	/** Cell formula string (without leading "=") */
	f?: string;
	/** Range of a shared/array formula (e.g. "A1:B2") */
	F?: string;
	/** If true, the formula is a dynamic array formula */
	D?: boolean;
	/** Rich text / XML representation */
	r?: any;
	/** HTML rendering of the cell (when cellHTML option is enabled) */
	h?: string;
	/** Comments attached to this cell */
	c?: Comments;
	/** Number format string or index */
	z?: NumberFormat;
	/** Hyperlink on this cell */
	l?: Hyperlink;
	/** Style object (when cellStyles option is enabled) */
	s?: any;
	/** Raw XF (extended format) record data */
	XF?: { numFmtId?: number };
}

/** Zero-based cell address with column (c) and row (r) indices */
export interface CellAddress {
	/** Zero-based column index */
	c: number;
	/** Zero-based row index */
	r: number;
}

/** Range defined by start (s) and end (e) cell addresses */
export interface Range {
	/** Start (top-left) cell address */
	s: CellAddress;
	/** End (bottom-right) cell address */
	e: CellAddress;
}

/** Column properties for worksheet column metadata */
export interface ColInfo {
	/** If true, column is hidden */
	hidden?: boolean;
	/** Column width in "Max Digit Width" units (Excel internal) */
	width?: number;
	/** Column width in pixels */
	wpx?: number;
	/** Column width in characters */
	wch?: number;
	/** Outline / grouping level (0-7) */
	level?: number;
	/** Maximum Digit Width in pixels (used for width calculations) */
	MDW?: number;
}

/** Row properties for worksheet row metadata */
export interface RowInfo {
	/** If true, row is hidden */
	hidden?: boolean;
	/** Row height in pixels */
	hpx?: number;
	/** Row height in points */
	hpt?: number;
	/** Outline / grouping level (0-7) */
	level?: number;
}

/** Sheet protection settings controlling what users can do on a protected sheet */
export interface ProtectInfo {
	/** Password hash for sheet protection */
	password?: string;
	/** If true, users can select locked cells */
	selectLockedCells?: boolean;
	/** If true, users can select unlocked cells */
	selectUnlockedCells?: boolean;
	/** If true, users can format cells */
	formatCells?: boolean;
	/** If true, users can format columns */
	formatColumns?: boolean;
	/** If true, users can format rows */
	formatRows?: boolean;
	/** If true, users can insert columns */
	insertColumns?: boolean;
	/** If true, users can insert rows */
	insertRows?: boolean;
	/** If true, users can insert hyperlinks */
	insertHyperlinks?: boolean;
	/** If true, users can delete columns */
	deleteColumns?: boolean;
	/** If true, users can delete rows */
	deleteRows?: boolean;
	/** If true, users can sort */
	sort?: boolean;
	/** If true, users can use autofilter */
	autoFilter?: boolean;
	/** If true, users can use pivot tables */
	pivotTables?: boolean;
	/** If true, users can edit objects (charts, shapes, etc.) */
	objects?: boolean;
	/** If true, users can edit scenarios */
	scenarios?: boolean;
}

/** Page margin settings in inches */
export interface MarginInfo {
	/** Left margin */
	left?: number;
	/** Right margin */
	right?: number;
	/** Top margin */
	top?: number;
	/** Bottom margin */
	bottom?: number;
	/** Header margin (distance from top of page) */
	header?: number;
	/** Footer margin (distance from bottom of page) */
	footer?: number;
}

/** AutoFilter definition for a worksheet */
export interface AutoFilterInfo {
	/** Range reference for the autofilter area (e.g. "A1:D10") */
	ref: string;
}

/** Dense (2D array) storage for worksheet data: rows of columns of optional cells */
export type DenseSheetData = ((CellObject | undefined)[] | undefined)[];

/**
 * Base sheet object supporting both sparse and dense storage modes.
 *
 * In sparse mode, cells are stored as properties keyed by A1 references (e.g. sheet["A1"]).
 * In dense mode, cells are stored in the "!data" 2D array.
 */
export interface Sheet {
	/** Sparse cell storage: cells keyed by A1-style references */
	[cell: string]: any;
	/** Dense cell storage: 2D array indexed by [row][col] */
	"!data"?: DenseSheetData;
	/** Sheet type: "sheet" for worksheets, "chart" for chart sheets */
	"!type"?: "sheet" | "chart";
	/** Used range reference (e.g. "A1:D10") */
	"!ref"?: string;
	/** Page margin settings */
	"!margins"?: MarginInfo;
}

/** Worksheet object with column, row, merge, protection, and filter metadata */
export interface WorkSheet extends Sheet {
	/** Column properties array (index corresponds to column index) */
	"!cols"?: ColInfo[];
	/** Row properties array (index corresponds to row index) */
	"!rows"?: RowInfo[];
	/** Array of merged cell ranges */
	"!merges"?: Range[];
	/** Sheet protection settings */
	"!protect"?: ProtectInfo;
	/** AutoFilter definition */
	"!autofilter"?: AutoFilterInfo;
}

/** Properties of a single sheet within the workbook */
export interface SheetProps {
	/** Sheet tab name */
	name?: string;
	/** Visibility: 0 = visible, 1 = hidden, 2 = very hidden (only accessible via VBA) */
	Hidden?: 0 | 1 | 2;
	/** VBA codename for the sheet module */
	CodeName?: string;
}

/** Defined Name (named range or named formula) */
export interface DefinedName {
	/** Name identifier (e.g. "MyRange") */
	Name: string;
	/** Reference formula (e.g. "Sheet1!$A$1:$B$10") */
	Ref: string;
	/** Sheet index this name is scoped to (undefined = workbook-scoped) */
	Sheet?: number;
	/** Descriptive comment */
	Comment?: string;
	/** If true, name is hidden from the UI */
	Hidden?: boolean;
}

/** Workbook view settings */
export interface WBView {
	/** If true, the workbook uses right-to-left layout */
	RTL?: boolean;
}

/** Workbook-level calculation and date system properties */
export interface WorkbookProperties {
	/** If true, use the 1904 date system (common in Mac Excel). Default is 1900 system. */
	date1904?: boolean;
	/** If true, personal information is stripped on save */
	filterPrivacy?: boolean;
	/** VBA codename for the workbook module */
	CodeName?: string;
}

/** Workbook-level attributes (sheets, names, views, properties) */
export interface WBProps {
	/** Sheet metadata array */
	Sheets?: SheetProps[];
	/** Defined names (named ranges, named formulas) */
	Names?: DefinedName[];
	/** Workbook view configurations */
	Views?: WBView[];
	/** Workbook-level properties (date system, etc.) */
	WBProps?: WorkbookProperties;
}

/** Top-level Workbook object containing all sheets, properties, and metadata */
export interface WorkBook {
	/** Map of sheet names to WorkSheet objects */
	Sheets: { [sheet: string]: WorkSheet };
	/** Ordered list of sheet names (determines tab order) */
	SheetNames: string[];
	/** File and document properties */
	Props?: FullProperties;
	/** Custom document properties (arbitrary key-value pairs) */
	Custprops?: Record<string, any>;
	/** Workbook-level attributes (names, views, sheet props) */
	Workbook?: WBProps;
	/** Raw VBA project binary (when bookVBA option is enabled) */
	vbaraw?: any;
	/** File format type identifier */
	bookType?: string;
}

/** Options for converting a worksheet to CSV */
export interface Sheet2CSVOpts {
	/** Field separator (default: ",") */
	FS?: string;
	/** Record separator / row delimiter (default: "\n") */
	RS?: string;
	/** If true, strip trailing field separators from each row */
	strip?: boolean;
	/** If true, include blank rows (default: true) */
	blankrows?: boolean;
	/** If true, skip hidden rows and columns */
	skipHidden?: boolean;
	/** If true, wrap all fields in quotes */
	forceQuotes?: boolean;
	/** If true, emit raw numeric values instead of formatted text */
	rawNumbers?: boolean;
	/** Override date format for date cells */
	dateNF?: NumberFormat;
}

/** Options for converting a worksheet to an HTML table string */
export interface Sheet2HTMLOpts {
	/** HTML id attribute for the table element */
	id?: string;
	/** If true, add contenteditable attribute to cells */
	editable?: boolean;
	/** HTML to prepend before the table */
	header?: string;
	/** HTML to append after the table */
	footer?: string;
	/** If true, sanitize hyperlink targets to prevent XSS */
	sanitizeLinks?: boolean;
}

/** Options for converting a worksheet to an array of JSON objects */
export interface Sheet2JSONOpts {
	/** "A" for column-letter keys, number for 1-indexed row keys, string[] for custom headers */
	header?: "A" | number | string[];
	/** Restrict output to a specific range (Range object, A1 string, or row number) */
	range?: any;
	/** If true, include blank rows in output */
	blankrows?: boolean;
	/** Default value for missing cells */
	defval?: any;
	/** If true, use raw values (v) instead of formatted text (w) */
	raw?: boolean;
	/** If true, skip hidden rows and columns */
	skipHidden?: boolean;
	/** If true, emit raw numeric values for non-date number cells */
	rawNumbers?: boolean;
	/** If true, interpret dates as UTC */
	UTC?: boolean;
	/** Override date format for date cells */
	dateNF?: NumberFormat;
}

/** Options for creating a worksheet from a 2D array (Array of Arrays) */
export interface AOA2SheetOpts extends CommonOptions {
	/** If true, use dense (2D array) storage mode */
	dense?: boolean;
	/** If true, include stub cells for null/undefined values */
	sheetStubs?: boolean;
	/** Date format string for date cells */
	dateNF?: NumberFormat;
	/** If true, store dates as Date objects instead of serial numbers */
	cellDates?: boolean;
	/** If true, interpret dates as UTC */
	UTC?: boolean;
	/** If true, use 1904 date system for date serial numbers */
	date1904?: boolean;
	/** Starting cell for data: row number, A1 reference, or CellAddress */
	origin?: number | string | CellAddress;
	/** If true, convert null values to #NULL! error cells */
	nullError?: boolean;
}

/** Options for creating a worksheet from an array of JSON objects */
export interface JSON2SheetOpts extends CommonOptions {
	/** Explicit header row keys (overrides object key order) */
	header?: string[];
	/** If true, do not emit a header row */
	skipHeader?: boolean;
	/** If true, use dense (2D array) storage mode */
	dense?: boolean;
	/** Date format string for date cells */
	dateNF?: NumberFormat;
	/** If true, store dates as Date objects instead of serial numbers */
	cellDates?: boolean;
	/** If true, interpret dates as UTC */
	UTC?: boolean;
	/** If true, use 1904 date system for date serial numbers */
	date1904?: boolean;
	/** Starting cell for data: row number, A1 reference, or CellAddress */
	origin?: number | string | CellAddress;
	/** If true, convert null values to #NULL! error cells */
	nullError?: boolean;
}

/**
 * Map of Excel error codes to their display strings.
 * Keys are the numeric error codes stored in XLSX files.
 */
export const BErr: Record<number, string> = {
	0x00: "#NULL!",
	0x07: "#DIV/0!",
	0x0f: "#VALUE!",
	0x17: "#REF!",
	0x1d: "#NAME?",
	0x24: "#NUM!",
	0x2a: "#N/A",
	0x2b: "#GETTING_DATA",
};

/** Reverse map from error display strings to numeric error codes */
export const RBErr: Record<string, number> = {};
for (const [k, v] of Object.entries(BErr)) {
	RBErr[v] = +k;
}
