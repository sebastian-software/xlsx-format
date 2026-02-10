/** Number Format (either a string or an index to the format table) */
export type NumberFormat = string | number;

/** Basic File Properties */
export interface Properties {
	Title?: string;
	Subject?: string;
	Author?: string;
	Manager?: string;
	Company?: string;
	Category?: string;
	Keywords?: string;
	Comments?: string;
	LastAuthor?: string;
	CreatedDate?: Date;
}

/** Extended File Properties */
export interface FullProperties extends Properties {
	ModifiedDate?: Date;
	Application?: string;
	AppVersion?: string;
	DocSecurity?: string;
	HyperlinksChanged?: boolean;
	SharedDoc?: boolean;
	LinksUpToDate?: boolean;
	ScaleCrop?: boolean;
	Worksheets?: number;
	SheetNames?: string[];
	ContentStatus?: string;
	LastPrinted?: string;
	Revision?: string | number;
	Version?: string;
	Identifier?: string;
	Language?: string;
}

export interface CommonOptions {
	WTF?: boolean;
	bookVBA?: boolean;
	cellDates?: boolean;
	sheetStubs?: boolean;
	cellStyles?: boolean;
	password?: string;
}

export interface ReadOptions extends CommonOptions {
	type?: "base64" | "buffer" | "array";
	cellFormula?: boolean;
	cellHTML?: boolean;
	cellNF?: boolean;
	cellText?: boolean;
	dateNF?: string;
	sheetRows?: number;
	bookDeps?: boolean;
	bookFiles?: boolean;
	bookProps?: boolean;
	bookSheets?: boolean;
	sheets?: number | string | Array<number | string>;
	nodim?: boolean;
	xlfn?: boolean;
	dense?: boolean;
	UTC?: boolean;
}

export interface WriteOptions extends CommonOptions {
	type?: "base64" | "buffer" | "array";
	bookSST?: boolean;
	compression?: boolean;
	themeXLSX?: string;
	ignoreEC?: boolean;
	Props?: Properties;
}

/** The Excel data type for a cell: b Boolean, n Number, e Error, s String, d Date, z Empty */
export type ExcelDataType = "b" | "n" | "e" | "s" | "d" | "z";

/** Comment element */
export interface Comment {
	a?: string;
	t: string;
	T?: boolean;
}

/** Cell comments */
export interface Comments extends Array<Comment> {
	hidden?: boolean;
}

/** Link object */
export interface Hyperlink {
	Target: string;
	Tooltip?: string;
}

/** Worksheet Cell Object */
export interface CellObject {
	v?: string | number | boolean | Date;
	w?: string;
	t: ExcelDataType;
	f?: string;
	F?: string;
	D?: boolean;
	r?: any;
	h?: string;
	c?: Comments;
	z?: NumberFormat;
	l?: Hyperlink;
	s?: any;
	XF?: { numFmtId?: number };
}

/** Simple Cell Address */
export interface CellAddress {
	c: number;
	r: number;
}

/** Range object (representing ranges like "A1:B2") */
export interface Range {
	s: CellAddress;
	e: CellAddress;
}

/** Column Properties Object */
export interface ColInfo {
	hidden?: boolean;
	width?: number;
	wpx?: number;
	wch?: number;
	level?: number;
	MDW?: number;
}

/** Row Properties Object */
export interface RowInfo {
	hidden?: boolean;
	hpx?: number;
	hpt?: number;
	level?: number;
}

/** Sheet Protection Properties */
export interface ProtectInfo {
	password?: string;
	selectLockedCells?: boolean;
	selectUnlockedCells?: boolean;
	formatCells?: boolean;
	formatColumns?: boolean;
	formatRows?: boolean;
	insertColumns?: boolean;
	insertRows?: boolean;
	insertHyperlinks?: boolean;
	deleteColumns?: boolean;
	deleteRows?: boolean;
	sort?: boolean;
	autoFilter?: boolean;
	pivotTables?: boolean;
	objects?: boolean;
	scenarios?: boolean;
}

/** Page Margins */
export interface MarginInfo {
	left?: number;
	right?: number;
	top?: number;
	bottom?: number;
	header?: number;
	footer?: number;
}

/** AutoFilter properties */
export interface AutoFilterInfo {
	ref: string;
}

export type DenseSheetData = ((CellObject | undefined)[] | undefined)[];

/** General object representing a Sheet */
export interface Sheet {
	[cell: string]: any;
	"!data"?: DenseSheetData;
	"!type"?: "sheet" | "chart";
	"!ref"?: string;
	"!margins"?: MarginInfo;
}

/** Worksheet Object */
export interface WorkSheet extends Sheet {
	"!cols"?: ColInfo[];
	"!rows"?: RowInfo[];
	"!merges"?: Range[];
	"!protect"?: ProtectInfo;
	"!autofilter"?: AutoFilterInfo;
}

/** Sheet Properties */
export interface SheetProps {
	name?: string;
	Hidden?: 0 | 1 | 2;
	CodeName?: string;
}

/** Defined Name Object */
export interface DefinedName {
	Name: string;
	Ref: string;
	Sheet?: number;
	Comment?: string;
	Hidden?: boolean;
}

/** Workbook View */
export interface WBView {
	RTL?: boolean;
}

/** Other Workbook Properties */
export interface WorkbookProperties {
	date1904?: boolean;
	filterPrivacy?: boolean;
	CodeName?: string;
}

/** Workbook-Level Attributes */
export interface WBProps {
	Sheets?: SheetProps[];
	Names?: DefinedName[];
	Views?: WBView[];
	WBProps?: WorkbookProperties;
}

/** Workbook Object */
export interface WorkBook {
	Sheets: { [sheet: string]: WorkSheet };
	SheetNames: string[];
	Props?: FullProperties;
	Custprops?: Record<string, any>;
	Workbook?: WBProps;
	vbaraw?: any;
	bookType?: string;
}

/** CSV output options */
export interface Sheet2CSVOpts {
	FS?: string;
	RS?: string;
	strip?: boolean;
	blankrows?: boolean;
	skipHidden?: boolean;
	forceQuotes?: boolean;
	rawNumbers?: boolean;
	dateNF?: NumberFormat;
}

/** HTML output options */
export interface Sheet2HTMLOpts {
	id?: string;
	editable?: boolean;
	header?: string;
	footer?: string;
	sanitizeLinks?: boolean;
}

/** JSON output options */
export interface Sheet2JSONOpts {
	header?: "A" | number | string[];
	range?: any;
	blankrows?: boolean;
	defval?: any;
	raw?: boolean;
	skipHidden?: boolean;
	rawNumbers?: boolean;
	UTC?: boolean;
	dateNF?: NumberFormat;
}

/** AOA to sheet options */
export interface AOA2SheetOpts extends CommonOptions {
	dense?: boolean;
	sheetStubs?: boolean;
	dateNF?: NumberFormat;
	cellDates?: boolean;
	UTC?: boolean;
	date1904?: boolean;
	origin?: number | string | CellAddress;
	nullError?: boolean;
}

/** JSON to sheet options */
export interface JSON2SheetOpts extends CommonOptions {
	header?: string[];
	skipHeader?: boolean;
	dense?: boolean;
	dateNF?: NumberFormat;
	cellDates?: boolean;
	UTC?: boolean;
	date1904?: boolean;
	origin?: number | string | CellAddress;
	nullError?: boolean;
}

/** Error codes */
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

/** Reverse error map */
export const RBErr: Record<string, number> = {};
for (const [k, v] of Object.entries(BErr)) {
	RBErr[v] = +k;
}
