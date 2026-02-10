/** Number Format (either a string or an index to the format table) */
type NumberFormat = string | number;
/** Basic File Properties */
interface Properties {
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
interface FullProperties extends Properties {
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
interface CommonOptions {
    WTF?: boolean;
    bookVBA?: boolean;
    cellDates?: boolean;
    sheetStubs?: boolean;
    cellStyles?: boolean;
    password?: string;
}
interface ReadOptions extends CommonOptions {
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
interface WriteOptions extends CommonOptions {
    type?: "base64" | "buffer" | "array";
    bookSST?: boolean;
    compression?: boolean;
    themeXLSX?: string;
    ignoreEC?: boolean;
    Props?: Properties;
}
/** The Excel data type for a cell: b Boolean, n Number, e Error, s String, d Date, z Empty */
type ExcelDataType = "b" | "n" | "e" | "s" | "d" | "z";
/** Comment element */
interface Comment {
    a?: string;
    t: string;
    T?: boolean;
}
/** Cell comments */
interface Comments extends Array<Comment> {
    hidden?: boolean;
}
/** Link object */
interface Hyperlink {
    Target: string;
    Tooltip?: string;
}
/** Worksheet Cell Object */
interface CellObject {
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
    XF?: {
        numFmtId?: number;
    };
}
/** Simple Cell Address */
interface CellAddress {
    c: number;
    r: number;
}
/** Range object (representing ranges like "A1:B2") */
interface Range {
    s: CellAddress;
    e: CellAddress;
}
/** Column Properties Object */
interface ColInfo {
    hidden?: boolean;
    width?: number;
    wpx?: number;
    wch?: number;
    level?: number;
    MDW?: number;
}
/** Row Properties Object */
interface RowInfo {
    hidden?: boolean;
    hpx?: number;
    hpt?: number;
    level?: number;
}
/** Sheet Protection Properties */
interface ProtectInfo {
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
interface MarginInfo {
    left?: number;
    right?: number;
    top?: number;
    bottom?: number;
    header?: number;
    footer?: number;
}
/** AutoFilter properties */
interface AutoFilterInfo {
    ref: string;
}
type DenseSheetData = ((CellObject | undefined)[] | undefined)[];
/** General object representing a Sheet */
interface Sheet {
    [cell: string]: any;
    "!data"?: DenseSheetData;
    "!type"?: "sheet" | "chart";
    "!ref"?: string;
    "!margins"?: MarginInfo;
}
/** Worksheet Object */
interface WorkSheet extends Sheet {
    "!cols"?: ColInfo[];
    "!rows"?: RowInfo[];
    "!merges"?: Range[];
    "!protect"?: ProtectInfo;
    "!autofilter"?: AutoFilterInfo;
}
/** Sheet Properties */
interface SheetProps {
    name?: string;
    Hidden?: 0 | 1 | 2;
    CodeName?: string;
}
/** Defined Name Object */
interface DefinedName {
    Name: string;
    Ref: string;
    Sheet?: number;
    Comment?: string;
    Hidden?: boolean;
}
/** Workbook View */
interface WBView {
    RTL?: boolean;
}
/** Other Workbook Properties */
interface WorkbookProperties {
    date1904?: boolean;
    filterPrivacy?: boolean;
    CodeName?: string;
}
/** Workbook-Level Attributes */
interface WBProps {
    Sheets?: SheetProps[];
    Names?: DefinedName[];
    Views?: WBView[];
    WBProps?: WorkbookProperties;
}
/** Workbook Object */
interface WorkBook {
    Sheets: {
        [sheet: string]: WorkSheet;
    };
    SheetNames: string[];
    Props?: FullProperties;
    Custprops?: Record<string, any>;
    Workbook?: WBProps;
    vbaraw?: any;
    bookType?: string;
}
/** CSV output options */
interface Sheet2CSVOpts {
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
interface Sheet2HTMLOpts {
    id?: string;
    editable?: boolean;
    header?: string;
    footer?: string;
    sanitizeLinks?: boolean;
}
/** JSON output options */
interface Sheet2JSONOpts {
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
interface AOA2SheetOpts extends CommonOptions {
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
interface JSON2SheetOpts extends CommonOptions {
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

/**
 * Read an XLSX file from a data source.
 *
 * @param data - File contents as Uint8Array, ArrayBuffer, Buffer, base64 string, or binary string
 * @param opts - Read options
 * @returns Parsed WorkBook object
 */
declare function read(data: any, opts?: ReadOptions): WorkBook;
/**
 * Read an XLSX file from the filesystem.
 *
 * @param filename - Path to the XLSX file
 * @param opts - Read options
 * @returns Parsed WorkBook object
 */
declare function readFile(filename: string, opts?: ReadOptions): WorkBook;

/**
 * Write a WorkBook to a Uint8Array (XLSX format).
 *
 * @param wb - WorkBook object to write
 * @param opts - Write options
 * @returns File contents as Uint8Array, base64 string, or Buffer depending on opts.type
 */
declare function write(wb: WorkBook, opts?: WriteOptions): any;
/**
 * Write a WorkBook to a file (XLSX format).
 *
 * @param wb - WorkBook object to write
 * @param filename - Output file path
 * @param opts - Write options
 */
declare function writeFile(wb: WorkBook, filename: string, opts?: WriteOptions): void;

/** Create a new blank workbook, optionally with a first sheet */
declare function book_new(ws?: WorkSheet, wsname?: string): WorkBook;
/** Add a worksheet to the end of a workbook */
declare function book_append_sheet(wb: WorkBook, ws: WorkSheet, name?: string, roll?: boolean): string;
/** Create a new empty worksheet */
declare function sheet_new(opts?: {
    dense?: boolean;
}): WorkSheet;
/** Find sheet index for given name or validate index */
declare function wb_sheet_idx(wb: WorkBook, sh: number | string): number;
/** Set sheet visibility (0=visible, 1=hidden, 2=veryHidden) */
declare function book_set_sheet_visibility(wb: WorkBook, sh: number | string, vis: 0 | 1 | 2): void;
/** Set a cell's number format */
declare function cell_set_number_format(cell: CellObject, fmt: string | number): CellObject;
/** Set a cell's hyperlink */
declare function cell_set_hyperlink(cell: CellObject, target?: string, tooltip?: string): CellObject;
/** Set an internal link (starts with #) on a cell */
declare function cell_set_internal_link(cell: CellObject, range: string, tooltip?: string): CellObject;
/** Add a comment to a cell */
declare function cell_add_comment(cell: CellObject, text: string, author?: string): void;
/** Set an array formula on a range of cells */
declare function sheet_set_array_formula(ws: WorkSheet, range: string | {
    s: {
        r: number;
        c: number;
    };
    e: {
        r: number;
        c: number;
    };
}, formula: string, dynamic?: boolean): WorkSheet;
/** Convert a worksheet to an array of formula strings */
declare function sheet_to_formulae(ws: WorkSheet): string[];

/** Add an array of arrays to an existing (or new) worksheet */
declare function sheet_add_aoa(_ws: WorkSheet | null, data: any[][], opts?: AOA2SheetOpts): WorkSheet;
/** Create a new worksheet from an array of arrays */
declare function aoa_to_sheet(data: any[][], opts?: AOA2SheetOpts): WorkSheet;

/** Convert a worksheet to an array of JSON objects */
declare function sheet_to_json<T = any>(sheet: WorkSheet, opts?: Sheet2JSONOpts): T[];
/** Add JSON data to a worksheet */
declare function sheet_add_json(_ws: WorkSheet | null, js: any[], opts?: JSON2SheetOpts): WorkSheet;
/** Create a new worksheet from JSON data */
declare function json_to_sheet(js: any[], opts?: JSON2SheetOpts): WorkSheet;

/** Convert a worksheet to CSV string */
declare function sheet_to_csv(sheet: WorkSheet, opts?: Sheet2CSVOpts): string;
/** Convert a worksheet to tab-separated text */
declare function sheet_to_txt(sheet: WorkSheet, opts?: Sheet2CSVOpts): string;

/** Convert a worksheet to an HTML table string */
declare function sheet_to_html(ws: WorkSheet, opts?: Sheet2HTMLOpts): string;

declare function format_cell(cell: CellObject, v?: any, o?: any): string;

declare function decode_row(rowstr: string): number;
declare function encode_row(row: number): string;
declare function decode_col(colstr: string): number;
declare function encode_col(col: number): string;
declare function decode_cell(cstr: string): CellAddress;
declare function encode_cell(cell: CellAddress): string;
declare function decode_range(range: string): Range;
declare function encode_range(cs: CellAddress | Range, ce?: CellAddress): string;

/** Format a value using an Excel number format string */
declare function SSF_format(fmt: string | number, v: any, o?: any): string;

declare const version = "1.0.0-alpha.0";

export { type AOA2SheetOpts, type AutoFilterInfo, type CellAddress, type CellObject, type ColInfo, type Comment, type Comments, type DefinedName, type DenseSheetData, type ExcelDataType, type FullProperties, type Hyperlink, type JSON2SheetOpts, type MarginInfo, type NumberFormat, type Properties, type ProtectInfo, type Range, type ReadOptions, type RowInfo, SSF_format, type Sheet, type Sheet2CSVOpts, type Sheet2HTMLOpts, type Sheet2JSONOpts, type SheetProps, type WBProps, type WBView, type WorkBook, type WorkSheet, type WorkbookProperties, type WriteOptions, aoa_to_sheet, book_append_sheet, book_new, book_set_sheet_visibility, cell_add_comment, cell_set_hyperlink, cell_set_internal_link, cell_set_number_format, decode_cell, decode_col, decode_range, decode_row, encode_cell, encode_col, encode_range, encode_row, format_cell, json_to_sheet, read, readFile, sheet_add_aoa, sheet_add_json, sheet_new, sheet_set_array_formula, sheet_to_csv, sheet_to_formulae, sheet_to_html, sheet_to_json, sheet_to_txt, version, wb_sheet_idx, write, writeFile };
