import type { WorkBook, WorkSheet, ReadOptions, FullProperties } from "../types.js";
import type { ZipArchive } from "../zip/index.js";
import type { ContentTypes } from "../opc/content-types.js";
import type { Relationships } from "../opc/relationships.js";
import type { SST } from "./shared-strings.js";
import type { StylesData } from "./styles.js";
import type { ThemeData } from "./theme.js";
import type { WorkbookFile, SheetEntry } from "./workbook.js";
import { zipReadString, zipHas } from "../zip/index.js";
import { parseContentTypes } from "../opc/content-types.js";
import { parseRelationships, getRelsPath } from "../opc/relationships.js";
import { parseCoreProperties } from "../opc/core-properties.js";
import { parseExtendedProperties } from "../opc/extended-properties.js";
import { parseCustomProperties } from "../opc/custom-properties.js";
import { parseSstXml } from "./shared-strings.js";
import { parseStylesXml } from "./styles.js";
import { parse_theme_xml } from "./theme.js";
import { parseWorkbookXml } from "./workbook.js";
import { parseWorksheetXml, resolveSharedStrings } from "./worksheet.js";
import { parseCommentsXml, parseTcmntXml, parsePeopleXml, insertCommentsIntoSheet } from "./comments.js";
import { parseVml } from "./vml.js";
import { parseMetadataXml } from "./metadata.js";
import { parseCalcChainXml } from "./calc-chain.js";
import { resetFormatTable, formatTable } from "../ssf/table.js";
import { utf8read } from "../utils/buffer.js";
import { RELS as RELTYPE } from "../xml/namespaces.js";

/** Strip a leading "/" from a path (ZIP entries don't use leading slashes) */
function stripLeadingSlash(x: string): string {
	return x.charAt(0) === "/" ? x.slice(1) : x;
}

/**
 * Resolve a relative path against a base path.
 * Handles ".." segments for paths like "../comments1.xml" relative to "xl/worksheets/sheet1.xml".
 */
function resolve_path(target: string, basePath: string): string {
	if (target.charAt(0) === "/") {
		return target;
	}
	const base = basePath.slice(0, basePath.lastIndexOf("/") + 1);
	const parts = (base + target).split("/");
	const resolved: string[] = [];
	for (const p of parts) {
		if (p === "..") {
			resolved.pop();
		} else if (p !== ".") {
			resolved.push(p);
		}
	}
	return resolved.join("/");
}

/**
 * Read a file from the ZIP as a string.
 * Throws if the file is not found and safe is not set.
 */
function getZipString(zip: ZipArchive, path: string, safe?: boolean): string | null {
	const p = zipReadString(zip, path);
	if (p == null && !safe) {
		throw new Error("Could not find " + path);
	}
	return p;
}

/** Read ZIP entry data as string (alias for XML-based files) */
function getZipData(zip: ZipArchive, path: string, safe?: boolean): string | null {
	return getZipString(zip, path, safe);
}

/** Recognized relationship types for worksheets (standard and transitional) */
const RELS_WS = [
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
	"http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet",
];

/** Determine the sheet type from a relationship type URI */
function get_sheet_type(n: string): string {
	if (RELS_WS.indexOf(n) > -1) {
		return "sheet";
	}
	return n && n.length ? n : "sheet";
}

/**
 * Safely map workbook sheet entries to their target paths and types using
 * the workbook relationships. Returns null if mapping fails.
 */
function safe_parse_wbrels(wbrels: Relationships, sheets: SheetEntry[]): [string, string, string][] | null {
	if (!wbrels) {
		return null;
	}
	try {
		const result: [string, string, string][] = sheets.map((sheetEntry) => {
			const id = (sheetEntry as any).id || (sheetEntry as any).strRelID;
			// [name, target path, sheet type]
			return [sheetEntry.name, wbrels["!id"][id].Target, get_sheet_type(wbrels["!id"][id].Type)];
		});
		return result.length === 0 ? null : result;
	} catch {
		return null;
	}
}

/**
 * Safely parse a single sheet from the ZIP, including its relationships,
 * comments, threaded comments, and VML drawings.
 */
function safe_parse_sheet(
	zip: ZipArchive,
	path: string,
	relsPath: string,
	sheetName: string,
	idx: number,
	sheetRels: Record<string, Relationships>,
	sheets: Record<string, WorkSheet>,
	stype: string,
	opts: any,
	wb: WorkbookFile,
	themes: ThemeData,
	styles: StylesData,
	strs: SST,
): void {
	try {
		sheetRels[sheetName] = parseRelationships(getZipString(zip, relsPath, true), path);
		const data = getZipData(zip, path);
		if (!data) {
			return;
		}

		let _ws: WorkSheet | undefined;
		switch (stype) {
			case "sheet":
				_ws = parseWorksheetXml(data, opts);
				break;
			default:
				return;
		}
		if (!_ws) {
			return;
		}

		// Replace SST index placeholders with actual string values
		resolveSharedStrings(_ws, strs, opts);

		sheets[sheetName] = _ws;

		// Scan sheet relationships for comments and threaded comments
		const comments: any[] = [];
		let tcomments: any[] = [];
		if (sheetRels[sheetName]) {
			for (const n of Object.keys(sheetRels[sheetName])) {
				// Skip internal keys
				if (n === "!id" || n === "!idx") {
					continue;
				}
				const rel = sheetRels[sheetName][n];
				if (!rel || !rel.Type) {
					continue;
				}

				// Parse legacy comments (ECMA-376 18.7)
				if (rel.Type === RELTYPE.CMNT) {
					const dfile = resolve_path(rel.Target, path);
					const cmntData = getZipData(zip, dfile, true);
					if (cmntData) {
						const parsedComments = parseCommentsXml(cmntData, opts);
						if (parsedComments && parsedComments.length > 0) {
							insertCommentsIntoSheet(_ws, parsedComments, false);
						}
					}
				}
				// Parse threaded comments (MS-XLSX extension)
				if (rel.Type === RELTYPE.TCMNT) {
					const dfile = resolve_path(rel.Target, path);
					const tcData = getZipData(zip, dfile, true);
					if (tcData) {
						tcomments = tcomments.concat(parseTcmntXml(tcData, opts));
					}
				}
			}
		}
		if (tcomments.length > 0) {
			insertCommentsIntoSheet(_ws, tcomments, true, opts.people || []);
		}

		// Parse legacy VML drawings (comment anchor shapes)
		if ((_ws as any)["!legdrawel"] && sheetRels[sheetName]) {
			const dfile = resolve_path((_ws as any)["!legdrawel"].Target, path);
			const draw = getZipString(zip, dfile, true);
			if (draw) {
				parseVml(utf8read(draw), _ws, comments);
			}
		}
	} catch (e) {
		if (opts.WTF) {
			throw e;
		}
	}
}

/**
 * Parse an XLSX ZIP archive into a WorkBook object.
 *
 * Orchestrates reading of all XLSX parts: content types, relationships,
 * shared strings, themes, styles, workbook, properties, metadata, people,
 * and individual worksheets with their comments and VML drawings.
 *
 * @param zip - ZipArchive containing the XLSX file parts
 * @param opts - Read options controlling parsing behavior
 * @returns Parsed WorkBook with sheets, properties, and metadata
 * @throws Error if the ZIP is not a valid XLSX file or the workbook is missing
 */
export function parseZip(zip: ZipArchive, opts?: ReadOptions): WorkBook {
	resetFormatTable();
	const options: any = opts || {};

	if (!zipHas(zip, "[Content_Types].xml")) {
		throw new Error("Unsupported ZIP file");
	}

	const dir: ContentTypes = parseContentTypes(getZipString(zip, "[Content_Types].xml"));

	// Fallback: if no workbook found in content types, try the default path
	if (dir.workbooks.length === 0) {
		const binname = "xl/workbook.xml";
		if (getZipData(zip, binname, true)) {
			dir.workbooks.push(binname);
		}
	}
	if (dir.workbooks.length === 0) {
		throw new Error("Could not find workbook");
	}

	const themes: ThemeData = { themeElements: { clrScheme: [] } };
	let styles: StylesData = { NumberFmt: {}, CellXf: [], Fonts: [], Fills: [], Borders: [] };
	let strs: SST = [] as any;

	// Parse shared resources unless only sheet names or properties are requested
	if (!options.bookSheets && !options.bookProps) {
		// Shared String Table
		if (dir.sst) {
			try {
				const sstData = getZipData(zip, stripLeadingSlash(dir.sst));
				if (sstData) {
					strs = parseSstXml(sstData, options);
				}
			} catch (e) {
				if (options.WTF) {
					throw e;
				}
			}
		}

		// Theme (color scheme for styled cells)
		if (dir.themes.length) {
			const themeData = getZipString(zip, dir.themes[0].replace(/^\//, ""), true);
			if (themeData) {
				const parsed = parse_theme_xml(themeData);
				Object.assign(themes, parsed);
			}
		}

		// Styles (number formats, cell xf entries)
		if (dir.style) {
			const styData = getZipData(zip, stripLeadingSlash(dir.style));
			if (styData) {
				styles = parseStylesXml(styData, themes, options);
			}
		}
	}

	// Parse the workbook XML (sheet list, defined names, properties)
	const wb: WorkbookFile = parseWorkbookXml(getZipData(zip, stripLeadingSlash(dir.workbooks[0]))!, options);

	// Parse document properties
	const props: any = {};
	if (dir.coreprops.length) {
		const propdata = getZipData(zip, stripLeadingSlash(dir.coreprops[0]), true);
		if (propdata) {
			Object.assign(props, parseCoreProperties(propdata));
		}
		if (dir.extprops.length) {
			const extdata = getZipData(zip, stripLeadingSlash(dir.extprops[0]), true);
			if (extdata) {
				parseExtendedProperties(extdata, props);
			}
		}
	}

	// Parse custom properties
	let custprops: Record<string, any> = {};
	if (!options.bookSheets || options.bookProps) {
		if (dir.custprops.length) {
			const custdata = getZipString(zip, stripLeadingSlash(dir.custprops[0]), true);
			if (custdata) {
				custprops = parseCustomProperties(custdata, options);
			}
		}
	}

	// Early return for bookSheets/bookProps-only mode
	const out: any = {};
	if (options.bookSheets || options.bookProps) {
		let sheets: string[] | undefined;
		if (wb.Sheets) {
			sheets = wb.Sheets.map((x: SheetEntry) => x.name);
		} else if (props.Worksheets && props.SheetNames?.length > 0) {
			sheets = props.SheetNames;
		}
		if (options.bookProps) {
			out.Props = props;
			out.Custprops = custprops;
		}
		if (options.bookSheets && sheets) {
			out.SheetNames = sheets;
		}
		if (options.bookSheets ? out.SheetNames : options.bookProps) {
			return out as WorkBook;
		}
	}

	const sheets: Record<string, WorkSheet> = {};

	// Calculation chain (dependency order for formula recalc)
	if (options.bookDeps && dir.calcchain) {
		parseCalcChainXml(getZipData(zip, stripLeadingSlash(dir.calcchain), true) || "");
	}

	// Build the sheet name list from the workbook
	const sheetRels: Record<string, Relationships> = {};
	const wbsheets = wb.Sheets;
	props.Worksheets = wbsheets.length;
	props.SheetNames = [];
	for (let j = 0; j < wbsheets.length; ++j) {
		props.SheetNames[j] = wbsheets[j].name;
	}

	// Locate the workbook relationships file
	const wbrelsi = dir.workbooks[0].lastIndexOf("/");
	let wbrelsfile = (
		dir.workbooks[0].slice(0, wbrelsi + 1) +
		"_rels/" +
		dir.workbooks[0].slice(wbrelsi + 1) +
		".rels"
	).replace(/^\//, "");
	if (!zipHas(zip, wbrelsfile)) {
		wbrelsfile = "xl/_rels/workbook.xml.rels";
	}
	const wbrels = parseRelationships(getZipString(zip, wbrelsfile, true), wbrelsfile.replace(/_rels.*/, "s5s"));

	// Parse cell metadata (for dynamic arrays, rich data, etc.)
	if ((dir.metadata || []).length >= 1) {
		options.xlmeta = parseMetadataXml(getZipData(zip, stripLeadingSlash(dir.metadata[0]), true) || "", options);
	}

	// Parse people list (for threaded comment author resolution)
	if ((dir.people || []).length >= 1) {
		options.people = parsePeopleXml(getZipData(zip, stripLeadingSlash(dir.people[0]), true) || "");
	}

	// Map sheet entries to their ZIP paths via workbook relationships
	const wbrelsArr = wbrels ? safe_parse_wbrels(wbrels, wb.Sheets) : null;

	// Detect legacy naming: some files use "sheet.xml" instead of "sheet1.xml"
	const nmode = getZipData(zip, "xl/worksheets/sheet.xml", true) ? 1 : 0;

	for (let i = 0; i < props.Worksheets; ++i) {
		let stype = "sheet";
		let path: string;
		if (wbrelsArr && wbrelsArr[i]) {
			// Resolve the sheet path from relationships, trying multiple fallback locations
			path = "xl/" + wbrelsArr[i][1].replace(/[/]?xl\//, "");
			if (!zipHas(zip, path)) {
				path = wbrelsArr[i][1];
			}
			if (!zipHas(zip, path)) {
				path = wbrelsfile.replace(/_rels\/[\S\s]*$/, "") + wbrelsArr[i][1];
			}
			stype = wbrelsArr[i][2];
		} else {
			// Fallback: construct path from sheet index
			path = "xl/worksheets/sheet" + (i + 1 - nmode) + ".xml";
			path = path.replace(/sheet0\./, "sheet.");
		}

		// Apply sheet filter (by index, name, or array of indices/names)
		if (options.sheets != null) {
			if (typeof options.sheets === "number" && i !== options.sheets) {
				continue;
			}
			if (typeof options.sheets === "string" && props.SheetNames[i].toLowerCase() !== options.sheets.toLowerCase()) {
				continue;
			}
			if (Array.isArray(options.sheets)) {
				let seen = false;
				for (const s of options.sheets) {
					if (typeof s === "number" && s === i) {
						seen = true;
					}
					if (typeof s === "string" && s.toLowerCase() === props.SheetNames[i].toLowerCase()) {
						seen = true;
					}
				}
				if (!seen) {
					continue;
				}
			}
		}

		// Derive the per-sheet .rels path from the sheet path
		const relsPath = path.replace(/^(.*)(\/)([^/]*)$/, "$1/_rels/$3.rels");
		safe_parse_sheet(
			zip,
			path,
			relsPath,
			props.SheetNames[i],
			i,
			sheetRels,
			sheets,
			stype,
			options,
			wb,
			themes,
			styles,
			strs,
		);
	}

	const result: WorkBook = {
		Sheets: sheets,
		SheetNames: props.SheetNames,
		Props: props,
		Custprops: custprops,
		bookType: "xlsx",
	};

	// Attach workbook-level metadata (properties, sheet visibility, defined names)
	if (wb.WBProps) {
		result.Workbook = {
			WBProps: wb.WBProps,
			Sheets: wb.Sheets as any,
			Names: wb.Names as any,
		};
	}

	return result;
}
