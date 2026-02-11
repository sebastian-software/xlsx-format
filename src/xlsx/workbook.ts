import { parseXmlTag, XML_TAG_REGEX, XML_HEADER, stripNamespace, parseXmlBoolean } from "../xml/parser.js";
import { unescapeXml, escapeXml } from "../xml/escape.js";
import { writeXmlElement } from "../xml/writer.js";
import { XMLNS_main, XMLNS } from "../xml/namespaces.js";
import { utf8read } from "../utils/buffer.js";
import type { WorkBook } from "../types.js";

/** Parsed workbook.xml structure */
export interface WorkbookFile {
	AppVersion: Record<string, any>;
	WBProps: Record<string, any>;
	WBView: Record<string, any>[];
	Sheets: SheetEntry[];
	CalcPr: Record<string, any>;
	Names: DefinedNameEntry[];
	xmlns: string;
}

/** A sheet entry from the <sheets> element in workbook.xml */
export interface SheetEntry {
	name: string;
	sheetId: string;
	/** 0 = visible, 1 = hidden, 2 = veryHidden */
	Hidden?: number;
	[key: string]: any;
}

/** A defined name entry from the <definedNames> element */
export interface DefinedNameEntry {
	Name: string;
	Ref?: string;
	/** Local sheet index (undefined = workbook-scoped) */
	Sheet?: number;
	Comment?: string;
	Hidden?: boolean;
}

/** Default values and types for workbook properties (workbookPr) */
const WBPropsDef: [string, any, string?][] = [
	["allowRefreshQuery", false, "bool"],
	["autoCompressPictures", true, "bool"],
	["backupFile", false, "bool"],
	["checkCompatibility", false, "bool"],
	["CodeName", ""],
	["date1904", false, "bool"],
	["defaultThemeVersion", 0, "int"],
	["filterPrivacy", false, "bool"],
	["hidePivotFieldList", false, "bool"],
	["promptedSolutions", false, "bool"],
	["publishItems", false, "bool"],
	["refreshAllConnections", false, "bool"],
	["saveExternalLinkValues", true, "bool"],
	["showBorderUnselectedTables", true, "bool"],
	["showInkAnnotation", true, "bool"],
	["showObjects", "all"],
	["showPivotChartFilter", false, "bool"],
	["updateLinks", "userSet"],
];

/** Default values and types for workbook view properties (workbookView) */
const WBViewDef: [string, any, string?][] = [
	["activeTab", 0, "int"],
	["autoFilterDateGrouping", true, "bool"],
	["firstSheet", 0, "int"],
	["minimized", false, "bool"],
	["showHorizontalScroll", true, "bool"],
	["showSheetTabs", true, "bool"],
	["showVerticalScroll", true, "bool"],
	["tabRatio", 600, "int"],
	["visibility", "visible"],
];

/** Default values for sheet entries (currently empty, extensible) */
const SheetDef: [string, any, string?][] = [];

/** Default values and types for calculation properties (calcPr) */
const CalcPrDef: [string, any, string?][] = [
	["calcCompleted", "true"],
	["calcMode", "auto"],
	["calcOnSave", "true"],
	["concurrentCalc", "true"],
	["fullCalcOnLoad", "false"],
	["fullPrecision", "true"],
	["iterate", "false"],
	["iterateCount", "100"],
	["iterateDelta", "0.001"],
	["refMode", "A1"],
];

/**
 * Apply default values to each entry in an array of objects.
 * Coerces string values to bool/int based on the type hint in the defaults definition.
 */
function applyDefaultsToArray(target: any[], defaults: [string, any, string?][]): void {
	for (let j = 0; j < target.length; ++j) {
		const entry = target[j];
		for (let i = 0; i < defaults.length; ++i) {
			const defaultDef = defaults[i];
			if (entry[defaultDef[0]] == null) {
				entry[defaultDef[0]] = defaultDef[1];
			} else {
				switch (defaultDef[2]) {
					case "bool":
						if (typeof entry[defaultDef[0]] === "string") {
							entry[defaultDef[0]] = parseXmlBoolean(entry[defaultDef[0]]);
						}
						break;
					case "int":
						if (typeof entry[defaultDef[0]] === "string") {
							entry[defaultDef[0]] = parseInt(entry[defaultDef[0]], 10);
						}
						break;
				}
			}
		}
	}
}

/**
 * Apply default values to a single object.
 * Coerces string values to bool/int based on the type hint in the defaults definition.
 */
function applyDefaults(target: Record<string, any>, defaults: [string, any, string?][]): void {
	for (let i = 0; i < defaults.length; ++i) {
		const defaultDef = defaults[i];
		if (target[defaultDef[0]] == null) {
			target[defaultDef[0]] = defaultDef[1];
		} else {
			switch (defaultDef[2]) {
				case "bool":
					if (typeof target[defaultDef[0]] === "string") {
						target[defaultDef[0]] = parseXmlBoolean(target[defaultDef[0]]);
					}
					break;
				case "int":
					if (typeof target[defaultDef[0]] === "string") {
						target[defaultDef[0]] = parseInt(target[defaultDef[0]], 10);
					}
					break;
			}
		}
	}
}

/**
 * Apply default values to all sections of a parsed workbook file.
 *
 * @param wb - Parsed workbook file to fill with defaults
 */
export function parse_wb_defaults(wb: WorkbookFile): void {
	applyDefaults(wb.WBProps, WBPropsDef);
	applyDefaults(wb.CalcPr, CalcPrDef);
	applyDefaultsToArray(wb.WBView, WBViewDef);
	applyDefaultsToArray(wb.Sheets, SheetDef);
}

/**
 * Determine whether the workbook uses the 1904 date system (Mac Excel legacy).
 *
 * @param wb - WorkBook to inspect
 * @returns "true" if 1904 date system is active, "false" otherwise
 */
export function is1904DateSystem(wb: WorkBook): string {
	if (!wb.Workbook) {
		return "false";
	}
	if (!wb.Workbook.WBProps) {
		return "false";
	}
	return parseXmlBoolean(wb.Workbook.WBProps.date1904) ? "true" : "false";
}

/** Characters forbidden in Excel sheet names */
const badchars = ":][*?/\\".split("");

/**
 * Validate a sheet name against Excel naming rules.
 *
 * @param n - Sheet name to validate
 * @param safe - If true, return false on invalid names instead of throwing
 * @returns true if valid
 * @throws Error describing the validation failure (unless safe=true)
 */
export function validateSheetName(n: string, safe?: boolean): boolean {
	try {
		if (n === "") {
			throw new Error("Sheet name cannot be blank");
		}
		if (n.length > 31) {
			throw new Error("Sheet name cannot exceed 31 chars");
		}
		// 0x27 = apostrophe (')
		if (n.charCodeAt(0) === 0x27 || n.charCodeAt(n.length - 1) === 0x27) {
			throw new Error("Sheet name cannot start or end with apostrophe (')");
		}
		if (n.toLowerCase() === "history") {
			throw new Error("Sheet name cannot be 'History'");
		}
		for (const c of badchars) {
			if (n.indexOf(c) !== -1) {
				throw new Error("Sheet name cannot contain : \\ / ? * [ ]");
			}
		}
	} catch (e) {
		if (safe) {
			return false;
		}
		throw e;
	}
	return true;
}

/**
 * Validate all sheet names in a workbook for correctness and uniqueness.
 *
 * @param sheetNames - Array of sheet names to validate
 * @param sheetEntries - Optional sheet entry metadata (reserved for future use)
 * @throws Error if any name is invalid or duplicated
 */
export function validateWorkbookNames(sheetNames: string[], _sheetEntries?: any[]): void {
	for (let i = 0; i < sheetNames.length; ++i) {
		validateSheetName(sheetNames[i]);
		for (let j = 0; j < i; ++j) {
			if (sheetNames[i] === sheetNames[j]) {
				throw new Error("Duplicate Sheet Name: " + sheetNames[i]);
			}
		}
	}
}

/**
 * Validate that a WorkBook object has the required structure.
 *
 * @param wb - WorkBook to validate
 * @throws Error if the workbook is missing required fields or has invalid sheet names
 */
export function validateWorkbook(wb: WorkBook): void {
	if (!wb || !wb.SheetNames || !wb.Sheets) {
		throw new Error("Invalid Workbook");
	}
	if (!wb.SheetNames.length) {
		throw new Error("Workbook is empty");
	}
	const Sheets = (wb.Workbook && wb.Workbook.Sheets) || [];
	validateWorkbookNames(wb.SheetNames, Sheets);
}

/** Detects whether the workbook XML uses a namespace prefix (e.g. <x:workbook>) */
const wbnsregex = /<\w+:workbook/;

/**
 * Parse a workbook.xml file into a WorkbookFile structure.
 *
 * Extracts file version, workbook properties, views, sheet list, defined names,
 * and calculation properties from the XML.
 *
 * @param data - Raw XML string of workbook.xml
 * @param opts - Parsing options
 * @returns Parsed workbook file structure
 * @throws Error if data is empty or the namespace is unrecognized
 */
export function parseWorkbookXml(data: string, _opts?: any): WorkbookFile {
	if (!data) {
		throw new Error("Could not find file");
	}
	const workbook: WorkbookFile = {
		AppVersion: {},
		WBProps: {},
		WBView: [],
		Sheets: [],
		CalcPr: {},
		Names: [],
		xmlns: "",
	};
	let xmlns = "xmlns";
	let dname: any = {};
	// Track the character offset where the defined name content starts
	let dnstart = 0;

	const ignoredTags = new Set([
		"<?xml",
		"</workbook>",
		"<fileVersion/>",
		"</fileVersion>",
		"<fileSharing",
		"<fileSharing/>",
		"</workbookPr>",
		"<workbookProtection",
		"<workbookProtection/>",
		"<bookViews",
		"<bookViews>",
		"</bookViews>",
		"</workbookView>",
		"<sheets",
		"<sheets>",
		"</sheets>",
		"</sheet>",
		"<functionGroups",
		"<functionGroups/>",
		"<functionGroup",
		"<externalReferences",
		"</externalReferences>",
		"<externalReferences>",
		"<externalReference",
		"<definedNames/>",
		"<definedNames>",
		"<definedNames",
		"</definedNames>",
		"<definedName/>",
		"</calcPr>",
		"<oleSize",
		"<customWorkbookViews>",
		"</customWorkbookViews>",
		"<customWorkbookViews",
		"<customWorkbookView",
		"</customWorkbookView>",
		"<pivotCaches>",
		"</pivotCaches>",
		"<pivotCaches",
		"<pivotCache",
		"<smartTagPr",
		"<smartTagPr/>",
		"<smartTagTypes",
		"<smartTagTypes>",
		"</smartTagTypes>",
		"<smartTagType",
		"<webPublishing",
		"<webPublishing/>",
		"<fileRecoveryPr",
		"<fileRecoveryPr/>",
		"<webPublishObjects>",
		"<webPublishObjects",
		"</webPublishObjects>",
		"<webPublishObject",
		"<extLst",
		"<extLst>",
		"</extLst>",
		"<extLst/>",
		"<ext",
		"</ext>",
		"<ArchID",
		"<AlternateContent",
		"<AlternateContent>",
		"</AlternateContent>",
		"<revisionPtr",
	]);

	data.replace(XML_TAG_REGEX, function xml_wb(xmlTag: string, idx: number): string {
		const parsedTag: any = parseXmlTag(xmlTag);
		const tag = stripNamespace(parsedTag[0]);
		if (ignoredTags.has(tag)) {
			return xmlTag;
		}
		switch (tag) {
			case "<workbook":
				// Detect namespace prefix (e.g. <x:workbook -> xmlns:x)
				if (xmlTag.match(wbnsregex)) {
					xmlns = "xmlns" + xmlTag.match(/<(\w+):/)?.[1];
				}
				workbook.xmlns = parsedTag[xmlns];
				break;

			case "<fileVersion":
				delete parsedTag[0];
				workbook.AppVersion = parsedTag;
				break;

			case "<workbookPr":
			case "<workbookPr/>":
				WBPropsDef.forEach((propDef) => {
					if (parsedTag[propDef[0]] == null) {
						return;
					}
					switch (propDef[2]) {
						case "bool":
							workbook.WBProps[propDef[0]] = parseXmlBoolean(parsedTag[propDef[0]]);
							break;
						case "int":
							workbook.WBProps[propDef[0]] = parseInt(parsedTag[propDef[0]], 10);
							break;
						default:
							workbook.WBProps[propDef[0]] = parsedTag[propDef[0]];
					}
				});
				if (parsedTag.codeName) {
					workbook.WBProps.CodeName = utf8read(parsedTag.codeName);
				}
				break;

			case "<workbookView":
			case "<workbookView/>":
				delete parsedTag[0];
				workbook.WBView.push(parsedTag);
				break;

			case "<sheet":
				// Map state attribute to numeric Hidden value
				switch (parsedTag.state) {
					case "hidden":
						parsedTag.Hidden = 1;
						break;
					case "veryHidden":
						parsedTag.Hidden = 2;
						break;
					default:
						parsedTag.Hidden = 0;
				}
				delete parsedTag.state;
				parsedTag.name = unescapeXml(utf8read(parsedTag.name));
				delete parsedTag[0];
				workbook.Sheets.push(parsedTag);
				break;

			case "<definedName": {
				dname = {};
				dname.Name = utf8read(parsedTag.name);
				if (parsedTag.comment) {
					dname.Comment = parsedTag.comment;
				}
				if (parsedTag.localSheetId) {
					dname.Sheet = +parsedTag.localSheetId;
				}
				if (parseXmlBoolean(parsedTag.hidden || "0")) {
					dname.Hidden = true;
				}
				// Record position after the opening tag to extract the Ref content later
				dnstart = idx + xmlTag.length;
				break;
			}
			case "</definedName>": {
				// Extract the defined name reference formula from between open/close tags
				dname.Ref = unescapeXml(utf8read(data.slice(dnstart, idx)));
				workbook.Names.push(dname);
				break;
			}

			case "<calcPr":
			case "<calcPr/>":
				delete parsedTag[0];
				workbook.CalcPr = parsedTag;
				break;
		}
		return xmlTag;
	});

	if (XMLNS_main.indexOf(workbook.xmlns) === -1) {
		throw new Error("Unknown Namespace: " + workbook.xmlns);
	}

	parse_wb_defaults(workbook);

	return workbook;
}

/**
 * Write the workbook.xml containing the sheet list, defined names, and properties.
 *
 * @param wb - WorkBook to serialize
 * @returns Complete workbook.xml string
 */
export function writeWorkbookXml(wb: WorkBook): string {
	const lines: string[] = [XML_HEADER];
	lines.push(
		writeXmlElement("workbook", null, {
			xmlns: XMLNS_main[0],
			"xmlns:r": XMLNS.r,
		}),
	);

	const write_names = !!(wb.Workbook && (wb.Workbook.Names || []).length > 0);

	const workbookPr: any = { codeName: "ThisWorkbook" };
	if (wb.Workbook && wb.Workbook.WBProps) {
		// Only write properties that differ from defaults
		WBPropsDef.forEach((x) => {
			if (!wb.Workbook || !wb.Workbook.WBProps) {
				return;
			}
			const wbp = wb.Workbook.WBProps as any;
			if (wbp[x[0]] == null) {
				return;
			}
			if (wbp[x[0]] === x[1]) {
				return;
			}
			workbookPr[x[0]] = wbp[x[0]];
		});
		if (wb.Workbook.WBProps.CodeName) {
			workbookPr.codeName = wb.Workbook.WBProps.CodeName;
			delete workbookPr.CodeName;
		}
	}
	lines.push(writeXmlElement("workbookPr", null, workbookPr));

	const sheets = (wb.Workbook && wb.Workbook.Sheets) || [];

	/* bookViews: only written if the first worksheet is hidden, to set activeTab to first visible sheet */
	if (sheets[0] && !!sheets[0].Hidden) {
		lines.push("<bookViews>");
		let i = 0;
		for (i = 0; i < wb.SheetNames.length; ++i) {
			if (!sheets[i]) {
				break;
			}
			if (!sheets[i].Hidden) {
				break;
			}
		}
		if (i === wb.SheetNames.length) {
			i = 0;
		}
		lines.push('<workbookView firstSheet="' + i + '" activeTab="' + i + '"/>');
		lines.push("</bookViews>");
	}

	// Sheet list
	lines.push("<sheets>");
	for (let i = 0; i < wb.SheetNames.length; ++i) {
		const sht: any = { name: escapeXml(wb.SheetNames[i].slice(0, 31)) };
		// sheetId is 1-based
		sht.sheetId = "" + (i + 1);
		sht["r:id"] = "rId" + (i + 1);
		if (sheets[i]) {
			switch (sheets[i].Hidden) {
				case 1:
					sht.state = "hidden";
					break;
				case 2:
					sht.state = "veryHidden";
					break;
			}
		}
		lines.push(writeXmlElement("sheet", null, sht));
	}
	lines.push("</sheets>");

	// Defined names
	if (write_names) {
		lines.push("<definedNames>");
		if (wb.Workbook && wb.Workbook.Names) {
			wb.Workbook.Names.forEach((n) => {
				const d: any = { name: n.Name };
				if (n.Comment) {
					d.comment = n.Comment;
				}
				if (n.Sheet != null) {
					d.localSheetId = "" + n.Sheet;
				}
				if (n.Hidden) {
					d.hidden = "1";
				}
				if (!n.Ref) {
					return;
				}
				lines.push(writeXmlElement("definedName", escapeXml(n.Ref), d));
			});
		}
		lines.push("</definedNames>");
	}

	if (lines.length > 2) {
		lines.push("</workbook>");
		// Convert self-closing <workbook .../> to opening tag <workbook ...>
		lines[1] = lines[1].replace("/>", ">");
	}
	return lines.join("");
}
