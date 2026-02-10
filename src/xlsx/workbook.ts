import { parseXmlTag, XML_TAG_REGEX, XML_HEADER, stripNamespace, parseXmlBoolean } from "../xml/parser.js";
import { unescapeXml, escapeXml } from "../xml/escape.js";
import { writeXmlElement } from "../xml/writer.js";
import { XMLNS_main, XMLNS } from "../xml/namespaces.js";
import { utf8read } from "../utils/buffer.js";
import type { WorkBook } from "../types.js";

export interface WorkbookFile {
	AppVersion: Record<string, any>;
	WBProps: Record<string, any>;
	WBView: Record<string, any>[];
	Sheets: SheetEntry[];
	CalcPr: Record<string, any>;
	Names: DefinedNameEntry[];
	xmlns: string;
}

export interface SheetEntry {
	name: string;
	sheetId: string;
	Hidden?: number;
	[key: string]: any;
}

export interface DefinedNameEntry {
	Name: string;
	Ref?: string;
	Sheet?: number;
	Comment?: string;
	Hidden?: boolean;
}

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

const SheetDef: [string, any, string?][] = [];

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

export function parse_wb_defaults(wb: WorkbookFile): void {
	applyDefaults(wb.WBProps, WBPropsDef);
	applyDefaults(wb.CalcPr, CalcPrDef);
	applyDefaultsToArray(wb.WBView, WBViewDef);
	applyDefaultsToArray(wb.Sheets, SheetDef);
}

export function is1904DateSystem(wb: WorkBook): string {
	if (!wb.Workbook) {
		return "false";
	}
	if (!wb.Workbook.WBProps) {
		return "false";
	}
	return parseXmlBoolean(wb.Workbook.WBProps.date1904) ? "true" : "false";
}

const badchars = ":][*?/\\".split("");

export function validateSheetName(n: string, safe?: boolean): boolean {
	try {
		if (n === "") {
			throw new Error("Sheet name cannot be blank");
		}
		if (n.length > 31) {
			throw new Error("Sheet name cannot exceed 31 chars");
		}
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

export function validateWorkbookNames(sheetNames: string[], sheetEntries?: any[]): void {
	for (let i = 0; i < sheetNames.length; ++i) {
		validateSheetName(sheetNames[i]);
		for (let j = 0; j < i; ++j) {
			if (sheetNames[i] === sheetNames[j]) {
				throw new Error("Duplicate Sheet Name: " + sheetNames[i]);
			}
		}
	}
}

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

const wbnsregex = /<\w+:workbook/;

/** Parse a workbook XML file */
export function parseWorkbookXml(data: string, opts?: any): WorkbookFile {
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
	let pass = false;
	let xmlns = "xmlns";
	let dname: any = {};
	let dnstart = 0;

	data.replace(XML_TAG_REGEX, function xml_wb(xmlTag: string, idx: number): string {
		const parsedTag: any = parseXmlTag(xmlTag);
		switch (stripNamespace(parsedTag[0])) {
			case "<?xml":
				break;

			case "<workbook":
				if (xmlTag.match(wbnsregex)) {
					xmlns = "xmlns" + xmlTag.match(/<(\w+):/)?.[1];
				}
				workbook.xmlns = parsedTag[xmlns];
				break;
			case "</workbook>":
				break;

			case "<fileVersion":
				delete parsedTag[0];
				workbook.AppVersion = parsedTag;
				break;
			case "<fileVersion/>":
			case "</fileVersion>":
				break;

			case "<fileSharing":
			case "<fileSharing/>":
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
			case "</workbookPr>":
				break;

			case "<workbookProtection":
			case "<workbookProtection/>":
				break;

			case "<bookViews":
			case "<bookViews>":
			case "</bookViews>":
				break;
			case "<workbookView":
			case "<workbookView/>":
				delete parsedTag[0];
				workbook.WBView.push(parsedTag);
				break;
			case "</workbookView>":
				break;

			case "<sheets":
			case "<sheets>":
			case "</sheets>":
				break;
			case "<sheet":
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
			case "</sheet>":
				break;

			case "<functionGroups":
			case "<functionGroups/>":
			case "<functionGroup":
				break;

			case "<externalReferences":
			case "</externalReferences>":
			case "<externalReferences>":
			case "<externalReference":
				break;

			case "<definedNames/>":
				break;
			case "<definedNames>":
			case "<definedNames":
				pass = true;
				break;
			case "</definedNames>":
				pass = false;
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
				dnstart = idx + xmlTag.length;
				break;
			}
			case "</definedName>": {
				dname.Ref = unescapeXml(utf8read(data.slice(dnstart, idx)));
				workbook.Names.push(dname);
				break;
			}
			case "<definedName/>":
				break;

			case "<calcPr":
			case "<calcPr/>":
				delete parsedTag[0];
				workbook.CalcPr = parsedTag;
				break;
			case "</calcPr>":
				break;

			case "<oleSize":
				break;

			case "<customWorkbookViews>":
			case "</customWorkbookViews>":
			case "<customWorkbookViews":
			case "<customWorkbookView":
			case "</customWorkbookView>":
				break;

			case "<pivotCaches>":
			case "</pivotCaches>":
			case "<pivotCaches":
			case "<pivotCache":
				break;

			case "<smartTagPr":
			case "<smartTagPr/>":
				break;

			case "<smartTagTypes":
			case "<smartTagTypes>":
			case "</smartTagTypes>":
			case "<smartTagType":
				break;

			case "<webPublishing":
			case "<webPublishing/>":
				break;

			case "<fileRecoveryPr":
			case "<fileRecoveryPr/>":
				break;

			case "<webPublishObjects>":
			case "<webPublishObjects":
			case "</webPublishObjects>":
			case "<webPublishObject":
				break;

			case "<extLst":
			case "<extLst>":
			case "</extLst>":
			case "<extLst/>":
				break;
			case "<ext":
				pass = true;
				break;
			case "</ext>":
				pass = false;
				break;

			case "<ArchID":
				break;
			case "<AlternateContent":
			case "<AlternateContent>":
				pass = true;
				break;
			case "</AlternateContent>":
				pass = false;
				break;

			case "<revisionPtr":
				break;

			default:
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

/** Write the workbook XML */
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

	/* bookViews only written if first worksheet is hidden */
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

	lines.push("<sheets>");
	for (let i = 0; i < wb.SheetNames.length; ++i) {
		const sht: any = { name: escapeXml(wb.SheetNames[i].slice(0, 31)) };
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
		lines[1] = lines[1].replace("/>", ">");
	}
	return lines.join("");
}
