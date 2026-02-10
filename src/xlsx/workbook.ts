import { parsexmltag, tagregex, XML_HEADER, strip_ns, parsexmlbool } from "../xml/parser.js";
import { unescapexml, escapexml } from "../xml/escape.js";
import { writextag } from "../xml/writer.js";
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

function push_defaults_array(target: any[], defaults: [string, any, string?][]): void {
	for (let j = 0; j < target.length; ++j) {
		const w = target[j];
		for (let i = 0; i < defaults.length; ++i) {
			const z = defaults[i];
			if (w[z[0]] == null) {
				w[z[0]] = z[1];
			} else {
				switch (z[2]) {
					case "bool":
						if (typeof w[z[0]] === "string") {
							w[z[0]] = parsexmlbool(w[z[0]]);
						}
						break;
					case "int":
						if (typeof w[z[0]] === "string") {
							w[z[0]] = parseInt(w[z[0]], 10);
						}
						break;
				}
			}
		}
	}
}

function push_defaults(target: Record<string, any>, defaults: [string, any, string?][]): void {
	for (let i = 0; i < defaults.length; ++i) {
		const z = defaults[i];
		if (target[z[0]] == null) {
			target[z[0]] = z[1];
		} else {
			switch (z[2]) {
				case "bool":
					if (typeof target[z[0]] === "string") {
						target[z[0]] = parsexmlbool(target[z[0]]);
					}
					break;
				case "int":
					if (typeof target[z[0]] === "string") {
						target[z[0]] = parseInt(target[z[0]], 10);
					}
					break;
			}
		}
	}
}

export function parse_wb_defaults(wb: WorkbookFile): void {
	push_defaults(wb.WBProps, WBPropsDef);
	push_defaults(wb.CalcPr, CalcPrDef);
	push_defaults_array(wb.WBView, WBViewDef);
	push_defaults_array(wb.Sheets, SheetDef);
}

export function safe1904(wb: WorkBook): string {
	if (!wb.Workbook) {
		return "false";
	}
	if (!wb.Workbook.WBProps) {
		return "false";
	}
	return parsexmlbool(wb.Workbook.WBProps.date1904) ? "true" : "false";
}

const badchars = ":][*?/\\".split("");

export function check_ws_name(n: string, safe?: boolean): boolean {
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

export function check_wb_names(N: string[], S?: any[]): void {
	for (let i = 0; i < N.length; ++i) {
		check_ws_name(N[i]);
		for (let j = 0; j < i; ++j) {
			if (N[i] === N[j]) {
				throw new Error("Duplicate Sheet Name: " + N[i]);
			}
		}
	}
}

export function check_wb(wb: WorkBook): void {
	if (!wb || !wb.SheetNames || !wb.Sheets) {
		throw new Error("Invalid Workbook");
	}
	if (!wb.SheetNames.length) {
		throw new Error("Workbook is empty");
	}
	const Sheets = (wb.Workbook && wb.Workbook.Sheets) || [];
	check_wb_names(wb.SheetNames, Sheets);
}

const wbnsregex = /<\w+:workbook/;

/** Parse a workbook XML file */
export function parse_wb_xml(data: string, opts?: any): WorkbookFile {
	if (!data) {
		throw new Error("Could not find file");
	}
	const wb: WorkbookFile = {
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

	data.replace(tagregex, function xml_wb(x: string, idx: number): string {
		const y: any = parsexmltag(x);
		switch (strip_ns(y[0])) {
			case "<?xml":
				break;

			case "<workbook":
				if (x.match(wbnsregex)) {
					xmlns = "xmlns" + x.match(/<(\w+):/)?.[1];
				}
				wb.xmlns = y[xmlns];
				break;
			case "</workbook>":
				break;

			case "<fileVersion":
				delete y[0];
				wb.AppVersion = y;
				break;
			case "<fileVersion/>":
			case "</fileVersion>":
				break;

			case "<fileSharing":
			case "<fileSharing/>":
				break;

			case "<workbookPr":
			case "<workbookPr/>":
				WBPropsDef.forEach((w) => {
					if (y[w[0]] == null) {
						return;
					}
					switch (w[2]) {
						case "bool":
							wb.WBProps[w[0]] = parsexmlbool(y[w[0]]);
							break;
						case "int":
							wb.WBProps[w[0]] = parseInt(y[w[0]], 10);
							break;
						default:
							wb.WBProps[w[0]] = y[w[0]];
					}
				});
				if (y.codeName) {
					wb.WBProps.CodeName = utf8read(y.codeName);
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
				delete y[0];
				wb.WBView.push(y);
				break;
			case "</workbookView>":
				break;

			case "<sheets":
			case "<sheets>":
			case "</sheets>":
				break;
			case "<sheet":
				switch (y.state) {
					case "hidden":
						y.Hidden = 1;
						break;
					case "veryHidden":
						y.Hidden = 2;
						break;
					default:
						y.Hidden = 0;
				}
				delete y.state;
				y.name = unescapexml(utf8read(y.name));
				delete y[0];
				wb.Sheets.push(y);
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
				dname.Name = utf8read(y.name);
				if (y.comment) {
					dname.Comment = y.comment;
				}
				if (y.localSheetId) {
					dname.Sheet = +y.localSheetId;
				}
				if (parsexmlbool(y.hidden || "0")) {
					dname.Hidden = true;
				}
				dnstart = idx + x.length;
				break;
			}
			case "</definedName>": {
				dname.Ref = unescapexml(utf8read(data.slice(dnstart, idx)));
				wb.Names.push(dname);
				break;
			}
			case "<definedName/>":
				break;

			case "<calcPr":
			case "<calcPr/>":
				delete y[0];
				wb.CalcPr = y;
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
		return x;
	});

	if (XMLNS_main.indexOf(wb.xmlns) === -1) {
		throw new Error("Unknown Namespace: " + wb.xmlns);
	}

	parse_wb_defaults(wb);

	return wb;
}

/** Write the workbook XML */
export function write_wb_xml(wb: WorkBook): string {
	const o: string[] = [XML_HEADER];
	o.push(
		writextag("workbook", null, {
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
	o.push(writextag("workbookPr", null, workbookPr));

	const sheets = (wb.Workbook && wb.Workbook.Sheets) || [];

	/* bookViews only written if first worksheet is hidden */
	if (sheets[0] && !!sheets[0].Hidden) {
		o.push("<bookViews>");
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
		o.push('<workbookView firstSheet="' + i + '" activeTab="' + i + '"/>');
		o.push("</bookViews>");
	}

	o.push("<sheets>");
	for (let i = 0; i < wb.SheetNames.length; ++i) {
		const sht: any = { name: escapexml(wb.SheetNames[i].slice(0, 31)) };
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
		o.push(writextag("sheet", null, sht));
	}
	o.push("</sheets>");

	if (write_names) {
		o.push("<definedNames>");
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
				o.push(writextag("definedName", escapexml(n.Ref), d));
			});
		}
		o.push("</definedNames>");
	}

	if (o.length > 2) {
		o.push("</workbook>");
		o[1] = o[1].replace("/>", ">");
	}
	return o.join("");
}
