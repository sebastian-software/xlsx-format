import type { WorkBook, WorkSheet, ReadOptions, FullProperties } from "../types.js";
import type { ZipArchive } from "../zip/index.js";
import type { ContentTypes } from "../opc/content-types.js";
import type { Relationships } from "../opc/relationships.js";
import type { SST } from "./shared-strings.js";
import type { StylesData } from "./styles.js";
import type { ThemeData } from "./theme.js";
import type { WorkbookFile, SheetEntry } from "./workbook.js";
import { zip_read_str, zip_read_bin, zip_has } from "../zip/index.js";
import { parse_ct } from "../opc/content-types.js";
import { parse_rels, get_rels_path } from "../opc/relationships.js";
import { parse_core_props } from "../opc/core-properties.js";
import { parse_ext_props } from "../opc/extended-properties.js";
import { parse_cust_props } from "../opc/custom-properties.js";
import { parse_sst_xml } from "./shared-strings.js";
import { parse_sty_xml } from "./styles.js";
import { parse_theme_xml } from "./theme.js";
import { parse_wb_xml } from "./workbook.js";
import { parse_ws_xml, resolve_sst } from "./worksheet.js";
import { parse_comments_xml, parse_tcmnt_xml, parse_people_xml, sheet_insert_comments } from "./comments.js";
import { parse_vml } from "./vml.js";
import { parse_xlmeta_xml } from "./metadata.js";
import { parse_cc_xml } from "./calc-chain.js";
import { make_ssf, table_fmt } from "../ssf/table.js";
import { dup, keys } from "../utils/helpers.js";
import { utf8read } from "../utils/buffer.js";
import { RELS as RELTYPE } from "../xml/namespaces.js";

function strip_front_slash(x: string): string {
	return x.charAt(0) === "/" ? x.slice(1) : x;
}

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

function getzipstr(zip: ZipArchive, path: string, safe?: boolean): string | null {
	const p = zip_read_str(zip, path);
	if (p == null && !safe) {
		throw new Error("Could not find " + path);
	}
	return p;
}

function getzipdata(zip: ZipArchive, path: string, safe?: boolean): string | null {
	// For XML-based files we just read as string
	return getzipstr(zip, path, safe);
}

const RELS_WS = [
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
	"http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet",
];

function get_sheet_type(n: string): string {
	if (RELS_WS.indexOf(n) > -1) {
		return "sheet";
	}
	return n && n.length ? n : "sheet";
}

function safe_parse_wbrels(wbrels: Relationships, sheets: SheetEntry[]): [string, string, string][] | null {
	if (!wbrels) {
		return null;
	}
	try {
		const result: [string, string, string][] = sheets.map((w) => {
			const id = (w as any).id || (w as any).strRelID;
			return [w.name, wbrels["!id"][id].Target, get_sheet_type(wbrels["!id"][id].Type)];
		});
		return result.length === 0 ? null : result;
	} catch {
		return null;
	}
}

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
		sheetRels[sheetName] = parse_rels(getzipstr(zip, relsPath, true), path);
		const data = getzipdata(zip, path);
		if (!data) {
			return;
		}

		let _ws: WorkSheet | undefined;
		switch (stype) {
			case "sheet":
				_ws = parse_ws_xml(data, opts);
				break;
			default:
				return;
		}
		if (!_ws) {
			return;
		}

		// Resolve shared string references
		resolve_sst(_ws, strs, opts);

		sheets[sheetName] = _ws;

		// Scan rels for comments and threaded comments
		const comments: any[] = [];
		let tcomments: any[] = [];
		if (sheetRels[sheetName]) {
			for (const n of keys(sheetRels[sheetName])) {
				if (n === "!id" || n === "!idx") {
					continue;
				}
				const rel = sheetRels[sheetName][n];
				if (!rel || !rel.Type) {
					continue;
				}

				if (rel.Type === RELTYPE.CMNT) {
					const dfile = resolve_path(rel.Target, path);
					const cmntData = getzipdata(zip, dfile, true);
					if (cmntData) {
						const parsedComments = parse_comments_xml(cmntData, opts);
						if (parsedComments && parsedComments.length > 0) {
							sheet_insert_comments(_ws, parsedComments, false);
						}
					}
				}
				if (rel.Type === RELTYPE.TCMNT) {
					const dfile = resolve_path(rel.Target, path);
					const tcData = getzipdata(zip, dfile, true);
					if (tcData) {
						tcomments = tcomments.concat(parse_tcmnt_xml(tcData, opts));
					}
				}
			}
		}
		if (tcomments.length > 0) {
			sheet_insert_comments(_ws, tcomments, true, opts.people || []);
		}

		// Parse legacy drawings (VML for comment boxes)
		if ((_ws as any)["!legdrawel"] && sheetRels[sheetName]) {
			const dfile = resolve_path((_ws as any)["!legdrawel"].Target, path);
			const draw = getzipstr(zip, dfile, true);
			if (draw) {
				parse_vml(utf8read(draw), _ws, comments);
			}
		}
	} catch (e) {
		if (opts.WTF) {
			throw e;
		}
	}
}

/** Parse an XLSX ZIP archive into a WorkBook */
export function parse_zip(zip: ZipArchive, opts?: ReadOptions): WorkBook {
	make_ssf();
	const o: any = opts || {};

	if (!zip_has(zip, "[Content_Types].xml")) {
		throw new Error("Unsupported ZIP file");
	}

	const dir: ContentTypes = parse_ct(getzipstr(zip, "[Content_Types].xml"));

	if (dir.workbooks.length === 0) {
		const binname = "xl/workbook.xml";
		if (getzipdata(zip, binname, true)) {
			dir.workbooks.push(binname);
		}
	}
	if (dir.workbooks.length === 0) {
		throw new Error("Could not find workbook");
	}

	const themes: ThemeData = { themeElements: { clrScheme: [] } };
	let styles: StylesData = { NumberFmt: {}, CellXf: [], Fonts: [], Fills: [], Borders: [] };
	let strs: SST = [] as any;

	if (!o.bookSheets && !o.bookProps) {
		if (dir.sst) {
			try {
				const sstData = getzipdata(zip, strip_front_slash(dir.sst));
				if (sstData) {
					strs = parse_sst_xml(sstData, o);
				}
			} catch (e) {
				if (o.WTF) {
					throw e;
				}
			}
		}

		if (dir.themes.length) {
			const themeData = getzipstr(zip, dir.themes[0].replace(/^\//, ""), true);
			if (themeData) {
				const parsed = parse_theme_xml(themeData);
				Object.assign(themes, parsed);
			}
		}

		if (dir.style) {
			const styData = getzipdata(zip, strip_front_slash(dir.style));
			if (styData) {
				styles = parse_sty_xml(styData, themes, o);
			}
		}
	}

	const wb: WorkbookFile = parse_wb_xml(getzipdata(zip, strip_front_slash(dir.workbooks[0]))!, o);

	const props: any = {};
	if (dir.coreprops.length) {
		const propdata = getzipdata(zip, strip_front_slash(dir.coreprops[0]), true);
		if (propdata) {
			Object.assign(props, parse_core_props(propdata));
		}
		if (dir.extprops.length) {
			const extdata = getzipdata(zip, strip_front_slash(dir.extprops[0]), true);
			if (extdata) {
				parse_ext_props(extdata, props);
			}
		}
	}

	let custprops: Record<string, any> = {};
	if (!o.bookSheets || o.bookProps) {
		if (dir.custprops.length) {
			const custdata = getzipstr(zip, strip_front_slash(dir.custprops[0]), true);
			if (custdata) {
				custprops = parse_cust_props(custdata, o);
			}
		}
	}

	const out: any = {};
	if (o.bookSheets || o.bookProps) {
		let sheets: string[] | undefined;
		if (wb.Sheets) {
			sheets = wb.Sheets.map((x: SheetEntry) => x.name);
		} else if (props.Worksheets && props.SheetNames?.length > 0) {
			sheets = props.SheetNames;
		}
		if (o.bookProps) {
			out.Props = props;
			out.Custprops = custprops;
		}
		if (o.bookSheets && sheets) {
			out.SheetNames = sheets;
		}
		if (o.bookSheets ? out.SheetNames : o.bookProps) {
			return out as WorkBook;
		}
	}

	const sheets: Record<string, WorkSheet> = {};

	if (o.bookDeps && dir.calcchain) {
		parse_cc_xml(getzipdata(zip, strip_front_slash(dir.calcchain), true) || "");
	}

	const sheetRels: Record<string, Relationships> = {};
	const wbsheets = wb.Sheets;
	props.Worksheets = wbsheets.length;
	props.SheetNames = [];
	for (let j = 0; j < wbsheets.length; ++j) {
		props.SheetNames[j] = wbsheets[j].name;
	}

	const wbrelsi = dir.workbooks[0].lastIndexOf("/");
	let wbrelsfile = (
		dir.workbooks[0].slice(0, wbrelsi + 1) +
		"_rels/" +
		dir.workbooks[0].slice(wbrelsi + 1) +
		".rels"
	).replace(/^\//, "");
	if (!zip_has(zip, wbrelsfile)) {
		wbrelsfile = "xl/_rels/workbook.xml.rels";
	}
	const wbrels = parse_rels(getzipstr(zip, wbrelsfile, true), wbrelsfile.replace(/_rels.*/, "s5s"));

	// Parse metadata
	if ((dir.metadata || []).length >= 1) {
		o.xlmeta = parse_xlmeta_xml(getzipdata(zip, strip_front_slash(dir.metadata[0]), true) || "", o);
	}

	// Parse people (for threaded comments)
	if ((dir.people || []).length >= 1) {
		o.people = parse_people_xml(getzipdata(zip, strip_front_slash(dir.people[0]), true) || "");
	}

	const wbrelsArr = wbrels ? safe_parse_wbrels(wbrels, wb.Sheets) : null;

	const nmode = getzipdata(zip, "xl/worksheets/sheet.xml", true) ? 1 : 0;

	for (let i = 0; i < props.Worksheets; ++i) {
		let stype = "sheet";
		let path: string;
		if (wbrelsArr && wbrelsArr[i]) {
			path = "xl/" + wbrelsArr[i][1].replace(/[/]?xl\//, "");
			if (!zip_has(zip, path)) {
				path = wbrelsArr[i][1];
			}
			if (!zip_has(zip, path)) {
				path = wbrelsfile.replace(/_rels\/[\S\s]*$/, "") + wbrelsArr[i][1];
			}
			stype = wbrelsArr[i][2];
		} else {
			path = "xl/worksheets/sheet" + (i + 1 - nmode) + ".xml";
			path = path.replace(/sheet0\./, "sheet.");
		}

		// Check sheet filter
		if (o.sheets != null) {
			if (typeof o.sheets === "number" && i !== o.sheets) {
				continue;
			}
			if (typeof o.sheets === "string" && props.SheetNames[i].toLowerCase() !== o.sheets.toLowerCase()) {
				continue;
			}
			if (Array.isArray(o.sheets)) {
				let seen = false;
				for (const s of o.sheets) {
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
			o,
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

	if (wb.WBProps) {
		result.Workbook = {
			WBProps: wb.WBProps,
			Sheets: wb.Sheets as any,
			Names: wb.Names as any,
		};
	}

	return result;
}
