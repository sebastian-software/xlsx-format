import type { WorkBook } from "../types.js";
import type { ZipArchive } from "../zip/index.js";
import type { Relationships } from "../opc/relationships.js";
import { zip_new, zip_add_str } from "../zip/index.js";
import { new_ct, write_ct } from "../opc/content-types.js";
import { write_rels, add_rels, get_rels_path } from "../opc/relationships.js";
import { write_core_props } from "../opc/core-properties.js";
import { write_ext_props } from "../opc/extended-properties.js";
import { write_cust_props } from "../opc/custom-properties.js";
import { write_sst_xml } from "./shared-strings.js";
import { write_sty_xml } from "./styles.js";
import { write_theme_xml } from "./theme.js";
import { write_wb_xml } from "./workbook.js";
import { write_ws_xml } from "./worksheet.js";
import { write_comments_xml, write_tcmnt_xml, write_people_xml } from "./comments.js";
import { write_vml } from "./vml.js";
import { write_xlmeta_xml } from "./metadata.js";
import { make_ssf, table_fmt, SSF_load_table } from "../ssf/table.js";
import { dup, keys } from "../utils/helpers.js";
import { RELS as RELTYPE } from "../xml/namespaces.js";

/** Write a WorkBook to a ZIP archive (XLSX format) */
export function write_zip_xlsx(wb: WorkBook, opts: any): ZipArchive {
	if (wb && !(wb as any).SSF) {
		(wb as any).SSF = dup(table_fmt);
	}
	if (wb && (wb as any).SSF) {
		make_ssf();
		SSF_load_table((wb as any).SSF);
	}

	opts.rels = { "!id": {} } as any;
	opts.wbrels = { "!id": {} } as any;
	opts.Strings = [] as any;
	opts.Strings.Count = 0;
	opts.Strings.Unique = 0;
	opts.revStrings = new Map();

	const ct = new_ct();
	const zip = zip_new();
	let f = "";

	opts.cellXfs = [];

	if (!wb.Props) {
		wb.Props = {};
	}

	// Core properties
	f = "docProps/core.xml";
	zip_add_str(zip, f, write_core_props(wb.Props, opts));
	ct.coreprops.push(f);
	add_rels(opts.rels, 2, f, RELTYPE.CORE_PROPS);

	// Extended properties
	f = "docProps/app.xml";
	if (wb.Props && (wb.Props as any).SheetNames) {
		/* already set */
	} else if (!wb.Workbook || !wb.Workbook.Sheets) {
		(wb.Props as any).SheetNames = wb.SheetNames;
	} else {
		const _sn: string[] = [];
		for (let _i = 0; _i < wb.SheetNames.length; ++_i) {
			if ((wb.Workbook.Sheets[_i] || ({} as any)).Hidden !== 2) {
				_sn.push(wb.SheetNames[_i]);
			}
		}
		(wb.Props as any).SheetNames = _sn;
	}
	(wb.Props as any).Worksheets = (wb.Props as any).SheetNames.length;
	zip_add_str(zip, f, write_ext_props(wb.Props));
	ct.extprops.push(f);
	add_rels(opts.rels, 3, f, RELTYPE.EXT_PROPS);

	// Custom properties
	if (wb.Custprops !== wb.Props && keys(wb.Custprops || {}).length > 0) {
		f = "docProps/custom.xml";
		zip_add_str(zip, f, write_cust_props(wb.Custprops));
		ct.custprops.push(f);
		add_rels(opts.rels, 4, f, RELTYPE.CUST_PROPS);
	}

	const people: string[] = ["SheetJ5"];
	opts.tcid = 0;

	// Sheets
	for (let rId = 1; rId <= wb.SheetNames.length; ++rId) {
		const wsrels: Relationships = { "!id": {} } as any;
		const ws = wb.Sheets[wb.SheetNames[rId - 1]];

		f = "xl/worksheets/sheet" + rId + ".xml";
		zip_add_str(zip, f, write_ws_xml(ws || ({} as any), opts, rId - 1, wsrels, wb));
		ct.sheets.push(f);
		add_rels(opts.wbrels, -1, "worksheets/sheet" + rId + ".xml", RELTYPE.SHEET);

		if (ws) {
			const comments = (ws as any)["!comments"];
			let need_vml = false;

			if (comments && comments.length > 0) {
				let needtc = false;
				comments.forEach((carr: any) => {
					carr[1].forEach((c: any) => {
						if (c.T === true) {
							needtc = true;
						}
					});
				});

				if (needtc) {
					const cf = "xl/threadedComments/threadedComment" + rId + ".xml";
					zip_add_str(zip, cf, write_tcmnt_xml(comments, people, opts));
					ct.threadedcomments.push(cf);
					add_rels(wsrels, -1, "../threadedComments/threadedComment" + rId + ".xml", RELTYPE.TCMNT);
				}

				const cf2 = "xl/comments" + rId + ".xml";
				zip_add_str(zip, cf2, write_comments_xml(comments));
				ct.comments.push(cf2);
				add_rels(wsrels, -1, "../comments" + rId + ".xml", RELTYPE.CMNT);
				need_vml = true;
			}

			if ((ws as any)["!legacy"]) {
				if (need_vml) {
					zip_add_str(zip, "xl/drawings/vmlDrawing" + rId + ".vml", write_vml(rId, (ws as any)["!comments"]));
				}
			}

			delete (ws as any)["!comments"];
			delete (ws as any)["!legacy"];
		}

		if ((wsrels["!id"] as any).rId1) {
			zip_add_str(zip, get_rels_path(f), write_rels(wsrels));
		}
	}

	// Shared strings
	if (opts.Strings != null && opts.Strings.length > 0) {
		f = "xl/sharedStrings.xml";
		zip_add_str(zip, f, write_sst_xml(opts.Strings, opts));
		ct.strs.push(f);
		add_rels(opts.wbrels, -1, "sharedStrings.xml", RELTYPE.SST);
	}

	// Workbook
	f = "xl/workbook.xml";
	zip_add_str(zip, f, write_wb_xml(wb));
	ct.workbooks.push(f);
	add_rels(opts.rels, 1, f, RELTYPE.WB);

	// Theme
	f = "xl/theme/theme1.xml";
	zip_add_str(zip, f, write_theme_xml());
	ct.themes.push(f);
	add_rels(opts.wbrels, -1, "theme/theme1.xml", RELTYPE.THEME);

	// Styles
	f = "xl/styles.xml";
	zip_add_str(zip, f, write_sty_xml(wb, opts));
	ct.styles.push(f);
	add_rels(opts.wbrels, -1, "styles.xml", RELTYPE.STY);

	// Metadata
	f = "xl/metadata.xml";
	zip_add_str(zip, f, write_xlmeta_xml());
	ct.metadata.push(f);
	add_rels(opts.wbrels, -1, "metadata.xml", RELTYPE.META);

	// People (threaded comments)
	if (people.length > 1) {
		f = "xl/persons/person.xml";
		zip_add_str(zip, f, write_people_xml(people));
		ct.people.push(f);
		add_rels(opts.wbrels, -1, "persons/person.xml", RELTYPE.PEOPLE);
	}

	// Content types and relationships
	zip_add_str(zip, "[Content_Types].xml", write_ct(ct, opts));
	zip_add_str(zip, "_rels/.rels", write_rels(opts.rels));
	zip_add_str(zip, "xl/_rels/workbook.xml.rels", write_rels(opts.wbrels));

	return zip;
}
