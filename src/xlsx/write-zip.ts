import type { WorkBook } from "../types.js";
import type { ZipArchive } from "../zip/index.js";
import type { Relationships } from "../opc/relationships.js";
import { zipCreate, zipAddString } from "../zip/index.js";
import { createContentTypes, writeContentTypes } from "../opc/content-types.js";
import { writeRelationships, addRelationship, getRelsPath } from "../opc/relationships.js";
import { writeCoreProperties } from "../opc/core-properties.js";
import { writeExtendedProperties } from "../opc/extended-properties.js";
import { writeCustomProperties } from "../opc/custom-properties.js";
import { writeSstXml } from "./shared-strings.js";
import { writeStylesXml } from "./styles.js";
import { write_theme_xml } from "./theme.js";
import { writeWorkbookXml } from "./workbook.js";
import { writeWorksheetXml } from "./worksheet.js";
import { writeCommentsXml, writeTcmntXml, writePeopleXml } from "./comments.js";
import { writeVml } from "./vml.js";
import { writeMetadataXml } from "./metadata.js";
import { resetFormatTable, formatTable, loadFormatTable } from "../ssf/table.js";
import { RELS as RELTYPE } from "../xml/namespaces.js";

/** Write a WorkBook to a ZIP archive (XLSX format) */
export function writeZipXlsx(wb: WorkBook, opts: any): ZipArchive {
	if (wb && !(wb as any).SSF) {
		(wb as any).SSF = { ...formatTable };
	}
	if (wb && (wb as any).SSF) {
		resetFormatTable();
		loadFormatTable((wb as any).SSF);
	}

	opts.rels = { "!id": {} } as any;
	opts.wbrels = { "!id": {} } as any;
	opts.Strings = [] as any;
	opts.Strings.Count = 0;
	opts.Strings.Unique = 0;
	opts.revStrings = new Map();

	const ct = createContentTypes();
	const zip = zipCreate();
	let f = "";

	opts.cellXfs = [];

	if (!wb.Props) {
		wb.Props = {};
	}

	// Core properties
	f = "docProps/core.xml";
	zipAddString(zip, f, writeCoreProperties(wb.Props, opts));
	ct.coreprops.push(f);
	addRelationship(opts.rels, 2, f, RELTYPE.CORE_PROPS);

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
	zipAddString(zip, f, writeExtendedProperties(wb.Props));
	ct.extprops.push(f);
	addRelationship(opts.rels, 3, f, RELTYPE.EXT_PROPS);

	// Custom properties
	if (wb.Custprops !== wb.Props && Object.keys(wb.Custprops || {}).length > 0) {
		f = "docProps/custom.xml";
		zipAddString(zip, f, writeCustomProperties(wb.Custprops));
		ct.custprops.push(f);
		addRelationship(opts.rels, 4, f, RELTYPE.CUST_PROPS);
	}

	const people: string[] = ["SheetJ5"];
	opts.tcid = 0;

	// Sheets
	for (let rId = 1; rId <= wb.SheetNames.length; ++rId) {
		const wsrels: Relationships = { "!id": {} } as any;
		const ws = wb.Sheets[wb.SheetNames[rId - 1]];

		f = "xl/worksheets/sheet" + rId + ".xml";
		zipAddString(zip, f, writeWorksheetXml(ws || ({} as any), opts, rId - 1, wsrels, wb));
		ct.sheets.push(f);
		addRelationship(opts.wbrels, -1, "worksheets/sheet" + rId + ".xml", RELTYPE.SHEET);

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
					zipAddString(zip, cf, writeTcmntXml(comments, people, opts));
					ct.threadedcomments.push(cf);
					addRelationship(wsrels, -1, "../threadedComments/threadedComment" + rId + ".xml", RELTYPE.TCMNT);
				}

				const cf2 = "xl/comments" + rId + ".xml";
				zipAddString(zip, cf2, writeCommentsXml(comments));
				ct.comments.push(cf2);
				addRelationship(wsrels, -1, "../comments" + rId + ".xml", RELTYPE.CMNT);
				need_vml = true;
			}

			if ((ws as any)["!legacy"]) {
				if (need_vml) {
					zipAddString(zip, "xl/drawings/vmlDrawing" + rId + ".vml", writeVml(rId, (ws as any)["!comments"]));
				}
			}

			delete (ws as any)["!comments"];
			delete (ws as any)["!legacy"];
		}

		if ((wsrels["!id"] as any).rId1) {
			zipAddString(zip, getRelsPath(f), writeRelationships(wsrels));
		}
	}

	// Shared strings
	if (opts.Strings != null && opts.Strings.length > 0) {
		f = "xl/sharedStrings.xml";
		zipAddString(zip, f, writeSstXml(opts.Strings, opts));
		ct.strs.push(f);
		addRelationship(opts.wbrels, -1, "sharedStrings.xml", RELTYPE.SST);
	}

	// Workbook
	f = "xl/workbook.xml";
	zipAddString(zip, f, writeWorkbookXml(wb));
	ct.workbooks.push(f);
	addRelationship(opts.rels, 1, f, RELTYPE.WB);

	// Theme
	f = "xl/theme/theme1.xml";
	zipAddString(zip, f, write_theme_xml());
	ct.themes.push(f);
	addRelationship(opts.wbrels, -1, "theme/theme1.xml", RELTYPE.THEME);

	// Styles
	f = "xl/styles.xml";
	zipAddString(zip, f, writeStylesXml(wb, opts));
	ct.styles.push(f);
	addRelationship(opts.wbrels, -1, "styles.xml", RELTYPE.STY);

	// Metadata
	f = "xl/metadata.xml";
	zipAddString(zip, f, writeMetadataXml());
	ct.metadata.push(f);
	addRelationship(opts.wbrels, -1, "metadata.xml", RELTYPE.META);

	// People (threaded comments)
	if (people.length > 1) {
		f = "xl/persons/person.xml";
		zipAddString(zip, f, writePeopleXml(people));
		ct.people.push(f);
		addRelationship(opts.wbrels, -1, "persons/person.xml", RELTYPE.PEOPLE);
	}

	// Content types and relationships
	zipAddString(zip, "[Content_Types].xml", writeContentTypes(ct, opts));
	zipAddString(zip, "_rels/.rels", writeRelationships(opts.rels));
	zipAddString(zip, "xl/_rels/workbook.xml.rels", writeRelationships(opts.wbrels));

	return zip;
}
