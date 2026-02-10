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
	let filePath = "";

	opts.cellXfs = [];

	if (!wb.Props) {
		wb.Props = {};
	}

	// Core properties
	filePath = "docProps/core.xml";
	zipAddString(zip, filePath, writeCoreProperties(wb.Props, opts));
	ct.coreprops.push(filePath);
	addRelationship(opts.rels, 2, filePath, RELTYPE.CORE_PROPS);

	// Extended properties
	filePath = "docProps/app.xml";
	if (wb.Props && (wb.Props as any).SheetNames) {
		/* already set */
	} else if (!wb.Workbook || !wb.Workbook.Sheets) {
		(wb.Props as any).SheetNames = wb.SheetNames;
	} else {
		const visibleSheetNames: string[] = [];
		for (let sheetIdx = 0; sheetIdx < wb.SheetNames.length; ++sheetIdx) {
			if ((wb.Workbook.Sheets[sheetIdx] || ({} as any)).Hidden !== 2) {
				visibleSheetNames.push(wb.SheetNames[sheetIdx]);
			}
		}
		(wb.Props as any).SheetNames = visibleSheetNames;
	}
	(wb.Props as any).Worksheets = (wb.Props as any).SheetNames.length;
	zipAddString(zip, filePath, writeExtendedProperties(wb.Props));
	ct.extprops.push(filePath);
	addRelationship(opts.rels, 3, filePath, RELTYPE.EXT_PROPS);

	// Custom properties
	if (wb.Custprops !== wb.Props && Object.keys(wb.Custprops || {}).length > 0) {
		filePath = "docProps/custom.xml";
		zipAddString(zip, filePath, writeCustomProperties(wb.Custprops));
		ct.custprops.push(filePath);
		addRelationship(opts.rels, 4, filePath, RELTYPE.CUST_PROPS);
	}

	const people: string[] = ["SheetJ5"];
	opts.tcid = 0;

	// Sheets
	for (let rId = 1; rId <= wb.SheetNames.length; ++rId) {
		const wsrels: Relationships = { "!id": {} } as any;
		const ws = wb.Sheets[wb.SheetNames[rId - 1]];

		filePath = "xl/worksheets/sheet" + rId + ".xml";
		zipAddString(zip, filePath, writeWorksheetXml(ws || ({} as any), opts, rId - 1, wsrels, wb));
		ct.sheets.push(filePath);
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
			zipAddString(zip, getRelsPath(filePath), writeRelationships(wsrels));
		}
	}

	// Shared strings
	if (opts.Strings != null && opts.Strings.length > 0) {
		filePath = "xl/sharedStrings.xml";
		zipAddString(zip, filePath, writeSstXml(opts.Strings, opts));
		ct.strs.push(filePath);
		addRelationship(opts.wbrels, -1, "sharedStrings.xml", RELTYPE.SST);
	}

	// Workbook
	filePath = "xl/workbook.xml";
	zipAddString(zip, filePath, writeWorkbookXml(wb));
	ct.workbooks.push(filePath);
	addRelationship(opts.rels, 1, filePath, RELTYPE.WB);

	// Theme
	filePath = "xl/theme/theme1.xml";
	zipAddString(zip, filePath, write_theme_xml());
	ct.themes.push(filePath);
	addRelationship(opts.wbrels, -1, "theme/theme1.xml", RELTYPE.THEME);

	// Styles
	filePath = "xl/styles.xml";
	zipAddString(zip, filePath, writeStylesXml(wb, opts));
	ct.styles.push(filePath);
	addRelationship(opts.wbrels, -1, "styles.xml", RELTYPE.STY);

	// Metadata
	filePath = "xl/metadata.xml";
	zipAddString(zip, filePath, writeMetadataXml());
	ct.metadata.push(filePath);
	addRelationship(opts.wbrels, -1, "metadata.xml", RELTYPE.META);

	// People (threaded comments)
	if (people.length > 1) {
		filePath = "xl/persons/person.xml";
		zipAddString(zip, filePath, writePeopleXml(people));
		ct.people.push(filePath);
		addRelationship(opts.wbrels, -1, "persons/person.xml", RELTYPE.PEOPLE);
	}

	// Content types and relationships
	zipAddString(zip, "[Content_Types].xml", writeContentTypes(ct, opts));
	zipAddString(zip, "_rels/.rels", writeRelationships(opts.rels));
	zipAddString(zip, "xl/_rels/workbook.xml.rels", writeRelationships(opts.wbrels));

	return zip;
}
