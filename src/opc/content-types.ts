import { parseXmlTag, XML_HEADER, XML_TAG_REGEX } from "../xml/parser.js";
import { writeXmlElement } from "../xml/writer.js";
import { XMLNS } from "../xml/namespaces.js";

// Extracts the namespace prefix from the first tag in a string (e.g., "w" from "<w:Types>")
const nsregex = /<(\w+):/;

/**
 * Map from OOXML content-type MIME strings to internal category names.
 * Used during parsing to classify each Override entry into the correct bucket.
 */
const CONTENT_TYPE_MAP: Record<string, string> = {
	"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml": "workbooks",
	"application/vnd.ms-excel.sheet.macroEnabled.main+xml": "workbooks",
	"application/vnd.ms-excel.sheet.binary.macroEnabled.main": "workbooks",
	"application/vnd.ms-excel.addin.macroEnabled.main+xml": "workbooks",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml": "workbooks",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml": "sheets",
	"application/vnd.ms-excel.worksheet": "sheets",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml": "charts",
	"application/vnd.ms-excel.chartsheet": "charts",
	"application/vnd.ms-excel.macrosheet+xml": "macros",
	"application/vnd.ms-excel.macrosheet": "macros",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml": "dialogs",
	"application/vnd.ms-excel.dialogsheet": "dialogs",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml": "strs",
	"application/vnd.ms-excel.sharedStrings": "strs",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml": "styles",
	"application/vnd.ms-excel.styles": "styles",
	"application/vnd.openxmlformats-package.core-properties+xml": "coreprops",
	"application/vnd.openxmlformats-officedocument.custom-properties+xml": "custprops",
	"application/vnd.openxmlformats-officedocument.extended-properties+xml": "extprops",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml": "comments",
	"application/vnd.ms-excel.comments": "comments",
	"application/vnd.ms-excel.threadedcomments+xml": "threadedcomments",
	"application/vnd.ms-excel.person+xml": "people",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml": "metadata",
	"application/vnd.ms-excel.sheetMetadata": "metadata",
	"application/vnd.ms-excel.calcChain": "calcchains",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml": "calcchains",
	"application/vnd.openxmlformats-officedocument.theme+xml": "themes",
	"application/vnd.ms-office.vbaProject": "vba",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml": "links",
	"application/vnd.ms-excel.externalLink": "links",
	"application/vnd.openxmlformats-officedocument.drawing+xml": "drawings",
	"application/vnd.openxmlformats-package.relationships+xml": "rels",
};

/**
 * Reverse lookup: maps internal category names to the preferred content-type strings
 * for writing, keyed by book type (e.g., "xlsx", "xlsm").
 */
const CT_LIST: Record<string, Record<string, string>> = {
	workbooks: {
		xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
		xlsm: "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
	},
	strs: {
		xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
	},
	comments: {
		xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml",
	},
	sheets: {
		xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
	},
	charts: {
		xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml",
	},
	dialogs: {
		xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml",
	},
	macros: {
		xlsx: "application/vnd.ms-excel.macrosheet+xml",
	},
	metadata: {
		xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml",
	},
	styles: {
		xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
	},
};

/** Parsed representation of the [Content_Types].xml file, with parts grouped by category */
export interface ContentTypes {
	workbooks: string[];
	sheets: string[];
	charts: string[];
	dialogs: string[];
	macros: string[];
	rels: string[];
	strs: string[];
	comments: string[];
	threadedcomments: string[];
	links: string[];
	coreprops: string[];
	extprops: string[];
	custprops: string[];
	themes: string[];
	styles: string[];
	calcchains: string[];
	vba: string[];
	drawings: string[];
	metadata: string[];
	people: string[];
	xmlns: string;
	/** Shortcut to the first calc chain part path */
	calcchain?: string;
	/** Shortcut to the first shared strings part path */
	sst?: string;
	/** Shortcut to the first styles part path */
	style?: string;
	/** Default content types by file extension (from <Default> elements) */
	defaults?: Record<string, string>;
}

/**
 * Create an empty ContentTypes object with all category arrays initialized.
 * @returns a fresh ContentTypes with empty arrays for every category
 */
export function createContentTypes(): ContentTypes {
	return {
		workbooks: [],
		sheets: [],
		charts: [],
		dialogs: [],
		macros: [],
		rels: [],
		strs: [],
		comments: [],
		threadedcomments: [],
		links: [],
		coreprops: [],
		extprops: [],
		custprops: [],
		themes: [],
		styles: [],
		calcchains: [],
		vba: [],
		drawings: [],
		metadata: [],
		people: [],
		xmlns: "",
	};
}

/**
 * Parse the [Content_Types].xml file from an OPC package.
 * Processes `<Default>` entries (file extension -> content type) and
 * `<Override>` entries (part path -> content type), classifying each
 * override into the appropriate category array.
 * @param data - raw XML string of the [Content_Types].xml file (may be null/undefined)
 * @returns a populated ContentTypes object
 * @throws if the root namespace is not the expected OPC content-types namespace
 */
export function parseContentTypes(data: string | null | undefined): ContentTypes {
	const ct = createContentTypes();
	if (!data) {
		return ct;
	}
	const ctext: Record<string, string> = {};
	const matches = data.match(XML_TAG_REGEX) || [];
	for (const x of matches) {
		const y = parseXmlTag(x);
		// Strip namespace prefix from tag name for uniform matching
		switch ((y[0] as string).replace(nsregex, "<")) {
			case "<?xml":
				break;
			case "<Types":
				// Extract the xmlns attribute, accounting for possible namespace prefix on the tag itself
				ct.xmlns = y["xmlns" + ((y[0] as string).match(/<(\w+):/) || ["", ""])[1]];
				break;
			case "<Default":
				// Map file extension (lowercased) to its content type
				ctext[y.Extension.toLowerCase()] = y.ContentType;
				break;
			case "<Override":
				// Classify the part into the correct category based on its content type
				if (CONTENT_TYPE_MAP[y.ContentType] && (ct as any)[CONTENT_TYPE_MAP[y.ContentType]] !== undefined) {
					(ct as any)[CONTENT_TYPE_MAP[y.ContentType]].push(y.PartName);
				}
				break;
		}
	}
	if (ct.xmlns !== XMLNS.CT) {
		throw new Error("Unknown Namespace: " + ct.xmlns);
	}
	// Set convenience shortcuts to the first entry in commonly-used categories
	ct.calcchain = ct.calcchains.length > 0 ? ct.calcchains[0] : "";
	ct.sst = ct.strs.length > 0 ? ct.strs[0] : "";
	ct.style = ct.styles.length > 0 ? ct.styles[0] : "";
	ct.defaults = ctext;
	return ct;
}

/**
 * Build a reverse map from category name to an array of content-type strings.
 * @param obj - the forward map (content-type -> category)
 * @returns a record mapping each category to all its content-type strings
 */
function invertToArrayMap(obj: Record<string, string>): Record<string, string[]> {
	const o: Record<string, string[]> = {};
	for (const [k, v] of Object.entries(obj)) {
		if (!o[v]) {
			o[v] = [];
		}
		o[v].push(k);
	}
	return o;
}

/**
 * Serialize a ContentTypes object to [Content_Types].xml format.
 * Emits `<Default>` entries for file extensions and `<Override>` entries
 * for each registered part, choosing the correct content-type string
 * based on the target book type.
 * @param ct - the ContentTypes object to serialize
 * @param opts - options containing the target bookType (e.g., "xlsx", "xlsm")
 * @returns the complete XML string for [Content_Types].xml
 */
export function writeContentTypes(ct: ContentTypes, opts: { bookType?: string }): string {
	const type2ct = invertToArrayMap(CONTENT_TYPE_MAP);
	const o: string[] = [];

	o.push(XML_HEADER);
	o.push(
		writeXmlElement("Types", null, {
			xmlns: XMLNS.CT,
			"xmlns:xsd": XMLNS.xsd,
			"xmlns:xsi": XMLNS.xsi,
		}),
	);

	// Default content types by file extension
	const defaults: [string, string][] = [
		["xml", "application/xml"],
		["bin", "application/vnd.ms-excel.sheet.binary.macroEnabled.main"],
		["vml", "application/vnd.openxmlformats-officedocument.vmlDrawing"],
		["data", "application/vnd.openxmlformats-officedocument.model+data"],
		["bmp", "image/bmp"],
		["png", "image/png"],
		["gif", "image/gif"],
		["emf", "image/x-emf"],
		["wmf", "image/x-wmf"],
		["jpg", "image/jpeg"],
		["jpeg", "image/jpeg"],
		["tif", "image/tiff"],
		["tiff", "image/tiff"],
		["pdf", "application/pdf"],
		["rels", "application/vnd.openxmlformats-package.relationships+xml"],
	];

	for (const [ext, contentType] of defaults) {
		o.push(writeXmlElement("Default", null, { Extension: ext, ContentType: contentType }));
	}

	// f1: write an Override for only the first entry in a category (singleton parts like workbook, SST, styles)
	const f1 = (w: string) => {
		if ((ct as any)[w] && (ct as any)[w].length > 0) {
			const v = (ct as any)[w][0];
			o.push(
				writeXmlElement("Override", null, {
					PartName: (v[0] === "/" ? "" : "/") + v,
					ContentType: CT_LIST[w]?.[opts.bookType || "xlsx"] || CT_LIST[w]?.["xlsx"],
				}),
			);
		}
	};

	// f2: write Overrides for every entry in a category (multi-instance parts like sheets, charts)
	const f2 = (w: string) => {
		for (const v of (ct as any)[w] || []) {
			o.push(
				writeXmlElement("Override", null, {
					PartName: (v[0] === "/" ? "" : "/") + v,
					ContentType: CT_LIST[w]?.[opts.bookType || "xlsx"] || CT_LIST[w]?.["xlsx"],
				}),
			);
		}
	};

	// f3: write Overrides using the reverse content-type map (for types not in CT_LIST)
	const f3 = (t: string) => {
		for (const v of (ct as any)[t] || []) {
			o.push(
				writeXmlElement("Override", null, {
					PartName: (v[0] === "/" ? "" : "/") + v,
					ContentType: type2ct[t]?.[0],
				}),
			);
		}
	};

	// Emit overrides in a specific order matching typical XLSX file structure
	f1("workbooks");
	f2("sheets");
	f2("charts");
	f3("themes");
	f1("strs");
	f1("styles");
	f3("coreprops");
	f3("extprops");
	f3("custprops");
	f3("vba");
	f3("comments");
	f3("threadedcomments");
	f3("drawings");
	f2("metadata");
	f3("people");

	// Only close the root element if child elements were added
	if (o.length > 2) {
		o.push("</Types>");
		// Convert the self-closing root tag to an opening tag
		o[1] = o[1].replace("/>", ">");
	}
	return o.join("");
}
