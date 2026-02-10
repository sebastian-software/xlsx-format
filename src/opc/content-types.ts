import { parseXmlTag, XML_HEADER, XML_TAG_REGEX } from "../xml/parser.js";
import { writeXmlElement } from "../xml/writer.js";
import { XMLNS } from "../xml/namespaces.js";

const nsregex = /<(\w+):/;

/** Map content types to internal categories */
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
	calcchain?: string;
	sst?: string;
	style?: string;
	defaults?: Record<string, string>;
}

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

export function parseContentTypes(data: string | null | undefined): ContentTypes {
	const ct = createContentTypes();
	if (!data) {
		return ct;
	}
	const ctext: Record<string, string> = {};
	const matches = data.match(XML_TAG_REGEX) || [];
	for (const x of matches) {
		const y = parseXmlTag(x);
		switch ((y[0] as string).replace(nsregex, "<")) {
			case "<?xml":
				break;
			case "<Types":
				ct.xmlns = y["xmlns" + ((y[0] as string).match(/<(\w+):/) || ["", ""])[1]];
				break;
			case "<Default":
				ctext[y.Extension.toLowerCase()] = y.ContentType;
				break;
			case "<Override":
				if (CONTENT_TYPE_MAP[y.ContentType] && (ct as any)[CONTENT_TYPE_MAP[y.ContentType]] !== undefined) {
					(ct as any)[CONTENT_TYPE_MAP[y.ContentType]].push(y.PartName);
				}
				break;
		}
	}
	if (ct.xmlns !== XMLNS.CT) {
		throw new Error("Unknown Namespace: " + ct.xmlns);
	}
	ct.calcchain = ct.calcchains.length > 0 ? ct.calcchains[0] : "";
	ct.sst = ct.strs.length > 0 ? ct.strs[0] : "";
	ct.style = ct.styles.length > 0 ? ct.styles[0] : "";
	ct.defaults = ctext;
	return ct;
}

/** Build reverse map from type category to content-type strings */
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

	if (o.length > 2) {
		o.push("</Types>");
		o[1] = o[1].replace("/>", ">");
	}
	return o.join("");
}
