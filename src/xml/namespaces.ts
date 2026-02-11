/**
 * Standard XML namespace URIs used in OPC (Open Packaging Conventions) and OOXML documents.
 * Keys are short identifiers used throughout the codebase; values are the full namespace URIs.
 */
export const XMLNS: Record<string, string> = {
	CORE_PROPS: "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
	CUST_PROPS: "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
	EXT_PROPS: "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
	CT: "http://schemas.openxmlformats.org/package/2006/content-types",
	RELS: "http://schemas.openxmlformats.org/package/2006/relationships",
	TCMNT: "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments",
	dc: "http://purl.org/dc/elements/1.1/",
	dcterms: "http://purl.org/dc/terms/",
	dcmitype: "http://purl.org/dc/dcmitype/",
	mx: "http://schemas.microsoft.com/office/mac/excel/2008/main",
	r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
	sjs: "http://schemas.openxmlformats.org/package/2006/sheetjs/core-properties",
	vt: "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
	xsi: "http://www.w3.org/2001/XMLSchema-instance",
	xsd: "http://www.w3.org/2001/XMLSchema",
};

/**
 * Recognized namespace URIs for the SpreadsheetML main namespace.
 * Multiple URIs exist because OOXML has both ECMA-376 and transitional/strict variants,
 * plus Microsoft-specific extensions.
 */
export const XMLNS_main = [
	"http://schemas.openxmlformats.org/spreadsheetml/2006/main",
	"http://purl.oclc.org/ooxml/spreadsheetml/main",
	"http://schemas.microsoft.com/office/excel/2006/main",
	"http://schemas.microsoft.com/office/excel/2006/2",
];

/**
 * OPC relationship type URIs.
 * Each key is a short identifier; each value is the full relationship type URI
 * used in .rels files to link package parts together.
 */
export const RELS = {
	WB: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
	SHEET: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
	CHARTSHEET: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet",
	HLINK: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
	VML: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
	CMNT: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
	CORE_PROPS: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
	EXT_PROPS: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
	CUST_PROPS: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties",
	SST: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
	STY: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
	THEME: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
	CHART: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
	CCHAIN: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain",
	TCMNT: "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment",
	PEOPLE: "http://schemas.microsoft.com/office/2017/10/relationships/person",
	DRAWING: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
	META: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata",
	XLINK: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink",
};
