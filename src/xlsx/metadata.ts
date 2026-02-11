import { parseXmlTag, XML_TAG_REGEX, XML_HEADER, stripNamespace } from "../xml/parser.js";

/** Parsed metadata structure from metadata.xml */
export interface XLMeta {
	/** Registered metadata type definitions */
	Types: { name: string; offsets?: number[] }[];
	/** Cell-level metadata references (metatype=1) */
	Cell: { type: string; index: number }[];
	/** Value-level metadata references (metatype=0) */
	Value: { type: string; index: number }[];
}

/**
 * Parse metadata XML (ECMA-376 18.9 / MS-XLSX extensions).
 *
 * Metadata provides additional cell-level information such as dynamic array
 * properties (XLDAPR) and rich data types. The XML contains:
 * - metadataTypes: type definitions
 * - futureMetadata: type-specific data (e.g. rich value offsets)
 * - cellMetadata / valueMetadata: per-cell type+index references
 *
 * @param data - Raw XML string of metadata.xml
 * @param opts - Parsing options
 * @returns Parsed metadata structure
 */
export function parseMetadataXml(data: string, _opts?: any): XLMeta {
	const out: XLMeta = { Types: [], Cell: [], Value: [] };
	if (!data) {
		return out;
	}

	// metatype tracks which section we're in: 2=none, 1=cellMetadata, 0=valueMetadata
	let metatype = 2;
	let lastmeta: any;

	const ignoredTags = new Set([
		"<?xml",
		"<metadata",
		"</metadata>",
		"<metadataTypes",
		"</metadataTypes>",
		"</metadataType>",
		"</futureMetadata>",
		"<bk>",
		"</bk>",
		"</rc>",
		"<extLst",
		"<extLst>",
		"</extLst>",
		"<extLst/>",
		"<ext",
		"</ext>",
	]);

	data.replace(XML_TAG_REGEX, function (x: string): string {
		const y: any = parseXmlTag(x);
		const tag = stripNamespace(y[0]);
		if (ignoredTags.has(tag)) {
			return x;
		}
		switch (tag) {
			case "<metadataType":
				out.Types.push({ name: y.name });
				break;
			case "<futureMetadata":
				// Associate future metadata with its type definition by name
				for (let j = 0; j < out.Types.length; ++j) {
					if (out.Types[j].name === y.name) {
						lastmeta = out.Types[j];
					}
				}
				break;
			case "<rc":
				// <rc t="N" v="M"> references type index N (1-based) and value index M
				if (metatype === 1) {
					out.Cell.push({ type: out.Types[y.t - 1].name, index: +y.v });
				} else if (metatype === 0) {
					out.Value.push({ type: out.Types[y.t - 1].name, index: +y.v });
				}
				break;
			case "<cellMetadata":
				metatype = 1;
				break;
			case "</cellMetadata>":
				metatype = 2;
				break;
			case "<valueMetadata":
				metatype = 0;
				break;
			case "</valueMetadata>":
				metatype = 2;
				break;
			case "<rvb":
				// Rich value block offset: records the index into the rich data store
				if (lastmeta) {
					if (!lastmeta.offsets) {
						lastmeta.offsets = [];
					}
					lastmeta.offsets.push(+y.i);
				}
				break;
		}
		return x;
	});
	return out;
}

/**
 * Write minimal metadata XML for dynamic array support.
 *
 * Generates the XLDAPR (Dynamic Array Properties) metadata type that Excel
 * requires for spill-range formulas. The metadata declares a single type
 * and a single cell metadata entry referencing it.
 *
 * @returns Complete metadata.xml string
 */
export function writeMetadataXml(): string {
	const o: string[] = [XML_HEADER];
	o.push(
		'<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">\n' +
			'  <metadataTypes count="1">\n' +
			'    <metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/>\n' +
			"  </metadataTypes>\n" +
			'  <futureMetadata name="XLDAPR" count="1">\n' +
			"    <bk>\n" +
			"      <extLst>\n" +
			'        <ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">\n' +
			'          <xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/>\n' +
			"        </ext>\n" +
			"      </extLst>\n" +
			"    </bk>\n" +
			"  </futureMetadata>\n" +
			'  <cellMetadata count="1">\n' +
			"    <bk>\n" +
			'      <rc t="1" v="0"/>\n' +
			"    </bk>\n" +
			"  </cellMetadata>\n" +
			"</metadata>",
	);
	return o.join("");
}
