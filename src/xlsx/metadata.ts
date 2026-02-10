import { parseXmlTag, XML_TAG_REGEX, XML_HEADER, stripNamespace } from "../xml/parser.js";

export interface XLMeta {
	Types: { name: string; offsets?: number[] }[];
	Cell: { type: string; index: number }[];
	Value: { type: string; index: number }[];
}

/** Parse metadata XML */
export function parseMetadataXml(data: string, opts?: any): XLMeta {
	const out: XLMeta = { Types: [], Cell: [], Value: [] };
	if (!data) {
		return out;
	}

	let pass = false;
	let metatype = 2;
	let lastmeta: any;

	data.replace(XML_TAG_REGEX, function (x: string): string {
		const y: any = parseXmlTag(x);
		switch (stripNamespace(y[0])) {
			case "<?xml":
				break;
			case "<metadata":
			case "</metadata>":
				break;
			case "<metadataTypes":
			case "</metadataTypes>":
				break;
			case "<metadataType":
				out.Types.push({ name: y.name });
				break;
			case "</metadataType>":
				break;
			case "<futureMetadata":
				for (let j = 0; j < out.Types.length; ++j) {
					if (out.Types[j].name === y.name) {
						lastmeta = out.Types[j];
					}
				}
				break;
			case "</futureMetadata>":
				break;
			case "<bk>":
			case "</bk>":
				break;
			case "<rc":
				if (metatype === 1) {
					out.Cell.push({ type: out.Types[y.t - 1].name, index: +y.v });
				} else if (metatype === 0) {
					out.Value.push({ type: out.Types[y.t - 1].name, index: +y.v });
				}
				break;
			case "</rc>":
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
			case "<rvb":
				if (!lastmeta) {
					break;
				}
				if (!lastmeta.offsets) {
					lastmeta.offsets = [];
				}
				lastmeta.offsets.push(+y.i);
				break;
			default:
				break;
		}
		return x;
	});
	return out;
}

/** Write minimal metadata XML for dynamic arrays */
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
