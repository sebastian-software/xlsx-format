import { parseXmlTag, XML_TAG_REGEX, stripNamespace } from "../xml/parser.js";
import { writeXmlTag, writeXmlElement } from "../xml/writer.js";
import { decodeCell, encodeCell } from "../utils/cell.js";
import { matchXmlTagGlobal } from "../utils/helpers.js";
import type { WorkSheet } from "../types.js";

/** VML XML namespace declarations for Microsoft Office drawing elements */
const XLMLNS: Record<string, string> = {
	v: "urn:schemas-microsoft-com:vml",
	o: "urn:schemas-microsoft-com:office:office",
	x: "urn:schemas-microsoft-com:office:excel",
	mv: "http://macVmlSchemaUri",
};

/**
 * Parse VML drawings to extract comment visibility and position.
 *
 * VML (Vector Markup Language) is the legacy drawing format used by Excel to
 * define comment box shapes. Each <v:shape> with ObjectType="Note" corresponds
 * to a comment, and its <Visible> element determines if the comment is shown.
 *
 * @param data - Raw VML XML string
 * @param sheet - Worksheet to update with comment visibility
 * @param comments - Array of comment references for fallback positioning
 */
export function parseVml(data: string, sheet: WorkSheet, comments: any[]): void {
	let cidx = 0;
	(matchXmlTagGlobal(data, "(?:shape|rect)") || []).forEach((m) => {
		let type = "";
		let hidden = true;
		let aidx = -1;
		let R = -1,
			C = -1;

		m.replace(XML_TAG_REGEX, function (x: string, idx: number): string {
			const y = parseXmlTag(x);
			switch (stripNamespace(y[0])) {
				case "<ClientData":
					if (y.ObjectType) {
						type = y.ObjectType;
					}
					break;
				case "<Visible":
				case "<Visible/>":
					hidden = false;
					break;
				// <Row> and <Column> contain the 0-based cell coordinates as text content
				case "<Row":
				case "<Row>":
					aidx = idx + x.length;
					break;
				case "</Row>":
					R = +m.slice(aidx, idx).trim();
					break;
				case "<Column":
				case "<Column>":
					aidx = idx + x.length;
					break;
				case "</Column>":
					C = +m.slice(aidx, idx).trim();
					break;
			}
			return "";
		});

		switch (type) {
			case "Note": {
				// Use parsed row/column if available, otherwise fall back to comment list order
				const ref = R >= 0 && C >= 0 ? encodeCell({ r: R, c: C }) : comments[cidx]?.ref;
				const dense = (sheet as any)["!data"] != null;
				let cell: any;
				if (dense) {
					const rows = (sheet as any)["!data"];
					cell = rows?.[R]?.[C];
				} else {
					cell = (sheet as any)[ref];
				}
				if (cell && cell.c) {
					cell.c.hidden = hidden;
				}
				++cidx;
				break;
			}
		}
	});
}

/** Format an object of attributes as XML attribute string (e.g. ' key="value"') */
function formatXmlAttributes(h: Record<string, string>): string {
	return Object.keys(h)
		.map((k) => " " + k + '="' + h[k] + '"')
		.join("");
}

/** Generate VML XML for a single comment shape */
function writeVmlComment(x: [string, any], _shapeid: number): string {
	const c = decodeCell(x[0]);
	// Gradient fill styling for the comment box
	const fillopts: any = { color2: "#BEFF82", type: "gradient" };
	if (fillopts.type === "gradient") {
		fillopts.angle = "-180";
	}
	const fillparm =
		fillopts.type === "gradient"
			? writeXmlElement("o:fill", null, { type: "gradientUnscaled", "v:ext": "view" })
			: null;
	const fillxml = writeXmlElement("v:fill", fillparm, fillopts);
	const shadata: any = { on: "t", obscured: "t" };

	return [
		"<v:shape" +
			formatXmlAttributes({
				id: "_x0000_s" + _shapeid,
				type: "#_x0000_t202", // TextBox shape type
				style:
					"position:absolute; margin-left:80pt;margin-top:5pt;width:104pt;height:64pt;z-index:10" +
					(x[1].hidden ? ";visibility:hidden" : ""),
				fillcolor: "#ECFAD4",
				strokecolor: "#edeaa1",
			}) +
			">",
		fillxml,
		writeXmlElement("v:shadow", null, shadata),
		writeXmlElement("v:path", null, { "o:connecttype": "none" }),
		'<v:textbox><div style="text-align:left"></div></v:textbox>',
		'<x:ClientData ObjectType="Note">',
		"<x:MoveWithCells/>",
		"<x:SizeWithCells/>",
		// Anchor: [startCol, colOffset, startRow, rowOffset, endCol, colOffset, endRow, rowOffset]
		writeXmlTag("x:Anchor", [c.c + 1, 0, c.r + 1, 0, c.c + 3, 20, c.r + 5, 20].join(",")),
		writeXmlTag("x:AutoFill", "False"),
		writeXmlTag("x:Row", String(c.r)),
		writeXmlTag("x:Column", String(c.c)),
		x[1].hidden ? "" : "<x:Visible/>",
		"</x:ClientData>",
		"</v:shape>",
	].join("");
}

/**
 * Write VML XML for all comment shapes on a sheet.
 *
 * VML is required for backward-compatible comment rendering in Excel.
 * Each comment gets a text-box shape positioned relative to its cell.
 *
 * @param rId - Sheet relationship ID (used for shape ID namespace partitioning)
 * @param comments - Array of [cell_ref, comment_data] tuples
 * @returns Complete VML XML string
 */
export function writeVml(rId: number, comments: [string, any][]): string {
	// 21600 x 21600 is the standard VML coordinate space
	const csize = [21600, 21600];
	// Define the shape path as a rectangle in VML path syntax
	const bbox = ["m0,0l0", csize[1], csize[0], csize[1], csize[0], "0xe"].join(",");
	const o: string[] = [
		writeXmlElement("xml", null, {
			"xmlns:v": XLMLNS.v,
			"xmlns:o": XLMLNS.o,
			"xmlns:x": XLMLNS.x,
			"xmlns:mv": XLMLNS.mv,
		}).replace(/\/>/, ">"),
		writeXmlElement("o:shapelayout", writeXmlElement("o:idmap", null, { "v:ext": "edit", data: String(rId) }), {
			"v:ext": "edit",
		}),
	];

	// Shape IDs are partitioned by sheet: 65536 * rId + sequential
	let _shapeid = 65536 * rId;
	const _comments = comments || [];

	// Define the shared shape type template (TextBox #202) if there are comments
	if (_comments.length > 0) {
		o.push(
			writeXmlElement(
				"v:shapetype",
				[
					writeXmlElement("v:stroke", null, { joinstyle: "miter" }),
					writeXmlElement("v:path", null, { gradientshapeok: "t", "o:connecttype": "rect" }),
				].join(""),
				{
					id: "_x0000_t202",
					coordsize: csize.join(","),
					"o:spt": "202", // Shape preset type 202 = TextBox
					path: bbox,
				},
			),
		);
	}

	_comments.forEach((x) => {
		++_shapeid;
		o.push(writeVmlComment(x, _shapeid));
	});
	o.push("</xml>");
	return o.join("");
}
