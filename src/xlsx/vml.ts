import { parsexmltag, tagregex, strip_ns } from "../xml/parser.js";
import { writetag, writextag } from "../xml/writer.js";
import { decode_cell, encode_cell } from "../utils/cell.js";
import { str_match_xml_ns_g } from "../utils/helpers.js";
import { keys } from "../utils/helpers.js";
import type { WorkSheet } from "../types.js";

const XLMLNS: Record<string, string> = {
	v: "urn:schemas-microsoft-com:vml",
	o: "urn:schemas-microsoft-com:office:office",
	x: "urn:schemas-microsoft-com:office:excel",
	mv: "http://macVmlSchemaUri",
};

/** Parse VML drawings (comment shapes) */
export function parse_vml(data: string, sheet: WorkSheet, comments: any[]): void {
	let cidx = 0;
	(str_match_xml_ns_g(data, "(?:shape|rect)") || []).forEach((m) => {
		let type = "";
		let hidden = true;
		let aidx = -1;
		let R = -1,
			C = -1;

		m.replace(tagregex, function (x: string, idx: number): string {
			const y = parsexmltag(x);
			switch (strip_ns(y[0])) {
				case "<ClientData":
					if (y.ObjectType) {
						type = y.ObjectType;
					}
					break;
				case "<Visible":
				case "<Visible/>":
					hidden = false;
					break;
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
				const ref = R >= 0 && C >= 0 ? encode_cell({ r: R, c: C }) : comments[cidx]?.ref;
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

function wxt_helper(h: Record<string, string>): string {
	return keys(h)
		.map((k) => " " + k + '="' + h[k] + '"')
		.join("");
}

function write_vml_comment(x: [string, any], _shapeid: number): string {
	const c = decode_cell(x[0]);
	const fillopts: any = { color2: "#BEFF82", type: "gradient" };
	if (fillopts.type === "gradient") {
		fillopts.angle = "-180";
	}
	const fillparm =
		fillopts.type === "gradient" ? writextag("o:fill", null, { type: "gradientUnscaled", "v:ext": "view" }) : null;
	const fillxml = writextag("v:fill", fillparm, fillopts);
	const shadata: any = { on: "t", obscured: "t" };

	return [
		"<v:shape" +
			wxt_helper({
				id: "_x0000_s" + _shapeid,
				type: "#_x0000_t202",
				style:
					"position:absolute; margin-left:80pt;margin-top:5pt;width:104pt;height:64pt;z-index:10" +
					(x[1].hidden ? ";visibility:hidden" : ""),
				fillcolor: "#ECFAD4",
				strokecolor: "#edeaa1",
			}) +
			">",
		fillxml,
		writextag("v:shadow", null, shadata),
		writextag("v:path", null, { "o:connecttype": "none" }),
		'<v:textbox><div style="text-align:left"></div></v:textbox>',
		'<x:ClientData ObjectType="Note">',
		"<x:MoveWithCells/>",
		"<x:SizeWithCells/>",
		writetag("x:Anchor", [c.c + 1, 0, c.r + 1, 0, c.c + 3, 20, c.r + 5, 20].join(",")),
		writetag("x:AutoFill", "False"),
		writetag("x:Row", String(c.r)),
		writetag("x:Column", String(c.c)),
		x[1].hidden ? "" : "<x:Visible/>",
		"</x:ClientData>",
		"</v:shape>",
	].join("");
}

/** Write VML for comment boxes */
export function write_vml(rId: number, comments: [string, any][]): string {
	const csize = [21600, 21600];
	const bbox = ["m0,0l0", csize[1], csize[0], csize[1], csize[0], "0xe"].join(",");
	const o: string[] = [
		writextag("xml", null, {
			"xmlns:v": XLMLNS.v,
			"xmlns:o": XLMLNS.o,
			"xmlns:x": XLMLNS.x,
			"xmlns:mv": XLMLNS.mv,
		}).replace(/\/>/, ">"),
		writextag("o:shapelayout", writextag("o:idmap", null, { "v:ext": "edit", data: String(rId) }), {
			"v:ext": "edit",
		}),
	];

	let _shapeid = 65536 * rId;
	const _comments = comments || [];

	if (_comments.length > 0) {
		o.push(
			writextag(
				"v:shapetype",
				[
					writextag("v:stroke", null, { joinstyle: "miter" }),
					writextag("v:path", null, { gradientshapeok: "t", "o:connecttype": "rect" }),
				].join(""),
				{
					id: "_x0000_t202",
					coordsize: csize.join(","),
					"o:spt": "202",
					path: bbox,
				},
			),
		);
	}

	_comments.forEach((x) => {
		++_shapeid;
		o.push(write_vml_comment(x, _shapeid));
	});
	o.push("</xml>");
	return o.join("");
}
