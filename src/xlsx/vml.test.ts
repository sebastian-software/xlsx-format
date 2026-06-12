import { describe, it, expect } from "vitest";
import { arrayToSheet } from "../index.js";
import { parseVml, writeVml } from "./vml.js";

describe("xlsx/vml", () => {
	it("writeVml should produce valid VML XML", () => {
		const comments: [string, any][] = [["A1", { hidden: false }]];
		const xml = writeVml(1, comments);
		expect(xml).toContain("<xml");
		expect(xml).toContain("v:shape");
		expect(xml).toContain("x:Row");
		expect(xml).toContain("x:Column");
		expect(xml).toContain("x:Visible");
	});

	it("writeVml should handle hidden comments", () => {
		const comments: [string, any][] = [["B2", { hidden: true }]];
		const xml = writeVml(1, comments);
		expect(xml).toContain("visibility:hidden");
		expect(xml).not.toContain("<x:Visible/>");
	});

	it("writeVml should handle empty comments", () => {
		const xml = writeVml(1, []);
		expect(xml).toContain("<xml");
		expect(xml).not.toContain("v:shapetype");
	});

	it("parseVml should set comment visibility", () => {
		const ws = arrayToSheet([[1]]);
		ws["A1"].c = [{ a: "Test", t: "Note" }];
		const vml = `<xml><v:shape><x:ClientData ObjectType="Note"><x:Row>0</x:Row><x:Column>0</x:Column><x:Visible/></x:ClientData></v:shape></xml>`;
		parseVml(vml, ws, [{ ref: "A1" }]);
		expect(ws["A1"].c.hidden).toBe(false);
	});
});
