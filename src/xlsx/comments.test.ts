import { describe, it, expect } from "vitest";
import {
	insertCommentsIntoSheet,
	writeCommentsXml,
	parseCommentsXml,
	parseTcmntXml,
	writeTcmntXml,
	parsePeopleXml,
	writePeopleXml,
} from "./comments.js";
import { arrayToSheet } from "../index.js";

describe("comments: insertCommentsIntoSheet", () => {
	it("inserts legacy comment into sheet", () => {
		const ws: any = { A1: { t: "n", v: 1 }, "!ref": "A1" };
		const comments = [{ ref: "A1", author: "Alice", t: "Hello", r: "<t>Hello</t>", h: "" }];
		insertCommentsIntoSheet(ws, comments as any, false);
		expect(ws["A1"].c).toBeDefined();
		expect(ws["A1"].c[0].t).toBe("Hello");
	});

	it("inserts threaded comment and removes legacy", () => {
		const ws: any = { A1: { t: "n", v: 1, c: [{ a: "Alice", t: "legacy", T: false }] }, "!ref": "A1" };
		const comments = [{ ref: "A1", author: "Bob", t: "threaded", r: "<t>threaded</t>", h: "" }];
		insertCommentsIntoSheet(ws, comments as any, true);
		// Threaded should override legacy
		expect(ws["A1"].c.some((c: any) => c.T === true)).toBe(true);
	});

	it("does not add legacy when threaded exists", () => {
		const ws: any = { A1: { t: "n", v: 1, c: [{ a: "Alice", t: "threaded", T: true }] }, "!ref": "A1" };
		const comments = [{ ref: "A1", author: "Bob", t: "legacy", r: "<t>legacy</t>", h: "" }];
		insertCommentsIntoSheet(ws, comments as any, false);
		// Should not add legacy comment when threaded exists
		expect(ws["A1"].c.length).toBe(1);
	});

	it("creates cell when it doesn't exist and expands range", () => {
		const ws: any = { "!ref": "A1" };
		const comments = [{ ref: "C3", author: "Alice", t: "New", r: "<t>New</t>", h: "" }];
		insertCommentsIntoSheet(ws, comments as any, false);
		expect(ws["C3"]).toBeDefined();
		expect(ws["C3"].c).toBeDefined();
		// Range should be expanded
		expect(ws["!ref"]).not.toBe("A1");
	});

	it("inserts into dense mode worksheet", () => {
		const ws: any = { "!data": [[{ t: "n", v: 1 }]], "!ref": "A1" };
		const comments = [{ ref: "A1", author: "Alice", t: "Dense", r: "<t>Dense</t>", h: "" }];
		insertCommentsIntoSheet(ws, comments as any, false);
		expect(ws["!data"][0][0].c).toBeDefined();
	});
});

describe("comments: writeCommentsXml", () => {
	it("writes basic comment XML", () => {
		const data: [string, any[]][] = [["A1", [{ a: "Alice", t: "Hello" }]]];
		const xml = writeCommentsXml(data);
		expect(xml).toContain("<authors>");
		expect(xml).toContain("Alice");
		expect(xml).toContain("Hello");
		expect(xml).toContain("<commentList>");
	});

	it("adds default author when no comments have authors", () => {
		// Empty data array triggers default author path
		const data: [string, any[]][] = [];
		const xml = writeCommentsXml(data);
		expect(xml).toContain("SheetJ5");
	});

	it("writes threaded comment with ID", () => {
		const data: [string, any[]][] = [["A1", [{ a: "Alice", t: "Thread", T: true, ID: "TC001" }]]];
		const xml = writeCommentsXml(data);
		expect(xml).toContain("tc=TC001");
	});
});

describe("comments: parseCommentsXml", () => {
	it("parses comment with empty text", () => {
		const xml = `<?xml version="1.0"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<authors><author>Alice</author></authors>
<commentList>
<comment ref="A1" authorId="0"><text><t></t></text></comment>
</commentList>
</comments>`;
		const comments = parseCommentsXml(xml, { cellHTML: true });
		expect(comments).toBeDefined();
		expect(comments.length).toBeGreaterThanOrEqual(1);
	});

	it("parses comment with sheetRows limit", () => {
		const xml = `<?xml version="1.0"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<authors><author>Alice</author></authors>
<commentList>
<comment ref="A100" authorId="0"><text><t>Hidden</t></text></comment>
</commentList>
</comments>`;
		const comments = parseCommentsXml(xml, { sheetRows: 10 });
		// Comment on row 100 should be excluded
		expect(comments.length).toBe(0);
	});
});

describe("xlsx/comments", () => {
	it("parseCommentsXml should roundtrip through write → parse", () => {
		const data: [string, any[]][] = [["A1", [{ a: "Alice", t: "Hello World" }]]];
		const xml = writeCommentsXml(data);
		const comments = parseCommentsXml(xml);
		expect(comments).toHaveLength(1);
		expect(comments[0].ref).toBe("A1");
		expect(comments[0].author).toBe("Alice");
		expect(comments[0].t).toBe("Hello World");
	});

	it("parseCommentsXml should parse hand-crafted XML", () => {
		const xml = `<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
			<authors><author>Bob</author></authors>
			<commentList>
				<comment ref="B2" authorId="0"><text><t>A note</t></text></comment>
			</commentList>
		</comments>`;
		const comments = parseCommentsXml(xml);
		expect(comments).toHaveLength(1);
		expect(comments[0].ref).toBe("B2");
		expect(comments[0].author).toBe("Bob");
		expect(comments[0].t).toBe("A note");
	});

	it("parseCommentsXml should handle empty comments", () => {
		expect(parseCommentsXml("<comments/>")).toEqual([]);
	});

	it("writeCommentsXml should produce valid XML", () => {
		const data: [string, any[]][] = [["A1", [{ a: "Bob", t: "Note here" }]]];
		const xml = writeCommentsXml(data);
		expect(xml).toContain("<comments");
		expect(xml).toContain("<author>Bob</author>");
		expect(xml).toContain('ref="A1"');
		expect(xml).toContain("Note here");
	});

	it("writeCommentsXml should handle threaded comments", () => {
		const data: [string, any[]][] = [
			[
				"B2",
				[
					{ a: "Alice", t: "Main comment", T: 1, ID: "{123}" },
					{ a: "Bob", t: "Reply", T: 1 },
				],
			],
		];
		const xml = writeCommentsXml(data);
		expect(xml).toContain("Comment:");
		expect(xml).toContain("Reply:");
	});

	it("parseTcmntXml should parse threaded comments", () => {
		const xml = `<?xml version="1.0"?><ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
			<threadedComment ref="A1" id="{111}" personId="{222}"><text>Thread text</text></threadedComment>
		</ThreadedComments>`;
		const result = parseTcmntXml(xml);
		expect(result).toHaveLength(1);
		expect(result[0].t).toBe("Thread text");
		expect(result[0].ref).toBe("A1");
		expect(result[0].T).toBe(1);
	});

	it("writeTcmntXml should produce valid XML", () => {
		const comments: [string, any[]][] = [["A1", [{ a: "Alice", t: "Hello", T: 1 }]]];
		const people: string[] = [];
		const xml = writeTcmntXml(comments, people, { tcid: 1 });
		expect(xml).toContain("<ThreadedComments");
		expect(xml).toContain("<threadedComment");
		expect(xml).toContain("Hello");
		expect(people).toContain("Alice");
	});

	it("parsePeopleXml should parse person list", () => {
		const xml = `<?xml version="1.0"?><personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
			<person displayname="Alice" id="{111}" userId="alice" providerId="None"/>
		</personList>`;
		const people = parsePeopleXml(xml);
		expect(people).toHaveLength(1);
		expect(people[0].name).toBe("Alice");
		expect(people[0].id).toBe("{111}");
	});

	it("writePeopleXml should produce valid XML", () => {
		const xml = writePeopleXml(["Alice", "Bob"]);
		expect(xml).toContain("<personList");
		expect(xml).toContain("Alice");
		expect(xml).toContain("Bob");
	});

	it("insertCommentsIntoSheet should attach comments to cells", () => {
		const ws = arrayToSheet([
			["A", "B"],
			[1, 2],
		]);
		const comments = [{ ref: "A1", author: "Test", t: "Comment text" }];
		insertCommentsIntoSheet(ws, comments, false);
		expect(ws["A1"].c).toBeDefined();
		expect(ws["A1"].c![0].t).toBe("Comment text");
	});

	it("insertCommentsIntoSheet should create cells if needed", () => {
		const ws = arrayToSheet([[1]]);
		const comments = [{ ref: "C3", author: "Test", t: "On empty cell" }];
		insertCommentsIntoSheet(ws, comments, false);
		expect((ws as any)["C3"]).toBeDefined();
		expect((ws as any)["C3"].c[0].t).toBe("On empty cell");
	});
});
