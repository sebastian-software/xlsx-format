import { parseXmlTag, XML_HEADER, XML_TAG_REGEX } from "../xml/parser.js";
import { unescapeXml } from "../xml/escape.js";
import { writeXmlElement } from "../xml/writer.js";
import { XMLNS, RELS } from "../xml/namespaces.js";

export interface RelEntry {
	Id: string;
	Type: string;
	Target: string;
	TargetMode?: string;
}

export interface Relationships {
	[target: string]: any;
	"!id": Record<string, RelEntry>;
	"!idx"?: number;
}

/** Resolve relative path against current file path */
function resolve_path(target: string, basePath: string): string {
	if (target.charAt(0) === "/") {
		return target;
	}
	const base = basePath.slice(0, basePath.lastIndexOf("/") + 1);
	const parts = (base + target).split("/");
	const resolved: string[] = [];
	for (const p of parts) {
		if (p === "..") {
			resolved.pop();
		} else if (p !== ".") {
			resolved.push(p);
		}
	}
	return resolved.join("/");
}

/** Get the .rels path for a given file */
export function getRelsPath(file: string): string {
	const n = file.lastIndexOf("/");
	return file.slice(0, n + 1) + "_rels/" + file.slice(n + 1) + ".rels";
}

/** Parse a .rels XML file */
export function parseRelationships(data: string | null | undefined, currentFilePath: string): Relationships {
	const rels = { "!id": {} } as any as Relationships;
	if (!data) {
		return rels;
	}
	if (currentFilePath.charAt(0) !== "/") {
		currentFilePath = "/" + currentFilePath;
	}

	const matches = data.match(XML_TAG_REGEX) || [];
	for (const x of matches) {
		const y = parseXmlTag(x);
		if (y[0] === "<Relationship") {
			const rel: RelEntry = {
				Type: y.Type,
				Target: unescapeXml(y.Target),
				Id: y.Id,
			};
			if (y.TargetMode) {
				rel.TargetMode = y.TargetMode;
			}
			const canonictarget = y.TargetMode === "External" ? y.Target : resolve_path(y.Target, currentFilePath);
			(rels as any)[canonictarget] = rel;
			rels["!id"][y.Id] = rel;
		}
	}
	return rels;
}

/** Serialize relationships to XML */
export function writeRelationships(rels: Relationships): string {
	const o: string[] = [
		XML_HEADER,
		writeXmlElement("Relationships", null, {
			xmlns: XMLNS.RELS,
		}),
	];
	for (const rid of Object.keys(rels["!id"])) {
		o.push(writeXmlElement("Relationship", null, rels["!id"][rid] as any));
	}
	if (o.length > 2) {
		o.push("</Relationships>");
		o[1] = o[1].replace("/>", ">");
	}
	return o.join("");
}

/** Add a relationship entry */
export function addRelationship(rels: Relationships, rId: number, f: string, type: string, targetmode?: string): number {
	if (!rels["!id"]) {
		(rels as any)["!id"] = {};
	}
	if (!(rels as any)["!idx"]) {
		(rels as any)["!idx"] = 1;
	}
	if (rId < 0) {
		for (rId = (rels as any)["!idx"]; rels["!id"]["rId" + rId]; ++rId) {
			/* find next free rId */
		}
	}
	(rels as any)["!idx"] = rId + 1;
	const relobj: RelEntry = {
		Id: "rId" + rId,
		Type: type,
		Target: f,
	};
	if (targetmode) {
		relobj.TargetMode = targetmode;
	} else if ([RELS.HLINK].indexOf(type) > -1) {
		relobj.TargetMode = "External";
	}
	if (rels["!id"][relobj.Id]) {
		throw new Error("Cannot rewrite rId " + rId);
	}
	rels["!id"][relobj.Id] = relobj;
	(rels as any)[("/" + relobj.Target).replace("//", "/")] = relobj;
	return rId;
}
