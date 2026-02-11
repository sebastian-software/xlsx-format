import { parseXmlTag, XML_HEADER, XML_TAG_REGEX } from "../xml/parser.js";
import { unescapeXml } from "../xml/escape.js";
import { writeXmlElement } from "../xml/writer.js";
import { XMLNS, RELS } from "../xml/namespaces.js";

/** A single relationship entry from a .rels file */
export interface RelEntry {
	Id: string;
	Type: string;
	Target: string;
	TargetMode?: string;
}

/** Collection of parsed relationships, indexed by resolved target path and by rId */
export interface Relationships {
	[target: string]: any;
	/** Lookup table mapping rId strings (e.g., "rId1") to their RelEntry objects */
	"!id": Record<string, RelEntry>;
	/** Auto-incrementing counter for assigning new relationship IDs */
	"!idx"?: number;
}

/**
 * Resolve a relative target path against a base file path.
 * Handles ".." segments to navigate up the directory tree.
 * Absolute paths (starting with "/") are returned as-is.
 * @param target - the relative or absolute target path
 * @param basePath - the path of the file that contains the relationship
 * @returns the resolved absolute path
 */
function resolve_path(target: string, basePath: string): string {
	if (target.charAt(0) === "/") {
		return target;
	}
	// Extract the directory portion of the base path
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

/**
 * Get the conventional .rels file path for a given OPC part.
 * For example, "xl/workbook.xml" becomes "xl/_rels/workbook.xml.rels".
 * @param file - the OPC part path
 * @returns the corresponding .rels file path
 */
export function getRelsPath(file: string): string {
	const n = file.lastIndexOf("/");
	return file.slice(0, n + 1) + "_rels/" + file.slice(n + 1) + ".rels";
}

/**
 * Parse a .rels XML file into a Relationships object.
 * Each Relationship element is stored both by its resolved target path
 * and by its rId in the "!id" lookup table.
 * @param data - raw XML string of the .rels file (may be null/undefined)
 * @param currentFilePath - the path of the file that owns this .rels (used for path resolution)
 * @returns the parsed Relationships object
 */
export function parseRelationships(data: string | null | undefined, currentFilePath: string): Relationships {
	const rels = { "!id": {} } as any as Relationships;
	if (!data) {
		return rels;
	}
	// Ensure the path starts with "/" for consistent resolution
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
			// External targets (e.g., hyperlinks) are stored as-is; internal targets are resolved
			const canonictarget = y.TargetMode === "External" ? y.Target : resolve_path(y.Target, currentFilePath);
			(rels as any)[canonictarget] = rel;
			rels["!id"][y.Id] = rel;
		}
	}
	return rels;
}

/**
 * Serialize a Relationships object back to .rels XML format.
 * @param rels - the Relationships object to serialize
 * @returns the complete XML string for the .rels file
 */
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
	// Only close the root element if child elements were added
	if (o.length > 2) {
		o.push("</Relationships>");
		// Convert the self-closing root tag to an opening tag
		o[1] = o[1].replace("/>", ">");
	}
	return o.join("");
}

/**
 * Add a new relationship entry to a Relationships object.
 * If rId is negative, automatically finds the next available rId.
 * Hyperlink-type relationships default to TargetMode="External".
 * @param rels - the Relationships object to modify
 * @param rId - the numeric relationship ID to use, or negative to auto-assign
 * @param f - the target path for the relationship
 * @param type - the relationship type URI
 * @param targetmode - optional TargetMode ("External" for hyperlinks, etc.)
 * @returns the numeric rId that was assigned
 * @throws if the specified rId is already in use
 */
export function addRelationship(
	rels: Relationships,
	rId: number,
	f: string,
	type: string,
	targetmode?: string,
): number {
	if (!rels["!id"]) {
		(rels as any)["!id"] = {};
	}
	if (!(rels as any)["!idx"]) {
		(rels as any)["!idx"] = 1;
	}
	// Auto-assign: scan forward from the last used index to find a free slot
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
		// Hyperlinks are always external by default
		relobj.TargetMode = "External";
	}
	if (rels["!id"][relobj.Id]) {
		throw new Error("Cannot rewrite rId " + rId);
	}
	rels["!id"][relobj.Id] = relobj;
	// Store by normalized target path (ensure single leading slash)
	(rels as any)[("/" + relobj.Target).replace("//", "/")] = relobj;
	return rId;
}
