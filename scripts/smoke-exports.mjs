import { createRequire } from "node:module";

const require = createRequire(import.meta.url);

const modules = [
	["ESM", await import("../dist/index.js")],
	["CJS", require("../dist/index.cjs")],
];

for (const [kind, mod] of modules) {
	for (const name of ["read", "write", "createWorkbook", "appendSheet"]) {
		if (typeof mod[name] !== "function") {
			throw new Error(`${kind} export ${name} is not available`);
		}
	}
}
