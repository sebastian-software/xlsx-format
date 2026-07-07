import { cpSync, existsSync, readFileSync, rmSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const currentDir = dirname(fileURLToPath(import.meta.url));
const docsDir = join(currentDir, "..");
const workspacePackage = JSON.parse(readFileSync(join(docsDir, "..", "package.json"), "utf8"));
const buildDir = join(docsDir, "build", "client");
const nestedDir = join(buildDir, workspacePackage.name);

if (!existsSync(nestedDir)) {
	console.log("[docs] No nested GitHub Pages output to flatten.");
	process.exit(0);
}

console.log(`[docs] Flattening GitHub Pages output: ${workspacePackage.name}/`);
cpSync(nestedDir, buildDir, { recursive: true });
rmSync(nestedDir, { force: true, recursive: true });
