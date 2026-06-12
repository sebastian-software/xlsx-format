import { defineConfig } from "tsdown";

export default defineConfig({
	entry: ["src/index.ts"],
	format: ["cjs", "esm"],
	dts: true,
	sourcemap: true,
	clean: true,
	fixedExtension: false,
	treeshake: true,
	minify: false,
	target: "es2020",
	outDir: "dist",
});
