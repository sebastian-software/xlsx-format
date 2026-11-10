import { defineArdoConfig } from "ardo/config";
import pkg from "../package.json" with { type: "json" };

export default defineArdoConfig({
	base: "/xlsx-format/",
	title: "xlsx-format",
	description:
		"The XLSX library your bundler will thank you for. Zero dependencies. Promise-based read/write APIs. TypeScript-first.",

	linkCheck: {
		level: "error",
	},

	markdown: {
		// Code panels are always dark (like the hero editor), so both
		// color schemes highlight with the same dark theme.
		theme: {
			light: "vesper",
			dark: "vesper",
		},
	},

	project: {
		name: pkg.name,
		version: pkg.version,
		homepage: pkg.homepage,
	},

	sidebar: {
		sectionOrder: ["guide", "api-reference"],
	},
	api: {
		entryPoints: ["../src/index.ts"],
		tsconfig: "./tsconfig.api.json",
		out: "api-reference",
		title: "API Reference",
	},
});
