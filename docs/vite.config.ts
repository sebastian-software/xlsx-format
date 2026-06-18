import { defineConfig } from "vite";
import { ardo } from "ardo/vite";
import pkg from "../package.json" with { type: "json" };

export default defineConfig({
	plugins: [
		ardo({
			title: "xlsx-format",
			description:
				"The XLSX library your bundler will thank you for. Zero dependencies. Fully async. TypeScript-first.",

			typedoc: true,

			project: {
				name: pkg.name,
				version: pkg.version,
				homepage: pkg.homepage,
			},

			sidebar: {
				sectionOrder: ["guide", "api-reference"],
			},
		}),
	],
});
