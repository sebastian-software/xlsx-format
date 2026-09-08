import { defineConfig } from "vite";
import { ardo } from "ardo/vite";
import pkg from "../package.json" with { type: "json" };

export default defineConfig({
	plugins: [
		ardo({
			title: "xlsx-format",
			description:
				"The XLSX library your bundler will thank you for. Zero dependencies. Promise-based read/write APIs. TypeScript-first.",

			linkCheck: {
				level: "error",
			},

			// Favicon set (favicon.ico, icon.svg, apple-touch-icon.png) from the brand mark.
			icons: {
				source: "app/brand-mark.svg",
			},

			markdown: {
				// Code panels are always dark (like the hero editor), so both
				// color schemes highlight with the same dark theme.
				theme: {
					light: "vesper",
					dark: "vesper",
				},
			},

			typedoc: {
				entryPoints: ["../src/index.ts"],
				markdown: {
					breadcrumbs: false,
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
		}),
	],
});
