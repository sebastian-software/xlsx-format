import { defineConfig } from "vite";
import { ardo } from "ardo/vite";

export default defineConfig({
	plugins: [
		ardo({
			title: "XLSX Format",
			description: "Modern XLSX reader/writer — TypeScript rewrite of SheetJS (XLSX only)",

			typedoc: true,

			// GitHub Pages: base path auto-detected from git remote

			themeConfig: {
				siteTitle: "XLSX Format",

				nav: [
					{ text: "Guide", link: "/guide/getting-started" },
					{ text: "API", link: "/api-reference" },
				],

				sidebar: [
					{
						text: "Guide",
						items: [{ text: "Getting Started", link: "/guide/getting-started" }],
					},
					{ text: "API Reference", link: "/api-reference" },
				],

				footer: {
					message: "Built with Ardo",
				},

				search: {
					enabled: true,
				},
			},
		}),
	],
});
