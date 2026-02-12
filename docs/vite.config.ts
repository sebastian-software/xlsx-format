import { defineConfig } from "vite";
import { ardo } from "ardo/vite";

export default defineConfig({
	define: {
		__BUILD_TIME__: JSON.stringify(new Date().toISOString()),
	},
	plugins: [
		ardo({
			title: "xlsx-format",
			description:
				"The XLSX library your bundler will thank you for. Zero dependencies. Fully async. TypeScript-first.",

			typedoc: true,

			// GitHub Pages: base path auto-detected from git remote

			themeConfig: {
				siteTitle: "xlsx-format",

				nav: [
					{ text: "Guide", link: "/guide/getting-started" },
					{ text: "API", link: "/api-reference" },
				],

				sidebar: [
					{
						text: "Guide",
						items: [
							{ text: "Getting Started", link: "/guide/getting-started" },
							{ text: "Why xlsx-format?", link: "/guide/why-xlsx-format" },
							{ text: "Migration from SheetJS", link: "/guide/migration" },
						],
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
