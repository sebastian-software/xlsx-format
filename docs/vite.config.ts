import { defineConfig } from "vite";
import { ardo } from "ardo/vite";

export default defineConfig({
	plugins: [
		ardo({
			title: "xlsx-format",
			description:
				"The XLSX library your bundler will thank you for. Zero dependencies. Fully async. TypeScript-first.",

			typedoc: true,

			themeConfig: {
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

				search: {
					enabled: true,
				},
			},
		}),
	],
});
