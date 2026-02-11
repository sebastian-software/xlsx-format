import { defineConfig } from "vite";
import { ardo } from "ardo/vite";

export default defineConfig({
	plugins: [
		ardo({
			title: "XLSX Docs",
			description: "Built with Ardo",

			typedoc: {
				entryPoints: ["../src/index.ts"],
			},

			// GitHub Pages: base path auto-detected from git remote

			themeConfig: {
				siteTitle: "XLSX Docs",

				nav: [{ text: "Guide", link: "/guide/getting-started" }],

				sidebar: [
					{
						text: "Guide",
						items: [{ text: "Getting Started", link: "/guide/getting-started" }],
					},
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
