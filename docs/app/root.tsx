import { ArdoNav, ArdoNavLink, ArdoRootLayout, ArdoRoot } from "ardo/ui";
import type { ArdoContextItem, SidebarItem } from "ardo";
import config from "virtual:ardo/config";
import type { MetaFunction } from "react-router";
import "ardo/ui/styles.css";
import "./custom.css";

export const meta: MetaFunction = () => [{ title: config.title }];

export function Layout({ children }: { children: React.ReactNode }) {
	return <ArdoRootLayout>{children}</ArdoRootLayout>;
}

const contexts = [
	{
		id: "guide",
		label: "Guide",
		href: "/guide/getting-started",
		match: "/guide",
	},
	{
		id: "api-reference",
		label: "API Reference",
		href: "/api-reference",
		match: "/api-reference",
	},
] satisfies ArdoContextItem[];

const sidebars = {
	guide: [
		{ text: "Getting Started", link: "/guide/getting-started" },
		{ text: "Styled Workbooks", link: "/guide/styled-workbooks" },
		{ text: "Why xlsx-format?", link: "/guide/why-xlsx-format" },
		{ text: "Migration from SheetJS", link: "/guide/migration" },
		{ text: "Security Considerations", link: "/guide/security" },
	],
	"api-reference": [
		{ text: "Classes", link: "/api-reference/classes" },
		{ text: "Functions", link: "/api-reference/functions" },
		{ text: "Interfaces", link: "/api-reference/interfaces" },
		{ text: "Types", link: "/api-reference/types" },
		{ text: "Version", link: "/api-reference/variables/version" },
	],
} satisfies Record<string, SidebarItem[]>;

export default function Root() {
	return (
		<ArdoRoot
			config={config}
			sidebar={sidebars}
			contexts={contexts}
			headerProps={{
				nav: (
					<ArdoNav>
						<ArdoNavLink to="/guide/getting-started" activeMatch="/guide">
							Guide
						</ArdoNavLink>
						<ArdoNavLink to="/api-reference" activeMatch="/api-reference">
							API
						</ArdoNavLink>
					</ArdoNav>
				),
			}}
		/>
	);
}
