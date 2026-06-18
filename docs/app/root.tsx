import { ArdoNav, ArdoNavLink, ArdoRootLayout, ArdoRoot } from "ardo/ui";
import config from "virtual:ardo/config";
import sidebar from "virtual:ardo/sidebar";
import type { MetaFunction } from "react-router";
import "ardo/ui/styles.css";
import "./custom.css";

export const meta: MetaFunction = () => [{ title: config.title }];

export function Layout({ children }: { children: React.ReactNode }) {
	return <ArdoRootLayout>{children}</ArdoRootLayout>;
}

export default function Root() {
	return (
		<ArdoRoot
			config={config}
			sidebar={sidebar}
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
