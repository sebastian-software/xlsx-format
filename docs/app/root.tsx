import {
	ArdoGeneratedSidebar,
	ArdoHeader,
	ArdoNav,
	ArdoNavLink,
	ArdoRoot,
	ArdoRootLayout,
	ArdoSidebar,
	ArdoSidebarGroup,
	ArdoSidebarLink,
	ArdoSidebarSection,
} from "ardo/ui";
import config from "virtual:ardo/config";
import type { MetaFunction } from "react-router";
import "ardo/ui/styles.css";
import "./custom.css";

export const meta: MetaFunction = () => [{ title: config.title }];

export function Layout({ children }: { children: React.ReactNode }) {
	return <ArdoRootLayout>{children}</ArdoRootLayout>;
}

export default function Root() {
	return (
		<ArdoRoot config={config}>
			<ArdoHeader>
				<ArdoNav>
					<ArdoNavLink to="/guide/getting-started">Guide</ArdoNavLink>
					<ArdoNavLink to="/api-reference">API</ArdoNavLink>
				</ArdoNav>
			</ArdoHeader>
			<ArdoSidebar>
				<ArdoSidebarSection id="guide" label="Guide" to="/guide/getting-started">
					<ArdoSidebarGroup title="Guide" collapsible={false}>
						<ArdoSidebarLink to="/guide/getting-started">Getting Started</ArdoSidebarLink>
						<ArdoSidebarLink to="/guide/styled-workbooks">Styled Workbooks</ArdoSidebarLink>
						<ArdoSidebarLink to="/guide/why-xlsx-format">Why xlsx-format?</ArdoSidebarLink>
						<ArdoSidebarLink to="/guide/migration">Migration from SheetJS</ArdoSidebarLink>
						<ArdoSidebarLink to="/guide/security">Security Considerations</ArdoSidebarLink>
					</ArdoSidebarGroup>
				</ArdoSidebarSection>
				<ArdoSidebarSection id="api-reference" label="API Reference" to="/api-reference">
					<ArdoGeneratedSidebar section="api-reference" />
				</ArdoSidebarSection>
			</ArdoSidebar>
		</ArdoRoot>
	);
}
