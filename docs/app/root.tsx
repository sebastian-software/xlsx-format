import { Links, Meta, Outlet, Scripts, ScrollRestoration, useLocation } from "react-router"
import {
  Layout as ArdoLayout,
  Header,
  Nav,
  NavLink,
  Sidebar,
  SidebarGroup,
  SidebarLink,
  Footer,
} from "ardo/ui"
import { PressProvider } from "ardo/runtime"
import config from "virtual:ardo/config"
import sidebar from "virtual:ardo/sidebar"
import "ardo/ui/styles.css"

export function Layout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en" suppressHydrationWarning>
      <head>
        <meta charSet="utf-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <Meta />
        <Links />
      </head>
      <body suppressHydrationWarning>
        {children}
        <ScrollRestoration />
        <Scripts />
      </body>
    </html>
  )
}

export default function Root() {
  const location = useLocation()
  const isHomePage = location.pathname === "/"

  return (
    <PressProvider config={config} sidebar={sidebar}>
      <ArdoLayout
        className={isHomePage ? "ardo-layout ardo-home" : "ardo-layout"}
        header={
          <Header
            title="XLSX Format"
            nav={
              <Nav>
                <NavLink to="/guide/getting-started">Guide</NavLink>
                <NavLink to="/api-reference">API</NavLink>
              </Nav>
            }
          />
        }
        sidebar={
          isHomePage ? undefined : (
            <Sidebar>
              <SidebarGroup title="Guide">
                <SidebarLink to="/guide/getting-started">Getting Started</SidebarLink>
              </SidebarGroup>
              <SidebarLink to="/api-reference">API Reference</SidebarLink>
            </Sidebar>
          )
        }
        footer={
          <Footer
            message={[
              config.project?.homepage
                ? `<a href="${config.project.homepage}">${config.title}</a>`
                : config.title,
              "Built with <a href='https://github.com/sebastian-software/ardo'>Ardo</a>",
            ].join(" &middot; ")}
            copyright={
              config.project?.author
                ? `Copyright &copy; ${new Date().getFullYear()} ${config.project.author}`
                : undefined
            }
          />
        }
      >
        <Outlet />
      </ArdoLayout>
    </PressProvider>
  )
}
