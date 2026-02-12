/// <reference types="vite/client" />

declare const __BUILD_TIME__: string;

declare module "virtual:ardo/config" {
	import type { PressConfig } from "ardo";
	const config: PressConfig;
	export default config;
}

declare module "virtual:ardo/sidebar" {
	import type { SidebarItem } from "ardo";
	const sidebar: SidebarItem[];
	export default sidebar;
}
