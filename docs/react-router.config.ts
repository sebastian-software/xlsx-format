import type { Config } from "@react-router/dev/config";
import { withArdoGitHubPages } from "ardo/vite";

export default withArdoGitHubPages({
	ssr: false,
	prerender: true,
} satisfies Config);
