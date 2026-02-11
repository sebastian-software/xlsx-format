import { defineConfig } from "vitest/config";

export default defineConfig({
	test: {
		globals: true,
		include: ["src/**/*.test.ts"],
		testTimeout: 30_000,
		coverage: {
			provider: "v8",
			include: ["src/**"],
			exclude: ["src/__fixtures__/**"],
		},
	},
});
