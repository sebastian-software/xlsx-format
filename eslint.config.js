import eslint from "@eslint/js";
import tseslint from "typescript-eslint";
import prettier from "eslint-config-prettier";

export default tseslint.config(
	{
		ignores: ["dist/", "node_modules/"],
	},
	eslint.configs.recommended,
	...tseslint.configs.strictTypeChecked,
	prettier,
	{
		languageOptions: {
			parserOptions: {
				project: "./tsconfig.eslint.json",
				tsconfigRootDir: import.meta.dirname,
			},
		},
		rules: {
			// Required for ported code — many intentional any casts
			"@typescript-eslint/no-explicit-any": "off",
			"@typescript-eslint/no-unsafe-argument": "off",
			"@typescript-eslint/no-unsafe-assignment": "off",
			"@typescript-eslint/no-unsafe-call": "off",
			"@typescript-eslint/no-unsafe-member-access": "off",
			"@typescript-eslint/no-unsafe-return": "off",

			// Ported code uses non-null assertions intentionally
			"@typescript-eslint/no-non-null-assertion": "off",

			// Switch fallthrough is intentional (SSF, shared-strings)
			"no-fallthrough": "off",

			// Allow empty catch blocks for safe_format etc.
			"@typescript-eslint/no-empty-function": "off",
			"no-empty": ["error", { allowEmptyCatch: true }],

			// Ported code uses parameter reassignment
			"@typescript-eslint/no-unnecessary-condition": "off",

			// Too strict for ported code
			"@typescript-eslint/restrict-template-expressions": "off",
			"@typescript-eslint/restrict-plus-operands": "off",

			// Always require curly braces
			curly: "error",

			// Unused vars as warning only
			"@typescript-eslint/no-unused-vars": [
				"warn",
				{ argsIgnorePattern: "^_", varsIgnorePattern: "^_" },
			],
		},
	},
);
