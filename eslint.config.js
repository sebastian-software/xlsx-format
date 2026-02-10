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
			// Nötig für den portierten Code — viele intentionale any-Casts
			"@typescript-eslint/no-explicit-any": "off",
			"@typescript-eslint/no-unsafe-argument": "off",
			"@typescript-eslint/no-unsafe-assignment": "off",
			"@typescript-eslint/no-unsafe-call": "off",
			"@typescript-eslint/no-unsafe-member-access": "off",
			"@typescript-eslint/no-unsafe-return": "off",

			// Portierter Code nutzt non-null assertions bewusst
			"@typescript-eslint/no-non-null-assertion": "off",

			// switch fallthrough ist intentional (SSF, shared-strings)
			"no-fallthrough": "off",

			// Erlaubt: leere catch-Blöcke für safe_format etc.
			"@typescript-eslint/no-empty-function": "off",
			"no-empty": ["error", { allowEmptyCatch: true }],

			// Portierter Code nutzt Parameter-Reassignment
			"@typescript-eslint/no-unnecessary-condition": "off",

			// Restrict template expressions ist zu streng für den portierten Code
			"@typescript-eslint/restrict-template-expressions": "off",
			"@typescript-eslint/restrict-plus-operands": "off",

			// Curly braces immer erforderlich
			curly: "error",

			// Unused vars als Warning
			"@typescript-eslint/no-unused-vars": [
				"warn",
				{ argsIgnorePattern: "^_", varsIgnorePattern: "^_" },
			],
		},
	},
);
