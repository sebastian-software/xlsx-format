import { disableRule, getEslintConfig } from "eslint-config-setup";

const config = await getEslintConfig({ node: true });

// Keep this PR focused on adopting the shared config. The ported SheetJS code can
// enable these stricter rule families incrementally in later source-focused PRs.
const portedCodeDisabledRules = [
	"@cspell/spellchecker",
	"@typescript-eslint/array-type",
	"@typescript-eslint/consistent-type-definitions",
	"@typescript-eslint/no-unsafe-type-assertion",
	"@typescript-eslint/prefer-for-of",
	"@typescript-eslint/prefer-includes",
	"@typescript-eslint/prefer-nullish-coalescing",
	"@typescript-eslint/prefer-optional-chain",
	"@typescript-eslint/prefer-regexp-exec",
	"@typescript-eslint/prefer-string-starts-ends-with",
	"@typescript-eslint/strict-boolean-expressions",
	"@typescript-eslint/switch-exhaustiveness-check",
	"complexity",
	"import/no-mutable-exports",
	"jsdoc/check-param-names",
	"jsdoc/require-throws-type",
	"max-depth",
	"max-lines",
	"max-lines-per-function",
	"max-params",
	"max-statements",
	"no-script-url",
	"perfectionist/sort-exports",
	"perfectionist/sort-imports",
	"perfectionist/sort-intersection-types",
	"perfectionist/sort-named-exports",
	"perfectionist/sort-named-imports",
	"perfectionist/sort-union-types",
	"prefer-template",
	"regexp/no-control-character",
	"regexp/no-super-linear-move",
	"regexp/no-unused-capturing-group",
	"security/detect-non-literal-regexp",
	"security/detect-non-literal-fs-filename",
	"security/detect-unsafe-regex",
	"sonarjs/cognitive-complexity",
	"sonarjs/no-duplicated-branches",
	"unicorn/consistent-function-scoping",
	"unicorn/no-for-loop",
	"unicorn/numeric-separators-style",
	"unicorn/prefer-array-find",
	"unicorn/prefer-includes",
	"unicorn/prefer-number-properties",
	"unicorn/prefer-regexp-test",
	"unicorn/prefer-spread",
	"unicorn/prefer-string-replace-all",
	"unicorn/prefer-string-slice",
	"vitest/expect-expect",
	"vitest/no-conditional-expect",
	"vitest/no-conditional-in-test",
	"vitest/no-identical-title",
	"vitest/prefer-strict-equal",
	"vitest/prefer-to-be",
	"vitest/prefer-to-have-length",
	"vitest/require-to-throw-message",
];

for (const rule of portedCodeDisabledRules) {
	disableRule(config, rule);
}

config.unshift({
	ignores: ["dist/", "node_modules/"],
});

config.push({
	name: "xlsx-format/ported-code",
	files: ["src/**/*.ts"],
	languageOptions: {
		parserOptions: {
			project: "./tsconfig.eslint.json",
			projectService: false,
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

		// Ported code keeps SheetJS-style mutation and state assignments
		"@typescript-eslint/no-unnecessary-condition": "off",
		"no-useless-assignment": "off",

		// Too strict for ported code
		"@typescript-eslint/restrict-template-expressions": "off",
		"@typescript-eslint/restrict-plus-operands": "off",

		// Always require curly braces
		curly: "error",

		// Unused vars as warning only
		"@typescript-eslint/no-unused-vars": ["warn", { argsIgnorePattern: "^_", varsIgnorePattern: "^_" }],
	},
});

export default config;
