import { describe, expect, it } from "vitest";
import pkg from "../package.json" with { type: "json" };
import { version } from "./index.js";

describe("version", () => {
	it("matches the package version", () => {
		expect(version).toBe(pkg.version);
	});
});
