import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { SERVER_NAME, VERSION } from "../version.js";

const pkg = JSON.parse(
  readFileSync(new URL("../../package.json", import.meta.url), "utf8")
) as { name: string; version: string };

describe("advertised version", () => {
  it("comes from package.json and cannot drift from it", () => {
    expect(VERSION).toBe(pkg.version);
  });

  it("is a valid semver", () => {
    expect(VERSION).toMatch(/^\d+\.\d+\.\d+(?:[-+][0-9A-Za-z.-]+)?$/);
  });

  it("keeps the server name the package name's basename", () => {
    expect(SERVER_NAME).toBe(pkg.name.split("/").pop());
  });
});
