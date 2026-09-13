/**
 * Single source of truth for the version this server advertises.
 *
 * Read from package.json at runtime instead of duplicated as a literal: the literal that
 * used to live in index.ts drifted (it still said 1.0.0 while the package was 1.1.0), so
 * clients were told the wrong version. `../package.json` resolves to the package root from
 * both layouts — `src/` under tsx and `dist/` once built — so dev and published agree.
 */

import { readFileSync } from "node:fs";

const pkg = JSON.parse(
  readFileSync(new URL("../package.json", import.meta.url), "utf8")
) as { version?: string };

if (typeof pkg.version !== "string" || pkg.version.length === 0) {
  throw new Error("package.json has no version field");
}

export const VERSION: string = pkg.version;

/** Server name as reported to MCP clients and in the startup banner. */
export const SERVER_NAME = "m365-mcp";
