import { readFileSync } from "node:fs";

// src/ and dist/ are both one level below the package manifest.
export const VERSION = (JSON.parse(readFileSync(new URL("../package.json", import.meta.url), "utf8")) as { version: string }).version;
