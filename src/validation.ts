import { AjvJsonSchemaValidator } from "@modelcontextprotocol/sdk/validation/ajv";
import type { JsonSchemaType } from "@modelcontextprotocol/sdk/validation";

/** Shared result limits also protect callers using the tool helpers directly. */
export function normalizeTop(value: number | undefined, fallback: number): number {
  if (value === undefined) return fallback;
  if (!Number.isInteger(value) || value < 1 || value > 500) {
    throw new Error("top must be an integer between 1 and 500");
  }
  return value;
}

/** Compile the advertised schemas once and enforce them before tool dispatch. */
export function createArgumentValidator(tools: readonly { name: string; inputSchema: unknown }[]) {
  const provider = new AjvJsonSchemaValidator();
  const validators = new Map(tools.map((tool) => [
    tool.name, provider.getValidator(tool.inputSchema as JsonSchemaType),
  ]));
  return (name: string, args: unknown): void => {
    const validator = validators.get(name);
    if (!validator) throw new Error(`Unknown tool: ${name}`);
    const result = validator(args);
    if (!result.valid) throw new Error(`Invalid arguments for ${name}: ${result.errorMessage}`);
  };
}
