import { describe, expect, it } from "vitest";
import {
  annotateTools,
  assertToolAllowed,
  envFlag,
  toolSafetyAnnotations,
  toolsForMode,
} from "../security.js";

describe("security controls", () => {
  it("parses only explicit enabled values", () => {
    expect(envFlag("true")).toBe(true);
    expect(envFlag("YES")).toBe(true);
    expect(envFlag("1")).toBe(true);
    expect(envFlag("false")).toBe(false);
    expect(envFlag(undefined)).toBe(false);
  });

  it("marks read, outbound, and destructive tools accurately", () => {
    expect(toolSafetyAnnotations("m365_mail_list")).toMatchObject({
      readOnlyHint: true,
      destructiveHint: false,
      openWorldHint: true,
    });
    expect(toolSafetyAnnotations("m365_contacts_people_list")).toMatchObject({
      readOnlyHint: true,
      destructiveHint: false,
      openWorldHint: true,
    });
    expect(toolSafetyAnnotations("m365_teams_send")).toMatchObject({
      readOnlyHint: false,
      destructiveHint: false,
      idempotentHint: false,
    });
    expect(toolSafetyAnnotations("m365_mail_delete")).toMatchObject({
      readOnlyHint: false,
      destructiveHint: true,
    });
    expect(toolSafetyAnnotations("m365_calendar_cancel")).toMatchObject({
      readOnlyHint: false,
      destructiveHint: true,
      idempotentHint: false,
    });
    expect(toolSafetyAnnotations("m365_calendar_delete")).toMatchObject({
      readOnlyHint: false, destructiveHint: true, idempotentHint: false,
    });
    expect(toolSafetyAnnotations("m365_calendar_update")).toMatchObject({
      readOnlyHint: false, destructiveHint: true, idempotentHint: false,
    });
    expect(toolSafetyAnnotations("m365_calendar_decline")).toMatchObject({
      readOnlyHint: false,
      destructiveHint: true,
      idempotentHint: false,
    });
    expect(toolSafetyAnnotations("m365_future_tool")).toMatchObject({
      readOnlyHint: false,
      destructiveHint: true,
      idempotentHint: false,
    });
  });

  it("filters and blocks mutating tools in read-only mode", () => {
    const tools = annotateTools([
      { name: "m365_mail_list" },
      { name: "m365_calendar_get" },
      { name: "m365_mail_send" },
      { name: "m365_mail_delete" },
      { name: "m365_calendar_cancel" },
      { name: "m365_calendar_decline" },
    ]);

    expect(toolsForMode(tools, true).map((tool) => tool.name)).toEqual([
      "m365_mail_list",
      "m365_calendar_get",
    ]);
    expect(() => assertToolAllowed("m365_mail_send", true)).toThrow(
      "M365_MCP_READ_ONLY",
    );
    expect(() => assertToolAllowed("m365_mail_delete", true)).toThrow(
      "M365_MCP_READ_ONLY",
    );
    expect(() => assertToolAllowed("m365_calendar_cancel", true)).toThrow(
      "M365_MCP_READ_ONLY",
    );
    expect(() => assertToolAllowed("m365_calendar_decline", true)).toThrow(
      "M365_MCP_READ_ONLY",
    );
    expect(() => assertToolAllowed("m365_future_tool", true)).toThrow(
      "M365_MCP_READ_ONLY",
    );
    expect(() => assertToolAllowed("m365_mail_list", true)).not.toThrow();
    expect(() => assertToolAllowed("m365_calendar_get", true)).not.toThrow();
    expect(toolSafetyAnnotations("m365_calendar_get")).toEqual({
      readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: true,
    });
  });

  it("preserves all tools when read-only mode is disabled", () => {
    const tools = [{ name: "m365_mail_list" }, { name: "m365_mail_send" }];
    expect(toolsForMode(tools, false)).toEqual(tools);
  });
});
