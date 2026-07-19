/**
 * Runtime safety controls and MCP tool annotations.
 */

export interface ToolDefinition {
  name: string;
  description?: string;
  [key: string]: unknown;
}

export interface ToolSafetyAnnotations {
  readOnlyHint: boolean;
  destructiveHint: boolean;
  idempotentHint: boolean;
  openWorldHint: boolean;
}

const READ_ONLY_TOOLS = new Set([
  "m365_mail_list",
  "m365_mail_read",
  "m365_mail_search",
  "m365_mail_folders",
  "m365_calendar_list",
  "m365_calendar_today",
  "m365_calendar_week",
  "m365_calendar_availability",
  "m365_contacts_list",
  "m365_contacts_search",
  "m365_contacts_read",
  "m365_contacts_folders_list",
  "m365_files_list",
  "m365_files_search",
  "m365_files_read",
  "m365_files_info",
  "m365_teams_chats",
  "m365_teams_messages",
  "m365_tasks_lists",
  "m365_tasks_list",
  "m365_users_list",
  "m365_users_profile",
  "m365_users_manager",
]);

const KNOWN_MUTATING_TOOLS = new Set([
  "m365_mail_send",
  "m365_mail_reply",
  "m365_mail_move",
  "m365_mail_delete",
  "m365_mail_mark_read",
  "m365_calendar_create",
  "m365_calendar_update",
  "m365_calendar_delete",
  "m365_contacts_create",
  "m365_contacts_update",
  "m365_contacts_delete",
  "m365_contacts_folder_create",
  "m365_contacts_folder_update",
  "m365_contacts_folder_delete",
  "m365_files_create_folder",
  "m365_teams_send",
  "m365_tasks_create",
  "m365_tasks_update",
  "m365_tasks_delete",
]);

const DESTRUCTIVE_TOOLS = new Set([
  "m365_mail_move",
  "m365_mail_delete",
  "m365_mail_mark_read",
  "m365_calendar_update",
  "m365_calendar_delete",
  "m365_contacts_update",
  "m365_contacts_delete",
  "m365_contacts_folder_update",
  "m365_contacts_folder_delete",
  "m365_tasks_update",
  "m365_tasks_delete",
]);

const IDEMPOTENT_TOOLS = new Set([
  "m365_mail_mark_read",
  "m365_calendar_update",
  "m365_calendar_delete",
  "m365_contacts_update",
  "m365_contacts_delete",
  "m365_contacts_folder_update",
  "m365_contacts_folder_delete",
  "m365_tasks_update",
  "m365_tasks_delete",
]);

export function envFlag(value: string | undefined): boolean {
  return /^(1|true|yes|on)$/i.test(value?.trim() ?? "");
}

export function isMutatingTool(name: string): boolean {
  return !READ_ONLY_TOOLS.has(name);
}

export function toolSafetyAnnotations(name: string): ToolSafetyAnnotations {
  const mutating = isMutatingTool(name);
  const unknown = mutating && !KNOWN_MUTATING_TOOLS.has(name);
  return {
    readOnlyHint: !mutating,
    destructiveHint: unknown || DESTRUCTIVE_TOOLS.has(name),
    idempotentHint: !mutating || IDEMPOTENT_TOOLS.has(name),
    openWorldHint: true,
  };
}

export function annotateTools<T extends ToolDefinition>(
  tools: readonly T[],
): Array<
  T & {
    annotations: ToolSafetyAnnotations;
  }
> {
  return tools.map((tool) => ({
    ...tool,
    annotations: toolSafetyAnnotations(tool.name),
  }));
}

export function toolsForMode<T extends ToolDefinition>(
  tools: readonly T[],
  readOnly: boolean,
): T[] {
  return readOnly
    ? tools.filter((tool) => !isMutatingTool(tool.name))
    : [...tools];
}

export function assertToolAllowed(name: string, readOnly: boolean): void {
  if (readOnly && isMutatingTool(name)) {
    throw new Error(
      `Tool "${name}" is disabled because M365_MCP_READ_ONLY is enabled`,
    );
  }
}
