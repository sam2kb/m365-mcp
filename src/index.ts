#!/usr/bin/env node
/**
 * m365-mcp — Production-grade Microsoft 365 MCP server
 *
 * Delegated OAuth device-code flow | Multi-account | Pagination | Rate-limit handling
 *
 * Covers: Email, Calendar, Contacts, OneDrive, Teams, Tasks, Users
 */

import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";

import { GraphClient } from "./graph.js";
import { registerMailTools, mailToolSchemas } from "./tools/mail.js";
import { registerCalendarTools, calendarToolSchemas } from "./tools/calendar.js";
import { registerContactsTools, contactsToolSchemas } from "./tools/contacts.js";
import { registerFilesTools, filesToolSchemas } from "./tools/files.js";
import { registerTeamsTools, teamsToolSchemas } from "./tools/teams.js";
import { registerTasksTools, tasksToolSchemas } from "./tools/tasks.js";
import { registerUsersTools, usersToolSchemas } from "./tools/users.js";

// ─── Config ──────────────────────────────────────────────────────────

const ACCOUNT_NAME = process.env.M365_ACCOUNT || undefined;
const TIMEZONE = process.env.M365_TIMEZONE || "UTC";

// ─── Init ────────────────────────────────────────────────────────────

const client = new GraphClient(TIMEZONE, ACCOUNT_NAME);

const mailTools = registerMailTools(client);
const calendarTools = registerCalendarTools(client);
const contactsTools = registerContactsTools(client);
const filesTools = registerFilesTools(client);
const teamsTools = registerTeamsTools(client);
const tasksTools = registerTasksTools(client);
const usersTools = registerUsersTools(client);

// ─── Tool catalog ────────────────────────────────────────────────────

const allTools = [
  ...mailToolSchemas,
  ...calendarToolSchemas,
  ...contactsToolSchemas,
  ...filesToolSchemas,
  ...teamsToolSchemas,
  ...tasksToolSchemas,
  ...usersToolSchemas,
];

// ─── Tool router ─────────────────────────────────────────────────────

async function handleToolCall(name: string, args: Record<string, unknown>): Promise<string> {
  // Mail
  if (name === "m365_mail_list") return mailTools.mail_list(args as any);
  if (name === "m365_mail_read") return mailTools.mail_read(args as any);
  if (name === "m365_mail_send") return mailTools.mail_send(args as any);
  if (name === "m365_mail_reply") return mailTools.mail_reply(args as any);
  if (name === "m365_mail_search") return mailTools.mail_search(args as any);
  if (name === "m365_mail_move") return mailTools.mail_move(args as any);
  if (name === "m365_mail_delete") return mailTools.mail_delete(args as any);
  if (name === "m365_mail_mark_read") return mailTools.mail_mark_read(args as any);
  if (name === "m365_mail_folders") return mailTools.mail_folders();

  // Calendar
  if (name === "m365_calendar_list") return calendarTools.calendar_list(args as any);
  if (name === "m365_calendar_today") return calendarTools.calendar_today();
  if (name === "m365_calendar_week") return calendarTools.calendar_week();
  if (name === "m365_calendar_create") return calendarTools.calendar_create(args as any);
  if (name === "m365_calendar_update") return calendarTools.calendar_update(args as any);
  if (name === "m365_calendar_delete") return calendarTools.calendar_delete(args as any);
  if (name === "m365_calendar_availability") return calendarTools.calendar_availability(args as any);

  // Contacts
  if (name === "m365_contacts_list") return contactsTools.contacts_list(args as any);
  if (name === "m365_contacts_search") return contactsTools.contacts_search(args as any);
  if (name === "m365_contacts_read") return contactsTools.contacts_read(args as any);
  if (name === "m365_contacts_create") return contactsTools.contacts_create(args as any);
  if (name === "m365_contacts_update") return contactsTools.contacts_update(args as any);
  if (name === "m365_contacts_delete") return contactsTools.contacts_delete(args as any);

  // Files
  if (name === "m365_files_list") return filesTools.files_list(args as any);
  if (name === "m365_files_search") return filesTools.files_search(args as any);
  if (name === "m365_files_read") return filesTools.files_read(args as any);
  if (name === "m365_files_info") return filesTools.files_info(args as any);
  if (name === "m365_files_create_folder") return filesTools.files_create_folder(args as any);

  // Teams
  if (name === "m365_teams_chats") return teamsTools.teams_chats(args as any);
  if (name === "m365_teams_messages") return teamsTools.teams_messages(args as any);
  if (name === "m365_teams_send") return teamsTools.teams_send(args as any);

  // Tasks
  if (name === "m365_tasks_lists") return tasksTools.tasks_lists();
  if (name === "m365_tasks_list") return tasksTools.tasks_list(args as any);
  if (name === "m365_tasks_create") return tasksTools.tasks_create(args as any);
  if (name === "m365_tasks_update") return tasksTools.tasks_update(args as any);
  if (name === "m365_tasks_delete") return tasksTools.tasks_delete(args as any);

  // Users
  if (name === "m365_users_list") return usersTools.users_list(args as any);
  if (name === "m365_users_profile") return usersTools.users_profile(args as any);
  if (name === "m365_users_manager") return usersTools.users_manager(args as any);

  throw new Error(`Unknown tool: ${name}`);
}

// ─── MCP Server ──────────────────────────────────────────────────────

const server = new Server(
  { name: "m365-mcp", version: "1.0.0" },
  { capabilities: { tools: {} } }
);

server.setRequestHandler(ListToolsRequestSchema, async () => ({
  tools: allTools,
}));

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  try {
    const result = await handleToolCall(
      request.params.name,
      (request.params.arguments ?? {}) as Record<string, unknown>
    );
    return { content: [{ type: "text", text: result }] };
  } catch (error: any) {
    return {
      content: [{ type: "text", text: `Error: ${error.message}` }],
      isError: true,
    };
  }
});

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error(
    `m365-mcp v1.0.0 running (account: ${ACCOUNT_NAME ?? "default"}, tz: ${TIMEZONE})`
  );
}

main().catch(console.error);
