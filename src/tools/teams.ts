/**
 * Teams tools — list chats, list messages, send message
 */

import type { TeamsChat, TeamsMessage } from "../types.js";
import type { GraphClient } from "../graph.js";

export function registerTeamsTools(client: GraphClient) {
  const user = "/me";

  return {
    async teams_chats(args: { top?: number }): Promise<string> {
      const path = `${user}/chats?$top=${args.top || 20}&$expand=members&$orderby=lastUpdatedDateTime desc`;
      const chats = await client.getAll<TeamsChat>(path);
      return JSON.stringify(
        chats.map((c) => ({
          id: c.id,
          topic: c.topic || "(no topic)",
          type: c.chatType,
          members: c.members?.map((m) => m.displayName) ?? [],
          lastUpdated: c.lastUpdatedDateTime,
        })),
        null,
        2
      );
    },

    async teams_messages(args: { chatId: string; top?: number }): Promise<string> {
      const path = `${user}/chats/${args.chatId}/messages?$top=${args.top || 30}&$orderby=createdDateTime desc`;
      const msgs = await client.getAll<TeamsMessage>(path);
      return JSON.stringify(
        msgs.map((m) => ({
          id: m.id,
          from: m.from?.user?.displayName ?? "Unknown",
          content: m.body?.content?.slice(0, 2000),
          contentType: m.body?.contentType,
          sent: m.createdDateTime,
        })),
        null,
        2
      );
    },

    async teams_send(args: { chatId: string; message: string }): Promise<string> {
      const data = await client.post<any>(
        `${user}/chats/${args.chatId}/messages`,
        { body: { contentType: "text", content: args.message } }
      );
      return JSON.stringify({ id: data.id, success: true }, null, 2);
    },
  };
}

export const teamsToolSchemas = [
  {
    name: "m365_teams_chats",
    description: "List your Teams chats",
    inputSchema: {
      type: "object",
      properties: {
        top: { type: "number", description: "Max chats", default: 20 },
      },
    },
  },
  {
    name: "m365_teams_messages",
    description: "Get messages from a Teams chat",
    inputSchema: {
      type: "object",
      properties: {
        chatId: { type: "string", description: "Chat ID (from m365_teams_chats)" },
        top: { type: "number", description: "Max messages", default: 30 },
      },
      required: ["chatId"],
    },
  },
  {
    name: "m365_teams_send",
    description: "Send a message to a Teams chat",
    inputSchema: {
      type: "object",
      properties: {
        chatId: { type: "string", description: "Chat ID" },
        message: { type: "string", description: "Message text" },
      },
      required: ["chatId", "message"],
    },
  },
];
