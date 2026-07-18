/**
 * Tests for Teams, Tasks, and Users tools
 */

import { describe, it, expect, vi, beforeEach } from "vitest";
import type { GraphClient } from "../graph.js";

function mockClient(overrides: Partial<GraphClient> = {}): GraphClient {
  return {
    timezone: "UTC",
    getAll: vi.fn().mockResolvedValue([]),
    get: vi.fn(),
    post: vi.fn(),
    patch: vi.fn(),
    delete: vi.fn(),
    paginate: vi.fn(),
    download: vi.fn(),
    dateTimeStr: vi.fn((iso: string) => ({ dateTime: iso, timeZone: "UTC" })),
    ...overrides,
  } as unknown as GraphClient;
}

// ─── Teams ───────────────────────────────────────────────────────────

import { registerTeamsTools, teamsToolSchemas } from "../tools/teams.js";

describe("teams tools", () => {
  let client: GraphClient;
  let tools: ReturnType<typeof registerTeamsTools>;

  beforeEach(() => {
    client = mockClient();
    tools = registerTeamsTools(client);
  });

  describe("teams_chats", () => {
    it("lists chats with member names", async () => {
      (client.getAll as any).mockResolvedValue([
        { id: "chat1", topic: "Project X", chatType: "group", members: [{ displayName: "Alice" }, { displayName: "Bob" }], lastUpdatedDateTime: "2026-07-18T10:00:00Z" },
        { id: "chat2", topic: null, chatType: "oneOnOne", members: [{ displayName: "Charlie" }], lastUpdatedDateTime: "2026-07-17T10:00:00Z" },
      ]);
      const result = await tools.teams_chats({ top: 10 });
      const parsed = JSON.parse(result);
      expect(parsed[0]).toMatchObject({ id: "chat1", topic: "Project X", type: "group" });
      expect(parsed[1]).toMatchObject({ id: "chat2", topic: "(no topic)", type: "oneOnOne" });
    });
  });

  describe("teams_messages", () => {
    it("lists messages with sender", async () => {
      (client.getAll as any).mockResolvedValue([
        { id: "msg1", from: { user: { displayName: "Alice" } }, body: { content: "Hello!", contentType: "text" }, createdDateTime: "2026-07-18T10:00:00Z" },
      ]);
      const result = await tools.teams_messages({ chatId: "chat1" });
      const parsed = JSON.parse(result);
      expect(parsed[0]).toMatchObject({ id: "msg1", from: "Alice", content: "Hello!" });
    });
  });

  describe("teams_send", () => {
    it("sends message with text content type", async () => {
      (client.post as any).mockResolvedValue({ id: "new-msg" });
      await tools.teams_send({ chatId: "chat1", message: "Hi team!" });
      expect(client.post).toHaveBeenCalledWith("/me/chats/chat1/messages", {
        body: { contentType: "text", content: "Hi team!" },
      });
    });
  });

  it("has 3 teams tool schemas", () => {
    expect(teamsToolSchemas).toHaveLength(3);
  });
});

// ─── Tasks ───────────────────────────────────────────────────────────

import { registerTasksTools, tasksToolSchemas } from "../tools/tasks.js";

describe("tasks tools", () => {
  let client: GraphClient;
  let tools: ReturnType<typeof registerTasksTools>;

  beforeEach(() => {
    client = mockClient();
    tools = registerTasksTools(client);
  });

  describe("tasks_lists", () => {
    it("lists task lists", async () => {
      (client.getAll as any).mockResolvedValue([
        { id: "list1", displayName: "Tasks", wellknownListName: "defaultList", isOwner: true },
      ]);
      const result = await tools.tasks_lists();
      const parsed = JSON.parse(result);
      expect(parsed[0]).toMatchObject({ id: "list1", name: "Tasks", wellknown: "defaultList" });
    });
  });

  describe("tasks_list", () => {
    it("lists tasks with filter", async () => {
      (client.getAll as any).mockResolvedValue([]);
      await tools.tasks_list({ listId: "list1", filter: "status ne 'completed'" });
      const path = (client.getAll as any).mock.calls[0][0] as string;
      expect(path).toContain("$filter=status%20ne%20'completed'");
    });

    it("formats tasks", async () => {
      (client.getAll as any).mockResolvedValue([
        { id: "t1", title: "Buy milk", status: "notStarted", importance: "normal", dueDateTime: { dateTime: "2026-07-20", timeZone: "UTC" }, body: { content: "2%" }, createdDateTime: "2026-07-18T10:00:00Z" },
      ]);
      const result = await tools.tasks_list({ listId: "list1" });
      const parsed = JSON.parse(result);
      expect(parsed[0]).toMatchObject({ id: "t1", title: "Buy milk", status: "notStarted" });
    });
  });

  describe("tasks_create", () => {
    it("creates task with due date and importance", async () => {
      (client.post as any).mockResolvedValue({ id: "new", title: "Task", status: "notStarted" });
      await tools.tasks_create({ listId: "list1", title: "Task", due: "2026-07-25T00:00:00", importance: "high", body: "Notes" });
      expect(client.post).toHaveBeenCalledWith("/me/todo/lists/list1/tasks", expect.objectContaining({
        title: "Task",
        importance: "high",
        dueDateTime: { dateTime: "2026-07-25T00:00:00", timeZone: "UTC" },
        body: { contentType: "text", content: "Notes" },
      }));
    });
  });

  describe("tasks_update", () => {
    it("updates task status", async () => {
      await tools.tasks_update({ listId: "list1", taskId: "t1", status: "completed" });
      expect(client.patch).toHaveBeenCalledWith("/me/todo/lists/list1/tasks/t1", { status: "completed" });
    });
  });

  describe("tasks_delete", () => {
    it("deletes task", async () => {
      await tools.tasks_delete({ listId: "list1", taskId: "t1" });
      expect(client.delete).toHaveBeenCalledWith("/me/todo/lists/list1/tasks/t1");
    });
  });

  it("has 5 tasks tool schemas", () => {
    expect(tasksToolSchemas).toHaveLength(5);
  });
});

// ─── Users ───────────────────────────────────────────────────────────

import { registerUsersTools, usersToolSchemas } from "../tools/users.js";

describe("users tools", () => {
  let client: GraphClient;
  let tools: ReturnType<typeof registerUsersTools>;

  beforeEach(() => {
    client = mockClient();
    tools = registerUsersTools(client);
  });

  describe("users_list", () => {
    it("lists org users with search", async () => {
      (client.getAll as any).mockResolvedValue([]);
      await tools.users_list({ search: "Alice", filter: "department eq 'Engineering'" });
      const path = (client.getAll as any).mock.calls[0][0] as string;
      expect(path).toContain('$search="Alice"');
      expect(path).toContain("$filter=department%20eq%20'Engineering'");
    });
  });

  describe("users_profile", () => {
    it("gets own profile by default", async () => {
      (client.get as any).mockResolvedValue({ id: "me", displayName: "Me", mail: "me@b.com" });
      await tools.users_profile({});
      expect(client.get).toHaveBeenCalledWith(expect.stringContaining("/me?"));
    });

    it("gets another user's profile", async () => {
      (client.get as any).mockResolvedValue({ id: "u1", displayName: "Alice" });
      await tools.users_profile({ user: "alice@b.com" });
      expect(client.get).toHaveBeenCalledWith(expect.stringContaining("/users/alice%40b.com?"));
    });
  });

  describe("users_manager", () => {
    it("gets own manager", async () => {
      (client.get as any).mockResolvedValue({ id: "mgr", displayName: "Boss" });
      const result = await tools.users_manager({});
      const parsed = JSON.parse(result);
      expect(parsed.displayName).toBe("Boss");
    });
  });

  it("has 3 users tool schemas", () => {
    expect(usersToolSchemas).toHaveLength(3);
  });
});
