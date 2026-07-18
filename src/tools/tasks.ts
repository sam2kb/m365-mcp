/**
 * Microsoft To Do tasks tools — list lists, list tasks, create, update, delete
 */

import type { TodoList, TodoTask } from "../types.js";
import type { GraphClient } from "../graph.js";

export function registerTasksTools(client: GraphClient) {
  const user = "/me";

  return {
    async tasks_lists(): Promise<string> {
      const lists = await client.getAll<TodoList>(
        `${user}/todo/lists?$select=id,displayName,isOwner,wellknownListName`
      );
      return JSON.stringify(
        lists.map((l) => ({
          id: l.id,
          name: l.displayName,
          wellknown: l.wellknownListName,
          isOwner: l.isOwner,
        })),
        null,
        2
      );
    },

    async tasks_list(args: { listId: string; top?: number; filter?: string }): Promise<string> {
      const top = args.top || 50;
      let path = `${user}/todo/lists/${encodeURIComponent(args.listId)}/tasks?$top=${top}&$orderby=createdDateTime desc`;
      if (args.filter) path += `&$filter=${encodeURIComponent(args.filter)}`;
      const tasks = await client.getAll<TodoTask>(path, 10, top);
      return JSON.stringify(
        tasks.map((t) => ({
          id: t.id,
          title: t.title,
          status: t.status,
          importance: t.importance,
          due: t.dueDateTime,
          body: t.body?.content?.slice(0, 500),
          created: t.createdDateTime,
        })),
        null,
        2
      );
    },

    async tasks_create(args: {
      listId: string;
      title: string;
      body?: string;
      due?: string;
      importance?: string;
    }): Promise<string> {
      const task: any = { title: args.title };
      if (args.importance) task.importance = args.importance;
      if (args.due) {
        task.dueDateTime = { dateTime: args.due, timeZone: "UTC" };
      }
      if (args.body) {
        task.body = { contentType: "text", content: args.body };
      }
      const data = await client.post<any>(
        `${user}/todo/lists/${encodeURIComponent(args.listId)}/tasks`,
        task
      );
      return JSON.stringify(
        { id: data.id, title: data.title, status: data.status },
        null,
        2
      );
    },

    async tasks_update(args: {
      listId: string;
      taskId: string;
      title?: string;
      status?: string;
      importance?: string;
      due?: string;
    }): Promise<string> {
      const patch: any = {};
      if (args.title) patch.title = args.title;
      if (args.status) patch.status = args.status;
      if (args.importance) patch.importance = args.importance;
      if (args.due) patch.dueDateTime = { dateTime: args.due, timeZone: "UTC" };

      const taskId = encodeURIComponent(args.taskId);
      await client.patch(
        `${user}/todo/lists/${encodeURIComponent(args.listId)}/tasks/${taskId}`,
        patch
      );
      return JSON.stringify({ success: true, taskId: args.taskId });
    },

    async tasks_delete(args: { listId: string; taskId: string }): Promise<string> {
      const taskId = encodeURIComponent(args.taskId);
      await client.delete(
        `${user}/todo/lists/${encodeURIComponent(args.listId)}/tasks/${taskId}`
      );
      return JSON.stringify({ success: true });
    },
  };
}

export const tasksToolSchemas = [
  {
    name: "m365_tasks_lists",
    description: "List your Microsoft To Do task lists",
    inputSchema: { type: "object", properties: {} },
  },
  {
    name: "m365_tasks_list",
    description: "List tasks in a task list",
    inputSchema: {
      type: "object",
      properties: {
        listId: { type: "string", description: "Task list ID (from m365_tasks_lists)" },
        top: { type: "number", description: "Max tasks", default: 50 },
        filter: { type: "string", description: "OData filter, e.g. status ne 'completed'" },
      },
      required: ["listId"],
    },
  },
  {
    name: "m365_tasks_create",
    description: "Create a new task",
    inputSchema: {
      type: "object",
      properties: {
        listId: { type: "string", description: "Task list ID" },
        title: { type: "string", description: "Task title" },
        body: { type: "string", description: "Task notes/description" },
        due: { type: "string", description: "Due date (ISO 8601)" },
        importance: { type: "string", description: "low, normal, or high" },
      },
      required: ["listId", "title"],
    },
  },
  {
    name: "m365_tasks_update",
    description: "Update an existing task",
    inputSchema: {
      type: "object",
      properties: {
        listId: { type: "string", description: "Task list ID" },
        taskId: { type: "string", description: "Task ID" },
        title: { type: "string", description: "New title" },
        status: { type: "string", description: "notStarted, inProgress, completed, waitingOnOthers, deferred" },
        importance: { type: "string", description: "low, normal, or high" },
        due: { type: "string", description: "New due date (ISO 8601)" },
      },
      required: ["listId", "taskId"],
    },
  },
  {
    name: "m365_tasks_delete",
    description: "Delete a task",
    inputSchema: {
      type: "object",
      properties: {
        listId: { type: "string", description: "Task list ID" },
        taskId: { type: "string", description: "Task ID" },
      },
      required: ["listId", "taskId"],
    },
  },
];
