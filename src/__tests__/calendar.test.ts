/**
 * Tests for calendar tools
 */

import { describe, it, expect, vi, beforeEach } from "vitest";
import type { GraphClient } from "../graph.js";
import { registerCalendarTools, calendarToolSchemas } from "../tools/calendar.js";

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

describe("calendar tools", () => {
  let client: GraphClient;
  let tools: ReturnType<typeof registerCalendarTools>;

  beforeEach(() => {
    client = mockClient();
    tools = registerCalendarTools(client);
  });

  describe("calendar_list", () => {
    it("uses calendarView when date range provided", async () => {
      (client.getAll as any).mockResolvedValue([]);
      await tools.calendar_list({ startDateTime: "2026-07-01T00:00:00Z", endDateTime: "2026-07-31T23:59:59Z" });
      const path = (client.getAll as any).mock.calls[0][0] as string;
      expect(path).toContain("calendarView");
      expect(path).toContain("startDateTime=");
      expect(path).toContain("endDateTime=");
    });

    it("uses events endpoint when no date range", async () => {
      (client.getAll as any).mockResolvedValue([]);
      await tools.calendar_list({});
      const path = (client.getAll as any).mock.calls[0][0] as string;
      expect(path).toContain("/events");
    });

    it("formats events with key fields", async () => {
      (client.getAll as any).mockResolvedValue([
        {
          id: "ev1",
          subject: "Meeting",
          start: { dateTime: "2026-07-18T10:00:00.0000000", timeZone: "UTC" },
          end: { dateTime: "2026-07-18T11:00:00.0000000", timeZone: "UTC" },
          isAllDay: false,
          location: { displayName: "Room 1" },
          isOnlineMeeting: true,
          onlineMeetingUrl: "https://teams.microsoft.com/...",
          organizer: { emailAddress: { name: "Alice", address: "alice@b.com" } },
          attendees: [{}, {}],
          webLink: "https://outlook.office.com/...",
        },
      ]);
      const result = await tools.calendar_list({});
      const parsed = JSON.parse(result);
      expect(parsed[0]).toMatchObject({
        id: "ev1",
        subject: "Meeting",
        isAllDay: false,
        location: "Room 1",
        isOnline: true,
        organizer: "Alice",
        attendees: 2,
      });
    });
  });

  describe("calendar_create", () => {
    it("creates event with attendees and online meeting", async () => {
      (client.post as any).mockResolvedValue({
        id: "new-ev",
        subject: "Sync",
        webLink: "https://...",
        onlineMeetingUrl: null,
      });
      const result = await tools.calendar_create({
        subject: "Sync",
        start: "2026-07-18T14:00:00",
        end: "2026-07-18T15:00:00",
        body: "Discuss Q3",
        location: "Board Room",
        attendees: "a@b.com, c@d.com",
        isOnline: true,
      });
      expect(client.post).toHaveBeenCalledWith("/me/events", expect.objectContaining({
        subject: "Sync",
        isOnlineMeeting: true,
        attendees: [
          { emailAddress: { address: "a@b.com" }, type: "required" },
          { emailAddress: { address: "c@d.com" }, type: "required" },
        ],
      }));
      const parsed = JSON.parse(result);
      expect(parsed.id).toBe("new-ev");
    });
  });

  describe("calendar_update", () => {
    it("patches only provided fields", async () => {
      await tools.calendar_update({ eventId: "ev1", subject: "Updated", location: "Room 2" });
      expect(client.patch).toHaveBeenCalledWith("/me/events/ev1", {
        subject: "Updated",
        location: { displayName: "Room 2" },
      });
    });
  });

  describe("calendar_delete", () => {
    it("cancels event with comment", async () => {
      await tools.calendar_delete({ eventId: "ev1", comment: "No longer needed" });
      expect(client.post).toHaveBeenCalledWith("/me/events/ev1/cancel", { comment: "No longer needed" });
    });
  });

  describe("calendar_availability", () => {
    it("queries free/busy for multiple users", async () => {
      (client.post as any).mockResolvedValue({ value: [] });
      await tools.calendar_availability({
        users: "a@b.com, c@d.com",
        startDateTime: "2026-07-18T00:00:00",
        endDateTime: "2026-07-18T23:59:59",
      });
      expect(client.post).toHaveBeenCalledWith("/me/calendar/getSchedule", expect.objectContaining({
        schedules: ["a@b.com", "c@d.com"],
      }));
    });
  });

  describe("tool schemas", () => {
    it("has 7 calendar tool schemas", () => {
      expect(calendarToolSchemas).toHaveLength(7);
    });
  });
});
