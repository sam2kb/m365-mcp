/**
 * Calendar tools — list, today, week, create, update, delete, availability
 */

import type { CalendarEvent } from "../types.js";
import type { GraphClient } from "../graph.js";

/** Build ISO 8601 start/end strings for a day in the given timezone. */
function tzDayBounds(timezone: string, offsetDays = 0): { startDateTime: string; endDateTime: string } {
  const d = new Date(Date.now() + offsetDays * 86400000);
  const dateStr = d.toLocaleDateString("en-CA", { timeZone: timezone }); // "2026-07-18"

  // Determine the UTC offset at noon (DST-safe) via formatToParts
  const parts = Intl.DateTimeFormat("en-US", {
    timeZone: timezone,
    hour12: false,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    timeZoneName: "longOffset",
  }).formatToParts(d);

  let offset = "+00:00";
  const tzPart = parts.find((p) => p.type === "timeZoneName")?.value ?? "";
  const m = tzPart.match(/GMT([+-]\d{1,2}):?(\d{2})?/);
  if (m) offset = `${m[1]}:${m[2] || "00"}`;

  return {
    startDateTime: `${dateStr}T00:00:00${offset}`,
    endDateTime: `${dateStr}T23:59:59${offset}`,
  };
}

export function registerCalendarTools(client: GraphClient) {
  const user = "/me";

  return {
    // ── list ─────────────────────────────────────────────────────
    async calendar_list(args: {
      startDateTime?: string;
      endDateTime?: string;
      top?: number;
    }): Promise<string> {
      const top = args.top || 50;
      let path: string;

      if (args.startDateTime && args.endDateTime) {
        path = `${user}/calendarView?startDateTime=${encodeURIComponent(args.startDateTime)}&endDateTime=${encodeURIComponent(args.endDateTime)}&$top=${top}&$orderby=start/dateTime&$select=id,subject,start,end,location,isOnlineMeeting,onlineMeetingUrl,organizer,attendees,isAllDay,webLink`;
      } else {
        path = `${user}/events?$top=${top}&$orderby=start/dateTime&$select=id,subject,start,end,location,isOnlineMeeting,onlineMeetingUrl,organizer,attendees,isAllDay,webLink`;
      }

      const events = await client.getAll<CalendarEvent>(path, 10, top);
      return JSON.stringify(
        events.map((e) => ({
          id: e.id,
          subject: e.subject,
          start: e.start,
          end: e.end,
          isAllDay: e.isAllDay,
          location: e.location?.displayName,
          isOnline: e.isOnlineMeeting,
          meetingUrl: e.onlineMeetingUrl,
          organizer: e.organizer?.emailAddress?.name ?? e.organizer?.emailAddress?.address,
          attendees: e.attendees?.length ?? 0,
          webLink: e.webLink,
        })),
        null,
        2
      );
    },

    // ── today ────────────────────────────────────────────────────
    async calendar_today(): Promise<string> {
      const bounds = tzDayBounds(client.timezone);

      const path = `${user}/calendarView?startDateTime=${encodeURIComponent(bounds.startDateTime)}&endDateTime=${encodeURIComponent(bounds.endDateTime)}&$top=50&$orderby=start/dateTime`;
      const events = await client.getAll<CalendarEvent>(path, 10, 50);

      if (events.length === 0) return "No events today.";

      const now = new Date();
      const lines = events.map((e) => {
        const startTime = new Date(e.start.dateTime + (e.start.dateTime.endsWith("Z") ? "" : "Z"));
        const endTime = new Date(e.end.dateTime + (e.end.dateTime.endsWith("Z") ? "" : "Z"));
        const time = e.isAllDay
          ? "All Day"
          : `${startTime.toLocaleTimeString("en-US", { hour: "numeric", minute: "2-digit" })}–${endTime.toLocaleTimeString("en-US", { hour: "numeric", minute: "2-digit" })}`;
        let line = `${time}  ${e.subject}`;
        if (e.location?.displayName) line += `  📍 ${e.location.displayName}`;
        if (e.isOnlineMeeting) line += `  🎥`;
        return line;
      });

      return `Today (${now.toLocaleDateString("en-US", { weekday: "long", month: "long", day: "numeric" })})\n${lines.join("\n")}\n\n${events.length} event${events.length !== 1 ? "s" : ""}`;
    },

    // ── week ─────────────────────────────────────────────────────
    async calendar_week(): Promise<string> {
      const startBounds = tzDayBounds(client.timezone, 0);
      const endBounds = tzDayBounds(client.timezone, 6);

      const path = `${user}/calendarView?startDateTime=${encodeURIComponent(startBounds.startDateTime)}&endDateTime=${encodeURIComponent(endBounds.endDateTime)}&$top=100&$orderby=start/dateTime`;
      const events = await client.getAll<CalendarEvent>(path, 10, 100);

      if (events.length === 0) return "No events this week.";

      const byDay = new Map<string, string[]>();
      for (const e of events) {
        const d = new Date(e.start.dateTime + (e.start.dateTime.endsWith("Z") ? "" : "Z"));
        const key = d.toLocaleDateString("en-US", { weekday: "short", month: "short", day: "numeric" });
        const startTime = d.toLocaleTimeString("en-US", { hour: "numeric", minute: "2-digit" });
        let line = `  ${e.isAllDay ? "All Day" : startTime}  ${e.subject}`;
        if (e.location?.displayName) line += `  📍 ${e.location.displayName}`;
        if (e.isOnlineMeeting) line += `  🎥`;
        if (!byDay.has(key)) byDay.set(key, []);
        byDay.get(key)!.push(line);
      }

      let out = "";
      for (const [day, lines] of byDay) {
        out += `\n${day}\n${lines.join("\n")}\n`;
      }
      return `This Week\n${out}\n${events.length} events`;
    },

    // ── create ───────────────────────────────────────────────────
    async calendar_create(args: {
      subject: string;
      start: string;
      end: string;
      body?: string;
      location?: string;
      attendees?: string;
      isOnline?: boolean;
    }): Promise<string> {
      const event: any = {
        subject: args.subject,
        start: client.dateTimeStr(args.start),
        end: client.dateTimeStr(args.end),
      };
      if (args.body) event.body = { contentType: "HTML", content: args.body };
      if (args.location) event.location = { displayName: args.location };
      if (args.attendees) {
        event.attendees = args.attendees.split(",").map((s) => ({
          emailAddress: { address: s.trim() },
          type: "required",
        }));
      }
      if (args.isOnline) event.isOnlineMeeting = true;

      const data = await client.post<any>(`${user}/events`, event);
      return JSON.stringify(
        {
          id: data.id,
          subject: data.subject,
          webLink: data.webLink,
          onlineMeetingUrl: data.onlineMeetingUrl,
        },
        null,
        2
      );
    },

    // ── update ───────────────────────────────────────────────────
    async calendar_update(args: {
      eventId: string;
      subject?: string;
      start?: string;
      end?: string;
      body?: string;
      location?: string;
    }): Promise<string> {
      const patch: any = {};
      if (args.subject) patch.subject = args.subject;
      if (args.start) patch.start = client.dateTimeStr(args.start);
      if (args.end) patch.end = client.dateTimeStr(args.end);
      if (args.body) patch.body = { contentType: "HTML", content: args.body };
      if (args.location) patch.location = { displayName: args.location };

      const eventId = encodeURIComponent(args.eventId);
      await client.patch(`${user}/events/${eventId}`, patch);
      return JSON.stringify({ success: true, eventId: args.eventId });
    },

    // ── delete ───────────────────────────────────────────────────
    async calendar_delete(args: {
      eventId: string;
      comment?: string;
    }): Promise<string> {
      const eventId = encodeURIComponent(args.eventId);
      await client.post(`${user}/events/${eventId}/cancel`, {
        comment: args.comment || "",
      });
      return JSON.stringify({ success: true, eventId: args.eventId });
    },

    // ── availability ─────────────────────────────────────────────
    async calendar_availability(args: {
      users: string;
      startDateTime: string;
      endDateTime: string;
    }): Promise<string> {
      const schedules = args.users.split(",").map((s) => s.trim());
      const data = await client.post<any>(`${user}/calendar/getSchedule`, {
        schedules,
        startTime: client.dateTimeStr(args.startDateTime),
        endTime: client.dateTimeStr(args.endDateTime),
        availabilityViewInterval: 30,
      });
      return JSON.stringify(data.value, null, 2);
    },
  };
}

export const calendarToolSchemas = [
  {
    name: "m365_calendar_list",
    description: "List calendar events in a date range (or upcoming if no range)",
    inputSchema: {
      type: "object",
      properties: {
        startDateTime: { type: "string", description: "Start datetime (ISO 8601)" },
        endDateTime: { type: "string", description: "End datetime (ISO 8601)" },
        top: { type: "number", description: "Max events", default: 50 },
      },
    },
  },
  {
    name: "m365_calendar_today",
    description: "Show today's calendar events",
    inputSchema: { type: "object", properties: {} },
  },
  {
    name: "m365_calendar_week",
    description: "Show this week's calendar events grouped by day",
    inputSchema: { type: "object", properties: {} },
  },
  {
    name: "m365_calendar_create",
    description: "Create a calendar event (with optional Teams meeting)",
    inputSchema: {
      type: "object",
      properties: {
        subject: { type: "string", description: "Event title" },
        start: { type: "string", description: "Start datetime (ISO 8601)" },
        end: { type: "string", description: "End datetime (ISO 8601)" },
        body: { type: "string", description: "Event description (HTML)" },
        location: { type: "string", description: "Event location" },
        attendees: { type: "string", description: "Attendee emails, comma-separated" },
        isOnline: { type: "boolean", description: "Create as Teams meeting" },
      },
      required: ["subject", "start", "end"],
    },
  },
  {
    name: "m365_calendar_update",
    description: "Update an existing calendar event",
    inputSchema: {
      type: "object",
      properties: {
        eventId: { type: "string", description: "Event ID" },
        subject: { type: "string", description: "New subject" },
        start: { type: "string", description: "New start time (ISO 8601)" },
        end: { type: "string", description: "New end time (ISO 8601)" },
        body: { type: "string", description: "New description" },
        location: { type: "string", description: "New location" },
      },
      required: ["eventId"],
    },
  },
  {
    name: "m365_calendar_delete",
    description: "Cancel/delete a calendar event",
    inputSchema: {
      type: "object",
      properties: {
        eventId: { type: "string", description: "Event ID" },
        comment: { type: "string", description: "Cancellation message to attendees" },
      },
      required: ["eventId"],
    },
  },
  {
    name: "m365_calendar_availability",
    description: "Check free/busy availability for users",
    inputSchema: {
      type: "object",
      properties: {
        users: { type: "string", description: "User emails to check, comma-separated" },
        startDateTime: { type: "string", description: "Start (ISO 8601)" },
        endDateTime: { type: "string", description: "End (ISO 8601)" },
      },
      required: ["users", "startDateTime", "endDateTime"],
    },
  },
];
