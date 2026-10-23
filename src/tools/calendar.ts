/**
 * Calendar tools — list, today, week, create, update, cancel, decline, availability
 */

import type { CalendarEvent } from "../types.js";
import type { GraphClient } from "../graph.js";
import { normalizeTop } from "../validation.js";

/** Build ISO 8601 start/end strings for a day in the given timezone. */
export function tzDayBounds(timezone: string, offsetDays = 0): { startDateTime: string; endDateTime: string } {
  const parts = new Intl.DateTimeFormat("en-US", {
    timeZone: timezone, year: "numeric", month: "2-digit", day: "2-digit",
  }).formatToParts(new Date());
  const part = (type: string) => Number(parts.find((value) => value.type === type)!.value);
  const date = Date.UTC(part("year"), part("month") - 1, part("day") + offsetDays);
  const offsetFormatter = new Intl.DateTimeFormat("en-US", { timeZone: timezone, timeZoneName: "longOffset" });
  const midnight = (localDate: number) => {
    let instant = localDate;
    let previous = instant;
    for (let attempt = 0; attempt < 4; attempt++) {
      const zone = offsetFormatter.formatToParts(new Date(instant)).find((value) => value.type === "timeZoneName")!.value;
      const match = zone.match(/GMT([+-])(\d{2}):(\d{2})/);
      const offset = match ? (Number(match[2]) * 60 + Number(match[3])) * 60_000 * (match[1] === "+" ? 1 : -1) : 0;
      const next = localDate - offset;
      if (next === instant) return instant;
      previous = instant;
      instant = next;
    }
    // If midnight is skipped by a zone transition, choose the first later instant.
    return Math.max(previous, instant);
  };
  return {
    startDateTime: new Date(midnight(date)).toISOString(),
    endDateTime: new Date(midnight(date + 86400000) - 1).toISOString(),
  };
}

// Today/week explicitly request UTC, then render in the configured timezone.
function utcEventDate(value: string): Date {
  return new Date(/(?:Z|[+-]\d{2}:\d{2})$/i.test(value) ? value : `${value}Z`);
}

export function registerCalendarTools(client: GraphClient) {
  const user = "/me";

  async function cancelOrganizerEvent(args: {
    eventId: string;
    comment?: string;
  }): Promise<string> {
    const eventId = encodeURIComponent(args.eventId);
    await client.post(`${user}/events/${eventId}/cancel`, {
      comment: args.comment || "",
    });
    return JSON.stringify({ success: true, eventId: args.eventId });
  }

  return {
    // ── list ─────────────────────────────────────────────────────
    async calendar_list(args: {
      startDateTime?: string;
      endDateTime?: string;
      top?: number;
    }): Promise<string> {
      const top = normalizeTop(args.top, 50);
      if (!!args.startDateTime !== !!args.endDateTime) throw new Error("Provide both startDateTime and endDateTime");
      let path: string;

      if (args.startDateTime && args.endDateTime) {
        path = `${user}/calendarView?startDateTime=${encodeURIComponent(args.startDateTime)}&endDateTime=${encodeURIComponent(args.endDateTime)}&$top=${top}&$orderby=start/dateTime&$select=id,subject,start,end,location,isOnlineMeeting,onlineMeeting,onlineMeetingUrl,organizer,attendees,isAllDay,isOrganizer,responseStatus,webLink`;
      } else {
        path = `${user}/events?$top=${top}&$orderby=start/dateTime&$select=id,subject,start,end,location,isOnlineMeeting,onlineMeeting,onlineMeetingUrl,organizer,attendees,isAllDay,isOrganizer,responseStatus,webLink`;
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
          meetingUrl: e.onlineMeeting?.joinUrl ?? e.onlineMeetingUrl,
          organizer: e.organizer?.emailAddress?.name ?? e.organizer?.emailAddress?.address,
          attendees: e.attendees?.length ?? 0,
          isOrganizer: e.isOrganizer,
          responseStatus: e.responseStatus?.response,
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
      const events = await client.getAll<CalendarEvent>(path, 10, 50, { Prefer: 'outlook.timezone="UTC"' });

      if (events.length === 0) return "No events today.";

      const now = new Date();
      const lines = events.map((e) => {
        const startTime = utcEventDate(e.start.dateTime);
        const endTime = utcEventDate(e.end.dateTime);
        const time = e.isAllDay
          ? "All Day"
          : `${startTime.toLocaleTimeString("en-US", { timeZone: client.timezone, hour: "numeric", minute: "2-digit" })}–${endTime.toLocaleTimeString("en-US", { timeZone: client.timezone, hour: "numeric", minute: "2-digit" })}`;
        let line = `${time}  ${e.subject}`;
        if (e.location?.displayName) line += `  📍 ${e.location.displayName}`;
        if (e.isOnlineMeeting) line += `  🎥`;
        return line;
      });

      return `Today (${now.toLocaleDateString("en-US", { timeZone: client.timezone, weekday: "long", month: "long", day: "numeric" })})\n${lines.join("\n")}\n\n${events.length} event${events.length !== 1 ? "s" : ""}`;
    },

    // ── week ─────────────────────────────────────────────────────
    async calendar_week(): Promise<string> {
      const startBounds = tzDayBounds(client.timezone, 0);
      const endBounds = tzDayBounds(client.timezone, 6);

      const path = `${user}/calendarView?startDateTime=${encodeURIComponent(startBounds.startDateTime)}&endDateTime=${encodeURIComponent(endBounds.endDateTime)}&$top=100&$orderby=start/dateTime`;
      const events = await client.getAll<CalendarEvent>(path, 10, 100, { Prefer: 'outlook.timezone="UTC"' });

      if (events.length === 0) return "No events this week.";

      const byDay = new Map<string, string[]>();
      for (const e of events) {
        const d = utcEventDate(e.start.dateTime);
        const key = d.toLocaleDateString("en-US", { timeZone: client.timezone, weekday: "short", month: "short", day: "numeric" });
        const startTime = d.toLocaleTimeString("en-US", { timeZone: client.timezone, hour: "numeric", minute: "2-digit" });
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
      if (args.isOnline) { event.isOnlineMeeting = true; event.onlineMeetingProvider = "teamsForBusiness"; }

      const data = await client.post<any>(`${user}/events`, event);
      return JSON.stringify(
        {
          id: data.id,
          subject: data.subject,
          webLink: data.webLink,
          onlineMeetingUrl: data.onlineMeeting?.joinUrl ?? data.onlineMeetingUrl,
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
      if (args.subject !== undefined) patch.subject = args.subject;
      if (args.start) patch.start = client.dateTimeStr(args.start);
      if (args.end) patch.end = client.dateTimeStr(args.end);
      if (args.body !== undefined) patch.body = { contentType: "HTML", content: args.body };
      if (args.location !== undefined) patch.location = { displayName: args.location };
      if (Object.keys(patch).length === 0) throw new Error("Provide at least one event field to update");

      const eventId = encodeURIComponent(args.eventId);
      await client.patch(`${user}/events/${eventId}`, patch);
      return JSON.stringify({ success: true, eventId: args.eventId });
    },

    // ── delete (legacy organizer-cancel alias) ──────────────────
    async calendar_delete(args: {
      eventId: string;
      comment?: string;
    }): Promise<string> {
      return cancelOrganizerEvent(args);
    },

    // ── cancel (organizer only) ──────────────────────────────────
    async calendar_cancel(args: {
      eventId: string;
      comment?: string;
    }): Promise<string> {
      return cancelOrganizerEvent(args);
    },

    // ── decline (attendee only) ──────────────────────────────────
    async calendar_decline(args: {
      eventId: string;
      comment?: string;
      sendResponse?: boolean;
    }): Promise<string> {
      const eventId = encodeURIComponent(args.eventId);
      await client.post(`${user}/events/${eventId}/decline`, {
        comment: args.comment || "",
        sendResponse: args.sendResponse ?? true,
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
      }, { retrySafe: true });
      return JSON.stringify(data.value, null, 2);
    },
  };
}

export const calendarToolSchemas = [
  {
    name: "m365_calendar_list",
    description: "List calendar events; supply both dates to expand recurring events within a range",
    inputSchema: {
      type: "object",
      properties: {
        startDateTime: { type: "string", description: "Start datetime (ISO 8601)" },
        endDateTime: { type: "string", description: "End datetime (ISO 8601)" },
        top: { type: "integer", minimum: 1, maximum: 500, description: "Max events", default: 50 },
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
    description: "Legacy alias: cancel a calendar event you organize and notify attendees",
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
    name: "m365_calendar_cancel",
    description: "Cancel a calendar event you organize and notify attendees",
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
    name: "m365_calendar_decline",
    description: "Decline a calendar invitation as an attendee",
    inputSchema: {
      type: "object",
      properties: {
        eventId: { type: "string", description: "Event ID" },
        comment: { type: "string", description: "Optional response to the organizer" },
        sendResponse: { type: "boolean", description: "Send the response to the organizer (default: true)" },
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
