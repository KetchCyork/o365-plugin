/**
 * m365 executor
 * -------------
 * Drops into paperclip-mesh-runner on the Windows work node. It reads the task,
 * figures out the intent (briefing / mail / calendar / files / teams), calls
 * Graph, and returns a concise summary. It builds its own Graph client from this
 * node's env, so work credentials never leave the work machine — only the summary
 * flows back through Paperclip.
 *
 * The Executor shape matches paperclip-mesh-runner's interface structurally, so
 * registering it there needs no shared dependency.
 */
import { loadConfig, type M365Config } from "./config.js";
import { M365Auth } from "./auth.js";
import { GraphClient } from "./graph.js";
import { buildBriefing } from "./briefing.js";

// Minimal structural mirror of the runner's Executor contract.
interface Task { title: string; body: string; capability?: string; raw: Record<string, unknown>; }
interface ExecutorResult { output: string; }
interface Executor {
  capability: string;
  description: string;
  run: (task: Task, ctx: unknown) => Promise<ExecutorResult>;
}

type Intent = "briefing" | "mail" | "calendar" | "files" | "teams" | "draft" | "meeting";

export function detectIntent(task: Task): { intent: Intent; query?: string } {
  const explicit = String(task.raw.intent ?? "").toLowerCase();
  const text = `${task.title} ${task.body}`.toLowerCase();
  const q = (task.raw.query as string) || task.body || "";

  const has = (...words: string[]) => words.some((w) => explicit === w || text.includes(w));

  // Write intents first (so "draft an email" doesn't fall through to read-mail).
  if (has("draft", "reply to", "draft a reply", "compose", "write an email", "write a reply"))
    return { intent: "draft", query: q };
  if (has("meeting", "invite", "schedule a", "create a meeting", "set up a meeting",
          "book a meeting", "calendar hold", "send an invite") &&
      !has("what meeting", "my meetings", "meetings today", "what's on"))
    return { intent: "meeting" };

  // Read intents.
  if (has("mail", "email", "inbox")) return { intent: "mail", query: q };
  if (has("calendar", "meetings", "schedule", "agenda")) return { intent: "calendar" };
  if (has("file", "files", "onedrive", "proposal", "document")) return { intent: "files", query: q };
  if (has("teams", "chat", "chats")) return { intent: "teams" };
  return { intent: "briefing" };
}

function toArray(v: unknown): string[] {
  if (Array.isArray(v)) return v.map(String);
  if (typeof v === "string") return v.split(",").map((s) => s.trim()).filter(Boolean);
  return [];
}

export function createM365Executor(getCfg: () => M365Config = loadConfig): Executor {
  let graph: GraphClient | null = null;
  const lazy = (): GraphClient => {
    if (!graph) { const cfg = getCfg(); graph = new GraphClient(new M365Auth(cfg), cfg); }
    return graph;
  };

  return {
    capability: "m365",
    description: "Read-only Microsoft 365 awareness: briefing, mail, calendar, OneDrive, Teams.",
    async run(task) {
      const graph = lazy();
      const { intent, query } = detectIntent(task);
      switch (intent) {
        case "mail": {
          const items = query?.trim()
            ? await graph.searchMail(query, 10)
            : await graph.recentMail(10, true);
          const out = items.length
            ? items.map((m) => `- ${m.isRead ? "" : "(unread) "}${m.from}: ${m.subject}` +
                (m.preview ? `\n    ${m.preview}` : "")).join("\n")
            : "No matching mail.";
          return { output: out };
        }
        case "calendar": {
          const now = new Date();
          const start = new Date(now.getFullYear(), now.getMonth(), now.getDate()).toISOString();
          const end = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1).toISOString();
          const events = await graph.events(start, end);
          const out = events.length
            ? events.map((e) => `- ${e.start} ${e.subject}${e.location ? ` @ ${e.location}` : ""}`).join("\n")
            : "Nothing on the calendar today.";
          return { output: out };
        }
        case "files": {
          const items = query?.trim() ? await graph.searchDrive(query) : await graph.listDrive();
          const out = items.length
            ? items.map((i) => `- ${i.isFolder ? "[dir] " : ""}${i.name}  ${i.webUrl}`).join("\n")
            : "No files found.";
          return { output: out };
        }
        case "teams": {
          const chats = await graph.recentChats(10);
          const out = chats.length
            ? chats.map((c) => `- ${c.topic} (${c.lastActivity})`).join("\n")
            : "No recent Teams chats (or scope not permitted).";
          return { output: out };
        }
        case "draft": {
          const replyText = String(task.raw.replyText ?? task.raw.draftBody ?? task.body ?? "").trim();
          if (!replyText) return { output: "Provide the draft text (raw.replyText or the task body)." };
          const replyTo = (task.raw.replyTo as string | undefined)?.trim();
          if (replyTo) {
            // Resolve to a message id: use as-is if it looks like one, else search.
            let messageId = replyTo;
            if (!/^[A-Za-z0-9_+/=\-]{60,}$/.test(replyTo)) {
              const hits = await graph.searchMail(replyTo, 1);
              if (!hits.length) return { output: `No message found matching "${replyTo}" to reply to.` };
              messageId = hits[0].id;
            }
            const ref = await graph.createReplyDraft(messageId, replyText, Boolean(task.raw.replyAll));
            return { output: `Draft reply saved to Drafts. Review and send from Outlook:\n${ref.webLink}` };
          }
          const ref = await graph.createDraft({
            to: toArray(task.raw.to),
            subject: String(task.raw.subject ?? task.title ?? "(no subject)"),
            body: replyText,
          });
          return { output: `Draft email saved to Drafts. Review and send from Outlook:\n${ref.webLink}` };
        }
        case "meeting": {
          const start = String(task.raw.start ?? "").trim();
          const end = String(task.raw.end ?? "").trim();
          if (!start || !end) {
            return { output: "To create a calendar hold, provide raw.start and raw.end (ISO datetimes)." };
          }
          const ref = await graph.createCalendarHold({
            subject: String(task.raw.subject ?? task.title ?? "Meeting"),
            start, end,
            agenda: task.raw.agenda ? String(task.raw.agenda) : task.body || undefined,
            location: task.raw.location ? String(task.raw.location) : undefined,
            proposedAttendees: toArray(task.raw.proposedAttendees),
          });
          return {
            output:
              `Created a tentative hold (no invites sent). Add attendees and send from Outlook:\n${ref.webLink}`,
          };
        }
        default:
          return { output: await buildBriefing(graph) };
      }
    },
  };
}

export default createM365Executor();
