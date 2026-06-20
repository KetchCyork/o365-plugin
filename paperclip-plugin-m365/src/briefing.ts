/**
 * Briefing
 * --------
 * Assembles a concise "what needs my attention" summary from Graph: unread mail,
 * the day's meetings, and recent Teams activity. Returns a short text block —
 * the kind of thing an agent reports back through Paperclip, or that you read
 * each morning.
 */
import { GraphClient } from "./graph.js";

function fmtTime(iso: string): string {
  if (!iso) return "";
  const d = new Date(iso);
  return isNaN(d.getTime()) ? iso : d.toLocaleString();
}

export async function buildBriefing(graph: GraphClient): Promise<string> {
  const lines: string[] = [];

  const [me, unread, chats] = await Promise.all([
    graph.me().catch(() => ({ displayName: "you", mail: "" })),
    graph.recentMail(8, true).catch(() => []),
    graph.recentChats(8).catch(() => []),
  ]);

  // Today's window in ISO (local midnight -> +24h).
  const now = new Date();
  const start = new Date(now.getFullYear(), now.getMonth(), now.getDate()).toISOString();
  const end = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1).toISOString();
  const events = await graph.events(start, end).catch(() => []);

  lines.push(`Briefing for ${me.displayName} — ${now.toLocaleString()}`);
  lines.push("");

  lines.push(`Unread email (${unread.length}):`);
  if (!unread.length) lines.push("  (inbox clear)");
  for (const m of unread.slice(0, 8)) lines.push(`  - ${m.from}: ${m.subject}`);
  lines.push("");

  lines.push(`Today's meetings (${events.length}):`);
  if (!events.length) lines.push("  (nothing scheduled)");
  for (const e of events) lines.push(`  - ${fmtTime(e.start)} ${e.subject}${e.location ? ` @ ${e.location}` : ""}`);
  lines.push("");

  lines.push(`Recent Teams chats (${chats.length}):`);
  if (!chats.length) lines.push("  (none / not permitted)");
  for (const c of chats.slice(0, 6)) lines.push(`  - ${c.topic} (${fmtTime(c.lastActivity)})`);

  return lines.join("\n");
}
