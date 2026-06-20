/**
 * hold — create a TENTATIVE calendar hold on your calendar (no invites sent).
 *   npm run hold -- "<subject>" <startISO> <endISO> ["<agenda>"] ["<attendee1,attendee2>"]
 * Add attendees and send the invite yourself from Outlook after reviewing.
 */
import "dotenv/config";
import { loadConfig } from "../config.js";
import { M365Auth } from "../auth.js";
import { GraphClient } from "../graph.js";

async function main() {
  const cfg = loadConfig();
  const graph = new GraphClient(new M365Auth(cfg), cfg);
  const [subject, start, end, agenda, attendees] = process.argv.slice(2);
  if (!subject || !start || !end) {
    throw new Error('Usage: hold -- "<subject>" <startISO> <endISO> ["<agenda>"] ["<a,b>"]');
  }
  const ref = await graph.createCalendarHold({
    subject, start, end, agenda,
    proposedAttendees: attendees ? attendees.split(",").map((s) => s.trim()) : [],
  });
  console.log(`Tentative hold created (no invites sent). Review/send in Outlook:\n${ref.webLink}`);
}
main().catch((e) => { console.error(e); process.exit(1); });
