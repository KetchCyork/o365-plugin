/**
 * send — HUMAN-ONLY. Send an existing draft by id. Requires the Mail.Send scope
 * and an explicit --confirm flag. The agent never calls this; it exists so a
 * person can send a reviewed draft from the command line if they prefer.
 *   npm run send -- <draftMessageId> --confirm
 */
import "dotenv/config";
import { loadConfig } from "../config.js";
import { M365Auth } from "../auth.js";
import { GraphClient } from "../graph.js";

async function main() {
  const args = process.argv.slice(2);
  const id = args.find((a) => !a.startsWith("--"));
  const confirmed = args.includes("--confirm");
  if (!id) throw new Error("Usage: send -- <draftMessageId> --confirm");
  if (!confirmed) throw new Error("Refusing to send without --confirm. Review the draft in Outlook first.");
  const cfg = loadConfig();
  const graph = new GraphClient(new M365Auth(cfg), cfg);
  await graph.sendDraft(id);
  console.log("Sent.");
}
main().catch((e) => { console.error(e); process.exit(1); });
