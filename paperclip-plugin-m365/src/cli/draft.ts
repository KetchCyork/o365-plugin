/**
 * draft — create a draft email or reply (nothing is sent).
 *   npm run draft -- reply "<search terms or message id>" "<your reply text>"
 *   npm run draft -- new   "<to1,to2>" "<subject>" "<body>"
 */
import "dotenv/config";
import { loadConfig } from "../config.js";
import { M365Auth } from "../auth.js";
import { GraphClient } from "../graph.js";

async function main() {
  const cfg = loadConfig();
  const graph = new GraphClient(new M365Auth(cfg), cfg);
  const [mode, ...rest] = process.argv.slice(2);

  if (mode === "reply") {
    const [target, ...textParts] = rest;
    const text = textParts.join(" ");
    if (!target || !text) throw new Error('Usage: draft -- reply "<search/id>" "<text>"');
    let id = target;
    if (!/^[A-Za-z0-9_+/=\-]{60,}$/.test(target)) {
      const hits = await graph.searchMail(target, 1);
      if (!hits.length) throw new Error(`No message found matching "${target}".`);
      id = hits[0].id;
    }
    const ref = await graph.createReplyDraft(id, text);
    console.log(`Draft reply saved. Review/send in Outlook:\n${ref.webLink}`);
  } else if (mode === "new") {
    const [to, subject, ...bodyParts] = rest;
    const body = bodyParts.join(" ");
    if (!subject || !body) throw new Error('Usage: draft -- new "<to>" "<subject>" "<body>"');
    const ref = await graph.createDraft({ to: to.split(",").map((s) => s.trim()).filter(Boolean), subject, body });
    console.log(`Draft email saved. Review/send in Outlook:\n${ref.webLink}`);
  } else {
    throw new Error('Usage: draft -- reply|new ...');
  }
}
main().catch((e) => { console.error(e); process.exit(1); });
