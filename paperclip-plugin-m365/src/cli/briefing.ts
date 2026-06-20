/** briefing — print a what-needs-attention summary. */
import "dotenv/config";
import { loadConfig } from "../config.js";
import { M365Auth } from "../auth.js";
import { GraphClient } from "../graph.js";
import { buildBriefing } from "../briefing.js";

async function main() {
  const cfg = loadConfig();
  const graph = new GraphClient(new M365Auth(cfg), cfg);
  console.log(await buildBriefing(graph));
}
main().catch((e) => { console.error(e); process.exit(1); });
