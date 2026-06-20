/** login — one-time interactive sign-in; caches tokens for silent refresh. */
import "dotenv/config";
import { loadConfig } from "../config.js";
import { M365Auth } from "../auth.js";

async function main() {
  const cfg = loadConfig();
  if (!cfg.clientId) throw new Error("Set M365_CLIENT_ID in .env (see docs/SETUP.md).");
  const auth = new M365Auth(cfg);
  const existing = await auth.getTokenSilent();
  if (existing) { console.log("Already signed in (cached token still valid)."); return; }
  await auth.login();
}
main().catch((e) => { console.error(e); process.exit(1); });
