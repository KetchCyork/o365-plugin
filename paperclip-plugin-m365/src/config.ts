/**
 * M365 plugin config
 * ------------------
 * Reads your Entra (Azure AD) app registration details + scopes from .env.
 * Delegated, read-only by default. See docs/SETUP.md to create the app.
 */
import { homedir } from "node:os";
import { join } from "node:path";

function env(name: string, fallback = ""): string {
  return process.env[name] ?? fallback;
}

export interface M365Config {
  clientId: string;
  /** Authority: use /organizations or your tenant id for work; /common for both. */
  authority: string;
  /** Loopback redirect for the auth-code flow (must match the app registration). */
  redirectUri: string;
  /** Delegated, read-only scopes. offline_access enables silent refresh. */
  scopes: string[];
  /** Where the encrypted-at-rest token cache lives (protect this file). */
  tokenCachePath: string;
  /** Prefer device code flow instead of auth-code (only if your tenant allows it). */
  useDeviceCode: boolean;
  /** If true, executors may include short body snippets; default false = metadata only. */
  allowSnippets: boolean;
  /** IANA/Windows timezone for created events, e.g. "America/Chicago". */
  timeZone: string;
}

export function loadConfig(): M365Config {
  const base = join(homedir(), ".paperclip-m365");
  const tenant = env("M365_TENANT_ID", "organizations");
  return {
    clientId: env("M365_CLIENT_ID"),
    authority: env("M365_AUTHORITY", `https://login.microsoftonline.com/${tenant}`),
    redirectUri: env("M365_REDIRECT_URI", "http://localhost:8385/redirect"),
    scopes: env(
      "M365_SCOPES",
      // ReadWrite enables drafting; Calendars.ReadWrite enables calendar holds.
      // Add "Mail.Send" only if you want the explicit, human-confirmed send path.
      "User.Read Mail.ReadWrite Calendars.ReadWrite Files.Read.All Chat.Read offline_access"
    ).split(/\s+/).filter(Boolean),
    tokenCachePath: env("M365_TOKEN_CACHE", join(base, "token-cache.json")),
    useDeviceCode: env("M365_USE_DEVICE_CODE", "false") === "true",
    allowSnippets: env("M365_ALLOW_SNIPPETS", "false") === "true",
    timeZone: env("M365_TIMEZONE", "UTC"),
  };
}
