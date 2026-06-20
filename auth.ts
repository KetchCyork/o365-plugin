/**
 * Auth (delegated, MSAL)
 * ----------------------
 * Acquires Microsoft Graph access tokens on behalf of the signed-in user.
 *
 * Primary flow: authorization code + PKCE via a localhost redirect (one browser
 * sign-in, then silent refresh using the cached refresh token). This is the
 * tenant-friendly path; device code flow is offered only as a fallback because
 * many organizations now block it.
 *
 * Tokens are cached to disk via MSAL's cache plugin. Protect that file — it
 * holds refresh tokens. See docs/SETUP.md.
 */
import { createServer } from "node:http";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import { dirname } from "node:path";
import {
  PublicClientApplication,
  CryptoProvider,
  type Configuration,
  type ICachePlugin,
  type TokenCacheContext,
  type AuthenticationResult,
} from "@azure/msal-node";
import type { M365Config } from "./config.js";

/** Minimal file-backed token cache so sign-in persists across runs. */
function filePlugin(path: string): ICachePlugin {
  return {
    async beforeCacheAccess(ctx: TokenCacheContext) {
      try { ctx.tokenCache.deserialize(await readFile(path, "utf8")); } catch { /* first run */ }
    },
    async afterCacheAccess(ctx: TokenCacheContext) {
      if (ctx.cacheHasChanged) {
        await mkdir(dirname(path), { recursive: true });
        await writeFile(path, ctx.tokenCache.serialize(), { mode: 0o600 });
      }
    },
  };
}

export class M365Auth {
  private app: PublicClientApplication;

  constructor(private cfg: M365Config) {
    const config: Configuration = {
      auth: { clientId: cfg.clientId, authority: cfg.authority },
      cache: { cachePlugin: filePlugin(cfg.tokenCachePath) },
    };
    this.app = new PublicClientApplication(config);
  }

  /** Get a token silently if possible; otherwise return null (caller triggers login). */
  async getTokenSilent(): Promise<string | null> {
    const accounts = await this.app.getTokenCache().getAllAccounts();
    if (!accounts.length) return null;
    try {
      const res = await this.app.acquireTokenSilent({ account: accounts[0], scopes: this.cfg.scopes });
      return res?.accessToken ?? null;
    } catch {
      return null; // refresh failed/expired -> need interactive login
    }
  }

  /** Interactive sign-in. Auth-code+PKCE via loopback, or device code if configured. */
  async login(): Promise<void> {
    if (this.cfg.useDeviceCode) return this.loginDeviceCode();
    return this.loginAuthCode();
  }

  /** One token call: silent, else throw a clear instruction to run `login`. */
  async requireToken(): Promise<string> {
    const t = await this.getTokenSilent();
    if (t) return t;
    throw new Error('Not signed in (or token expired). Run: npm run login');
  }

  private async loginAuthCode(): Promise<void> {
    const crypto = new CryptoProvider();
    const { verifier, challenge } = await crypto.generatePkceCodes();
    const redirect = new URL(this.cfg.redirectUri);

    const authUrl = await this.app.getAuthCodeUrl({
      scopes: this.cfg.scopes,
      redirectUri: this.cfg.redirectUri,
      codeChallenge: challenge,
      codeChallengeMethod: "S256",
    });

    const code: string = await new Promise((resolve, reject) => {
      const server = createServer((req, res) => {
        try {
          const u = new URL(req.url ?? "/", `http://${req.headers.host}`);
          if (u.pathname !== redirect.pathname) { res.writeHead(404); res.end(); return; }
          const c = u.searchParams.get("code");
          const err = u.searchParams.get("error_description") || u.searchParams.get("error");
          res.writeHead(200, { "Content-Type": "text/html" });
          res.end("<html><body><h3>Signed in. You can close this tab.</h3></body></html>");
          server.close();
          if (c) resolve(c); else reject(new Error(err || "No authorization code returned."));
        } catch (e) { reject(e as Error); }
      });
      server.listen(Number(redirect.port || 80), redirect.hostname, () => {
        console.log("\nOpen this URL to sign in:\n" + authUrl + "\n");
      });
    });

    const result: AuthenticationResult = await this.app.acquireTokenByCode({
      code,
      scopes: this.cfg.scopes,
      redirectUri: this.cfg.redirectUri,
      codeVerifier: verifier,
    });
    console.log(`Signed in as ${result.account?.username ?? "(unknown)"}. Token cached.`);
  }

  private async loginDeviceCode(): Promise<void> {
    const result = await this.app.acquireTokenByDeviceCode({
      scopes: this.cfg.scopes,
      deviceCodeCallback: (info) => console.log("\n" + info.message + "\n"),
    });
    console.log(`Signed in as ${result?.account?.username ?? "(unknown)"}. Token cached.`);
  }
}
