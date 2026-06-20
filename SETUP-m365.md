# Setup: Microsoft 365 (Entra app registration)

This plugin reads **your own** Outlook, calendar, OneDrive, and Teams chats using
Microsoft Graph with **delegated** permissions — it can only see what you can see,
never the whole tenant. To do that it needs an Entra (Azure AD) app registration.

## Read this first (it's your firm's M365)

This touches TSP's tenant and your work mailbox, which contains client data. Before
connecting it:

- **Involve TSP IT.** Registering the app, granting consent, and Conditional Access
  are tenant-controlled. Depending on your tenant's user-consent settings, an admin
  may need to consent to the app. Do this with IT, not around them.
- **Delegated + read-only by default.** Scopes below are read-only. The plugin
  returns summaries (subjects, senders, times), not full message bodies, unless you
  set `M365_ALLOW_SNIPPETS=true`.
- **The token cache is a secret.** It holds a refresh token to your mailbox.
  It's written with `0600` perms and gitignored; keep it on the work machine only.

## 1. Register the app

In the Microsoft Entra admin center → **App registrations → New registration**:

- Name: e.g. `paperclip-m365` (anything).
- Supported account types: **single tenant** (this org only) is the safe default
  for work; use multi-tenant/`common` only if you need personal accounts too.
- Click Register, then copy the **Application (client) ID** and
  **Directory (tenant) ID**.

## 2. Add the redirect (public client)

Under **Authentication → Add a platform → Mobile and desktop applications**, add:

```
http://localhost:8385/redirect
```

This is a public client using PKCE — **no client secret is created or needed**.
Leave "Allow public client flows" set to **No** unless you intend to use device
code flow (see caveat below).

## 3. Grant Graph permissions (delegated, read-only)

Under **API permissions → Add a permission → Microsoft Graph → Delegated**, add:

```
User.Read   Mail.ReadWrite   Calendars.ReadWrite   Files.Read.All   Chat.Read   offline_access
```

`Mail.ReadWrite` lets the plugin create draft emails and reply drafts;
`Calendars.ReadWrite` lets it create tentative calendar holds. Neither sends
anything. Add `Mail.Send` **only** if you want the explicit, human-confirmed
`send` command — it's not required for drafting and is intentionally omitted by
default.

If your tenant requires it, click **Grant admin consent** (or ask an admin to).

## 4. Configure and sign in

```bash
cp .env.example .env     # paste client id + tenant id
npm install
npm run login            # opens a sign-in URL; approve in the browser, once
npm run briefing         # what needs your attention today
npm run query -- files "proposal"   # search OneDrive
```

## Honest caveats

- **Device code flow is often blocked.** Microsoft now recommends restricting it,
  and many tenants block it via Conditional Access. This plugin defaults to the
  auth-code + PKCE browser flow for that reason. Only set `M365_USE_DEVICE_CODE=true`
  if IT confirms it's allowed.
- **Conditional Access** can still block or step up the sign-in (MFA, managed
  device). That's expected and correct; complete it in the browser.
- **Teams is limited.** Reading your own chats works with `Chat.Read`. Channel
  messages need broader, admin-consented scopes; this plugin intentionally stays
  on your own chats.
- **Consent ≠ scope restriction.** Delegated permissions never exceed your own
  access, but they aren't limited to a subset of your mail — the app can read all
  of your mailbox once you consent. Keep scopes minimal and read-only.

## Drafting & calendar invites (the write path)

The plugin can prepare work for you, but it never sends on its own:

- **Draft a reply / new email** → saved to your **Drafts** folder. You review and
  send from Outlook.
  ```
  npm run draft -- reply "Hilmar SOW" "Thanks — proposed timeline attached..."
  npm run draft -- new "client@example.com" "Follow-up" "Good speaking today..."
  ```
- **Calendar invite** → created as a **tentative hold on your own calendar with no
  attendees**, so no invitations go out. Proposed attendees + agenda are written
  into the event for you to review; you add attendees and send from Outlook.
  ```
  npm run hold -- "TSP / Acme kickoff" 2026-07-01T15:00:00 2026-07-01T16:00:00 "Agenda..." "a@x.com,b@y.com"
  ```
- **Sending** is deliberately a separate, human-only command that requires the
  `Mail.Send` scope and an explicit `--confirm`. The agent never sends; the
  simplest approval is just to send the reviewed draft from Outlook.
  ```
  npm run send -- <draftMessageId> --confirm
  ```

Why this shape: creating a draft or an attendee-less hold is reversible and
notifies no one, so it's safe for an agent to do. Anything that actually leaves
your outbox or hits someone's calendar stays a human decision — which is also the
defense against a malicious email trying to get the agent to send on its behalf.
