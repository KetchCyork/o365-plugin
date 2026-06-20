# paperclip-plugin-m365

Read-only **Microsoft 365 awareness** for your agents — Outlook mail, calendar,
OneDrive, and Teams chats — via Microsoft Graph with **delegated** auth. It runs as
a capability executor on your **Windows work node** in the mesh, so work
credentials and raw data stay on the work machine and only summaries flow back.

## What it does

- A **briefing**: unread mail, today's meetings, recent Teams chats.
- **Mail** search/triage (metadata by default).
- **Calendar** for today.
- **OneDrive** list/search (great for finding proposals).
- **Draft** replies and new emails (saved to Drafts — never sent).
- **Calendar holds** (tentative, no attendees notified — you add + send in Outlook).
- Stays within your own access; read-first; **nothing leaves your outbox or hits a
  calendar without your explicit action**.

## Commands

```
npm run briefing
npm run query -- mail "budget"
npm run draft -- reply "SOW thread" "Thanks, here's the timeline..."
npm run draft -- new "client@x.com" "Follow-up" "Good speaking today..."
npm run hold  -- "Kickoff" 2026-07-01T15:00:00 2026-07-01T16:00:00 "Agenda"
npm run send  -- <draftId> --confirm     # human-only; needs Mail.Send scope
```

## Quickstart

```bash
cp .env.example .env      # client id + tenant id from your Entra app
npm install
npm run login             # one-time browser sign-in
npm run briefing
npm run query -- mail "budget"
```

Full setup (Entra app registration + the important tenant caveats) is in
`docs/SETUP.md`. Wiring it into the mesh runner is in `docs/INTEGRATION.md`.

## What's verified

Compiles cleanly against the real `@azure/msal-node` v2 types (auth-code+PKCE,
device-code fallback, file token cache) and the Graph client. Intent routing is
unit-tested. Live Graph calls require your own Entra app registration and sign-in,
which only you can do against TSP's tenant.

## Make it yours (clean GitHub authoring history)

Ships without git history so your first commit is yours:

```bash
git init && git add -A && git commit -m "Initial commit: m365 plugin"
git remote add origin git@github.com:<you>/paperclip-plugin-m365.git
git push -u origin main
```

## License

MIT (recommended).
