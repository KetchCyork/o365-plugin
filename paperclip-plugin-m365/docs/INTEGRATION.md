# Wiring m365 into the mesh runner

The m365 executor runs on the **Windows work node**, so credentials and raw work
data never leave that machine — only summaries flow back through Paperclip.

## Register the executor in paperclip-mesh-runner

The executor's shape matches the runner's `Executor` interface structurally, so no
shared dependency is required. Two ways to wire it:

**Option A — npm dependency (clean):**
```bash
# in paperclip-mesh-runner, on the Windows node
npm install /path/to/paperclip-plugin-m365   # or your published package
```
Then in `src/executors/registry.ts` add it to the BUILTIN list:
```ts
import m365 from "paperclip-plugin-m365/dist/executor.js"; // or src in dev
const BUILTIN = [shell, memory, localModel, m365];
```

**Option B — copy the executor file** into the runner's `src/executors/` and add
it to BUILTIN. Simplest for a quick start.

Then advertise the capability on this node:
```
RUNNER_CAPABILITIES=m365,memory
```

## How work routes here

In Paperclip, give the work node's agent the `m365` capability (and a role like
"Work Context"). When any agent needs work awareness, it delegates a task to that
agent; Paperclip wakes the Windows runner, the m365 executor calls Graph locally,
and a summary comes back as the run result / a comment.

Example tasks the executor understands (intent is auto-detected, or pass
`raw.intent`):
- "morning briefing" -> unread mail + today's meetings + recent chats
- "any urgent emails about the Hilmar SOW" -> mail search
- "what meetings do I have today" -> calendar
- "find the latest proposal in OneDrive" -> file search

## Data boundary (by design)

- Auth + token cache live only on the Windows node.
- Graph calls run on the Windows node.
- Default output is metadata (subjects/senders/times), not bodies. Turn on short
  snippets with `M365_ALLOW_SNIPPETS=true` only if you accept the exposure.
- Nothing copies your mailbox or files onto the personal Mac. If you later index
  OneDrive proposals into the memory brain, run that indexer **on the work node**
  against the locally-synced OneDrive folder, and expose only scoped query results.

## Write intents (drafting)

The executor also understands write intents, all of which produce drafts/holds
only — never sends:

- "draft a reply to <message>" (pass `raw.replyTo` = id or search terms, and the
  text as `raw.replyText` or the task body) -> reply draft in Drafts
- "draft an email" (pass `raw.to`, `raw.subject`, body) -> draft in Drafts
- "schedule a meeting" / "create a hold" (pass `raw.start`, `raw.end`, optional
  `raw.subject`, `raw.agenda`, `raw.location`, `raw.proposedAttendees`) -> tentative
  calendar hold, no invites sent

The model writes the content; the executor performs the deterministic action and
returns the Outlook web link for your review. Sending is never an executor action —
it's a human step (send from Outlook, or the explicit `send --confirm` command).
