# O365 Plugin

This repository contains the Microsoft 365 connector for the Agent OS ecosystem.

It provides delegated Graph access to Outlook, calendar, OneDrive, and Teams data from a Windows work node.

## What this repo does

- Provides read-only Microsoft 365 awareness for agents.
- Exposes briefings, mail search, calendar data, OneDrive search, and draft support.
- Keeps credentials and raw work data on the Windows node.
- Supports remote execution in a mesh-enabled environment.

## Capabilities

- Briefings: unread mail, today’s meetings, recent Teams chats.
- Mail search and triage.
- OneDrive file discovery and search.
- Draft replies and new email drafts.
- Calendar hold creation.

## Installation

```bash
cd "o365 plugin/paperclip-plugin-m365"
cp .env.example .env
npm install
```

## Usage

```bash
npm run login
npm run briefing
npm run query -- mail "budget"
npm run draft -- reply "SOW thread" "Thanks, here’s the timeline..."
```

## Documentation

- `paperclip-plugin-m365/docs/SETUP.md` — tenant and auth setup.
- `paperclip-plugin-m365/docs/INTEGRATION.md` — integration into Agent OS / mesh runner.

## Notes

This repo is designed to run on a work machine where Graph credentials are available, and to keep raw data local to that machine.
