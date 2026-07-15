---
name: classic-outlook-todo
description: Build and operate a local Windows workflow that reads Classic Outlook through COM/MAPI when Microsoft Graph or admin-approved APIs are blocked, scans inboxes and sent mail, classifies actionable emails, and writes candidate data for an AI agent to turn into a daily Markdown todo list. Use when users need Outlook-to-todo automation under enterprise admin restrictions, configurable recipient scope, folder scope, reminder times, and category filtering.
---

# Classic Outlook Todo

Use this skill when an organization blocks Microsoft Graph, Outlook add-ins, Zapier, Make, or other cloud integrations, but Classic Outlook is installed and signed in on a Windows PC.

## Core Workflow

1. Run `scripts/install.ps1` to verify Classic Outlook, create a config file, and register a Windows scheduled task.
2. Run `scripts/scan-outlook.ps1` to read Classic Outlook locally through COM/MAPI.
3. Ask the user to choose reminder scope, mailbox folders, and categories in `config/outlook-todo.config.json`.
4. Read `work/outlook-action-candidates.md` and `state/outlook-action-candidates.json`.
5. Generate a human-readable Markdown todo list in the user's preferred language.
6. Respect `state/closed-decisions.json` so dismissed or completed items do not reappear.

## Required Choices

Confirm or configure:

- Recipient scope: direct recipient, CC, mention in body or subject, unread only, important only.
- Folder scope: default Inbox, all Inbox folders across mailboxes, named folders such as Important or Unread Mail.
- Sent mail scan: enabled by default to detect threads the user already answered.
- Categories: customer issues, meetings, logistics, technical work, marketing, ads, system notifications, and user-specific detected categories.
- Reminder times: recommended `07:00` and `12:00`; store scan times a few minutes earlier.

## Safety Rules

- Do not send, delete, move, archive, or mark emails.
- Do not open links or attachments.
- Use Classic Outlook local COM/MAPI, not Microsoft Graph, unless the user explicitly changes the architecture.
- If Classic Outlook is missing, use `scripts/install.ps1 -InstallIfMissing` for best-effort installation or tell the user what failed.

## Agent Output Rules

When summarizing candidates:

- Keep only items where the user is clearly expected to act.
- Treat already replied threads as lower priority unless the latest inbound message still asks for action.
- Exclude marketing, ads, newsletters, automatic system notifications, and FYI-only CC mail unless the user opted into them.
- Use stable IDs such as `A1`, `A2`, `W1`, `F1`, and `C1` so the user can batch close items.
- Store closed items in `state/closed-decisions.json`; do not keep closed items in the visible todo list.
