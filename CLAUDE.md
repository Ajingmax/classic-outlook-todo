# Claude Guide

This repository automates local Classic Outlook scanning on Windows for users whose company Microsoft 365 tenant blocks Microsoft Graph or add-in permissions.

Read `SKILL.md` first, then inspect `scripts/install.ps1` and `scripts/scan-outlook.ps1`.

Rules:

- Do not use Microsoft Graph by default.
- Do not send, delete, move, archive, or mark email.
- Do not open links or attachments.
- Use `config/outlook-todo.config.json` to control recipient scope, mailbox scope, category selection, reminder times, and Sent Items scanning.
- Summarize `work/outlook-action-candidates.md` and `state/outlook-action-candidates.json` into a daily todo file.
- Keep closed decisions in `state/closed-decisions.json`.
