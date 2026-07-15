# Summary Prompt

Use this prompt with any AI agent after `scripts/scan-outlook.ps1` has generated candidate files.

```text
Read:
- work/outlook-action-candidates.md
- state/outlook-action-candidates.json
- state/closed-decisions.json if present

Create a daily todo Markdown file.

Only keep messages where the user is clearly expected to act. Use recipient scope, category labels, and sent-mail reply status as evidence.

Lower priority when `alreadyReplied` is true, unless the newest inbound message still asks for action.

Exclude marketing, advertising, newsletters, system notifications, FYI-only CC mail, and items already listed in closed-decisions.json.

Use stable IDs:
- A1, A2 for action items
- W1, W2 for waiting or monitor items
- F1, F2 for filtered/not actionable items
- C1, C2 for suggested close items

For each action item, write:
- Source email subject
- Sender
- Latest time
- Why it needs action
- What the user should do next
- Evidence

Never send, delete, move, archive, mark email, click links, or open attachments.
```
