# 读取 Outlook 文件建立代办 / Classic Outlook Todo

一个面向企业 Windows 电脑的本地 Outlook 邮件待办生成工具。  
It turns locally available Classic Outlook mail into AI-readable todo candidates when Microsoft Graph, Outlook add-ins, or cloud automation tools are blocked by company admin policies.

## 中文介绍

很多公司会限制 Microsoft Graph、Outlook Add-ins、Zapier、Make、Power Automate 云连接器等权限。这个项目提供另一条路径：

- 使用 **Classic Outlook** 本地 COM/MAPI 接口读取邮件。
- 不需要 Microsoft Graph 管理员授权。
- 不发送、不删除、不移动、不归档邮件。
- 扫描收件箱和发件箱，判断哪些邮件可能需要你处理。
- 输出 Markdown 和 JSON，让 Codex、Claude 或其他 agent 生成每日待办。

适合场景：

- 公司 Outlook 账号有管理员限制，无法授权第三方应用。
- 电脑上已经能登录 Classic Outlook。
- 你希望每天早上/中午自动生成待办。
- 你希望选择提醒范围，例如只看收件人是自己、CC、正文 @ 自己、未读、重要邮件。
- 你希望按自己的邮箱结构选择收件箱，例如所有收件箱、重要、未读、自定义文件夹。
- 你希望自动识别邮件类别，例如客户问题、会议、物流、技术、营销、广告、系统通知。
- 你希望同时检查发件箱，降低“已经回复过的邮件”重复提醒。

## English Overview

This project helps people create a daily email todo workflow on locked-down enterprise Windows machines.

When admin approval blocks Microsoft Graph, cloud connectors, or Outlook add-ins, this workflow reads **Classic Outlook** locally through Windows COM/MAPI and writes AI-readable candidate files.

It is designed for:

- Company Outlook accounts with admin approval restrictions.
- Users who can sign in to Classic Outlook on their PC.
- Daily morning/noon todo reminders.
- Configurable recipient scope: direct recipient, CC, mention, unread, important.
- Configurable mailbox scope: default Inbox, all Inbox folders, Important, Unread Mail, or custom folders.
- Auto category detection: customer issues, meetings, logistics, technical work, marketing, ads, system notifications.
- Sent Items scanning to detect threads already answered by the user.

## What It Does

The scanner writes:

- `work/outlook-action-candidates.md`
- `state/outlook-action-candidates.json`
- `work/detected-categories.md`
- `logs/scan-outlook-last-run.log`

An AI agent can then read those files and generate a daily todo list.

## What It Does Not Do

This tool does **not**:

- send email
- delete email
- move email
- archive email
- mark email as read
- click links
- open attachments
- bypass company security controls

It only reads locally available Outlook data from the signed-in Windows user profile.

## Requirements / 前置条件

- Windows 10 or Windows 11
- PowerShell 5.1+
- Classic Outlook installed and signed in
- A local AI/coding agent such as Codex, Claude, Cursor, or another tool that can read Markdown/JSON files

If Classic Outlook is missing, the installer can try a best-effort installation:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\install.ps1 -InstallIfMissing
```

Automatic installation depends on `winget`, Microsoft 365 licensing, and company IT policy. If it fails, install Microsoft 365 Apps / Classic Outlook manually, open Outlook once, sign in, then rerun the installer.

## Quick Start / 快速开始

Clone or download this project, then run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\install.ps1
```

Default runtime folder:

```text
%USERPROFILE%\Documents\OutlookTodo
```

The installer will:

1. Check Classic Outlook.
2. Create runtime folders.
3. Create `config/outlook-todo.config.json`.
4. Copy the scanner script.
5. Register a Windows scheduled task.
6. Run one test scan.

Recommended reminder design:

- scan mail at `06:55`
- generate AI summary at `07:00`
- scan mail at `11:55`
- generate AI summary at `12:00`

## Configuration / 配置选项

Edit:

```text
%USERPROFILE%\Documents\OutlookTodo\config\outlook-todo.config.json
```

### 1. Recipient Scope / 提醒范围

```json
"recipientScope": {
  "directTo": true,
  "cc": false,
  "mentionInSubjectOrBody": true,
  "unreadOnly": false,
  "importantOnly": false
}
```

Common choices:

- Only mail sent directly to me: `directTo=true`, `cc=false`
- Include CC mail: `cc=true`
- Include messages mentioning my name: `mentionInSubjectOrBody=true`
- Only unread mail: `unreadOnly=true`
- Only high importance mail: `importantOnly=true`

### 2. Mailbox Scope / 收件箱范围

```json
"folderScope": {
  "defaultInbox": true,
  "allInboxes": false,
  "folderNames": ["Inbox", "Important", "Unread Mail"]
}
```

Options:

- `defaultInbox`: default Outlook Inbox.
- `allInboxes`: all Inbox folders across mailboxes.
- `folderNames`: custom folder names such as `Important`, `Unread Mail`, project folders, or shared mailbox folders.

### 3. Category Detection / 邮件类型识别

```json
"categories": {
  "include": ["customer_issue", "meeting", "logistics", "technical", "deadline"],
  "exclude": ["marketing", "advertising", "newsletter", "system_notification"],
  "autoDetect": true
}
```

Built-in categories:

- `customer_issue`
- `meeting`
- `logistics`
- `technical`
- `deadline`
- `marketing`
- `advertising`
- `newsletter`
- `system_notification`

After each scan, review:

```text
work/detected-categories.md
```

Use it to decide which categories should be included or excluded for your job.

### 4. Reminder Times / 每日提醒时间

```json
"reminder": {
  "suggestedSummaryTimes": ["07:00", "12:00"],
  "scanTimes": ["06:55", "11:55"]
}
```

The scanner should run a few minutes before the AI summary.

### 5. Sent Items / 发件箱检测

```json
"sentMail": {
  "scanSentItems": true,
  "lowerPriorityIfAlreadyReplied": true
}
```

This checks Sent Items for matching conversations. If the user has already replied, the candidate will be marked:

```text
Already replied from Sent Items: True
```

The AI agent should usually lower its priority unless the latest inbound email still asks for action.

## Agent Usage / 给 Agent 的使用方式

For Codex, Claude, or another coding agent:

1. Read `SKILL.md`.
2. Run or inspect `scripts/install.ps1`.
3. Read the runtime output files:
   - `work/outlook-action-candidates.md`
   - `state/outlook-action-candidates.json`
   - `state/closed-decisions.json` if present
4. Generate a concise daily todo Markdown file.
5. Use stable item numbers such as `A1`, `A2`, `W1`, `F1`, `C1`.
6. When the user says an item is done or irrelevant, write it to `state/closed-decisions.json` so it does not reappear.

Recommended todo sections:

```markdown
# Outlook Daily Todo

## Actions
- [ ] A1 ...

## Waiting / Monitor
- [ ] W1 ...

## Filtered / Not Actionable
- F1 ...

## Suggested Close
- C1 ...
```

## Security Notes / 安全说明

This project does not bypass admin controls. It uses Outlook data already available to the signed-in Windows user through Classic Outlook.

If your company prohibits local automation against Outlook, do not use this tool.

## Project Name

Chinese project name:

```text
读取outlook文件建立代办
```

Suggested English repository name:

```text
classic-outlook-todo
```
