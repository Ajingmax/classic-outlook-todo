# Classic Outlook Reader / Outlook 邮件读取

A read-only skill for local agents on Windows. Read Classic Outlook through COM/MAPI and return JSON. No Microsoft Graph or cloud email permissions are needed.

Windows 本地邮件读取 skill：仅通过 Classic Outlook COM/MAPI 读取元数据和分页纯文本正文，输出 JSON。没有邮件汇总、待办、Case 写入、定时任务或附件导出功能。

## Requirements / 前置条件

- Windows 10/11 and Windows PowerShell 5.1+.
- Classic Outlook installed, signed in, online, and containing the requested SMTP account. New Outlook is unsupported.
- Run under the same Windows user session. Outlook permissions and enterprise policy still apply.
- An agent able to run PowerShell, or run the script directly.

## Install / 安装

Clone this repository or choose **Code → Download ZIP**, then extract it. The repository name remains `classic-outlook-todo` for existing links; its current contents are a reader only. The former todo tool has been removed.

克隆或下载解压即可运行。仓库保留旧名称以保留链接，当前内容仅为读取工具，旧待办功能已删除。无需运行安装程序，不会创建定时任务。

```powershell
git clone https://github.com/Ajingmax/classic-outlook-todo.git
cd classic-outlook-todo
```

For Codex, copy the extracted repository folder to `~/.codex/skills/classic-outlook-reader`. Other agents can read `SKILL.md` directly. Keep mail outputs outside the repository.

作为 Codex skill 使用时，将解压目录复制为 `~/.codex/skills/classic-outlook-reader`；其他 agent 可直接读取 `SKILL.md`。邮件输出应保存在仓库外。

## Scan metadata / 读取列表

Replace the sample account and dates. Include explicit UTC offsets. Scan reads metadata only. Failures stop the scan rather than claiming completeness. Check Truncated; if true, split the time interval and scan again. The interval includes From and excludes Until. Results are sorted by time after collecting up to MaxMessages; they are not necessarily the globally newest messages.

替换示例账号和日期，时间必须包含时区偏移。扫描只读取列表，不加载正文。读取失败会报错；检查 Truncated，若为 true 则缩小时间范围重读。时间范围包含 From、不包含 Until。达到上限时不保证返回全库最新邮件。

```powershell
.\scripts\read-outlook.ps1 -Mode Scan -Account 'you@example.com' `
  -From '2026-11-24T00:00:00+11:00' -Until '2026-11-25T00:00:00+11:00' `
  -MaxMessages 100
```

## Read one message / 读取正文

Copy the exact identifiers from a Scan row. Continue at NextOffset until End=true.

从列表复制原样的 EntryId、StoreId 和 Key；长正文用 NextOffset 继续读取，直到 End=true。

```powershell
.\scripts\read-outlook.ps1 -Mode Read -Account 'you@example.com' `
  -EntryId '<EntryId from Scan>' -StoreId '<StoreId from Scan>' `
  -Key '<Key from Scan>' -Offset 0 -Limit 6000
```

Only JSON is written to stdout. Redirecting output saves private metadata/body to your chosen local file. Do not commit such files. There is no processed ledger; repeated reads may return the same message.

脚本只向标准输出返回 JSON，不保存处理状态，不记录“已处理”。重定向会在指定位置保存私人邮件内容，请勿提交到 GitHub。

## Boundaries / 读取边界

The selected account's accessible delivery store is scanned, excluding deleted/junk/drafts/outbox, non-mail system folders and virtual search folders. Protected/private bodies are refused. Cached data may be incomplete; scans do not establish full server, PST or shared-mailbox coverage. Message-ID hashes deduplicate within each scan; StoreID/EntryID fallback keys may change if moved.

仅读取选定账号本地可访问的邮件库，排除已删除、垃圾、草稿、发件中、非邮件系统文件夹及虚拟搜索文件夹。受保护或私人正文不读取。本地缓存不等于服务器全量。没有发送、移动、删除、归档、标为已读、附件解析或汇总动作。

## Validation / 验证

```powershell
.\scripts\test-reader.ps1
```

Offline tests use synthetic objects and do not connect to Outlook. Real mailbox access depends on the local Outlook profile and permissions.

离线测试使用模拟数据，不连接 Outlook。实际访问取决于本机 Outlook 配置及权限。
