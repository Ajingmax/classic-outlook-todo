param(
    [string]$TodoRoot = "$env:USERPROFILE\Documents\OutlookTodo",
    [string]$ConfigPath = "",
    [int]$ScanLimit = 0
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
    $ConfigPath = Join-Path $TodoRoot "config\outlook-todo.config.json"
}

$WorkPath = Join-Path $TodoRoot "work\outlook-action-candidates.md"
$StatePath = Join-Path $TodoRoot "state\outlook-action-candidates.json"
$DetectedPath = Join-Path $TodoRoot "work\detected-categories.md"
$LogPath = Join-Path $TodoRoot "logs\scan-outlook-last-run.log"

foreach ($path in @($WorkPath, $StatePath, $DetectedPath, $LogPath)) {
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $path) | Out-Null
}

function Read-Config {
    if (Test-Path -LiteralPath $ConfigPath) {
        return Get-Content -LiteralPath $ConfigPath -Encoding UTF8 -Raw | ConvertFrom-Json
    }
    throw "Config file not found: $ConfigPath. Run scripts/install.ps1 first."
}

function Clip-Text([string]$Value, [int]$MaxLength) {
    if ([string]::IsNullOrWhiteSpace($Value)) { return "" }
    $clean = ($Value -replace "\s+", " ").Trim()
    if ($clean.Length -le $MaxLength) { return $clean }
    return $clean.Substring(0, $MaxLength).Trim() + "..."
}

function Normalize-Subject([string]$Subject) {
    if ([string]::IsNullOrWhiteSpace($Subject)) { return "no-subject" }
    $s = $Subject.ToLowerInvariant()
    $s = $s -replace "^\s*(re|fw|fwd)\s*:\s*", ""
    $s = $s -replace "\[ext\]", ""
    $s = $s -replace "\s+", " "
    return $s.Trim()
}

function Get-MailKey($Mail, [string]$Subject) {
    try {
        $conversationId = [string]$Mail.ConversationID
        if (-not [string]::IsNullOrWhiteSpace($conversationId)) {
            return "conversation:" + $conversationId
        }
    } catch { }
    return "subject:" + (Normalize-Subject $Subject)
}

function Test-ContainsAny([string]$Text, $Words) {
    if ([string]::IsNullOrWhiteSpace($Text) -or $null -eq $Words) { return $false }
    $lower = $Text.ToLowerInvariant()
    foreach ($word in @($Words)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$word) -and $lower.Contains(([string]$word).ToLowerInvariant())) {
            return $true
        }
    }
    return $false
}

function Get-CategoryDefinitions {
    return [ordered]@{
        customer_issue = @("issue", "problem", "fault", "failure", "site issue", "customer", "urgent", "complaint")
        meeting = @("meeting", "invite", "reschedule", "cancelled", "agenda", "teams meeting", "appointment")
        logistics = @("delivery", "shipment", "shipping", "warehouse", "courier", "eta", "logistics", "tracking")
        technical = @("technical", "parameter", "firmware", "inverter", "battery", "fault code", "commissioning")
        deadline = @("deadline", "due", "overdue", "today", "tomorrow", "asap", "required by")
        marketing = @("webinar", "campaign", "promotion", "offer", "marketing")
        advertising = @("advertisement", "unsubscribe", "sale", "discount")
        newsletter = @("newsletter", "digest", "subscription", "weekly update")
        system_notification = @("automatic notification", "system notification", "no-reply", "noreply", "do not reply")
    }
}

function Get-MatchingCategories([string]$Text, $Definitions) {
    $matches = New-Object System.Collections.Generic.List[string]
    foreach ($name in $Definitions.Keys) {
        if (Test-ContainsAny $Text $Definitions[$name]) { $matches.Add($name) }
    }
    return @($matches)
}

function Find-FoldersByName($Folder, [string[]]$Names, [System.Collections.Generic.List[object]]$Results) {
    foreach ($child in $Folder.Folders) {
        if ($Names -contains [string]$child.Name) { $Results.Add($child) }
        Find-FoldersByName $child $Names $Results
    }
}

function Add-Folder([System.Collections.Generic.List[object]]$Folders, $Folder) {
    if ($null -eq $Folder) { return }
    foreach ($existing in $Folders) {
        try {
            if ([string]$existing.FolderPath -eq [string]$Folder.FolderPath) { return }
        } catch { }
    }
    $Folders.Add($Folder)
}

$config = Read-Config
$limit = if ($ScanLimit -gt 0) { $ScanLimit } else { [int]$config.limits.scanLimitPerFolder }
$previewChars = [int]$config.limits.bodyPreviewChars
$definitions = Get-CategoryDefinitions

$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")

$targetFolders = New-Object System.Collections.Generic.List[object]
if ($config.folderScope.defaultInbox) {
    Add-Folder $targetFolders ($namespace.GetDefaultFolder(6))
}
if ($config.folderScope.allInboxes) {
    foreach ($store in $namespace.Stores) {
        try {
            $root = $store.GetRootFolder()
            $found = New-Object System.Collections.Generic.List[object]
            Find-FoldersByName $root @("Inbox") $found
            foreach ($folder in $found) { Add-Folder $targetFolders $folder }
        } catch { }
    }
}
if ($config.folderScope.folderNames.Count -gt 0) {
    foreach ($store in $namespace.Stores) {
        try {
            $root = $store.GetRootFolder()
            $found = New-Object System.Collections.Generic.List[object]
            Find-FoldersByName $root @($config.folderScope.folderNames) $found
            foreach ($folder in $found) { Add-Folder $targetFolders $folder }
        } catch { }
    }
}

$sentKeys = @{}
if ($config.sentMail.scanSentItems) {
    try {
        $sent = $namespace.GetDefaultFolder(5)
        $sentItems = $sent.Items
        $sentItems.Sort("[SentOn]", $true)
        $sentCount = [Math]::Min($limit, [int]$sentItems.Count)
        for ($i = 1; $i -le $sentCount; $i++) {
            $mail = $sentItems.Item($i)
            if ($null -eq $mail) { continue }
            $key = Get-MailKey $mail ([string]$mail.Subject)
            $sentKeys[$key] = [datetime]$mail.SentOn
        }
    } catch { }
}

$records = New-Object System.Collections.Generic.List[object]
$categoryCounts = @{}
$scanned = 0

foreach ($folder in $targetFolders) {
    $items = $folder.Items
    try { $items.Sort("[ReceivedTime]", $true) } catch { }
    $count = [Math]::Min($limit, [int]$items.Count)

    for ($i = 1; $i -le $count; $i++) {
        $mail = $items.Item($i)
        if ($null -eq $mail) { continue }
        $messageClass = [string]$mail.MessageClass
        if ($messageClass -notlike "IPM.Note*" -and $messageClass -notlike "IPM.Schedule*") { continue }
        $scanned++

        $subject = [string]$mail.Subject
        $sender = [string]$mail.SenderName
        $toText = [string]$mail.To
        $ccText = [string]$mail.CC
        $body = Clip-Text ([string]$mail.Body) $previewChars
        $content = "$subject $sender $toText $ccText $body"
        $categories = Get-MatchingCategories $content $definitions

        foreach ($cat in $categories) {
            if (-not $categoryCounts.ContainsKey($cat)) { $categoryCounts[$cat] = 0 }
            $categoryCounts[$cat]++
        }

        $excluded = $false
        foreach ($cat in @($config.categories.exclude)) {
            if ($categories -contains $cat) { $excluded = $true; break }
        }
        if ($excluded) { continue }

        if ($config.categories.include.Count -gt 0) {
            $hasIncluded = $false
            foreach ($cat in @($config.categories.include)) {
                if ($categories -contains $cat) { $hasIncluded = $true; break }
            }
            if (-not $hasIncluded -and $messageClass -notlike "IPM.Schedule*") { continue }
        }

        if ($config.recipientScope.unreadOnly -and -not [bool]$mail.UnRead) { continue }
        if ($config.recipientScope.importantOnly -and [int]$mail.Importance -lt 2) { continue }

        $aliases = @($config.userAliases)
        $directHit = Test-ContainsAny $toText $aliases
        $ccHit = Test-ContainsAny $ccText $aliases
        $mentionHit = Test-ContainsAny "$subject $body" $aliases

        $scopeHit = $false
        if ($config.recipientScope.directTo -and $directHit) { $scopeHit = $true }
        if ($config.recipientScope.cc -and $ccHit) { $scopeHit = $true }
        if ($config.recipientScope.mentionInSubjectOrBody -and $mentionHit) { $scopeHit = $true }
        if (-not $scopeHit) { continue }

        $key = Get-MailKey $mail $subject
        $alreadyReplied = $sentKeys.ContainsKey($key)

        $records.Add([pscustomobject]@{
            key = $key
            subject = $subject
            sender = $sender
            received = ([datetime]$mail.ReceivedTime).ToString("o")
            folder = [string]$folder.FolderPath
            entryId = [string]$mail.EntryID
            to = $toText
            cc = $ccText
            directRecipientMatch = $directHit
            ccMatch = $ccHit
            mentionMatch = $mentionHit
            unread = [bool]$mail.UnRead
            importance = [int]$mail.Importance
            categories = $categories
            alreadyReplied = $alreadyReplied
            bodyPreview = $body
        })
    }
}

$sorted = @($records | Sort-Object received -Descending)
$sorted | ConvertTo-Json -Depth 8 | Set-Content -Encoding UTF8 -LiteralPath $StatePath

$lines = New-Object System.Collections.Generic.List[string]
$lines.Add("# Outlook action candidates")
$lines.Add("")
$lines.Add("Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm')")
$lines.Add("Account: $($namespace.CurrentUser.Name)")
$lines.Add("Folders: $($targetFolders.Count)")
$lines.Add("Scanned: $scanned")
$lines.Add("Candidates: $($sorted.Count)")
$lines.Add("")

$n = 1
foreach ($item in $sorted) {
    $lines.Add("## M$n $($item.subject)")
    $lines.Add("")
    $lines.Add("- Sender: $($item.sender)")
    $lines.Add("- Received: $($item.received)")
    $lines.Add("- Folder: $($item.folder)")
    $lines.Add("- To: $($item.to)")
    $lines.Add("- CC: $($item.cc)")
    $lines.Add("- Scope: direct=$($item.directRecipientMatch), cc=$($item.ccMatch), mention=$($item.mentionMatch)")
    $lines.Add("- Categories: $(@($item.categories) -join ', ')")
    $lines.Add("- Already replied from Sent Items: $($item.alreadyReplied)")
    $lines.Add("")
    $lines.Add('```text')
    $lines.Add($item.bodyPreview)
    $lines.Add('```')
    $lines.Add("")
    $n++
}

($lines -join [Environment]::NewLine) | Set-Content -Encoding UTF8 -LiteralPath $WorkPath

$detected = New-Object System.Collections.Generic.List[string]
$detected.Add("# Detected category hints")
$detected.Add("")
$detected.Add("Review these counts and update config/outlook-todo.config.json if needed.")
$detected.Add("")
foreach ($key in ($categoryCounts.Keys | Sort-Object)) {
    $detected.Add("- ${key}: $($categoryCounts[$key])")
}
($detected -join [Environment]::NewLine) | Set-Content -Encoding UTF8 -LiteralPath $DetectedPath

"$(Get-Date -Format o) OK folders=$($targetFolders.Count) scanned=$scanned candidates=$($sorted.Count) output=$WorkPath state=$StatePath" | Set-Content -Encoding UTF8 -LiteralPath $LogPath
Write-Output $WorkPath
