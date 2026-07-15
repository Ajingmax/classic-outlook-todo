param(
    [string]$TodoRoot = "$env:USERPROFILE\Documents\OutlookTodo",
    [string]$TaskName = "Classic Outlook Todo Scan",
    [string[]]$ScanTimes = @("06:55", "11:55"),
    [switch]$InstallIfMissing,
    [switch]$SkipTask
)

$ErrorActionPreference = "Stop"

function Test-ClassicOutlook {
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        $null = $namespace.CurrentUser.Name
        return $true
    } catch {
        return $false
    }
}

function Install-ClassicOutlookBestEffort {
    $winget = Get-Command winget -ErrorAction SilentlyContinue
    if (-not $winget) {
        throw "Classic Outlook is not available and winget is not installed. Install Microsoft 365 Apps or classic Outlook manually, then rerun this script."
    }

    $packages = @(
        "Microsoft.Office",
        "Microsoft.OutlookForWindows"
    )

    foreach ($package in $packages) {
        try {
            winget install --id $package --silent --accept-source-agreements --accept-package-agreements
            if (Test-ClassicOutlook) { return }
        } catch {
            Write-Warning "Install attempt failed for ${package}: $($_.Exception.Message)"
        }
    }

    throw "Automatic Outlook installation did not complete. Install Microsoft 365 Apps with classic Outlook, sign in once, then rerun this script."
}

function Write-Utf8Json($Path, $Object) {
    $json = $Object | ConvertTo-Json -Depth 10
    $dir = Split-Path -Parent $Path
    if ($dir) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText($Path, $json, $utf8Bom)
}

New-Item -ItemType Directory -Force -Path $TodoRoot | Out-Null
foreach ($dir in @("config", "work", "state", "logs", "scripts")) {
    New-Item -ItemType Directory -Force -Path (Join-Path $TodoRoot $dir) | Out-Null
}

$repoRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$scanSource = Join-Path $repoRoot "scripts\scan-outlook.ps1"
$scanTarget = Join-Path $TodoRoot "scripts\scan-outlook.ps1"
Copy-Item -LiteralPath $scanSource -Destination $scanTarget -Force

if (-not (Test-ClassicOutlook)) {
    if ($InstallIfMissing) {
        Install-ClassicOutlookBestEffort
    } else {
        throw "Classic Outlook is not ready. Install or open classic Outlook, sign in, then rerun. To try automatic installation, run with -InstallIfMissing."
    }
}

$configPath = Join-Path $TodoRoot "config\outlook-todo.config.json"
if (-not (Test-Path -LiteralPath $configPath)) {
    $config = [ordered]@{
        version = 1
        userAliases = @($env:USERNAME)
        outputLanguage = "zh-CN"
        recipientScope = [ordered]@{
            directTo = $true
            cc = $false
            mentionInSubjectOrBody = $true
            unreadOnly = $false
            importantOnly = $false
        }
        folderScope = [ordered]@{
            defaultInbox = $true
            allInboxes = $false
            folderNames = @("Inbox", "Important", "Unread Mail")
        }
        sentMail = [ordered]@{
            scanSentItems = $true
            lowerPriorityIfAlreadyReplied = $true
        }
        categories = [ordered]@{
            include = @("customer_issue", "meeting", "logistics", "technical", "deadline")
            exclude = @("marketing", "advertising", "newsletter", "system_notification")
            autoDetect = $true
        }
        reminder = [ordered]@{
            suggestedSummaryTimes = @("07:00", "12:00")
            scanTimes = $ScanTimes
        }
        limits = [ordered]@{
            scanLimitPerFolder = 300
            bodyPreviewChars = 1800
        }
    }
    Write-Utf8Json $configPath $config
}

if (-not $SkipTask) {
    $actions = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$scanTarget`" -TodoRoot `"$TodoRoot`""
    $triggers = @()
    foreach ($time in $ScanTimes) {
        $triggers += New-ScheduledTaskTrigger -Daily -At $time
    }
    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
    Register-ScheduledTask -TaskName $TaskName -Action $actions -Trigger $triggers -Settings $settings -Description "Scan Classic Outlook locally and write todo candidate files." -Force | Out-Null
}

& powershell -NoProfile -ExecutionPolicy Bypass -File $scanTarget -TodoRoot $TodoRoot

[pscustomobject]@{
    TodoRoot = $TodoRoot
    ConfigPath = $configPath
    ScanScript = $scanTarget
    TaskName = if ($SkipTask) { "" } else { $TaskName }
}
