# Run offline using a PowerShell stand-in for the Outlook COM object.
$ErrorActionPreference = 'Stop'
$reader = Join-Path $PSScriptRoot 'read-outlook.ps1'
$tokens = $null
$errors = $null
$null = [System.Management.Automation.Language.Parser]::ParseFile($reader, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw ($errors -join [Environment]::NewLine) }

$now = [DateTimeOffset]::Now
$store = [pscustomobject]@{StoreID='delivery'}
$accessor = [pscustomobject]@{}
$accessor | Add-Member ScriptMethod GetProperty {
    param($property)
    if ($property -like '*1035001F') { return '<synthetic@invalid.example>' }
    return 1
}
$mail = [pscustomobject]@{
    MessageClass='IPM.Note'
    EntryID='message'
    Parent=$store
    PropertyAccessor=$accessor
    ReceivedTime=$now.LocalDateTime
    SentOn=$now.LocalDateTime
    Permission=0
    Sensitivity=0
    ConversationID='test-thread'
    Subject='Synthetic'
    SenderName='Example'
    To='Example'
    CC=''
    Body=('A' + [char]::ConvertFromUtf32(0x1F600) + 'BC')
}
$items = [pscustomobject]@{Count=1; Mail=$mail}
$items | Add-Member ScriptMethod Sort { param($field,$descending) }
$items | Add-Member ScriptMethod Item { param($index) $this.Mail }
$folder = [pscustomobject]@{
    EntryID='folder-6'
    StoreID='delivery'
    PropertyAccessor=$accessor
    DefaultItemType=0
    Items=$items
    Folders=@()
}
$store | Add-Member NoteProperty Inbox $folder
$store | Add-Member ScriptMethod GetDefaultFolder {
    param($id)
    if ($id -eq 6) { return $this.Inbox }
    [pscustomobject]@{EntryID="folder-$id"}
}
$store | Add-Member ScriptMethod GetRootFolder { [pscustomobject]@{Folders=@($this.Inbox)} }
$session = [pscustomobject]@{
    Offline=$false
    Accounts=@([pscustomobject]@{SmtpAddress='mail@example.invalid'; DeliveryStore=$store})
    Mail=$mail
}
$session | Add-Member ScriptMethod GetItemFromID {
    param($id,$storeId)
    if ($id -eq $this.Mail.EntryID) { $this.Mail }
}
$app = [pscustomobject]@{Session=$session}
$app | Add-Member ScriptMethod GetNamespace { param($namespace) $this.Session }
function New-Object {
    param([string]$ComObject)
    if ($ComObject -ne 'Outlook.Application') { throw 'Unexpected test call.' }
    return $app
}
function Assert-Error([scriptblock]$Action, [string]$Pattern) {
    try { & $Action | Out-Null }
    catch {
        if ($_.Exception.Message -match $Pattern) { return }
        throw
    }
    throw "Expected a failure matching '$Pattern'."
}
$scanArgs = @{
    Mode='Scan'
    Account='mail@example.invalid'
    From=$now.AddMinutes(-1).ToString('o')
    Until=$now.AddMinutes(1).ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss.fffffff'Z'")
}
$scan = & $reader @scanArgs | ConvertFrom-Json
if ($scan.Count -ne 1 -or $scan.Truncated -or $scan.Rows[0].PSObject.Properties['Body']) {
    throw 'Scan metadata mismatch.'
}
$readArgs = @{
    Mode='Read'
    Account='mail@example.invalid'
    EntryId=$scan.Rows[0].EntryId
    StoreId=$scan.Rows[0].StoreId
    Key=$scan.Rows[0].Key
    Limit=2
}
$first = & $reader @readArgs | ConvertFrom-Json
$second = & $reader @readArgs -Offset $first.NextOffset | ConvertFrom-Json
$third = & $reader @readArgs -Offset $second.NextOffset | ConvertFrom-Json
if ($first.NextOffset -ne 1 -or $second.NextOffset -ne 3 -or -not $third.End -or
    "$($first.Body)$($second.Body)$($third.Body)" -cne $mail.Body) {
    throw 'Body pagination mismatch.'
}
$mail.Permission = 1
Assert-Error { & $reader @readArgs } 'protected or private'
$mail.Permission = 0
$readArgs.StoreId = 'another-delivery-store'
Assert-Error { & $reader @readArgs } 'does not match the selected account'
$readArgs.StoreId = $scan.Rows[0].StoreId
$readArgs.Key = '0' * 64
Assert-Error { & $reader @readArgs } 'does not match the scanned message'
$scanArgs.Until = $scanArgs.From
Assert-Error { & $reader @scanArgs } 'must precede'
$scanArgs.Until = $now.AddMinutes(1).ToString('s')
Assert-Error { & $reader @scanArgs } 'explicit UTC offset'
'PASS: syntax, metadata scan, Unicode paging, protected body, store/key checks, time windows and offsets. No real Outlook access.'
