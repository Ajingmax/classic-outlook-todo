# Read a configured Classic Outlook mailbox via its local COM/MAPI profile.
[CmdletBinding(DefaultParameterSetName = 'Scan')]
param(
    [Parameter(ParameterSetName = 'Scan')]
    [Parameter(Mandatory, ParameterSetName = 'Read')]
    [ValidateSet('Scan','Read')][string]$Mode = 'Scan',
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Account,
    [Parameter(Mandatory, ParameterSetName = 'Scan')][string]$From,
    [Parameter(Mandatory, ParameterSetName = 'Scan')][string]$Until,
    [Parameter(ParameterSetName = 'Scan')][ValidateRange(1,10000)][int]$MaxMessages = 500,
    [Parameter(Mandatory, ParameterSetName = 'Read')][ValidateNotNullOrEmpty()][string]$EntryId,
    [Parameter(Mandatory, ParameterSetName = 'Read')][ValidateNotNullOrEmpty()][string]$StoreId,
    [Parameter(Mandatory, ParameterSetName = 'Read')][ValidatePattern('^[0-9a-fA-F]{64}$')][string]$Key,
    [Parameter(ParameterSetName = 'Read')][ValidateRange(0,2147483647)][int]$Offset = 0,
    [Parameter(ParameterSetName = 'Read')][ValidateRange(1,20000)][int]$Limit = 6000
)
$ErrorActionPreference = 'Stop'
if ($Mode -ne $PSCmdlet.ParameterSetName) { throw 'Mode does not match the supplied parameters.' }
$Account = $Account.Trim()
if ([string]::IsNullOrWhiteSpace($Account)) { throw 'Specify the SMTP address of a configured Outlook account.' }
function Convert-ScanDate([string]$Value, [string]$Name) {
    $parsed = [DateTimeOffset]::MinValue
    if ($Value -notmatch '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(:\d{2}(\.\d{1,7})?)?(Z|[+-]\d{2}:\d{2})$' -or
        -not [DateTimeOffset]::TryParse($Value, [Globalization.CultureInfo]::InvariantCulture,
            [Globalization.DateTimeStyles]::None, [ref]$parsed)) {
        throw "$Name must be an ISO 8601 timestamp with Z or an explicit UTC offset."
    }
    return $parsed
}
if ($Mode -eq 'Scan') {
    $fromDate = Convert-ScanDate $From 'From'
    $untilDate = Convert-ScanDate $Until 'Until'
    if ($fromDate -ge $untilDate) { throw 'From must precede Until.' }
}

function Get-MailStore {
    for ($attempt = 0; $attempt -lt 5; $attempt++) {
        try { $app = New-Object -ComObject Outlook.Application; break }
        catch [System.Runtime.InteropServices.COMException] {
            if ($_.Exception.HResult -ne -2147418111 -or $attempt -eq 4) { throw }
            Start-Sleep -Milliseconds 1200
        }
    }
    $script:session = $app.GetNamespace('MAPI')
    if ($session.Offline) { throw 'Outlook is offline. Reconnect and retry.' }
    foreach ($profileAccount in $session.Accounts) {
        if ([string]$profileAccount.SmtpAddress -eq $Account) {
            $deliveryStore = $profileAccount.DeliveryStore
            if (-not $deliveryStore) { throw 'The selected account has no delivery store.' }
            return ,$deliveryStore
        }
    }
    throw 'Classic Outlook does not have the configured SMTP account.'
}

function Test-MailItem($mail) {
    return $mail.MessageClass -like 'IPM.Note*' -or $mail.MessageClass -like 'IPM.Schedule.Meeting*'
}

function Get-Key($mail) {
    $accessor = $mail.PropertyAccessor
    try { $messageId = [string]$accessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x1035001F') }
    catch [System.Runtime.InteropServices.COMException] {
        if ($_.Exception.GetBaseException().HResult -ne -2147221233) { throw }
        $messageId = ''
    }
    if (-not $messageId) { $messageId = 'entry:' + [string]$mail.Parent.StoreID + ':' + [string]$mail.EntryID }
    $hash = [Security.Cryptography.SHA256]::Create()
    try { return [BitConverter]::ToString($hash.ComputeHash([Text.Encoding]::UTF8.GetBytes($messageId))).Replace('-','').ToLowerInvariant() }
    finally { $hash.Dispose() }
}

function Test-Restricted($mail) {
    # Meeting items do not always implement the MailItem permission property.
    $permission = $mail.Permission
    return $mail.MessageClass -like '*.RPMSG*' -or $permission -gt 0 -or $mail.Sensitivity -ge 2
}

function Get-DatedRows($folder, [bool]$inSent, [string]$field) {
    $items = $folder.Items
    $items.Sort("[$field]", $true)
    for ($itemIndex = 1; $itemIndex -le $items.Count; $itemIndex++) {
        $mail = $items.Item($itemIndex)
        if (-not (Test-MailItem $mail)) { continue }
        $stamp = $mail.$field
        if (-not $stamp) { continue }
        $at = [DateTimeOffset]([datetime]$stamp)
        if ($at -lt $fromDate) { break }
        if ($at -ge $untilDate) { continue }
        $messageKey = Get-Key $mail
        if (-not $known.Add($messageKey)) { continue }
        if ($rows.Count -ge $MaxMessages) {
            $script:truncated = $true
            return
        }
        $rows.Add([ordered]@{
            Key = $messageKey
            EntryId = [string]$mail.EntryID
            StoreId = [string]$folder.StoreID
            AtUtc = $at.ToUnixTimeSeconds()
            TimeField = $field
            InSent = $inSent
            Thread = [string]$mail.ConversationID
            Subject = [string]$mail.Subject
            Sender = [string]$mail.SenderName
            To = [string]$mail.To
            Cc = [string]$mail.CC
            Restricted = Test-Restricted $mail
        })
    }
}

function Get-FolderRows($folder, [bool]$inSent) {
    if ($script:truncated -or $excluded.Contains([string]$folder.EntryID)) { return }
    if ($folder.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x36010003') -eq 2) { return }
    $inSent = $inSent -or ([string]$folder.EntryID -eq $sentId)
    if ($folder.DefaultItemType -eq 0) {
        if ($inSent) { Get-DatedRows $folder $true 'SentOn' }
        else {
            Get-DatedRows $folder $false 'ReceivedTime'
            # A manually filed sent item may not have a meaningful ReceivedTime.
            if (-not $script:truncated -and [string]$folder.EntryID -ne $inboxId) { Get-DatedRows $folder $false 'SentOn' }
        }
    }
    foreach ($child in $folder.Folders) { Get-FolderRows $child $inSent }
}

$store = Get-MailStore
if ($Mode -eq 'Read') {
    if ([string]$store.StoreID -ne $StoreId) { throw 'StoreId does not match the selected account.' }
    $mail = $session.GetItemFromID($EntryId, $StoreId)
    if (-not (Test-MailItem $mail) -or [string]$mail.Parent.StoreID -ne $StoreId -or (Get-Key $mail) -ne $Key) {
        throw 'The Outlook item does not match the scanned message.'
    }
    if (Test-Restricted $mail) { throw 'Outlook marks this message protected or private. Body is unavailable.' }
    $body = [string]$mail.Body
    if ($Offset -gt $body.Length) { throw 'Offset exceeds the body length.' }
    $length = [Math]::Min($Limit, $body.Length - $Offset)
    # Keep character offsets stable and avoid splitting a UTF-16 surrogate pair.
    if ($Offset -gt 0 -and $Offset -lt $body.Length -and
        [char]::IsHighSurrogate($body[$Offset - 1]) -and [char]::IsLowSurrogate($body[$Offset])) {
        throw 'Offset splits a Unicode character. Use the previous page NextOffset.'
    }
    if ($length -gt 0 -and $Offset + $length -lt $body.Length -and
        [char]::IsHighSurrogate($body[$Offset + $length - 1]) -and [char]::IsLowSurrogate($body[$Offset + $length])) {
        if ($length -eq 1) { throw 'Limit is too small for the next Unicode character. Use a limit of at least 2.' }
        $length--
    }
    $next = $Offset + $length
    $result = [ordered]@{
        Key = $Key.ToLowerInvariant()
        EntryId = $EntryId
        StoreId = $StoreId
        Thread = [string]$mail.ConversationID
        Subject = [string]$mail.Subject
        Body = $body.Substring($Offset, $length)
        NextOffset = $next
        End = $next -eq $body.Length
    }
} else {
    $known = [Collections.Generic.HashSet[string]]::new()
    $rows = [Collections.Generic.List[object]]::new()
    $excluded = [Collections.Generic.HashSet[string]]::new()
    $script:truncated = $false
    # Deleted, Outbox, Calendar, Contacts, Journal, Notes, Tasks, Drafts,
    # Junk, Conflicts, Sync Issues, Local Failures, Server Failures, RSS.
    foreach ($id in @(3,4,9,10,11,12,13,16,23,19,20,21,22,25)) {
        try { $systemFolder = $store.GetDefaultFolder($id) }
        catch [System.Runtime.InteropServices.COMException] {
            if ($_.Exception.GetBaseException().HResult -in @(-2147221233,-2147024809)) { continue }
            throw
        }
        if ($systemFolder) { [void]$excluded.Add([string]$systemFolder.EntryID) }
    }
    $sentId = [string]$store.GetDefaultFolder(5).EntryID
    $inboxId = [string]$store.GetDefaultFolder(6).EntryID
    foreach ($folder in $store.GetRootFolder().Folders) { Get-FolderRows $folder $false }
    $result = [ordered]@{
        FromUtc = $fromDate.ToUniversalTime().ToString('o')
        UntilUtc = $untilDate.ToUniversalTime().ToString('o')
        Count = $rows.Count
        Truncated = $script:truncated
        Rows = @($rows.ToArray() | Sort-Object AtUtc, Key)
    }
}
ConvertTo-Json -InputObject $result -Depth 4 -Compress
