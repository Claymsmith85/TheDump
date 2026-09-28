#Requires -Version 7.0
<#
.SYNOPSIS
    Inventories and copies mailbox content between two Exchange Online mailboxes using EWS
    ExportItems/UploadItems (full fidelity). Progress is tracked per item in an inventory CSV.

.DESCRIPTION
    Phase 1 - Inventory (-InventoryOnly):
        Scans the source mailbox and writes one row per item to the inventory CSV with Status = Pending.
        Running it again later adds new items and leaves existing rows untouched.

    Phase 2 - Transfer (run without -InventoryOnly):
        Copies every Pending item. Each row is marked Copied (with timestamp and target item ID)
        or Failed (with the error), and the script moves on to the next item.
        Failed items are left alone on later runs unless -RetryFailed is used.
        If no inventory exists yet, one is built first.

    Throttling (per Microsoft's EWS throttling guidance):
        Exchange Online does not publish or expose its EWS budget values, so the script adapts to
        them: it runs one request at a time (well under the 27-connection limit), keeps batches
        small enough to finish inside the server's one-minute batch limit, and honors
        BackOffMilliseconds. On every throttle it doubles the gap between requests and halves the
        batch size; after 20 clean requests in a row it eases back toward full speed. This keeps
        the run just under whatever limit the service is enforcing at the time.

    Duplicate protection (safe to rerun):
        - Target folders are matched by path and only created if missing.
        - Each copied item is tagged in the target with a hidden property holding its source key.
          Before copying into a folder, the script reads that folder and skips any item already
          present (by tag, or by matching search key), marking it Exists. This holds even if the
          inventory file is lost or rebuilt.
        - An upload interrupted by a connection error is not blindly resent; it is marked Failed and
          rechecked against the target on the next -RetryFailed run.

    Crash safety:
        Status changes are written immediately to <inventory>.journal and merged into the inventory
        every -CheckpointSeconds and at exit. If the inventory is open in Excel, the merge is
        deferred and nothing is lost.

    Auth: delegated EWS.AccessAsUser.All via username/password (ROPC). The signed-in account needs
    FullAccess on BOTH mailboxes. Use the same -TargetSubfolder / -ExcludeFolders for both phases.

.EXAMPLE
    # 1. Inventory
    .\Copy-MailboxEws.ps1 -TenantId contoso.onmicrosoft.com -ClientId <appId> `
        -SourceMailbox old@contoso.com -TargetMailbox new@contoso.com -TargetSubfolder 'Migrated - old' -InventoryOnly

    # 2. Transfer (same parameters, without -InventoryOnly)
    .\Copy-MailboxEws.ps1 -TenantId contoso.onmicrosoft.com -ClientId <appId> `
        -SourceMailbox old@contoso.com -TargetMailbox new@contoso.com -TargetSubfolder 'Migrated - old'
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$TenantId,
    [Parameter(Mandatory)][string]$ClientId,
    [Parameter(Mandatory)][string]$SourceMailbox,
    [Parameter(Mandatory)][string]$TargetMailbox,
    [string]$UserName,
    [string]$InventoryPath,
    [switch]$InventoryOnly,
    [switch]$RetryFailed,
    [string]$TargetSubfolder,
    [ValidateSet('msgfolderroot', 'archivemsgfolderroot')][string]$SourceRoot = 'msgfolderroot',
    [ValidateSet('msgfolderroot', 'archivemsgfolderroot')][string]$TargetRoot = 'msgfolderroot',
    [string[]]$ExcludeFolders = @('\Sync Issues', '\Conversation History', '\Outbox'),
    [ValidateRange(1, 100)][int]$BatchSize = 10,
    [long]$MaxBatchBytes = 10MB,
    [int]$MaxRetries = 8,
    [int]$CheckpointSeconds = 60
)

$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'

$EwsUrl   = 'https://outlook.office365.com/EWS/Exchange.asmx'
$AuthBase = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0"
$Scope    = 'https://outlook.office365.com/EWS.AccessAsUser.All offline_access'

# Hidden property stamped on every copied item in the target (fixed GUID for this tool)
$TagUri       = '<t:ExtendedFieldURI PropertySetId="6a3e8f1c-2b7d-4c9e-9f0a-5d1b3c7e2a41" PropertyName="MigrationSourceKey" PropertyType="String"/>'
$SearchKeyUri = '<t:ExtendedFieldURI PropertyTag="0x300B" PropertyType="Binary"/>'

if (-not $InventoryPath) {
    $InventoryPath = 'MailboxCopy_{0}_to_{1}.csv' -f ($SourceMailbox -replace '[^\w.-]', '_'), ($TargetMailbox -replace '[^\w.-]', '_')
}
$InventoryPath = [IO.Path]::GetFullPath($InventoryPath, $PWD.ProviderPath)
$JournalPath   = "$InventoryPath.journal"

#region Helpers
function Esc([string]$s) { [Security.SecurityElement]::Escape($s) }

function Get-Node($Node, [string]$Name) { $Node.SelectSingleNode("*[local-name()='$Name']") }

function Get-Text($Node, [string]$Name) {
    $n = Get-Node $Node $Name
    if ($null -ne $n) { $n.InnerText }
}

function Get-DistinguishedXml([string]$Mailbox, [string]$Id) {
    '<t:DistinguishedFolderId Id="{0}"><t:Mailbox><t:EmailAddress>{1}</t:EmailAddress></t:Mailbox></t:DistinguishedFolderId>' -f $Id, (Esc $Mailbox)
}

function Get-BackOffMs($Msg) {
    $v = $Msg.SelectSingleNode(".//*[local-name()='Value'][@Name='BackOffMilliseconds']")
    if ($null -ne $v) { [int]$v.InnerText } else { 30000 }
}

function Test-Excluded([string]$Path, [string[]]$List) {
    foreach ($x in $List) {
        if ($Path -eq $x -or $Path.StartsWith("$x\", [StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    $false
}

function Get-TargetPath([string]$Path) {
    if ($TargetSubfolder) { '\' + $TargetSubfolder.Trim('\') + $Path } else { $Path }
}

# --- Adaptive pacing: back off hard on throttling, recover slowly when healthy ---
$ThrottleCodes = @('ErrorServerBusy', 'ErrorExceededConnectionCount', 'ErrorExceededFindCountLimit', 'ErrorTooManyObjectsOpened')

function Test-Throttled($Msg) {
    $code = Get-Text $Msg 'ResponseCode'
    ($code -in $ThrottleCodes) -or ($code -eq 'ErrorInternalServerError' -and $Msg.OuterXml -match 'ServerBusy')
}

function Find-ThrottleMessage([xml]$Xml) {
    foreach ($m in $Xml.SelectNodes("//*[@ResponseClass='Error']")) { if (Test-Throttled $m) { return $m } }
}

function Register-Throttle {
    $p = $script:Pace
    $p.DelayMs     = [math]::Min(10000, [math]::Max(500, $p.DelayMs * 2))
    $p.BatchSize   = [math]::Max(1, [int][math]::Floor($p.BatchSize / 2))
    $p.BatchBytes  = [math]::Max(1MB, [long]($p.BatchBytes / 2))
    $p.CleanStreak = 0
}

function Register-Success {
    $p = $script:Pace
    $p.CleanStreak++
    if ($p.CleanStreak -ge 20) {
        $p.CleanStreak = 0
        $p.DelayMs     = [int]($p.DelayMs * 0.75); if ($p.DelayMs -lt 50) { $p.DelayMs = 0 }
        $p.BatchSize   = [math]::Min($BatchSize, $p.BatchSize + 1)
        $p.BatchBytes  = [math]::Min($MaxBatchBytes, $p.BatchBytes + 2MB)
    }
}

function Write-Advisory([string]$Reason, [double]$Seconds) {
    $now = Get-Date
    $p   = $script:Pace
    Write-Host ('[{0}] BACKOFF: {1}. Pausing {2:N0} s, resuming about {3}. Pace now: {4} ms between requests, batch {5} items / {6:N0} MB.' -f `
        $now.ToString('HH:mm:ss'), $Reason, $Seconds, $now.AddSeconds($Seconds).ToString('HH:mm:ss'),
        $p.DelayMs, $p.BatchSize, ($p.BatchBytes / 1MB)) `
        -ForegroundColor Black -BackgroundColor Yellow
    $script:Stats.Backoffs++
    $script:Stats.BackoffSeconds += $Seconds
}
#endregion

#region Auth (username/password + refresh)
function Set-Token($r) {
    $script:Token = [pscustomobject]@{
        AccessToken  = $r.access_token
        RefreshToken = if ($r.refresh_token) { $r.refresh_token } else { $script:Token.RefreshToken }
        ExpiresOn    = (Get-Date).AddSeconds([int]$r.expires_in)
    }
}

function Connect-Ews {
    # Username/password sign-in (ROPC). Does not work if the account must complete MFA.
    # Credential is cached in $global:EwsCred for this PowerShell session.
    $msg  = 'Account with FullAccess on both mailboxes (UPN)'
    $cred = $global:EwsCred
    if (-not $cred -or ($UserName -and $cred.UserName -ne $UserName)) {
        $cred = if ($UserName) { Get-Credential -UserName $UserName -Message $msg } else { Get-Credential -Message $msg }
    }
    $resp = Invoke-WebRequest -Method Post -Uri "$AuthBase/token" -SkipHttpErrorCheck -Body @{
        grant_type = 'password'
        client_id  = $ClientId
        scope      = $Scope
        username   = $cred.UserName
        password   = $cred.GetNetworkCredential().Password
    }
    $json = $resp.Content | ConvertFrom-Json
    if ([int]$resp.StatusCode -ne 200) {
        Remove-Variable EwsCred -Scope Global -ErrorAction SilentlyContinue   # don't keep a bad password
        throw "Sign-in failed: $($json.error_description)"
    }
    $global:EwsCred = $cred
    Set-Token $json
    Write-Host "Signed in as $($cred.UserName)."
}

function Get-AccessToken {
    if ((Get-Date).AddMinutes(5) -ge $script:Token.ExpiresOn) {
        $r = Invoke-RestMethod -Method Post -Uri "$AuthBase/token" -Body @{
            grant_type    = 'refresh_token'
            client_id     = $ClientId
            refresh_token = $script:Token.RefreshToken
            scope         = $Scope
        }
        Set-Token $r
    }
    $script:Token.AccessToken
}
#endregion

#region EWS transport
function Invoke-Ews {
    param([Parameter(Mandatory)][string]$Body, [Parameter(Mandatory)][string]$AnchorMailbox,
          [switch]$NoNetworkRetry,    # don't resend on connection loss (uploads)
          [switch]$PerItemThrottle)   # caller handles per-item throttle errors in a 200 response

    $envelope = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
  <soap:Header><t:RequestServerVersion Version="Exchange2013_SP1"/></soap:Header>
  <soap:Body>$Body</soap:Body>
</soap:Envelope>
"@
    $bytes = [Text.Encoding]::UTF8.GetBytes($envelope)

    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        $headers = @{
            Authorization     = "Bearer $(Get-AccessToken)"
            'X-AnchorMailbox' = $AnchorMailbox
        }
        if ($script:Pace.DelayMs -gt 0) { Start-Sleep -Milliseconds $script:Pace.DelayMs }
        $timer = [Diagnostics.Stopwatch]::StartNew()
        try {
            $resp = Invoke-WebRequest -Uri $EwsUrl -Method Post -Headers $headers -Body $bytes `
                -ContentType 'text/xml; charset=utf-8' -SkipHttpErrorCheck -TimeoutSec 900
        }
        catch {
            # For uploads the server may have created the items before the connection dropped,
            # so resending could duplicate them. Fail instead; the next run checks the target first.
            if ($NoNetworkRetry) { throw "Connection lost, outcome unknown ($($_.Exception.Message))" }
            $wait = [math]::Min(300, 15 * $attempt)
            Write-Advisory "Connection error ($($_.Exception.Message)), attempt $attempt of $MaxRetries" $wait
            Start-Sleep -Seconds $wait
            continue
        }

        $code    = [int]$resp.StatusCode
        $content = [string]$resp.Content

        if ($code -eq 200) {
            $xml = [xml]$content.TrimStart([char]0xFEFF)

            # Throttling can also arrive inside a 200 response (ErrorServerBusy, or ErrorInternalServerError
            # with an inner ServerBusy). Retry the whole request unless the caller handles items itself.
            if (-not $PerItemThrottle) {
                $busy = Find-ThrottleMessage $xml
                if ($null -ne $busy -and $attempt -lt $MaxRetries) {
                    Register-Throttle
                    $wait = (Get-BackOffMs $busy) / 1000
                    Write-Advisory ("Server busy ({0}), attempt {1} of {2}" -f (Get-Text $busy 'ResponseCode'), $attempt, $MaxRetries) $wait
                    Start-Sleep -Seconds $wait
                    continue
                }
            }

            # EWS stops executing a batch after one minute. If a request is running long, shrink batches early.
            if ($timer.Elapsed.TotalSeconds -gt 40) {
                $p = $script:Pace
                $p.BatchSize   = [math]::Max(1, [int][math]::Floor($p.BatchSize / 2))
                $p.BatchBytes  = [math]::Max(1MB, [long]($p.BatchBytes / 2))
                $p.CleanStreak = 0
                Write-Host ('[{0}] PACE: request took {1:N0} s (server limit is 60 s). Batch reduced to {2} items / {3:N0} MB.' -f `
                    (Get-Date).ToString('HH:mm:ss'), $timer.Elapsed.TotalSeconds, $p.BatchSize, ($p.BatchBytes / 1MB)) `
                    -ForegroundColor Black -BackgroundColor Yellow
            }
            else { Register-Success }
            return $xml
        }

        if ($code -eq 401) { $script:Token.ExpiresOn = Get-Date; continue }   # force token refresh

        if (($code -in @(429, 503)) -or ($code -eq 500 -and $content -match 'ErrorServerBusy|ErrorExceededConnectionCount')) {
            Register-Throttle
            $wait = 30 * $attempt
            if ($content -match 'Name="BackOffMilliseconds">(\d+)<') { $wait = [math]::Ceiling([int]$Matches[1] / 1000) }
            elseif ($resp.Headers['Retry-After']) { $wait = [int]($resp.Headers['Retry-After'] | Select-Object -First 1) }
            Write-Advisory "Server throttling (HTTP $code), attempt $attempt of $MaxRetries" $wait
            Start-Sleep -Seconds $wait
            continue
        }

        $fault = if ($content -match '<faultstring[^>]*>(.*?)</faultstring>') { $Matches[1] }
                 else { $content.Substring(0, [math]::Min(300, $content.Length)) }
        if ($code -eq 403) { $fault += ' (EWS may be blocked for this tenant, mailbox, or app ID - check EWSEnabled / EWSAllowedAppIDs)' }
        throw "EWS HTTP $code : $fault"
    }
    throw "EWS request failed after $MaxRetries attempts."
}

function Assert-Success([xml]$Xml, [string]$Name) {
    $rm = $Xml.SelectSingleNode("//*[local-name()='$Name']")
    if ($null -eq $rm) { throw "Unexpected EWS response (no $Name)." }
    if ($rm.GetAttribute('ResponseClass') -ne 'Success') {
        throw ('{0}: {1} {2}' -f $Name, (Get-Text $rm 'ResponseCode'), (Get-Text $rm 'MessageText'))
    }
    $rm
}
#endregion

#region Folders
function Get-RootFolderId([string]$Mailbox, [string]$Root) {
    $body = "<m:GetFolder><m:FolderShape><t:BaseShape>IdOnly</t:BaseShape></m:FolderShape><m:FolderIds>$(Get-DistinguishedXml $Mailbox $Root)</m:FolderIds></m:GetFolder>"
    $rm = Assert-Success (Invoke-Ews -Body $body -AnchorMailbox $Mailbox) 'GetFolderResponseMessage'
    $rm.SelectSingleNode(".//*[local-name()='FolderId']").GetAttribute('Id')
}

function Get-FolderTree([string]$Mailbox, [string]$Root) {
    $offset = 0
    do {
        $body = @"
<m:FindFolder Traversal="Deep">
  <m:FolderShape>
    <t:BaseShape>IdOnly</t:BaseShape>
    <t:AdditionalProperties>
      <t:FieldURI FieldURI="folder:DisplayName"/>
      <t:FieldURI FieldURI="folder:FolderClass"/>
      <t:FieldURI FieldURI="folder:TotalCount"/>
      <t:ExtendedFieldURI PropertyTag="0x66B5" PropertyType="String"/>
      <t:ExtendedFieldURI PropertyTag="0x10F4" PropertyType="Boolean"/>
    </t:AdditionalProperties>
  </m:FolderShape>
  <m:IndexedPageFolderView MaxEntriesReturned="500" Offset="$offset" BasePoint="Beginning"/>
  <m:ParentFolderIds>$(Get-DistinguishedXml $Mailbox $Root)</m:ParentFolderIds>
</m:FindFolder>
"@
        $rm       = Assert-Success (Invoke-Ews -Body $body -AnchorMailbox $Mailbox) 'FindFolderResponseMessage'
        $rootNode = Get-Node $rm 'RootFolder'
        $list     = Get-Node $rootNode 'Folders'
        if ($null -ne $list) {
            foreach ($f in $list.ChildNodes) {
                $ext = @{}
                foreach ($ep in $f.SelectNodes("*[local-name()='ExtendedProperty']")) {
                    $ext[(Get-Node $ep 'ExtendedFieldURI').GetAttribute('PropertyTag')] = Get-Text $ep 'Value'
                }
                [pscustomobject]@{
                    Id     = (Get-Node $f 'FolderId').GetAttribute('Id')
                    Type   = $f.LocalName
                    Class  = Get-Text $f 'FolderClass'
                    Count  = [int](Get-Text $f 'TotalCount')
                    Path   = ([string]$ext['0x66B5']).Replace([string][char]0xFFFE, '\')
                    Hidden = $ext['0x10F4'] -eq 'true'
                }
            }
        }
        $offset = [int]$rootNode.GetAttribute('IndexedPagingOffset')
    } while ($rootNode.GetAttribute('IncludesLastItemInRange') -ne 'true')
}

function Get-SourceFolders {
    $all     = @(Get-FolderTree $SourceMailbox $SourceRoot)
    $exclude = @($ExcludeFolders) + @($all | Where-Object Hidden | ForEach-Object Path)
    $all | Where-Object { $_.Type -ne 'SearchFolder' -and $_.Path -and -not (Test-Excluded $_.Path $exclude) } |
        Sort-Object Path
}

function Resolve-TargetFolder([string]$Path, [string]$Class) {
    if ($script:TargetMap.ContainsKey($Path)) { return $script:TargetMap[$Path] }

    $cut      = $Path.LastIndexOf('\')
    $parent   = $Path.Substring(0, $cut)
    $name     = $Path.Substring($cut + 1)
    $parentId = if ($parent) { Resolve-TargetFolder $parent 'IPF.Note' } else { $script:TargetRootId }
    $classXml = if ($Class) { "<t:FolderClass>$(Esc $Class)</t:FolderClass>" } else { '' }

    $body = @"
<m:CreateFolder>
  <m:ParentFolderId><t:FolderId Id="$parentId"/></m:ParentFolderId>
  <m:Folders><t:Folder>$classXml<t:DisplayName>$(Esc $name)</t:DisplayName></t:Folder></m:Folders>
</m:CreateFolder>
"@
    $rm = Assert-Success (Invoke-Ews -Body $body -AnchorMailbox $TargetMailbox) 'CreateFolderResponseMessage'
    $id = $rm.SelectSingleNode(".//*[local-name()='FolderId']").GetAttribute('Id')
    Write-Host "  Created target folder $Path"
    $script:TargetMap[$Path] = $id
    $id
}
#endregion

#region Items
function Get-FolderItems([string]$Mailbox, [string]$FolderId) {
    $offset = 0
    do {
        $body = @"
<m:FindItem Traversal="Shallow">
  <m:ItemShape>
    <t:BaseShape>IdOnly</t:BaseShape>
    <t:AdditionalProperties>
      <t:FieldURI FieldURI="item:Subject"/>
      <t:FieldURI FieldURI="item:ItemClass"/>
      <t:FieldURI FieldURI="item:Size"/>
      <t:FieldURI FieldURI="item:DateTimeReceived"/>
      $SearchKeyUri
      $TagUri
    </t:AdditionalProperties>
  </m:ItemShape>
  <m:IndexedPageItemView MaxEntriesReturned="500" Offset="$offset" BasePoint="Beginning"/>
  <m:ParentFolderIds><t:FolderId Id="$FolderId"/></m:ParentFolderIds>
</m:FindItem>
"@
        $rm       = Assert-Success (Invoke-Ews -Body $body -AnchorMailbox $Mailbox) 'FindItemResponseMessage'
        $rootNode = Get-Node $rm 'RootFolder'
        $list     = Get-Node $rootNode 'Items'
        if ($null -ne $list) {
            foreach ($i in $list.ChildNodes) {
                $id = (Get-Node $i 'ItemId').GetAttribute('Id')
                $sk = $null; $tag = $null
                foreach ($ep in $i.SelectNodes("*[local-name()='ExtendedProperty']")) {
                    if ((Get-Node $ep 'ExtendedFieldURI').GetAttribute('PropertyName') -eq 'MigrationSourceKey') { $tag = Get-Text $ep 'Value' }
                    else { $sk = Get-Text $ep 'Value' }
                }
                [pscustomobject]@{
                    Id        = $id
                    Subject   = Get-Text $i 'Subject'
                    Class     = Get-Text $i 'ItemClass'
                    Size      = [long](Get-Text $i 'Size')
                    Received  = Get-Text $i 'DateTimeReceived'
                    SearchKey = $sk
                    Tag       = $tag
                    Key       = if ($sk) { "sk:$sk" } else { "id:$id" }
                }
            }
        }
        $offset = [int]$rootNode.GetAttribute('IndexedPagingOffset')
    } while ($rootNode.GetAttribute('IncludesLastItemInRange') -ne 'true')
}

# Keys of items already in a target folder: tags written by this script, plus search keys
function Get-TargetKeys([string]$FolderId) {
    $set = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($it in (Get-FolderItems $TargetMailbox $FolderId)) {
        if ($it.Tag)       { [void]$set.Add($it.Tag) }
        if ($it.SearchKey) { [void]$set.Add("sk:$($it.SearchKey)") }
    }
    return ,$set
}

# Stamp uploaded items with their source key; returns one response message per item
function Set-TargetTag([object[]]$Uploaded) {
    $changes = -join ($Uploaded | ForEach-Object {
        '<t:ItemChange><t:ItemId Id="{0}"/><t:Updates><t:SetItemField>{1}<t:Item><t:ExtendedProperty>{1}<t:Value>{2}</t:Value></t:ExtendedProperty></t:Item></t:SetItemField></t:Updates></t:ItemChange>' -f `
            $_.NewId, $TagUri, (Esc $_.Row.MigrationKey)
    })
    $body = "<m:UpdateItem ConflictResolution=`"AlwaysOverwrite`" MessageDisposition=`"SaveOnly`" SendMeetingInvitationsOrCancellations=`"SendToNone`"><m:ItemChanges>$changes</m:ItemChanges></m:UpdateItem>"
    $xml  = Invoke-Ews -AnchorMailbox $TargetMailbox -Body $body
    $xml.SelectNodes("//*[local-name()='UpdateItemResponseMessage']")
}
#endregion

#region Inventory file + journal
function Open-Journal {
    $script:JournalHasHeader = (Test-Path $JournalPath) -and ((Get-Item $JournalPath).Length -gt 0)
    $script:Journal = [IO.StreamWriter]::new($JournalPath, $true, [Text.UTF8Encoding]::new($false))
    $script:Journal.AutoFlush = $true
}

function Import-Inventory {
    if (Test-Path $InventoryPath) {
        foreach ($r in (Import-Csv $InventoryPath)) {
            if (-not $r.PSObject.Properties['MigrationKey']) { $r | Add-Member -NotePropertyName MigrationKey -NotePropertyValue "id:$($r.SourceItemId)" }
            $script:Rows.Add($r); $script:Index[$r.SourceItemId] = $r
        }
    }
    $pendingJournal = (Test-Path $JournalPath) -and ((Get-Item $JournalPath).Length -gt 0)
    if ($pendingJournal) {
        foreach ($j in (Import-Csv $JournalPath)) {
            $r = $null
            if ($script:Index.TryGetValue($j.SourceItemId, [ref]$r)) {
                $r.Status = $j.Status; $r.ProcessedAt = $j.ProcessedAt; $r.TargetItemId = $j.TargetItemId; $r.Error = $j.Error
            }
        }
    }
    if ($script:Rows.Count) { Write-Host "Loaded inventory: $($script:Rows.Count) item(s) from $InventoryPath" }
    $pendingJournal
}

function Save-Inventory {
    if ($script:Journal) { $script:Journal.Dispose(); $script:Journal = $null }
    try {
        $tmp = "$InventoryPath.tmp"
        $script:Rows | Export-Csv -Path $tmp -NoTypeInformation -Encoding utf8BOM
        Move-Item -LiteralPath $tmp -Destination $InventoryPath -Force
        Remove-Item -LiteralPath $JournalPath -ErrorAction SilentlyContinue
    }
    catch {
        Write-Warning "Could not update $InventoryPath ($($_.Exception.Message)). Progress is kept in the journal and will be merged later."
    }
    Open-Journal
    $script:LastCheckpoint = Get-Date
}

function Set-ItemStatus($Row, [string]$Status, [string]$TargetItemId, [string]$Message) {
    $Row.Status       = $Status
    $Row.ProcessedAt  = (Get-Date).ToString('s')
    $Row.TargetItemId = $TargetItemId
    $Row.Error        = $Message -replace '[\r\n]+', ' '

    $csv = [pscustomobject]@{
        SourceItemId = $Row.SourceItemId
        Status       = $Row.Status
        ProcessedAt  = $Row.ProcessedAt
        TargetItemId = $Row.TargetItemId
        Error        = $Row.Error
    } | ConvertTo-Csv -NoTypeInformation
    if (-not $script:JournalHasHeader) { $script:Journal.WriteLine($csv[0]); $script:JournalHasHeader = $true }
    $script:Journal.WriteLine($csv[1])

    $script:Stats[$Status]++
    if ($Status -eq 'Copied') {
        Write-Host ('[{0}] Copied #{1}: {2} | {3} ({4:N0} KB)' -f (Get-Date).ToString('HH:mm:ss'), $script:Stats.Copied,
            $Row.SourceFolder, $Row.Subject, ([long]$Row.SizeBytes / 1KB)) -ForegroundColor Green
    }
    if ($Status -eq 'Failed') { Write-Warning ('Failed [{0}] {1}: {2}' -f $Row.SourceFolder, $Row.Subject, $Row.Error) }

    if (((Get-Date) - $script:LastCheckpoint).TotalSeconds -ge $CheckpointSeconds) { Save-Inventory }
}
#endregion

#region Phase 1 - Inventory
function Update-Inventory {
    Write-Host 'Scanning source mailbox...'
    $added      = 0
    $scanErrors = [Collections.Generic.List[string]]::new()

    foreach ($folder in (Get-SourceFolders)) {
        if ($folder.Count -eq 0) { continue }
        $targetPath = Get-TargetPath $folder.Path
        $n = 0
        try {
            foreach ($it in (Get-FolderItems $SourceMailbox $folder.Id)) {
                if ($script:Index.ContainsKey($it.Id)) { continue }
                $row = [pscustomobject]@{
                    Status       = 'Pending'
                    SourceFolder = $folder.Path
                    TargetFolder = $targetPath
                    Subject      = ([string]$it.Subject) -replace '[\r\n]+', ' '
                    Received     = $it.Received
                    ItemClass    = $it.Class
                    SizeBytes    = $it.Size
                    ProcessedAt  = ''
                    Error        = ''
                    TargetItemId = ''
                    SourceItemId = $it.Id
                    MigrationKey = $it.Key
                }
                $script:Rows.Add($row); $script:Index[$it.Id] = $row; $n++
            }
        }
        catch {
            $scanErrors.Add("$($folder.Path): $($_.Exception.Message)")
            Write-Warning "Could not list $($folder.Path): $($_.Exception.Message)"
        }
        $added += $n
        Write-Host ('  {0}: {1} item(s), {2} new' -f $folder.Path, $folder.Count, $n)
    }

    Save-Inventory

    $totalBytes = 0
    $folderSet  = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($r in $script:Rows) { $totalBytes += [long]$r.SizeBytes; [void]$folderSet.Add($r.SourceFolder) }
    $byStatus = ($script:Rows | Group-Object Status | ForEach-Object { '{0}: {1:N0}' -f $_.Name, $_.Count }) -join '   '

    Write-Host ('Inventory: {0:N0} item(s), {1:N1} MB, {2} folder(s). {3:N0} added this scan.' -f `
        $script:Rows.Count, ($totalBytes / 1MB), $folderSet.Count, $added) -ForegroundColor Green
    Write-Host "  $byStatus"
    if ($scanErrors.Count) {
        Write-Warning "$($scanErrors.Count) folder(s) could not be scanned:"
        $scanErrors | ForEach-Object { Write-Warning "  $_" }
    }
}
#endregion

#region Phase 2 - Transfer
function Copy-Batch([object[]]$Batch, [string]$TargetFolderId) {
    $queue = $Batch
    for ($attempt = 1; $queue.Count -gt 0; $attempt++) {
        $retry    = [Collections.Generic.List[object]]::new()
        $exported = [Collections.Generic.List[object]]::new()
        $backoff  = 0
        $canRetry = $attempt -lt $MaxRetries

        # --- Export from source ---
        try {
            $ids  = -join ($queue | ForEach-Object { '<t:ItemId Id="{0}"/>' -f $_.SourceItemId })
            $xml  = Invoke-Ews -AnchorMailbox $SourceMailbox -PerItemThrottle -Body "<m:ExportItems><m:ItemIds>$ids</m:ItemIds></m:ExportItems>"
            $msgs = @($xml.SelectNodes("//*[local-name()='ExportItemsResponseMessage']"))
            if ($msgs.Count -ne $queue.Count) { throw "ExportItems returned $($msgs.Count) results for $($queue.Count) items." }
        }
        catch {
            $err = $_.Exception.Message
            foreach ($r in $queue) { Set-ItemStatus $r 'Failed' '' "Export: $err" }
            return
        }

        for ($i = 0; $i -lt $queue.Count; $i++) {
            $r = $queue[$i]; $m = $msgs[$i]
            $rc = Get-Text $m 'ResponseCode'
            if ($m.GetAttribute('ResponseClass') -eq 'Success') {
                $exported.Add([pscustomobject]@{ Row = $r; Data = Get-Text $m 'Data' })
            }
            elseif ((Test-Throttled $m) -and $canRetry) {
                $retry.Add($r); $backoff = [math]::Max($backoff, (Get-BackOffMs $m))
            }
            else {
                Set-ItemStatus $r 'Failed' '' "Export: $rc $(Get-Text $m 'MessageText')"
            }
        }

        # --- Upload to target ---
        if ($exported.Count) {
            try {
                $parts = foreach ($e in $exported) {
                    '<t:Item CreateAction="CreateNew"><t:ParentFolderId Id="{0}"/><t:Data>{1}</t:Data></t:Item>' -f $TargetFolderId, $e.Data
                }
                $xml  = Invoke-Ews -AnchorMailbox $TargetMailbox -NoNetworkRetry -PerItemThrottle -Body "<m:UploadItems><m:Items>$(-join $parts)</m:Items></m:UploadItems>"
                $msgs = @($xml.SelectNodes("//*[local-name()='UploadItemsResponseMessage']"))
                if ($msgs.Count -ne $exported.Count) { throw "UploadItems returned $($msgs.Count) results for $($exported.Count) items." }
            }
            catch {
                $err = $_.Exception.Message
                foreach ($e in $exported) { Set-ItemStatus $e.Row 'Failed' '' "Upload: $err" }
                $msgs = @()
            }

            $uploaded = [Collections.Generic.List[object]]::new()
            for ($i = 0; $i -lt $msgs.Count; $i++) {
                $r = $exported[$i].Row; $m = $msgs[$i]
                $rc = Get-Text $m 'ResponseCode'
                if ($m.GetAttribute('ResponseClass') -eq 'Success') {
                    $uploaded.Add([pscustomobject]@{ Row = $r; NewId = (Get-Node $m 'ItemId').GetAttribute('Id') })
                }
                elseif ((Test-Throttled $m) -and $canRetry) {
                    $retry.Add($r); $backoff = [math]::Max($backoff, (Get-BackOffMs $m))
                }
                else {
                    Set-ItemStatus $r 'Failed' '' "Upload: $rc $(Get-Text $m 'MessageText')"
                }
            }

            # --- Tag copies so reruns can recognize them ---
            if ($uploaded.Count) {
                $tagMsgs = @(); $tagErr = ''
                try { $tagMsgs = @(Set-TargetTag $uploaded.ToArray()) } catch { $tagErr = $_.Exception.Message }
                for ($i = 0; $i -lt $uploaded.Count; $i++) {
                    $u = $uploaded[$i]; $note = ''
                    if ($tagErr) { $note = "Copied, but duplicate-check tag not set: $tagErr" }
                    elseif ($i -ge $tagMsgs.Count -or $tagMsgs[$i].GetAttribute('ResponseClass') -ne 'Success') {
                        $note = "Copied, but duplicate-check tag not set: $(if ($i -lt $tagMsgs.Count) { Get-Text $tagMsgs[$i] 'ResponseCode' })"
                    }
                    Set-ItemStatus $u.Row 'Copied' $u.NewId $note
                }
            }
        }

        if ($retry.Count) {
            Register-Throttle
            Write-Advisory "$($retry.Count) item(s) throttled by server, attempt $attempt of $MaxRetries" ($backoff / 1000)
            Start-Sleep -Milliseconds $backoff
        }
        $queue = $retry.ToArray()
    }
}

function Start-Transfer {
    $statuses = if ($RetryFailed) { @('Pending', 'Failed') } else { @('Pending') }
    $todo = @($script:Rows | Where-Object { $_.Status -in $statuses })
    Write-Host "$($todo.Count) item(s) to transfer."
    if (-not $todo.Count) { return }

    Write-Host 'Preparing target folders...'
    $script:TargetRootId = Get-RootFolderId $TargetMailbox $TargetRoot
    $script:TargetMap    = @{}
    foreach ($f in (Get-FolderTree $TargetMailbox $TargetRoot)) { $script:TargetMap[$f.Path] = $f.Id }
    foreach ($f in (Get-SourceFolders)) {   # also creates empty folders, with the correct folder type
        try { [void](Resolve-TargetFolder (Get-TargetPath $f.Path) $f.Class) }
        catch { Write-Warning "Could not create target folder for $($f.Path): $($_.Exception.Message)" }
    }

    foreach ($group in ($todo | Group-Object TargetFolder)) {
        try { $targetId = Resolve-TargetFolder $group.Name '' }
        catch {
            $err = $_.Exception.Message
            foreach ($r in $group.Group) { Set-ItemStatus $r 'Failed' '' "Target folder: $err" }
            continue
        }

        # Skip anything already in the target folder
        try { $existing = Get-TargetKeys $targetId }
        catch {
            $err = $_.Exception.Message
            foreach ($r in $group.Group) { Set-ItemStatus $r 'Failed' '' "Target check: $err" }
            continue
        }
        $toCopy = foreach ($r in $group.Group) {
            if (-not $existing.Add($r.MigrationKey)) { Set-ItemStatus $r 'Exists' '' 'Already in target; not copied again' }
            else { $r }
        }
        $toCopy = @($toCopy)

        Write-Host ('{0}: {1} item(s), {2} to copy' -f $group.Name, $group.Count, $toCopy.Count)
        $batch = [Collections.Generic.List[object]]::new(); $size = 0
        foreach ($r in $toCopy) {
            $s = [long]$r.SizeBytes
            if ($batch.Count -and ($batch.Count -ge $script:Pace.BatchSize -or ($size + $s) -gt $script:Pace.BatchBytes)) {
                Copy-Batch $batch.ToArray() $targetId
                $batch.Clear(); $size = 0
            }
            $batch.Add($r); $size += $s
        }
        if ($batch.Count) { Copy-Batch $batch.ToArray() $targetId }
    }
}
#endregion

#region Main
$script:Stats          = @{ Copied = 0; Exists = 0; Failed = 0; Backoffs = 0; BackoffSeconds = 0 }
$script:Rows           = [Collections.Generic.List[object]]::new()
$script:Index          = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::Ordinal)
$script:LastCheckpoint = Get-Date
$script:Pace           = [pscustomobject]@{ DelayMs = 0; BatchSize = $BatchSize; BatchBytes = $MaxBatchBytes; CleanStreak = 0 }

try {
    if (Import-Inventory) { Save-Inventory } else { Open-Journal }

    Connect-Ews

    if ($InventoryOnly -or $script:Rows.Count -eq 0) { Update-Inventory }

    if ($InventoryOnly) {
        Write-Host "Review $InventoryPath, then rerun without -InventoryOnly to transfer."
    }
    else {
        Start-Transfer
        Write-Host ('Transfer finished. This run: {0} copied, {1} already in target, {2} failed, {3} backoff(s) totaling {4:N0} s.' -f `
            $script:Stats.Copied, $script:Stats.Exists, $script:Stats.Failed, $script:Stats.Backoffs, $script:Stats.BackoffSeconds) -ForegroundColor Green
        if ($script:Stats.Failed) { Write-Host 'Filter the inventory on Status = Failed for details. Use -RetryFailed to try them again.' }
    }
}
finally {
    if ($script:Rows.Count) { Save-Inventory }
    if ($script:Journal) { $script:Journal.Dispose() }
    if ((Test-Path $JournalPath) -and (Get-Item $JournalPath).Length -eq 0) { Remove-Item -LiteralPath $JournalPath }
}
#endregion