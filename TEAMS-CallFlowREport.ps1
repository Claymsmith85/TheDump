<#
.SYNOPSIS
    Traces inbound Teams calls through auto attendants, call queues, and agents,
    and uploads per-hop and per-call CSV reports to SharePoint.

.DESCRIPTION
    Runs in Azure Automation or from a local PowerShell 7 session. Automation
    uses the system-assigned identity federated to the environment application
    registration. Local execution retrieves that application's certificate from
    Delinea Secret Server through the Clayton_M365_Scripts helpers. Microsoft
    Graph is called directly with an app-only bearer token; MicrosoftTeams is not
    loaded.

    For every audio call record that started in the reporting window, the
    runbook reads the record's sessions and segments and rebuilds the hop-by-hop
    path, for example:

        PSTN caller -> [Auto Attendant] Main AA -> [Call Queue] Claims CQ
                    -> [Agent/User] Jane Doe (RANG - no answer)
                    -> [Agent/User] John Smith (ANSWERED)

    Three CSV files are built in temporary storage and uploaded to the
    SharePoint document-library folder named by SharePointFolderUrl:
      - CallFlow-Summary_<window>.csv: one row per call with the entry point,
        auto attendants and queues traversed, outcome (Answered, Voicemail,
        AbandonedInQueue, AbandonedInAA, Unanswered), who answered, time to
        answer, and queue wait.
      - CallFlow-Hops_<window>.csv: one row per hop of each call.
      - CallFlow-Queues_<window>.csv: one row per call queue with offered,
        transferred-out, handled, answered, voicemail, and abandoned counts,
        answer rate, and average and maximum answer wait.
    Re-running the same window replaces the files and relies on SharePoint
    version history. All times in the CSV files are wall-clock times in
    TimeZoneId; durations and waits are in seconds.

    When StartDateTime and EndDateTime are omitted, the window is the previous
    full calendar day in TimeZoneId.

    Known limits of the callRecords API:
      - Records appear 30-60 minutes after a call ends and can keep updating for
        several hours. Run a daily report a few hours after midnight so the
        previous day is complete.
      - Records are retained for 30 days.
      - DTMF menu choices inside an auto attendant are not exposed. The menu path
        is inferred from where the call moved next, not from the digit pressed.
      - A transfer by a human agent can start a second call record, which is
        reported as a separate call.
      - A caller hanging up and a queue or auto attendant timeout that
        disconnects look the same; both are reported as abandoned.
      - The number the caller dialed is usually not in the call record.
        DialedNumber is filled only when it is; EntryPoint names the first
        auto attendant, queue, or person the call reached.

    The environment application registration requires Microsoft Graph
    CallRecords.Read.All and Sites.ReadWrite.All application permissions, or the
    accepted broader alternatives checked by the runbook.
#>

param(
    # Destination document library or folder, pasted from the browser address
    # bar. Library view URLs (.../Forms/AllItems.aspx?id=...) are accepted. The
    # folder must already exist.
    [string]$SharePointFolderUrl = 'PASTE_SHAREPOINT_LIBRARY_URL_HERE',

    # Wall-clock times in TimeZoneId without a UTC offset, such as
    # '2026-08-20 00:00'. Omit both for the previous full calendar day.
    [string]$StartDateTime = '',

    [string]$EndDateTime = '',

    [string]$TimeZoneId = 'America/New_York',

    [string]$CallerNumberFilter = '',

    [bool]$OnlyVoiceAppCalls = $true,

    [bool]$WriteCallFlowTree = $false,

    [switch]$DryRun,

    [ValidateSet('Auto', 'Automation', 'Delinea')]
    [string]$AuthenticationMode = 'Auto',

    [string]$LocalSupportRoot = $env:M365_SCRIPTS_ROOT
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Terraform replaces this marker with the shared authentication source. It
# remains a comment locally, where the same source is dot-sourced below.
# __SHARED_FUNCTIONS__

$sharePointFolderUrlPlaceholder = 'PASTE_SHAREPOINT_LIBRARY_URL_HERE'
$script:GraphAccessToken = ''
$script:GraphTokenExpiresUtc = [datetime]::MinValue
$script:ConnectionSettings = $null
$script:ReportTimeZone = $null
$script:TemporaryRoot = ''
$script:MaxGraphAttempts = 6
$script:RetryableStatusCodes = @(429, 502, 503, 504)
$script:VoiceAppKinds = @('AutoAttendant', 'CallQueue', 'VoiceApp', 'Voicemail')

# First-party bot application IDs seen on voice-app legs that carry no
# service role.
$script:BotApplicationKinds = @{
    'ce933385-9390-45d1-9512-c8d228074e07' = 'AutoAttendant'
    '11cd3e2e-fccb-42ad-ad00-878b93575e07' = 'CallQueue'
}

$script:CsvColumns = @{
    Hops = @('CallStart', 'Caller', 'Hop', 'LegStart', 'From', 'FromKind', 'To', 'ToKind', 'LegDuration', 'HopOutcome', 'Failure', 'CallOutcome', 'AnsweredBy', 'CallRecordId')
    Summary = @('CallStart', 'Caller', 'DialedNumber', 'EntryPoint', 'EntryPointKind', 'AutoAttendants', 'CallQueues', 'FinalQueue', 'Outcome', 'AnsweredBy', 'TimeToAnswerSeconds', 'QueueWaitSeconds', 'Duration', 'CallRecordId')
    Queues = @('CallQueue', 'Offered', 'TransferredOut', 'Handled', 'Answered', 'Voicemail', 'Abandoned', 'AnswerRatePercent', 'AvgAnswerWaitSeconds', 'MaxAnswerWaitSeconds')
}

function Write-Status {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Level = 'INFO'
    )

    $prefix = "[$Level]"
    switch ($Level) {
        'WARN' { Write-Host "$prefix $Message" -ForegroundColor Yellow }
        'ERROR' { Write-Host "$prefix $Message" -ForegroundColor Red }
        default { Write-Host "$prefix $Message" -ForegroundColor Cyan }
    }
}

function Get-GraphField {
    param(
        $Object,

        [Parameter(Mandatory)]
        [string]$Name
    )

    if ($null -eq $Object) {
        return $null
    }
    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name)) {
            return $Object[$Name]
        }
        return $null
    }

    $property = $Object.PSObject.Properties[$Name]
    if ($null -ne $property) {
        return $property.Value
    }
    return $null
}

function Get-JwtPayload {
    param(
        [Parameter(Mandatory)]
        [string]$AccessToken
    )

    $parts = $AccessToken.Split('.')
    if ($parts.Count -lt 2) {
        throw 'The access token is not a valid JWT.'
    }

    try {
        $payload = $parts[1].Replace('-', '+').Replace('_', '/')
        switch ($payload.Length % 4) {
            2 { $payload += '==' }
            3 { $payload += '=' }
        }
        return [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($payload)) | ConvertFrom-Json
    } catch {
        throw "The access token claims could not be decoded: $($_.Exception.Message)"
    }
}

function Assert-ApplicationRole {
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [string[]]$ActualRoles,

        [Parameter(Mandatory)]
        [string[]]$RequiredRoles,

        [Parameter(Mandatory)]
        [string]$Purpose
    )

    if (@($RequiredRoles | Where-Object { $ActualRoles -contains $_ }).Count -eq 0) {
        throw "Microsoft Graph permission preflight failed for $Purpose. The environment application requires one of: $($RequiredRoles -join ', ')."
    }
}

function Connect-GraphAccess {
    # Called at startup and again when the token nears expiry: detail requests
    # for a busy day of calls can outlive a single Graph token.
    $settings = $script:ConnectionSettings
    if ($settings.Mode -eq 'Automation') {
        $connection = Connect-M365ServicesWithManagedIdentity -TenantId $settings.TenantId -ClientId $settings.ClientId -SkipTeams
    } else {
        $connection = Connect-M365ServicesWithDelinea -SupportRoot $settings.SupportRoot -SkipTeams
    }

    $accessToken = [string]$connection.GraphAccessToken
    if ([string]::IsNullOrWhiteSpace($accessToken)) {
        throw 'The shared connection profile did not return a Microsoft Graph access token.'
    }
    $expiry = Get-GraphField -Object (Get-JwtPayload -AccessToken $accessToken) -Name 'exp'
    if ($null -eq $expiry) {
        throw 'The Microsoft Graph access token has no expiry claim.'
    }

    $script:GraphAccessToken = $accessToken
    $script:GraphTokenExpiresUtc = [DateTimeOffset]::FromUnixTimeSeconds([int64]$expiry).UtcDateTime
}

function Get-GraphAuthorizationHeader {
    if ([string]::IsNullOrWhiteSpace($script:GraphAccessToken)) {
        throw 'The Microsoft Graph access token has not been initialized.'
    }
    if ([datetime]::UtcNow -ge $script:GraphTokenExpiresUtc.AddMinutes(-5)) {
        Write-Status 'Refreshing the Microsoft Graph access token.'
        Connect-GraphAccess
    }

    return @{ Authorization = "Bearer $script:GraphAccessToken" }
}

function Get-HttpStatusCode {
    param(
        [Parameter(Mandatory)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    $exception = $ErrorRecord.Exception
    if ($exception -is [Microsoft.PowerShell.Commands.HttpResponseException] -and $null -ne $exception.Response) {
        return [int]$exception.Response.StatusCode
    }
    return 0
}

function Invoke-GraphRequest {
    param(
        [Parameter(Mandatory)]
        [string]$Uri,

        [ValidateSet('Get', 'Put')]
        [string]$Method = 'Get',

        [string]$InFile = '',

        [string]$ContentType = ''
    )

    $attempt = 0
    while ($true) {
        $attempt++
        $parameters = @{
            Method = $Method
            Uri = $Uri
            Headers = (Get-GraphAuthorizationHeader)
            ErrorAction = 'Stop'
        }
        if (-not [string]::IsNullOrWhiteSpace($InFile)) {
            $parameters.InFile = $InFile
            $parameters.ContentType = $ContentType
        }

        try {
            return Invoke-RestMethod @parameters
        } catch {
            $statusCode = Get-HttpStatusCode -ErrorRecord $_
            if ($script:RetryableStatusCodes -notcontains $statusCode -or $attempt -ge $script:MaxGraphAttempts) {
                throw
            }

            $waitSeconds = 5 * $attempt
            $retryAfter = $_.Exception.Response.Headers.RetryAfter
            if ($null -ne $retryAfter -and $null -ne $retryAfter.Delta) {
                $waitSeconds = [Math]::Max(1, [int][Math]::Ceiling($retryAfter.Delta.TotalSeconds))
            }
            Write-Status "Microsoft Graph returned HTTP $statusCode; retrying in $waitSeconds second(s) (attempt $attempt of $script:MaxGraphAttempts)." -Level WARN
            Start-Sleep -Seconds $waitSeconds
        }
    }
}

function Get-GraphCollection {
    param(
        [Parameter(Mandatory)]
        [string]$Uri
    )

    $items = New-Object System.Collections.Generic.List[object]
    $nextLink = $Uri
    while (-not [string]::IsNullOrWhiteSpace($nextLink)) {
        $response = Invoke-GraphRequest -Uri $nextLink
        foreach ($item in @((Get-GraphField -Object $response -Name 'value'))) {
            if ($null -ne $item) {
                $items.Add($item)
            }
        }
        $nextLink = [string](Get-GraphField -Object $response -Name '@odata.nextLink')
    }

    return [object[]]$items.ToArray()
}

function ConvertTo-UtcDateTime {
    # Invoke-RestMethod may return Graph timestamps as local-kind DateTime values
    # or as strings, depending on the PowerShell version.
    param($Value)

    if ($null -eq $Value) {
        return $null
    }
    if ($Value -is [datetime]) {
        if ($Value.Kind -eq [System.DateTimeKind]::Unspecified) {
            return [datetime]::SpecifyKind($Value, [System.DateTimeKind]::Utc)
        }
        return $Value.ToUniversalTime()
    }
    if ($Value -is [DateTimeOffset]) {
        return $Value.UtcDateTime
    }
    if ([string]::IsNullOrWhiteSpace([string]$Value)) {
        return $null
    }
    return [DateTimeOffset]::Parse([string]$Value, [System.Globalization.CultureInfo]::InvariantCulture).UtcDateTime
}

function Format-ReportTime {
    param(
        [Parameter(Mandatory)]
        [datetime]$UtcValue
    )

    return [System.TimeZoneInfo]::ConvertTimeFromUtc($UtcValue, $script:ReportTimeZone).ToString('yyyy-MM-dd HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture)
}

function Format-CallDuration {
    param(
        [Parameter(Mandatory)]
        [datetime]$From,

        [Parameter(Mandatory)]
        [datetime]$To
    )

    if ($To -le $From) {
        return '00:00'
    }
    $span = $To - $From
    if ($span.TotalHours -ge 1) {
        return $span.ToString('hh\:mm\:ss')
    }
    return $span.ToString('mm\:ss')
}

function ConvertFrom-WindowText {
    param(
        [Parameter(Mandatory)]
        [string]$Value,

        [Parameter(Mandatory)]
        [string]$Name
    )

    $parsed = [datetime]::MinValue
    if (-not [datetime]::TryParse($Value, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) {
        throw "$Name '$Value' is not a valid date and time. Use a value such as '2026-08-20 00:00'."
    }
    if ($parsed.Kind -ne [System.DateTimeKind]::Unspecified) {
        throw "$Name '$Value' must not include a UTC offset; it is interpreted in TimeZoneId '$TimeZoneId'."
    }
    return $parsed
}

function Resolve-ReportWindow {
    param(
        [AllowEmptyString()]
        [string]$Start,

        [AllowEmptyString()]
        [string]$End
    )

    $hasStart = -not [string]::IsNullOrWhiteSpace($Start)
    $hasEnd = -not [string]::IsNullOrWhiteSpace($End)
    if ($hasStart -ne $hasEnd) {
        throw 'Supply both StartDateTime and EndDateTime, or neither for the previous full day.'
    }

    if ($hasStart) {
        $startLocal = ConvertFrom-WindowText -Value $Start -Name 'StartDateTime'
        $endLocal = ConvertFrom-WindowText -Value $End -Name 'EndDateTime'
    } else {
        $todayLocal = [System.TimeZoneInfo]::ConvertTimeFromUtc([datetime]::UtcNow, $script:ReportTimeZone).Date
        $startLocal = $todayLocal.AddDays(-1)
        $endLocal = $todayLocal
    }

    $startUtc = [System.TimeZoneInfo]::ConvertTimeToUtc([datetime]::SpecifyKind($startLocal, [System.DateTimeKind]::Unspecified), $script:ReportTimeZone)
    $endUtc = [System.TimeZoneInfo]::ConvertTimeToUtc([datetime]::SpecifyKind($endLocal, [System.DateTimeKind]::Unspecified), $script:ReportTimeZone)
    if ($endUtc -le $startUtc) {
        throw "EndDateTime must be later than StartDateTime."
    }

    return [pscustomobject]@{
        StartLocal = $startLocal
        EndLocal = $endLocal
        StartUtc = $startUtc
        EndUtc = $endUtc
    }
}

function Resolve-SharePointFolderDestination {
    param(
        [Parameter(Mandatory)]
        [string]$FolderUrl
    )

    $parsedUrl = $null
    if (-not [uri]::TryCreate($FolderUrl.Trim(), [System.UriKind]::Absolute, [ref]$parsedUrl)) {
        throw "SharePointFolderUrl '$FolderUrl' is not a valid absolute URL."
    }
    $sharePointHost = $parsedUrl.Host.ToLowerInvariant()
    if ($parsedUrl.Scheme -ne 'https' -or $sharePointHost -notmatch '^[a-z0-9-]+\.sharepoint\.com$' -or $sharePointHost -match '-(my|admin)\.sharepoint\.com$') {
        throw "SharePointFolderUrl '$FolderUrl' must be an HTTPS SharePoint site URL, not a OneDrive or admin center URL."
    }

    # Browser "copy link" URLs prefix the path with /:f:/r; library view URLs
    # carry the open folder in the id query value.
    $path = ([uri]::UnescapeDataString($parsedUrl.AbsolutePath) -replace '^/:[a-z]:/r(?=/)', '').TrimEnd('/')
    if ($path -match '/Forms/[^/]+\.aspx$') {
        $viewFolder = [string]([System.Web.HttpUtility]::ParseQueryString($parsedUrl.Query)['id'])
        $path = if (-not [string]::IsNullOrWhiteSpace($viewFolder)) {
            $viewFolder.TrimEnd('/')
        } else {
            $path -replace '/Forms/[^/]+\.aspx$', ''
        }
    }

    $segments = @($path.Trim('/') -split '/')
    if ($segments.Count -lt 3 -or @('sites', 'teams') -notcontains $segments[0]) {
        throw "SharePointFolderUrl '$FolderUrl' must point to a document library or folder under /sites/<site>/ or /teams/<site>/."
    }

    $sitePath = '/{0}/{1}' -f $segments[0], [uri]::EscapeDataString($segments[1])
    $site = Invoke-GraphRequest -Uri ('https://graph.microsoft.com/v1.0/sites/{0}:{1}?$select=id,webUrl' -f $sharePointHost, $sitePath)
    $siteId = [string](Get-GraphField -Object $site -Name 'id')
    if ([string]::IsNullOrWhiteSpace($siteId)) {
        throw "Microsoft Graph did not return an ID for SharePoint site 'https://$sharePointHost$sitePath'."
    }

    $drives = @(Get-GraphCollection -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/drives?`$select=id,name,webUrl,driveType")
    $matchedDrive = $null
    $matchedDrivePath = ''
    foreach ($drive in $drives) {
        $driveWebUrl = [string](Get-GraphField -Object $drive -Name 'webUrl')
        if ([string](Get-GraphField -Object $drive -Name 'driveType') -ne 'documentLibrary' -or [string]::IsNullOrWhiteSpace($driveWebUrl)) {
            continue
        }

        $drivePath = [uri]::UnescapeDataString(([uri]$driveWebUrl).AbsolutePath).TrimEnd('/')
        $isMatch = $path.Equals($drivePath, [System.StringComparison]::OrdinalIgnoreCase) -or
            $path.StartsWith("$drivePath/", [System.StringComparison]::OrdinalIgnoreCase)
        if ($isMatch -and $drivePath.Length -gt $matchedDrivePath.Length) {
            $matchedDrive = $drive
            $matchedDrivePath = $drivePath
        }
    }
    if ($null -eq $matchedDrive) {
        throw "No document library in '$([string](Get-GraphField -Object $site -Name 'webUrl'))' contains '$path'. Libraries in subsites are not supported."
    }

    $driveId = [string](Get-GraphField -Object $matchedDrive -Name 'id')
    $relativeFolder = $path.Substring($matchedDrivePath.Length).Trim('/')
    $folderRequestUri = if ([string]::IsNullOrWhiteSpace($relativeFolder)) {
        "https://graph.microsoft.com/v1.0/drives/$driveId/root?`$select=id,webUrl,folder"
    } else {
        $encodedFolder = (@($relativeFolder -split '/') | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/'
        "https://graph.microsoft.com/v1.0/drives/$driveId/root:/$encodedFolder`?`$select=id,webUrl,folder"
    }

    try {
        $folderItem = Invoke-GraphRequest -Uri $folderRequestUri
    } catch {
        if ((Get-HttpStatusCode -ErrorRecord $_) -eq 404) {
            throw "Folder '$relativeFolder' does not exist in library '$([string](Get-GraphField -Object $matchedDrive -Name 'name'))'. Create it before running this runbook."
        }
        throw
    }
    if ($null -eq (Get-GraphField -Object $folderItem -Name 'folder')) {
        throw "SharePointFolderUrl '$FolderUrl' resolves to a file, not a folder."
    }

    return [pscustomobject]@{
        DriveId = $driveId
        FolderItemId = [string](Get-GraphField -Object $folderItem -Name 'id')
        FolderWebUrl = [string](Get-GraphField -Object $folderItem -Name 'webUrl')
    }
}

function Resolve-CallEndpoint {
    # Returns { Kind; Label; Id } for a callRecords caller or callee endpoint.
    # Kind: Phone | AutoAttendant | CallQueue | Voicemail | VoiceApp | User |
    # Guest | App | Service | Unknown
    param($Endpoint)

    if ($null -eq $Endpoint) {
        return [pscustomobject]@{ Kind = 'Unknown'; Label = '(none)'; Id = '' }
    }

    $identity = Get-GraphField -Object $Endpoint -Name 'identity'
    $kind = 'Unknown'
    $label = ''
    $endpointId = ''

    # The service user-agent role is the most reliable auto attendant and call
    # queue signal.
    $role = [string](Get-GraphField -Object (Get-GraphField -Object $Endpoint -Name 'userAgent') -Name 'role')

    $applicationInstance = Get-GraphField -Object $identity -Name 'applicationInstance'
    $phone = Get-GraphField -Object $identity -Name 'phone'
    $user = Get-GraphField -Object $identity -Name 'user'
    $guest = Get-GraphField -Object $identity -Name 'guest'
    $application = Get-GraphField -Object $identity -Name 'application'
    if ($null -ne $applicationInstance) {
        $kind = 'Service'
        $label = [string](Get-GraphField -Object $applicationInstance -Name 'displayName')
        $endpointId = [string](Get-GraphField -Object $applicationInstance -Name 'id')
    } elseif ($null -ne $phone) {
        $kind = 'Phone'
        $endpointId = [string](Get-GraphField -Object $phone -Name 'id')
        $label = $endpointId
        if ([string]::IsNullOrWhiteSpace($label)) {
            $label = [string](Get-GraphField -Object $phone -Name 'displayName')
        }
    } elseif ($null -ne $user) {
        $kind = 'User'
        $label = [string](Get-GraphField -Object $user -Name 'displayName')
        $endpointId = [string](Get-GraphField -Object $user -Name 'id')
    } elseif ($null -ne $guest) {
        $kind = 'Guest'
        $label = [string](Get-GraphField -Object $guest -Name 'displayName')
        $endpointId = [string](Get-GraphField -Object $guest -Name 'id')
    } elseif ($null -ne $application) {
        $kind = 'App'
        $label = [string](Get-GraphField -Object $application -Name 'displayName')
        $endpointId = [string](Get-GraphField -Object $application -Name 'id')
    }

    if ($role -match 'AutoAttendant') {
        $kind = 'AutoAttendant'
    } elseif ($role -match 'CallQueue') {
        $kind = 'CallQueue'
    } elseif ($role -match 'Voicemail|UnifiedMessaging') {
        $kind = 'Voicemail'
    } elseif (-not [string]::IsNullOrWhiteSpace($endpointId) -and $script:BotApplicationKinds.ContainsKey($endpointId)) {
        $kind = $script:BotApplicationKinds[$endpointId]
    }

    if (($kind -eq 'Service' -or $kind -eq 'App') -and -not [string]::IsNullOrWhiteSpace($label)) {
        # A named resource account without a recognizable role.
        $kind = 'VoiceApp'
    }

    if ([string]::IsNullOrWhiteSpace($label)) {
        $label = if (-not [string]::IsNullOrWhiteSpace($endpointId)) {
            $endpointId
        } else {
            '(unidentified {0})' -f ([string](Get-GraphField -Object $Endpoint -Name '@odata.type') -replace '#microsoft.graph.callRecords.', '')
        }
    }

    return [pscustomobject]@{ Kind = $kind; Label = $label; Id = $endpointId }
}

function Get-SessionFailure {
    # Returns the session failure reason, then its last segment's, or '' when
    # the leg completed normally.
    param($Session)

    $reason = [string](Get-GraphField -Object (Get-GraphField -Object $Session -Name 'failureInfo') -Name 'reason')
    if (-not [string]::IsNullOrWhiteSpace($reason)) {
        return $reason
    }

    $segments = @((Get-GraphField -Object $Session -Name 'segments') | Where-Object { $null -ne $_ })
    if ($segments.Count -eq 0) {
        return ''
    }
    $lastSegment = @($segments | Sort-Object { ConvertTo-UtcDateTime (Get-GraphField -Object $_ -Name 'startDateTime') })[-1]
    return [string](Get-GraphField -Object (Get-GraphField -Object $lastSegment -Name 'failureInfo') -Name 'reason')
}

function Get-CallFlow {
    # Returns the rebuilt call flow for one call record, or $null when the call
    # is excluded by OnlyVoiceAppCalls or CallerNumberFilter.
    param(
        [Parameter(Mandatory)]
        $CallRecord,

        [bool]$VoiceAppCallsOnly = $true,

        [AllowEmptyString()]
        [string]$CallerFilter = ''
    )

    $callRecordId = [string](Get-GraphField -Object $CallRecord -Name 'id')
    $detailUri = 'https://graph.microsoft.com/v1.0/communications/callRecords/{0}?$expand=sessions($expand=segments)' -f [uri]::EscapeDataString($callRecordId)
    $detail = Invoke-GraphRequest -Uri $detailUri

    $sessions = New-Object System.Collections.Generic.List[object]
    foreach ($session in @((Get-GraphField -Object $detail -Name 'sessions'))) {
        if ($null -ne $session) {
            $sessions.Add($session)
        }
    }
    $sessionsNextLink = [string](Get-GraphField -Object $detail -Name 'sessions@odata.nextLink')
    if (-not [string]::IsNullOrWhiteSpace($sessionsNextLink)) {
        foreach ($session in @(Get-GraphCollection -Uri $sessionsNextLink)) {
            $sessions.Add($session)
        }
    }
    if ($sessions.Count -eq 0) {
        return $null
    }

    $callStartUtc = ConvertTo-UtcDateTime (Get-GraphField -Object $CallRecord -Name 'startDateTime')
    $callEndUtc = ConvertTo-UtcDateTime (Get-GraphField -Object $CallRecord -Name 'endDateTime')
    if ($null -eq $callEndUtc) {
        $callEndUtc = $callStartUtc
    }

    $legs = @(
        foreach ($session in @($sessions | Sort-Object { ConvertTo-UtcDateTime (Get-GraphField -Object $_ -Name 'startDateTime') })) {
            $legStartUtc = ConvertTo-UtcDateTime (Get-GraphField -Object $session -Name 'startDateTime')
            if ($null -eq $legStartUtc) {
                $legStartUtc = $callStartUtc
            }
            $legEndUtc = ConvertTo-UtcDateTime (Get-GraphField -Object $session -Name 'endDateTime')
            if ($null -eq $legEndUtc) {
                $legEndUtc = $legStartUtc
            }

            [pscustomobject]@{
                StartUtc = $legStartUtc
                From = Resolve-CallEndpoint (Get-GraphField -Object $session -Name 'caller')
                To = Resolve-CallEndpoint (Get-GraphField -Object $session -Name 'callee')
                Failure = Get-SessionFailure $session
                Duration = Format-CallDuration -From $legStartUtc -To $legEndUtc
            }
        }
    )

    $touchedVoiceApp = @($legs | Where-Object { $script:VoiceAppKinds -contains $_.To.Kind -or $script:VoiceAppKinds -contains $_.From.Kind }).Count -gt 0
    if ($VoiceAppCallsOnly -and -not $touchedVoiceApp) {
        return $null
    }

    # The caller is the phone endpoint on the earliest leg; fall back to the
    # record organizer for calls that did not start from the PSTN.
    $callerLeg = @($legs | Where-Object { $_.From.Kind -eq 'Phone' }) | Select-Object -First 1
    $caller = if ($null -ne $callerLeg) {
        $callerLeg.From.Label
    } else {
        (Resolve-CallEndpoint @{ identity = (Get-GraphField -Object $CallRecord -Name 'organizer') }).Label
    }
    if (-not [string]::IsNullOrWhiteSpace($CallerFilter) -and $caller.IndexOf($CallerFilter, [System.StringComparison]::OrdinalIgnoreCase) -lt 0) {
        return $null
    }

    # The call was answered by the last user leg that completed without a
    # failure.
    $answeredLeg = @($legs | Where-Object { $_.To.Kind -eq 'User' -and -not $_.Failure }) | Select-Object -Last 1
    $voicemailLeg = @($legs | Where-Object { $_.To.Kind -eq 'Voicemail' -and -not $_.Failure }) | Select-Object -Last 1
    $queueLegs = @($legs | Where-Object { $_.To.Kind -eq 'CallQueue' })
    $autoAttendantLegs = @($legs | Where-Object { $_.To.Kind -eq 'AutoAttendant' })

    # callRecords cannot distinguish a caller hanging up from a queue or auto
    # attendant timeout that disconnects, so both are reported as abandoned.
    $answeredBy = ''
    if ($null -ne $answeredLeg) {
        $outcome = 'Answered'
        $answeredBy = $answeredLeg.To.Label
    } elseif ($null -ne $voicemailLeg) {
        $outcome = 'Voicemail'
        $answeredBy = 'Voicemail'
    } elseif ($queueLegs.Count -gt 0) {
        $outcome = 'AbandonedInQueue'
    } elseif ($autoAttendantLegs.Count -gt 0) {
        $outcome = 'AbandonedInAA'
    } else {
        $outcome = 'Unanswered'
    }

    # Answer times are measured to the start of the answering agent's leg,
    # which excludes that agent's own ring time. Queue wait runs from the
    # first queue the call reached until it was answered, reached voicemail,
    # or ended.
    $timeToAnswerSeconds = $null
    if ($null -ne $answeredLeg) {
        $timeToAnswerSeconds = [int][Math]::Max(0, [Math]::Round(($answeredLeg.StartUtc - $callStartUtc).TotalSeconds))
    }
    $queueWaitSeconds = $null
    if ($queueLegs.Count -gt 0) {
        $waitEndUtc = if ($null -ne $answeredLeg) {
            $answeredLeg.StartUtc
        } elseif ($null -ne $voicemailLeg) {
            $voicemailLeg.StartUtc
        } else {
            $callEndUtc
        }
        $queueWaitSeconds = [int][Math]::Max(0, [Math]::Round(($waitEndUtc - $queueLegs[0].StartUtc).TotalSeconds))
    }

    # The entry point is the first auto attendant, queue, or person the caller
    # reached. A dialed number is present only when the record carries the
    # called party as a phone identity.
    $entryLeg = @($legs | Where-Object { $_.To.Kind -ne 'Phone' }) | Select-Object -First 1
    $dialedLeg = @($legs | Where-Object { $_.From.Kind -eq 'Phone' -and $_.To.Kind -eq 'Phone' -and $_.To.Label -ne $caller }) | Select-Object -First 1

    return [pscustomobject]@{
        CallStartUtc = $callStartUtc
        Caller = $caller
        DialedNumber = if ($null -ne $dialedLeg) { $dialedLeg.To.Label } else { '' }
        EntryPoint = if ($null -ne $entryLeg) { $entryLeg.To.Label } else { '' }
        EntryPointKind = if ($null -ne $entryLeg) { $entryLeg.To.Kind } else { '' }
        AutoAttendants = (@($autoAttendantLegs | ForEach-Object { $_.To.Label } | Select-Object -Unique) -join '; ')
        CallQueues = (@($queueLegs | ForEach-Object { $_.To.Label } | Select-Object -Unique) -join '; ')
        QueueNames = @($queueLegs | ForEach-Object { $_.To.Label } | Select-Object -Unique)
        FinalQueue = if ($queueLegs.Count -gt 0) { $queueLegs[-1].To.Label } else { '' }
        Outcome = $outcome
        AnsweredBy = $answeredBy
        TimeToAnswerSeconds = $timeToAnswerSeconds
        QueueWaitSeconds = $queueWaitSeconds
        Duration = Format-CallDuration -From $callStartUtc -To $callEndUtc
        Legs = $legs
        CallRecordId = $callRecordId
    }
}

function Get-QueueRollupRows {
    # One row per call queue. Offered counts every call that reached the queue;
    # a call that moved on to another queue is TransferredOut here and is
    # handled by the last queue it reached.
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [array]$Calls
    )

    $queueNames = @($Calls | ForEach-Object { $_.QueueNames } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
    foreach ($queueName in $queueNames) {
        $offered = @($Calls | Where-Object { $_.QueueNames -contains $queueName })
        $handled = @($offered | Where-Object { $_.FinalQueue -eq $queueName })
        $answered = @($handled | Where-Object { $_.Outcome -eq 'Answered' })
        $answerWaits = @($answered | Where-Object { $null -ne $_.QueueWaitSeconds } | ForEach-Object { [int]$_.QueueWaitSeconds })
        $answerWaitStats = if ($answerWaits.Count -gt 0) { $answerWaits | Measure-Object -Average -Maximum } else { $null }

        [pscustomobject]@{
            CallQueue = $queueName
            Offered = $offered.Count
            TransferredOut = $offered.Count - $handled.Count
            Handled = $handled.Count
            Answered = $answered.Count
            Voicemail = @($handled | Where-Object { $_.Outcome -eq 'Voicemail' }).Count
            Abandoned = @($handled | Where-Object { $_.Outcome -eq 'AbandonedInQueue' }).Count
            AnswerRatePercent = if ($handled.Count -gt 0) { [Math]::Round([double]100 * $answered.Count / $handled.Count, 1) } else { $null }
            AvgAnswerWaitSeconds = if ($null -ne $answerWaitStats) { [int][Math]::Round($answerWaitStats.Average) } else { $null }
            MaxAnswerWaitSeconds = if ($null -ne $answerWaitStats) { [int]$answerWaitStats.Maximum } else { $null }
        }
    }
}

function Get-HopOutcome {
    param(
        [Parameter(Mandatory)]
        $Leg
    )

    if ($Leg.Failure) {
        if ($Leg.To.Kind -eq 'User') {
            return "RANG - no answer ($($Leg.Failure))"
        }
        return "ended: $($Leg.Failure)"
    }
    if ($Leg.To.Kind -eq 'User') {
        return "ANSWERED - connected $($Leg.Duration)"
    }
    return "connected $($Leg.Duration)"
}

function Write-CallFlowTree {
    param(
        [Parameter(Mandatory)]
        [int]$Number,

        [Parameter(Mandatory)]
        $Call
    )

    Write-Host ''
    Write-Host ('Call {0}:  {1}  at {2}  (total {3})' -f $Number, $Call.Caller, (Format-ReportTime -UtcValue $Call.CallStartUtc), $Call.Duration) -ForegroundColor Cyan
    foreach ($leg in $Call.Legs) {
        $toTag = switch ($leg.To.Kind) {
            'AutoAttendant' { '[Auto Attendant]' }
            'CallQueue' { '[Call Queue]' }
            'Voicemail' { '[Voicemail]' }
            'VoiceApp' { '[Voice App]' }
            'User' { '[Agent/User]' }
            'Phone' { '[Phone]' }
            default { "[$($leg.To.Kind)]" }
        }
        $color = if ($leg.Failure) { 'DarkYellow' } elseif ($leg.To.Kind -eq 'User') { 'Green' } else { 'Gray' }
        $legTime = (Format-ReportTime -UtcValue $leg.StartUtc).Substring(11)
        Write-Host ('  {0}  {1} -> {2} {3,-40} {4}' -f $legTime, $leg.From.Label, $toTag, $leg.To.Label, (Get-HopOutcome -Leg $leg)) -ForegroundColor $color
    }
    Write-Host ('  Outcome: {0}  Answered by: {1}' -f $Call.Outcome, $Call.AnsweredBy) -ForegroundColor White
}

function Write-CsvReport {
    param(
        [Parameter(Mandatory)]
        [string]$FileName,

        [Parameter(Mandatory)]
        [string[]]$Columns,

        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [array]$Rows
    )

    $path = Join-Path $script:TemporaryRoot $FileName
    # Excel detects UTF-8 in a CSV file only when it has a byte order mark.
    if ($Rows.Count -eq 0) {
        $emptyRow = [ordered]@{}
        foreach ($column in $Columns) {
            $emptyRow[$column] = ''
        }
        $header = ([pscustomobject]$emptyRow | ConvertTo-Csv -NoTypeInformation)[0]
        Set-Content -LiteralPath $path -Value $header -Encoding utf8BOM
    } else {
        $Rows | Select-Object -Property $Columns | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding utf8BOM
    }
    Write-Status "Wrote $($Rows.Count) rows to $FileName."
    return $path
}

function Publish-CsvReport {
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        $Destination
    )

    $fileName = Split-Path -Path $Path -Leaf
    $uploadUri = 'https://graph.microsoft.com/v1.0/drives/{0}/items/{1}:/{2}:/content' -f $Destination.DriveId, $Destination.FolderItemId, [uri]::EscapeDataString($fileName)

    Write-Status "Uploading $fileName to $($Destination.FolderWebUrl)."
    $driveItem = Invoke-GraphRequest -Uri $uploadUri -Method Put -InFile $Path -ContentType 'text/csv'

    $expectedSize = (Get-Item -LiteralPath $Path).Length
    $uploadedSize = [int64](Get-GraphField -Object $driveItem -Name 'size')
    $webUrl = [string](Get-GraphField -Object $driveItem -Name 'webUrl')
    if ($uploadedSize -ne $expectedSize -or [string]::IsNullOrWhiteSpace($webUrl)) {
        throw "Microsoft Graph did not confirm the upload of '$fileName' (expected $expectedSize bytes, received $uploadedSize)."
    }
    return $webUrl
}

try {
    $scriptDirectory = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        $PSScriptRoot
    } else {
        $psEditorVariable = Get-Variable -Name psEditor -ErrorAction SilentlyContinue
        if ($null -ne $psEditorVariable -and $null -ne $psEditorVariable.Value) {
            Split-Path $psEditorVariable.Value.GetEditorContext().CurrentFile.Path
        } else {
            $PWD.Path
        }
    }

    if ($null -eq (Get-Command Resolve-AuthenticationMode -ErrorAction SilentlyContinue)) {
        $functionPath = Join-Path (Join-Path (Split-Path $scriptDirectory -Parent) 'functions') 'M365Authentication.ps1'
        if (-not (Test-Path -LiteralPath $functionPath -PathType Leaf)) {
            throw "Shared authentication functions were not injected and '$functionPath' is unavailable for local dot-sourcing."
        }
        . $functionPath
    }

    if ([string]::IsNullOrWhiteSpace($SharePointFolderUrl) -or $SharePointFolderUrl.Trim() -eq $sharePointFolderUrlPlaceholder) {
        throw 'SharePointFolderUrl has not been set. Paste the destination document library or folder URL into the parameter default, or pass it when starting the runbook.'
    }

    try {
        $script:ReportTimeZone = [System.TimeZoneInfo]::FindSystemTimeZoneById($TimeZoneId)
    } catch {
        throw "TimeZoneId '$TimeZoneId' is not a recognized time zone. Use an IANA name such as 'America/New_York'."
    }

    $window = Resolve-ReportWindow -Start $StartDateTime -End $EndDateTime
    $windowText = '{0:yyyy-MM-dd HH:mm} to {1:yyyy-MM-dd HH:mm} ({2})' -f $window.StartLocal, $window.EndLocal, $TimeZoneId
    if ($window.StartUtc -lt [datetime]::UtcNow.AddDays(-30)) {
        Write-Status 'Call records are retained for 30 days; calls older than that are missing from this report.' -Level WARN
    }
    if ($window.EndUtc -gt [datetime]::UtcNow.AddHours(-1)) {
        Write-Status 'Call records lag 30-60 minutes or more behind live calls; recent calls may be missing or incomplete.' -Level WARN
    }

    $resolvedAuthenticationMode = Resolve-AuthenticationMode -RequestedMode $AuthenticationMode
    Write-Status "Selected authentication mode: $resolvedAuthenticationMode."

    if ($resolvedAuthenticationMode -eq 'Automation') {
        $tenantId = Get-AutomationVariable -Name 'M365_TenantId'
        $appClientId = Get-AutomationVariable -Name 'M365_AppClientId'
        if ([string]::IsNullOrWhiteSpace($tenantId)) {
            throw 'M365_TenantId is empty.'
        }
        if ([string]::IsNullOrWhiteSpace($appClientId)) {
            throw 'M365_AppClientId is empty. The environment application registration has not been configured.'
        }
        $script:ConnectionSettings = @{ Mode = 'Automation'; TenantId = $tenantId; ClientId = $appClientId }
    } else {
        $resolvedSupportRoot = Resolve-LocalSupportRoot -ConfiguredRoot $LocalSupportRoot -ScriptDirectory $scriptDirectory
        $script:ConnectionSettings = @{ Mode = 'Delinea'; SupportRoot = $resolvedSupportRoot }
    }
    Connect-GraphAccess

    $graphRoles = @((Get-GraphField -Object (Get-JwtPayload -AccessToken $script:GraphAccessToken) -Name 'roles'))
    Assert-ApplicationRole -ActualRoles $graphRoles -RequiredRoles @('CallRecords.Read.All') -Purpose 'call record collection'
    Assert-ApplicationRole -ActualRoles $graphRoles -RequiredRoles @('Sites.ReadWrite.All', 'Sites.Manage.All', 'Sites.FullControl.All') -Purpose 'SharePoint report upload'

    Write-Status 'Validating the SharePoint destination folder.'
    $destination = Resolve-SharePointFolderDestination -FolderUrl $SharePointFolderUrl
    Write-Status "Validated SharePoint destination: $($destination.FolderWebUrl)"

    Write-Status "Listing call records from $windowText."
    $utcFormat = "yyyy-MM-dd'T'HH:mm:ss'Z'"
    $filter = [uri]::EscapeDataString(('startDateTime ge {0} and startDateTime lt {1}' -f $window.StartUtc.ToString($utcFormat, [System.Globalization.CultureInfo]::InvariantCulture), $window.EndUtc.ToString($utcFormat, [System.Globalization.CultureInfo]::InvariantCulture)))
    $callRecords = @(Get-GraphCollection -Uri "https://graph.microsoft.com/v1.0/communications/callRecords?`$filter=$filter")
    $audioRecords = @($callRecords | Where-Object { @((Get-GraphField -Object $_ -Name 'modalities')) -contains 'audio' })
    Write-Status "Found $($callRecords.Count) call records, $($audioRecords.Count) with audio. Reading session detail."

    $calls = New-Object System.Collections.Generic.List[object]
    $detailFailures = 0
    $processed = 0
    $progressTimer = [System.Diagnostics.Stopwatch]::StartNew()
    foreach ($callRecord in $audioRecords) {
        $processed++
        if ($processed % 100 -eq 0 -or $progressTimer.Elapsed.TotalSeconds -ge 60) {
            Write-Status "Read session detail for $processed of $($audioRecords.Count) call records."
            $progressTimer.Restart()
        }

        try {
            $call = Get-CallFlow -CallRecord $callRecord -VoiceAppCallsOnly $OnlyVoiceAppCalls -CallerFilter $CallerNumberFilter
        } catch {
            $detailFailures++
            Write-Status "Could not read call record $([string](Get-GraphField -Object $callRecord -Name 'id')): $($_.Exception.Message)" -Level WARN
            continue
        }
        if ($null -ne $call) {
            $calls.Add($call)
        }
    }

    $sortedCalls = @($calls | Sort-Object CallStartUtc)
    $hopRows = New-Object System.Collections.Generic.List[object]
    $callNumber = 0
    foreach ($call in $sortedCalls) {
        $callNumber++
        if ($WriteCallFlowTree) {
            Write-CallFlowTree -Number $callNumber -Call $call
        }

        $callStart = Format-ReportTime -UtcValue $call.CallStartUtc
        $hop = 0
        foreach ($leg in $call.Legs) {
            $hop++
            $hopRows.Add([pscustomobject]@{
                CallStart = $callStart
                Caller = $call.Caller
                Hop = $hop
                LegStart = Format-ReportTime -UtcValue $leg.StartUtc
                From = $leg.From.Label
                FromKind = $leg.From.Kind
                To = $leg.To.Label
                ToKind = $leg.To.Kind
                LegDuration = $leg.Duration
                HopOutcome = Get-HopOutcome -Leg $leg
                Failure = $leg.Failure
                CallOutcome = $call.Outcome
                AnsweredBy = $call.AnsweredBy
                CallRecordId = $call.CallRecordId
            })
        }
    }
    $summaryRows = @(
        foreach ($call in $sortedCalls) {
            [pscustomobject]@{
                CallStart = Format-ReportTime -UtcValue $call.CallStartUtc
                Caller = $call.Caller
                DialedNumber = $call.DialedNumber
                EntryPoint = $call.EntryPoint
                EntryPointKind = $call.EntryPointKind
                AutoAttendants = $call.AutoAttendants
                CallQueues = $call.CallQueues
                FinalQueue = $call.FinalQueue
                Outcome = $call.Outcome
                AnsweredBy = $call.AnsweredBy
                TimeToAnswerSeconds = $call.TimeToAnswerSeconds
                QueueWaitSeconds = $call.QueueWaitSeconds
                Duration = $call.Duration
                CallRecordId = $call.CallRecordId
            }
        }
    )

    $queueRows = @(Get-QueueRollupRows -Calls $sortedCalls)
    $outcomeCounts = [ordered]@{}
    foreach ($outcomeName in @('Answered', 'Voicemail', 'AbandonedInQueue', 'AbandonedInAA', 'Unanswered')) {
        $outcomeCounts[$outcomeName] = @($sortedCalls | Where-Object { $_.Outcome -eq $outcomeName }).Count
    }

    $script:TemporaryRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('CallFlowReport-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $script:TemporaryRoot -Force | Out-Null
    $fileSuffix = '{0:yyyyMMdd-HHmm}_{1:yyyyMMdd-HHmm}' -f $window.StartLocal, $window.EndLocal
    $hopCsvPath = Write-CsvReport -FileName "CallFlow-Hops_$fileSuffix.csv" -Columns $script:CsvColumns.Hops -Rows ([object[]]$hopRows.ToArray())
    $summaryCsvPath = Write-CsvReport -FileName "CallFlow-Summary_$fileSuffix.csv" -Columns $script:CsvColumns.Summary -Rows $summaryRows
    $queueCsvPath = Write-CsvReport -FileName "CallFlow-Queues_$fileSuffix.csv" -Columns $script:CsvColumns.Queues -Rows $queueRows

    $hopCsvUrl = ''
    $summaryCsvUrl = ''
    $queueCsvUrl = ''
    if ($DryRun) {
        Write-Status "DryRun enabled; CSV files remain in $script:TemporaryRoot and were not uploaded."
    } else {
        $hopCsvUrl = Publish-CsvReport -Path $hopCsvPath -Destination $destination
        $summaryCsvUrl = Publish-CsvReport -Path $summaryCsvPath -Destination $destination
        $queueCsvUrl = Publish-CsvReport -Path $queueCsvPath -Destination $destination
    }

    $outcomeText = ($outcomeCounts.Keys | ForEach-Object { '{0}={1}' -f $_, $outcomeCounts[$_] }) -join ', '
    Write-Status "Report complete for $windowText. Calls=$($sortedCalls.Count), $outcomeText, Queues=$($queueRows.Count), DetailFailures=$detailFailures."
    Write-Status 'DTMF menu choices are not exposed by callRecords; auto attendant menu paths are inferred from the next hop.'

    [pscustomobject][ordered]@{
        WindowStart = $window.StartLocal.ToString('yyyy-MM-dd HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)
        WindowEnd = $window.EndLocal.ToString('yyyy-MM-dd HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)
        TimeZoneId = $TimeZoneId
        CallRecordsScanned = $callRecords.Count
        AudioCallRecords = $audioRecords.Count
        CallsReported = $sortedCalls.Count
        Answered = $outcomeCounts['Answered']
        Voicemail = $outcomeCounts['Voicemail']
        AbandonedInQueue = $outcomeCounts['AbandonedInQueue']
        AbandonedInAA = $outcomeCounts['AbandonedInAA']
        Unanswered = $outcomeCounts['Unanswered']
        DetailFailures = $detailFailures
        HopCsvUrl = $hopCsvUrl
        SummaryCsvUrl = $summaryCsvUrl
        QueueCsvUrl = $queueCsvUrl
    }
} catch {
    Write-Status $_.Exception.Message -Level ERROR
    throw
} finally {
    if (-not $DryRun -and -not [string]::IsNullOrWhiteSpace($script:TemporaryRoot) -and (Test-Path -LiteralPath $script:TemporaryRoot)) {
        Remove-Item -LiteralPath $script:TemporaryRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
    $script:GraphAccessToken = ''
}
