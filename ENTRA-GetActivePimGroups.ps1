<#
.SYNOPSIS
    Reports the PIM for Groups assignments currently active for a user.

.DESCRIPTION
    Read-only diagnostic runbook. Resolves the target user, then lists the
    user's Privileged Identity Management (PIM) for Groups assignment schedule
    instances from Microsoft Graph and keeps only those active now. An active
    instance is either a just-in-time activation of an eligible assignment
    (AssignmentType 'activated') or a standing active assignment
    (AssignmentType 'assigned').

    Instances are filtered by the user's object ID. MemberType reports whether
    Graph returned the instance as a direct assignment or one inherited through
    group membership.

    The runbook makes no changes and is not Jira-gated.

    Required application roles on the Microsoft Graph token:
    - PIM read: PrivilegedAssignmentSchedule.Read.AzureADGroup (least
      privileged), or PrivilegedAssignmentSchedule.ReadWrite.AzureADGroup,
      PrivilegedAccess.Read.AzureADGroup, or
      PrivilegedAccess.ReadWrite.AzureADGroup.
    - User lookup: User.Read.All, User.ReadWrite.All, Directory.Read.All, or
      Directory.ReadWrite.All.
    - Group names: Group.Read.All, Group.ReadWrite.All, Directory.Read.All, or
      Directory.ReadWrite.All.

.PARAMETER User
    User principal name or Entra object ID of the user to inspect.

.OUTPUTS
    PSCustomObject with the resolved user, the evaluation time, the active
    group count, and an ActiveGroups array of group, access, and schedule
    details.

.EXAMPLE
    Start-AzAutomationRunbook -ResourceGroupName $resourceGroup `
        -AutomationAccountName $automationAccount `
        -Name 'ENTRA-GetActivePimGroups' `
        -Parameters @{ User = 'clayton.smith@corespecialty.com' }
#>

param(
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$User,

    [ValidateSet('Auto', 'Automation', 'Delinea')]
    [string]$AuthenticationMode = 'Auto',

    [string]$LocalSupportRoot = $env:M365_SCRIPTS_ROOT
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Terraform replaces this marker with the shared authentication source. It
# remains a comment locally, where the same source is dot-sourced below.
# __SHARED_FUNCTIONS__

$script:GraphAccessToken = ''

$User = $User.Trim()
$parsedUserId = [guid]::Empty
$userIsObjectId = [guid]::TryParse($User, [ref]$parsedUserId)
if (-not $userIsObjectId -and $User -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
    throw "User '$User' is neither a valid user principal name nor an object ID."
}

function Write-Status {
    param(
        [Parameter(Mandatory)]
        [string]$Message
    )

    Write-Host "[STATUS] $Message" -ForegroundColor Cyan
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

function Get-JwtApplicationRole {
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
        $claims = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($payload)) | ConvertFrom-Json
    } catch {
        throw "The access token claims could not be decoded: $($_.Exception.Message)"
    }

    return @((Get-GraphField -Object $claims -Name 'roles'))
}

function Assert-AnyApplicationRole {
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [string[]]$ActualRoles,

        [Parameter(Mandatory)]
        [string[]]$RequiredRoles,

        [Parameter(Mandatory)]
        [string]$Purpose,

        [Parameter(Mandatory)]
        [string]$Audience
    )

    if (@($RequiredRoles | Where-Object { $ActualRoles -contains $_ }).Count -eq 0) {
        throw "$Audience permission preflight failed for $Purpose. The environment application requires one of: $($RequiredRoles -join ', ')."
    }
}

function Invoke-GraphRequest {
    param(
        [Parameter(Mandatory)]
        [string]$Uri
    )

    if ([string]::IsNullOrWhiteSpace($script:GraphAccessToken)) {
        throw 'The Microsoft Graph access token has not been initialized.'
    }

    return Invoke-RestMethod -Method Get -Uri $Uri -Headers @{ Authorization = "Bearer $script:GraphAccessToken" } -ErrorAction Stop
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

function Resolve-GraphUser {
    param(
        [Parameter(Mandatory)]
        [string]$UserIdentifier
    )

    $requestUri = "https://graph.microsoft.com/v1.0/users/$([System.Uri]::EscapeDataString($UserIdentifier))?`$select=id,userPrincipalName,displayName"
    $graphUser = Invoke-GraphRequest -Uri $requestUri

    $userId = [string](Get-GraphField -Object $graphUser -Name 'id')
    if ([string]::IsNullOrWhiteSpace($userId)) {
        throw "Microsoft Graph did not return an object ID for user '$UserIdentifier'."
    }

    return [pscustomobject]@{
        Id = $userId
        UserPrincipalName = [string](Get-GraphField -Object $graphUser -Name 'userPrincipalName')
        DisplayName = [string](Get-GraphField -Object $graphUser -Name 'displayName')
    }
}

function Get-GroupDisplayName {
    param(
        [Parameter(Mandatory)]
        [string]$GroupId,

        [Parameter(Mandatory)]
        [hashtable]$Cache
    )

    if (-not $Cache.ContainsKey($GroupId)) {
        try {
            $group = Invoke-GraphRequest -Uri "https://graph.microsoft.com/v1.0/groups/$([System.Uri]::EscapeDataString($GroupId))?`$select=id,displayName"
            $Cache[$GroupId] = [string](Get-GraphField -Object $group -Name 'displayName')
        } catch {
            # A deleted group can keep a schedule instance briefly; report the
            # instance rather than failing the whole diagnostic.
            Write-Warning "Group '$GroupId' could not be read: $($_.Exception.Message)"
            $Cache[$GroupId] = ''
        }
    }

    return $Cache[$GroupId]
}

function ConvertTo-NullableDateTimeOffset {
    param(
        $Value
    )

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        return $null
    }
    if ($Value -is [datetime]) {
        return [datetimeoffset]$Value.ToUniversalTime()
    }
    if ($Value -is [datetimeoffset]) {
        return $Value
    }

    return [datetimeoffset]::Parse([string]$Value, [System.Globalization.CultureInfo]::InvariantCulture)
}

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

    $m365Connection = Connect-M365ServicesWithManagedIdentity -TenantId $tenantId -ClientId $appClientId -SkipTeams
} else {
    $resolvedSupportRoot = Resolve-LocalSupportRoot -ConfiguredRoot $LocalSupportRoot -ScriptDirectory $scriptDirectory
    $m365Connection = Connect-M365ServicesWithDelinea -SupportRoot $resolvedSupportRoot -SkipTeams
}

$script:GraphAccessToken = $m365Connection.GraphAccessToken
if ([string]::IsNullOrWhiteSpace($script:GraphAccessToken)) {
    throw 'The shared authentication profile did not return a Microsoft Graph access token.'
}

$graphRoles = @(Get-JwtApplicationRole -AccessToken $script:GraphAccessToken)
Assert-AnyApplicationRole -ActualRoles $graphRoles -RequiredRoles @(
    'PrivilegedAssignmentSchedule.Read.AzureADGroup',
    'PrivilegedAssignmentSchedule.ReadWrite.AzureADGroup',
    'PrivilegedAccess.Read.AzureADGroup',
    'PrivilegedAccess.ReadWrite.AzureADGroup'
) -Purpose 'reading PIM for Groups assignments' -Audience 'Microsoft Graph'
Assert-AnyApplicationRole -ActualRoles $graphRoles -RequiredRoles @(
    'User.Read.All',
    'User.ReadWrite.All',
    'Directory.Read.All',
    'Directory.ReadWrite.All'
) -Purpose 'user lookup' -Audience 'Microsoft Graph'
Assert-AnyApplicationRole -ActualRoles $graphRoles -RequiredRoles @(
    'Group.Read.All',
    'Group.ReadWrite.All',
    'Directory.Read.All',
    'Directory.ReadWrite.All'
) -Purpose 'group name lookup' -Audience 'Microsoft Graph'

Write-Status "Resolving user '$User'."
$targetUser = Resolve-GraphUser -UserIdentifier $User
Write-Status "Resolved '$($targetUser.UserPrincipalName)' ($($targetUser.Id))."

Write-Status 'Reading PIM for Groups assignment schedule instances.'
$instanceUri = "https://graph.microsoft.com/v1.0/identityGovernance/privilegedAccess/group/assignmentScheduleInstances?`$filter=principalId eq '$($targetUser.Id)'"
$instances = @(Get-GraphCollection -Uri $instanceUri)

$evaluatedAt = [datetimeoffset]::UtcNow
$groupNameCache = @{}
$activeGroups = New-Object System.Collections.Generic.List[object]

foreach ($instance in $instances) {
    $startDateTime = ConvertTo-NullableDateTimeOffset (Get-GraphField -Object $instance -Name 'startDateTime')
    $endDateTime = ConvertTo-NullableDateTimeOffset (Get-GraphField -Object $instance -Name 'endDateTime')

    if ($null -ne $startDateTime -and $startDateTime -gt $evaluatedAt) {
        continue
    }
    if ($null -ne $endDateTime -and $endDateTime -le $evaluatedAt) {
        continue
    }

    $groupId = [string](Get-GraphField -Object $instance -Name 'groupId')
    $activeGroups.Add([pscustomobject]@{
        GroupName = Get-GroupDisplayName -GroupId $groupId -Cache $groupNameCache
        GroupId = $groupId
        AccessId = [string](Get-GraphField -Object $instance -Name 'accessId')
        AssignmentType = [string](Get-GraphField -Object $instance -Name 'assignmentType')
        MemberType = [string](Get-GraphField -Object $instance -Name 'memberType')
        StartDateTime = if ($null -ne $startDateTime) { $startDateTime.ToString('o') } else { '' }
        EndDateTime = if ($null -ne $endDateTime) { $endDateTime.ToString('o') } else { 'Permanent' }
        MinutesRemaining = if ($null -ne $endDateTime) { [int][math]::Floor(($endDateTime - $evaluatedAt).TotalMinutes) } else { $null }
        InstanceId = [string](Get-GraphField -Object $instance -Name 'id')
    })
}

$sortedGroups = @($activeGroups | Sort-Object GroupName, AccessId)

Write-Status "Found $($sortedGroups.Count) active PIM group assignment(s) for '$($targetUser.UserPrincipalName)' out of $($instances.Count) schedule instance(s)."
foreach ($group in $sortedGroups) {
    Write-Status "  $($group.GroupName) [$($group.GroupId)] access=$($group.AccessId) type=$($group.AssignmentType) member=$($group.MemberType) ends=$($group.EndDateTime)"
}

[pscustomobject]@{
    UserPrincipalName = $targetUser.UserPrincipalName
    UserId = $targetUser.Id
    DisplayName = $targetUser.DisplayName
    EvaluatedAt = $evaluatedAt.ToString('o')
    ActiveGroupCount = $sortedGroups.Count
    ActiveGroups = $sortedGroups
}
