<#
.SYNOPSIS
    Reconciles SendGrid parent-account SSO teammates from Entra security groups.

.DESCRIPTION
    Runs in Azure Automation or from a local PowerShell 7 session. Automation
    uses the system-assigned identity federated to the environment application
    registration and reads the SendGrid API key from the encrypted
    M365_SendGridApiKey Automation variable. Local execution retrieves the
    application certificate and the SendGrid API key from Delinea Secret Server
    through the Clayton_M365_Scripts helpers.

    Entra groups are the source of truth. Membership in cs-sendgrid-* groups is
    reconciled onto SendGrid parent-account SSO teammates:
      cs-sendgrid-admin            -> is_admin = true
      cs-sendgrid-admin-ro         -> parent Observer scopes
      cs-sendgrid-<subuser>-<role> -> one subuser_access entry per subuser
        role admin                                 -> permission_type = admin
        role accountant|developer|marketer|observer -> permission_type =
        restricted with the subuser persona template (personas with campaigns
        access — marketer, developer — additionally carry marketing.read +
        marketing.automation.read, which the UI requires to show the subuser
        "Marketing" tab)

    Baseline: every SendGrid subuser gets its 5 role groups in Entra (plus the
    admin / admin-ro core groups); missing groups are created empty when
    -ApplyEntraGroups is true. Teammates whose Entra user is DISABLED or
    DELETED are removed when -ApplyTeammateRemovals is true. A Microsoft Graph
    lookup that fails for any reason other than a genuine 404 marks the user
    'unknown' and the teammate is left alone; lookup failure is never treated
    as permission to delete.

    Conflict rules: admin beats admin-ro beats subuser access. admin-ro
    together with any subuser group is a conflict and nothing is applied for
    that user. An unknown subuser token or unparseable group name is a
    warning, not a failure. A user in multiple role groups for the SAME
    subuser resolves to admin (if present) or the union of persona scopes.

    This runbook is a continuous reconciler intended to run on a recurring
    schedule. It does not correspond to a single tracked change request, so it
    does not require or publish a Jira ticket.

    Microsoft Graph is called directly with an app-only bearer token (no
    Microsoft.Graph PowerShell modules). Required application permissions:
    Group.ReadWrite.All (group discovery, creation, and membership reads;
    GroupMember.ReadWrite.All is also sufficient for membership writes) and
    User.ReadWrite.All (to validate Entra account status). The SendGrid API key
    requires teammate and subuser management scopes.
#>

param(
    [ValidateSet('Global', 'EU')]
    [string]$Region = 'Global',

    [ValidateNotNullOrEmpty()]
    [string]$GroupPrefix = 'cs-sendgrid-',

    [ValidateNotNullOrEmpty()]
    [string]$SsoAccessGroupName = 'cs-ea-Sendgrid-users',

    [string[]]$ProtectedTeammates = @(),

    [bool]$ApplyEntraGroups = $true,

    [bool]$ApplySsoAccessGroup = $true,

    [bool]$ApplyTeammateChanges = $true,

    [bool]$ApplyTeammateRemovals = $true,

    [bool]$ExportStateCsv = $false,

    [string]$ExportFolder = '',

    [ValidateSet('Auto', 'Automation', 'Delinea')]
    [string]$AuthenticationMode = 'Auto',

    [string]$LocalSupportRoot = $env:M365_SCRIPTS_ROOT
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Terraform replaces this marker with the shared authentication and SendGrid
# API sources. It remains a comment locally, where those sources are
# dot-sourced below.
# __SHARED_FUNCTIONS__

$script:GraphAccessToken = ''

$GroupPrefix = $GroupPrefix.Trim()
$SsoAccessGroupName = $SsoAccessGroupName.Trim()
if ([string]::IsNullOrWhiteSpace($SsoAccessGroupName)) {
    throw 'SsoAccessGroupName must not be empty or whitespace.'
}

# Scopes SendGrid adds on its own; ignored when comparing desired vs current.
# (stats.read is injected on restricted subuser entries even when the persona
# template omits it.)
$SendGridImplicitScopes = @(
    '2fa_exempt',
    '2fa_required',
    'sender_verification_eligible',
    'sender_verification_legacy',
    'stats.read',
    'user.profile.read',
    'user.profile.update'
)

# Scopes SendGrid may silently refuse to persist on restricted subuser access
# (seen with the Marketing-tab pair). They are still requested on every update,
# but never count as drift on their own, so the sync cannot re-patch forever
# when SendGrid drops them.
$SendGridBestEffortScopes = @(
    'marketing.read',
    'marketing.automation.read'
)

function Write-Status {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('INFO', 'WARN')]
        [string]$Level = 'INFO'
    )

    if ($Level -eq 'WARN') {
        Write-Host "[WARN] $Message" -ForegroundColor Yellow
    } else {
        Write-Host "[STATUS] $Message" -ForegroundColor Cyan
    }
}

$warnings = New-Object System.Collections.Generic.List[string]
function Add-SyncWarning {
    param(
        [Parameter(Mandatory)]
        [string]$Message
    )

    [void]$warnings.Add($Message)
    Write-Status $Message -Level WARN
}

function Get-OptionalObjectProperty {
    param(
        $InputObject,
        [string]$Name,
        $Default = $null
    )

    if ($null -eq $InputObject) {
        return $Default
    }

    if ($InputObject -is [System.Collections.IDictionary]) {
        if ($InputObject.Contains($Name)) {
            return $InputObject[$Name]
        }
        return $Default
    }

    $prop = $InputObject.PSObject.Properties[$Name]
    if ($null -ne $prop) {
        return $prop.Value
    }

    return $Default
}

function Get-JwtApplicationRole {
    param(
        [Parameter(Mandatory)]
        [string]$AccessToken
    )

    $parts = $AccessToken.Split('.')
    if ($parts.Count -lt 2) {
        throw 'The Microsoft Graph access token is not a valid JWT.'
    }

    try {
        $payload = $parts[1].Replace('-', '+').Replace('_', '/')
        switch ($payload.Length % 4) {
            2 { $payload += '==' }
            3 { $payload += '=' }
        }
        $claims = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($payload)) | ConvertFrom-Json
    } catch {
        throw "The Microsoft Graph access token claims could not be decoded: $($_.Exception.Message)"
    }

    return @((Get-OptionalObjectProperty -InputObject $claims -Name 'roles' -Default @()))
}

function Assert-AnyApplicationRole {
    param(
        [Parameter(Mandatory)]
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

function Test-GraphNotFoundError {
    param(
        [Parameter(Mandatory)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    try {
        if ($ErrorRecord.Exception.Response -and $ErrorRecord.Exception.Response.StatusCode) {
            return ([int]$ErrorRecord.Exception.Response.StatusCode -eq 404)
        }
    } catch {
    }

    return $false
}

function Invoke-GraphRequest {
    param(
        [Parameter(Mandatory)]
        [string]$Uri,

        [ValidateSet('Get', 'Post', 'Patch', 'Delete')]
        [string]$Method = 'Get',

        $Body
    )

    if ([string]::IsNullOrWhiteSpace($script:GraphAccessToken)) {
        throw 'The Microsoft Graph access token has not been initialized.'
    }

    $parameters = @{
        Method      = $Method
        Uri         = $Uri
        Headers     = @{ Authorization = "Bearer $script:GraphAccessToken" }
        ErrorAction = 'Stop'
    }
    if ($null -ne $Body) {
        $parameters.ContentType = 'application/json'
        $parameters.Body = $Body | ConvertTo-Json -Depth 10
    }

    return Invoke-RestMethod @parameters
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
        foreach ($item in @((Get-OptionalObjectProperty -InputObject $response -Name 'value' -Default @()))) {
            if ($null -ne $item) {
                $items.Add($item)
            }
        }
        $nextLink = [string](Get-OptionalObjectProperty -InputObject $response -Name '@odata.nextLink' -Default '')
    }

    return [object[]]$items.ToArray()
}

$implicitScopeSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($implicitScope in $SendGridImplicitScopes) {
    [void]$implicitScopeSet.Add($implicitScope)
}
foreach ($bestEffortScope in $SendGridBestEffortScopes) {
    [void]$implicitScopeSet.Add($bestEffortScope)
}

function ConvertTo-ScopeSet {
    # Normalized scope set with SendGrid's implicit scopes removed.
    [CmdletBinding()]
    param(
        [string[]]$Scopes
    )

    $set = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($scopeName in @($Scopes)) {
        $trimmed = ([string]$scopeName).Trim()
        if (-not [string]::IsNullOrWhiteSpace($trimmed) -and -not $implicitScopeSet.Contains($trimmed)) {
            [void]$set.Add($trimmed)
        }
    }

    return , $set
}

function Get-SubuserAccessDrift {
    # $null when the current access satisfies the desired entries; otherwise a
    # short description of the first difference found. Implicit and best-effort
    # scopes are excluded from the comparison.
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Access,

        [object[]]$DesiredEntries
    )

    if (-not [bool]$Access.HasRestrictedSubuserAccess) {
        return 'teammate is not in restricted subuser mode'
    }

    $currentById = New-Object 'System.Collections.Generic.Dictionary[int, object]'
    foreach ($currentEntry in @($Access.SubuserAccess)) {
        $currentId = Get-OptionalObjectProperty -InputObject $currentEntry -Name 'id' -Default $null
        if ($null -ne $currentId) {
            $currentById[[int]$currentId] = $currentEntry
        }
    }

    if (@($DesiredEntries).Count -ne $currentById.Count) {
        return "has $($currentById.Count) subuser entries, expected $(@($DesiredEntries).Count)"
    }

    foreach ($desiredEntry in @($DesiredEntries)) {
        $desiredId = [int]$desiredEntry['id']
        if (-not $currentById.ContainsKey($desiredId)) {
            return "no entry for subuser id $desiredId"
        }

        $currentEntry = $currentById[$desiredId]
        $currentType = ([string](Get-OptionalObjectProperty -InputObject $currentEntry -Name 'permission_type' -Default '')).Trim().ToLowerInvariant()
        if ($currentType -ne [string]$desiredEntry['permission_type']) {
            return "subuser id ${desiredId}: permission_type is '$currentType', expected '$($desiredEntry['permission_type'])'"
        }

        if ($currentType -eq 'restricted') {
            $desiredSet = ConvertTo-ScopeSet -Scopes @($desiredEntry['scopes'])
            $currentSet = ConvertTo-ScopeSet -Scopes @((Get-OptionalObjectProperty -InputObject $currentEntry -Name 'scopes' -Default @()))
            $missingScopes = @($desiredSet | Where-Object { -not $currentSet.Contains($_) } | Sort-Object)
            $extraScopes = @($currentSet | Where-Object { -not $desiredSet.Contains($_) } | Sort-Object)
            if ($missingScopes.Count -gt 0 -or $extraScopes.Count -gt 0) {
                $parts = New-Object System.Collections.Generic.List[string]
                if ($missingScopes.Count -gt 0) { [void]$parts.Add("missing scope(s): $($missingScopes -join ', ')") }
                if ($extraScopes.Count -gt 0) { [void]$parts.Add("extra scope(s): $($extraScopes -join ', ')") }
                return "subuser id ${desiredId}: $($parts -join '; ')"
            }
        }
    }

    return $null
}

function Resolve-EntraUserAccountState {
    # Hardened: only a genuine 404 counts as not-found; any other Graph
    # failure returns Status 'unknown' so callers never act destructively on
    # flaky lookups.
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Email
    )

    $normalized = $Email.Trim().ToLowerInvariant()
    $selectClause = 'id,userPrincipalName,mail,accountEnabled,givenName,surname,displayName'
    $result = [ordered]@{
        Status      = 'unknown'
        Upn         = $null
        UserId      = $null
        GivenName   = ''
        Surname     = ''
        DisplayName = ''
        Detail      = ''
    }

    $user = $null
    $notFoundById = $false
    try {
        $user = Invoke-GraphRequest -Uri "https://graph.microsoft.com/v1.0/users/$([uri]::EscapeDataString($normalized))?`$select=$selectClause"
    } catch {
        if (Test-GraphNotFoundError -ErrorRecord $_) {
            $notFoundById = $true
        } else {
            $result.Detail = $_.Exception.Message
            return [pscustomobject]$result
        }
    }

    if ($null -eq $user -and $notFoundById) {
        $escaped = $normalized.Replace("'", "''")
        $filterUri = "https://graph.microsoft.com/v1.0/users?`$filter=mail eq '$escaped' or userPrincipalName eq '$escaped'&`$select=$selectClause"
        try {
            $candidates = @(Get-GraphCollection -Uri $filterUri)
        } catch {
            $result.Detail = $_.Exception.Message
            return [pscustomobject]$result
        }

        if ($candidates.Count -gt 0) {
            $exact = $candidates | Where-Object {
                (([string](Get-OptionalObjectProperty -InputObject $_ -Name 'userPrincipalName' -Default '')).Trim().ToLowerInvariant() -eq $normalized) -or
                (([string](Get-OptionalObjectProperty -InputObject $_ -Name 'mail' -Default '')).Trim().ToLowerInvariant() -eq $normalized)
            } | Select-Object -First 1
            $user = if ($exact) { $exact } else { $candidates[0] }
        }
    }

    if ($null -eq $user) {
        $result.Status = 'not-found'
        return [pscustomobject]$result
    }

    $result.Status = if ([bool](Get-OptionalObjectProperty -InputObject $user -Name 'accountEnabled' -Default $true)) { 'enabled' } else { 'disabled' }
    $result.Upn = ([string](Get-OptionalObjectProperty -InputObject $user -Name 'userPrincipalName' -Default $normalized)).Trim().ToLowerInvariant()
    $result.UserId = [string](Get-OptionalObjectProperty -InputObject $user -Name 'id' -Default '')
    $result.GivenName = [string](Get-OptionalObjectProperty -InputObject $user -Name 'givenName' -Default '')
    $result.Surname = [string](Get-OptionalObjectProperty -InputObject $user -Name 'surname' -Default '')
    $result.DisplayName = [string](Get-OptionalObjectProperty -InputObject $user -Name 'displayName' -Default '')
    return [pscustomobject]$result
}

function Get-TeammateNameParts {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$UserRecord,

        [Parameter(Mandatory)]
        [string]$Upn
    )

    $firstName = ([string](Get-OptionalObjectProperty -InputObject $UserRecord -Name 'givenName' -Default '')).Trim()
    $lastName = ([string](Get-OptionalObjectProperty -InputObject $UserRecord -Name 'surname' -Default '')).Trim()
    if ([string]::IsNullOrWhiteSpace($firstName) -or [string]::IsNullOrWhiteSpace($lastName)) {
        $displayParts = @((([string](Get-OptionalObjectProperty -InputObject $UserRecord -Name 'displayName' -Default '')).Trim() -split '\s+') | Where-Object { $_ })
        if ([string]::IsNullOrWhiteSpace($firstName)) { $firstName = if ($displayParts.Count -gt 0) { $displayParts[0] } else { ($Upn -split '@')[0] } }
        if ([string]::IsNullOrWhiteSpace($lastName)) { $lastName = if ($displayParts.Count -gt 1) { ($displayParts[1..($displayParts.Count - 1)] -join ' ') } else { 'User' } }
    }

    return [pscustomobject]@{ FirstName = $firstName; LastName = $lastName }
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

$functionsDirectory = Join-Path (Split-Path $scriptDirectory -Parent) 'functions'
$sharedFunctions = [ordered]@{
    'Resolve-AuthenticationMode' = 'M365Authentication.ps1'
    'New-SendGridClient'         = 'Sendgrid.ps1'
}
foreach ($commandName in $sharedFunctions.Keys) {
    if ($null -ne (Get-Command $commandName -ErrorAction SilentlyContinue)) {
        continue
    }

    $functionPath = Join-Path $functionsDirectory $sharedFunctions[$commandName]
    if (-not (Test-Path -LiteralPath $functionPath -PathType Leaf)) {
        throw "Shared function '$commandName' was not injected and '$functionPath' is unavailable for local dot-sourcing."
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

    $sendGridApiKey = Get-AutomationVariable -Name 'M365_SendGridApiKey'
    if ([string]::IsNullOrWhiteSpace($sendGridApiKey) -or $sendGridApiKey -eq 'REPLACE_SECRET_NOT_SUPPLIED') {
        throw 'M365_SendGridApiKey is not set. Populate the SENDGRID_API_KEY GitHub environment secret and rerun the deployment workflow.'
    }
    $sendGridClient = New-SendGridClient -ApiKey $sendGridApiKey -Region $Region -DefaultPageSize 500
} else {
    $resolvedSupportRoot = Resolve-LocalSupportRoot -ConfiguredRoot $LocalSupportRoot -ScriptDirectory $scriptDirectory
    $m365Connection = Connect-M365ServicesWithDelinea -SupportRoot $resolvedSupportRoot -SkipTeams

    Write-Status 'Resolving the SendGrid API key from Delinea Secret Server.'
    $sendGridClient = Get-SendGridClientFromLocalSupport -SupportRoot $resolvedSupportRoot -Region $Region -DefaultPageSize 500
}

$script:GraphAccessToken = $m365Connection.GraphAccessToken
if ([string]::IsNullOrWhiteSpace($script:GraphAccessToken)) {
    throw 'The shared authentication profile did not return a Microsoft Graph access token.'
}

$graphRoles = @(Get-JwtApplicationRole -AccessToken $script:GraphAccessToken)
Assert-AnyApplicationRole -ActualRoles $graphRoles -RequiredRoles @('Group.ReadWrite.All') -Purpose 'Entra group discovery, creation, and membership reads' -Audience 'Microsoft Graph'
Assert-AnyApplicationRole -ActualRoles $graphRoles -RequiredRoles @('GroupMember.ReadWrite.All', 'Group.ReadWrite.All') -Purpose 'SSO access group membership writes' -Audience 'Microsoft Graph'
Assert-AnyApplicationRole -ActualRoles $graphRoles -RequiredRoles @('User.ReadWrite.All') -Purpose 'Entra user account-status validation' -Audience 'Microsoft Graph'

Write-Host ''
Write-Host 'Apply flags:' -ForegroundColor Cyan
Write-Host ("  ApplyEntraGroups      : {0}" -f $ApplyEntraGroups)
Write-Host ("  ApplySsoAccessGroup   : {0}" -f $ApplySsoAccessGroup)
Write-Host ("  ApplyTeammateChanges  : {0}" -f $ApplyTeammateChanges)
Write-Host ("  ApplyTeammateRemovals : {0}" -f $ApplyTeammateRemovals)

$protectedSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($protectedEmail in @($ProtectedTeammates)) {
    if (-not [string]::IsNullOrWhiteSpace([string]$protectedEmail)) {
        [void]$protectedSet.Add(([string]$protectedEmail).Trim().ToLowerInvariant())
    }
}

# --- 1) SendGrid subusers + Entra baseline groups ----------------------------
Write-Host ''
Write-Host '1) Baseline role groups' -ForegroundColor Cyan

$subuserByToken = New-Object 'System.Collections.Generic.Dictionary[string, object]' ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($subuser in @(Get-SendGridSubusers -Client $sendGridClient -PageSize 500)) {
    $subuserName = [string](Get-OptionalObjectProperty -InputObject $subuser -Name 'username' -Default '')
    $subuserId = Get-OptionalObjectProperty -InputObject $subuser -Name 'id' -Default $null
    if ([string]::IsNullOrWhiteSpace($subuserName) -or $null -eq $subuserId) {
        continue
    }

    $token = ConvertTo-GroupToken -Value $subuserName
    if ($subuserByToken.ContainsKey($token)) {
        Add-SyncWarning "Subusers '$($subuserByToken[$token].Name)' and '$subuserName' collapse to the same group token '$token'; skipping the second."
        continue
    }

    $subuserByToken[$token] = [pscustomobject]@{ Name = $subuserName; Id = [int]$subuserId }
}

Write-Host ("  SendGrid subusers: {0}" -f $subuserByToken.Count)

$prefixFilterUri = "https://graph.microsoft.com/v1.0/groups?`$filter=startswith(displayName,'$($GroupPrefix.Replace("'", "''"))')&`$select=id,displayName"
$prefixGroups = @(Get-GraphCollection -Uri $prefixFilterUri)
$groupIdByName = New-Object 'System.Collections.Generic.Dictionary[string, string]' ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($g in $prefixGroups) {
    $gName = ([string](Get-OptionalObjectProperty -InputObject $g -Name 'displayName' -Default '')).Trim()
    if (-not [string]::IsNullOrWhiteSpace($gName)) {
        $groupIdByName[$gName] = [string](Get-OptionalObjectProperty -InputObject $g -Name 'id' -Default '')
    }
}

$roleSlugs = @('admin') + @(Get-SupportedPersonaSlugs)
$requiredGroupNames = New-Object System.Collections.Generic.List[string]
[void]$requiredGroupNames.Add((Get-SendGridAdminRoleSpec).GroupName)
[void]$requiredGroupNames.Add((Get-SendGridAdminReadOnlyRoleSpec).GroupName)
foreach ($token in @($subuserByToken.Keys | Sort-Object)) {
    foreach ($roleSlug in $roleSlugs) {
        [void]$requiredGroupNames.Add(('{0}{1}-{2}' -f $GroupPrefix, $token, $roleSlug))
    }
}

$createdGroupCount = 0
foreach ($requiredName in $requiredGroupNames) {
    if ($groupIdByName.ContainsKey($requiredName)) {
        continue
    }

    if (-not $ApplyEntraGroups) {
        Write-Host ("  [WhatIf] Would create Entra group: {0}" -f $requiredName) -ForegroundColor Yellow
        continue
    }

    $nickname = ($requiredName -replace '[^A-Za-z0-9-]', '')
    if ($nickname.Length -gt 64) { $nickname = $nickname.Substring(0, 64) }
    try {
        $newGroup = Invoke-GraphRequest -Uri 'https://graph.microsoft.com/v1.0/groups' -Method Post -Body @{
            displayName     = $requiredName
            mailEnabled     = $false
            mailNickname    = $nickname
            securityEnabled = $true
            groupTypes      = @()
            description     = 'SendGrid access group (managed by SendGrid group sync).'
        }
        $newGroupId = [string](Get-OptionalObjectProperty -InputObject $newGroup -Name 'id' -Default '')
        if ([string]::IsNullOrWhiteSpace($newGroupId)) {
            throw "Microsoft Graph did not confirm creation of Entra group '$requiredName'."
        }
        $groupIdByName[$requiredName] = $newGroupId
        $createdGroupCount++
        Write-Host ("  Created Entra group: {0}" -f $requiredName) -ForegroundColor Green
    } catch {
        Add-SyncWarning "Failed to create group '$requiredName': $($_.Exception.Message)"
    }
}

if ($createdGroupCount -eq 0 -and $ApplyEntraGroups) {
    Write-Host '  All baseline groups already exist.' -ForegroundColor Green
}

# --- 2) Desired state from Entra group membership ----------------------------
Write-Host ''
Write-Host '2) Desired state from Entra groups' -ForegroundColor Cyan

# upn -> @{ IsAdmin; IsAdminRo; SubuserRoles (token -> role HashSet); Groups }
$desiredByUpn = New-Object 'System.Collections.Generic.Dictionary[string, object]' ([System.StringComparer]::OrdinalIgnoreCase)
# upn -> Entra user record from group membership (id, names, accountEnabled)
$userRecordByUpn = New-Object 'System.Collections.Generic.Dictionary[string, object]' ([System.StringComparer]::OrdinalIgnoreCase)
# One row per (group, member) pair, for the CSV export.
$membershipExportRows = New-Object System.Collections.Generic.List[object]

foreach ($groupName in @($groupIdByName.Keys | Sort-Object)) {
    $role = Resolve-SendGridRoleFromGroupName -GroupName $groupName
    if ($null -eq $role) {
        Add-SyncWarning "Group '$groupName' does not match any known role pattern; ignored."
        continue
    }

    if ($role.RoleKey -eq 'subuser-access' -and -not $subuserByToken.ContainsKey($role.Subuser)) {
        Add-SyncWarning "Group '$groupName' references unknown subuser token '$($role.Subuser)'; ignored."
        continue
    }

    $membersUri = "https://graph.microsoft.com/v1.0/groups/$($groupIdByName[$groupName])/members/microsoft.graph.user?`$select=id,userPrincipalName,accountEnabled,givenName,surname,displayName"
    $members = @(Get-GraphCollection -Uri $membersUri)
    foreach ($member in $members) {
        $upn = ([string](Get-OptionalObjectProperty -InputObject $member -Name 'userPrincipalName' -Default '')).Trim().ToLowerInvariant()
        if ([string]::IsNullOrWhiteSpace($upn)) {
            continue
        }

        $userRecordByUpn[$upn] = $member

        [void]$membershipExportRows.Add([pscustomobject]@{
            GroupName = $groupName
            MemberUpn = $upn
            RoleKind  = [string]$role.RoleKey
            Subuser   = if ($role.RoleKey -eq 'subuser-access') { [string]$subuserByToken[$role.Subuser].Name } else { '' }
            Role      = [string]$role.Role
            Enabled   = [bool](Get-OptionalObjectProperty -InputObject $member -Name 'accountEnabled' -Default $true)
        })

        if (-not $desiredByUpn.ContainsKey($upn)) {
            $desiredByUpn[$upn] = [pscustomobject]@{
                IsAdmin      = $false
                IsAdminRo    = $false
                SubuserRoles = (New-Object 'System.Collections.Generic.Dictionary[string, object]' ([System.StringComparer]::OrdinalIgnoreCase))
                Groups       = (New-Object System.Collections.Generic.List[string])
            }
        }

        $desired = $desiredByUpn[$upn]
        [void]$desired.Groups.Add($groupName)

        switch ($role.RoleKey) {
            'admin' { $desired.IsAdmin = $true }
            'admin-ro' { $desired.IsAdminRo = $true }
            'subuser-access' {
                if (-not $desired.SubuserRoles.ContainsKey($role.Subuser)) {
                    $desired.SubuserRoles[$role.Subuser] = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
                }
                [void]$desired.SubuserRoles[$role.Subuser].Add($role.Role)
            }
        }
    }
}

# Resolve each user's effective access per the conflict rules.
# upn -> @{ Kind = admin|admin-ro|subuser|conflict; Entries = @(payload hashtables); Summary }
$effectiveByUpn = New-Object 'System.Collections.Generic.Dictionary[string, object]' ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($upn in @($desiredByUpn.Keys | Sort-Object)) {
    $desired = $desiredByUpn[$upn]

    if ($desired.IsAdmin) {
        $effectiveByUpn[$upn] = [pscustomobject]@{ Kind = 'admin'; Entries = @(); Summary = 'account admin' }
        continue
    }

    if ($desired.IsAdminRo -and $desired.SubuserRoles.Count -gt 0) {
        Add-SyncWarning "User '$upn' is in admin-ro AND subuser group(s) ($(@($desired.Groups) -join ', ')); cannot coexist on one teammate. Applying nothing."
        $effectiveByUpn[$upn] = [pscustomobject]@{ Kind = 'conflict'; Entries = @(); Summary = 'conflict' }
        continue
    }

    if ($desired.IsAdminRo) {
        $effectiveByUpn[$upn] = [pscustomobject]@{ Kind = 'admin-ro'; Entries = @(); Summary = 'parent read-only (observer)' }
        continue
    }

    $entries = New-Object System.Collections.Generic.List[object]
    $summaryParts = New-Object System.Collections.Generic.List[string]
    foreach ($token in @($desired.SubuserRoles.Keys | Sort-Object)) {
        $roles = $desired.SubuserRoles[$token]
        $subuserInfo = $subuserByToken[$token]

        if ($roles.Contains('admin')) {
            if ($roles.Count -gt 1) {
                Add-SyncWarning "User '$upn' holds multiple roles on subuser '$($subuserInfo.Name)'; admin wins."
            }
            [void]$entries.Add(@{ id = $subuserInfo.Id; permission_type = 'admin' })
            [void]$summaryParts.Add(('{0} (admin)' -f $subuserInfo.Name))
            continue
        }

        $scopeUnion = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($personaRole in @($roles | Sort-Object)) {
            foreach ($scopeName in @(Get-SendGridPersonaScopes -Persona $personaRole -Scope Subuser)) {
                [void]$scopeUnion.Add($scopeName)
            }
        }

        if ($roles.Count -gt 1) {
            Add-SyncWarning "User '$upn' holds multiple personas on subuser '$($subuserInfo.Name)' ($(@($roles | Sort-Object) -join ', ')); granting the union of their scopes."
        }

        [void]$entries.Add(@{ id = $subuserInfo.Id; permission_type = 'restricted'; scopes = @($scopeUnion | Sort-Object) })
        [void]$summaryParts.Add(('{0} ({1})' -f $subuserInfo.Name, (@($roles | Sort-Object) -join '+')))
    }

    $effectiveByUpn[$upn] = [pscustomobject]@{
        Kind    = 'subuser'
        Entries = $entries.ToArray()
        Summary = ($summaryParts -join ', ')
    }
}

Write-Host ("  Users with desired SendGrid access: {0}" -f $effectiveByUpn.Count)

# --- 3) Current state from SendGrid ------------------------------------------
Write-Host ''
Write-Host '3) Current SendGrid teammates + Entra account states' -ForegroundColor Cyan

$teammateStates = New-Object System.Collections.Generic.List[object]
$teammateByUpn = New-Object 'System.Collections.Generic.Dictionary[string, object]' ([System.StringComparer]::OrdinalIgnoreCase)

foreach ($teammate in @(Get-SendGridTeammates -Client $sendGridClient -PageSize 500)) {
    $email = ([string](Get-OptionalObjectProperty -InputObject $teammate -Name 'email' -Default '')).Trim().ToLowerInvariant()
    $username = ([string](Get-OptionalObjectProperty -InputObject $teammate -Name 'username' -Default '')).Trim()
    $lookupKey = if (-not [string]::IsNullOrWhiteSpace($username)) { $username } else { $email }

    if ([string]::IsNullOrWhiteSpace($lookupKey)) {
        Add-SyncWarning 'Teammate with no username or email found; skipped.'
        continue
    }

    $entraState = if (-not [string]::IsNullOrWhiteSpace($email)) {
        Resolve-EntraUserAccountState -Email $email
    } else {
        [pscustomobject]@{ Status = 'unknown'; Upn = $null; UserId = $null; GivenName = ''; Surname = ''; DisplayName = ''; Detail = 'teammate has no email' }
    }

    $state = [pscustomobject]@{
        LookupKey  = $lookupKey
        Email      = $email
        IsAdmin    = [bool](Get-OptionalObjectProperty -InputObject $teammate -Name 'is_admin' -Default $false)
        EntraState = $entraState
        Removed    = $false
    }
    [void]$teammateStates.Add($state)

    if (-not [string]::IsNullOrWhiteSpace([string]$entraState.Upn)) {
        $teammateByUpn[$entraState.Upn] = $state
    } elseif (-not [string]::IsNullOrWhiteSpace($email)) {
        # Keep an email-keyed fallback so desired users still match if Graph was flaky.
        if (-not $teammateByUpn.ContainsKey($email)) {
            $teammateByUpn[$email] = $state
        }
    }
}

Write-Host ("  Parent teammates: {0}" -f $teammateStates.Count)

# --- 4) Cleanup: disabled / deleted Entra users ------------------------------
Write-Host ''
Write-Host '4) Cleanup of disabled/deleted Entra users' -ForegroundColor Cyan

$removalCount = 0
foreach ($state in $teammateStates) {
    $status = [string]$state.EntraState.Status

    if ($status -eq 'enabled') {
        continue
    }

    if ($protectedSet.Contains($state.Email) -or $protectedSet.Contains($state.LookupKey)) {
        Write-Host ("  Protected: {0} (Entra status: {1}) - left alone." -f $state.LookupKey, $status) -ForegroundColor DarkGray
        continue
    }

    if ($status -eq 'unknown') {
        Add-SyncWarning "Could not determine Entra status for teammate '$($state.LookupKey)' ($($state.EntraState.Detail)); left alone."
        continue
    }

    $reason = if ($status -eq 'disabled') { 'entra-user-disabled' } else { 'not-found-in-entra' }

    if (-not $ApplyTeammateRemovals) {
        Write-Host ("  [WhatIf] Would remove SendGrid teammate: {0} ({1})" -f $state.LookupKey, $reason) -ForegroundColor Yellow
        continue
    }

    try {
        Remove-SendGridTeammate -Client $sendGridClient -TeammateName $state.LookupKey -Confirm:$false
        $state.Removed = $true
        $removalCount++
        Write-Host ("  Removed teammate: {0} ({1})" -f $state.LookupKey, $reason) -ForegroundColor Green
    } catch {
        Add-SyncWarning "Failed to remove teammate '$($state.LookupKey)': $(Get-SendGridErrorMessage -ErrorRecord $_)"
    }
}

if ($removalCount -eq 0 -and $ApplyTeammateRemovals) {
    Write-Host '  Nothing removed.' -ForegroundColor Green
}

# --- 5) Reconcile teammates with group-derived roles -------------------------
Write-Host ''
Write-Host '5) Reconcile teammates (create/update from groups)' -ForegroundColor Cyan

$teammateCreatedCount = 0
$teammateUpdatedCount = 0

foreach ($upn in @($effectiveByUpn.Keys | Sort-Object)) {
    $effective = $effectiveByUpn[$upn]
    if ($effective.Kind -eq 'conflict') {
        continue
    }

    if ($protectedSet.Contains($upn)) {
        Write-Host ("  Protected: {0} - left alone." -f $upn) -ForegroundColor DarkGray
        continue
    }

    $state = if ($teammateByUpn.ContainsKey($upn)) { $teammateByUpn[$upn] } else { $null }

    # --- Create ---
    if ($null -eq $state -or $state.Removed) {
        $userRecord = if ($userRecordByUpn.ContainsKey($upn)) { $userRecordByUpn[$upn] } else { $null }
        if ($null -eq $userRecord) {
            Add-SyncWarning "Desired user '$upn' has no Entra record; skipped."
            continue
        }

        if (-not [bool](Get-OptionalObjectProperty -InputObject $userRecord -Name 'accountEnabled' -Default $true)) {
            Add-SyncWarning "Desired user '$upn' is disabled in Entra but still in role group(s); remove them from the groups. Skipped."
            continue
        }

        if (-not $ApplyTeammateChanges) {
            Write-Host ("  [WhatIf] Would create SSO teammate {0}: {1}" -f $upn, $effective.Summary) -ForegroundColor Yellow
            continue
        }

        $names = Get-TeammateNameParts -UserRecord $userRecord -Upn $upn
        try {
            switch ($effective.Kind) {
                'admin' {
                    New-SendGridSsoTeammate -Client $sendGridClient -Email $upn -FirstName $names.FirstName -LastName $names.LastName -IsAdmin $true | Out-Null
                }
                'admin-ro' {
                    New-SendGridSsoTeammate -Client $sendGridClient -Email $upn -FirstName $names.FirstName -LastName $names.LastName -Scopes @(Get-SendGridPersonaScopes -Persona observer -Scope Parent) | Out-Null
                }
                'subuser' {
                    New-SendGridSsoTeammate -Client $sendGridClient -Email $upn -FirstName $names.FirstName -LastName $names.LastName -SubuserAccess $effective.Entries | Out-Null
                }
            }
            $teammateCreatedCount++
            Write-Host ("  Created SSO teammate {0}: {1}" -f $upn, $effective.Summary) -ForegroundColor Green
        } catch {
            $createError = Get-SendGridErrorMessage -ErrorRecord $_
            if ($createError -match 'username exists') {
                Add-SyncWarning "Failed to create teammate '$upn': $createError. The username is likely held by a legacy subuser-context teammate that must be migrated first."
            } else {
                Add-SyncWarning "Failed to create teammate '$upn': $createError"
            }
        }

        continue
    }

    # --- Update ---
    $teammateName = $state.LookupKey
    $encoded = [uri]::EscapeDataString($teammateName)

    try {
        if ($effective.Kind -eq 'admin') {
            if ($state.IsAdmin) {
                Write-Host ("  OK: {0} is admin." -f $upn) -ForegroundColor DarkGray
                continue
            }

            if (-not $ApplyTeammateChanges) {
                Write-Host ("  [WhatIf] Would promote {0} to admin." -f $upn) -ForegroundColor Yellow
                continue
            }

            Invoke-SendGridRequest -Client $sendGridClient -Method PATCH -Path "/v3/sso/teammates/$encoded" -Body @{ is_admin = $true } | Out-Null
            $teammateUpdatedCount++
            Write-Host ("  Promoted {0} to admin." -f $upn) -ForegroundColor Green
            continue
        }

        if ($state.IsAdmin) {
            # Downgrade from admin to the desired non-admin access.
            if (-not $ApplyTeammateChanges) {
                Write-Host ("  [WhatIf] Would downgrade {0} from admin to: {1}" -f $upn, $effective.Summary) -ForegroundColor Yellow
                continue
            }

            if ($effective.Kind -eq 'admin-ro') {
                Invoke-SendGridRequest -Client $sendGridClient -Method PATCH -Path "/v3/sso/teammates/$encoded" -Body @{ is_admin = $false; has_restricted_subuser_access = $false; scopes = @(Get-SendGridPersonaScopes -Persona observer -Scope Parent) } | Out-Null
            } else {
                Set-SendGridTeammateSubuserAccess -Client $sendGridClient -TeammateName $teammateName -SubuserAccess $effective.Entries | Out-Null
            }
            $teammateUpdatedCount++
            Write-Host ("  Downgraded {0} from admin to: {1}" -f $upn, $effective.Summary) -ForegroundColor Green
            continue
        }

        $access = Get-SendGridTeammateSubuserAccess -Client $sendGridClient -TeammateName $teammateName

        if ($effective.Kind -eq 'admin-ro') {
            $compliant = $false
            if (-not $access.HasRestrictedSubuserAccess) {
                $detail = Get-SendGridTeammate -Client $sendGridClient -TeammateName $teammateName
                $parentScopes = @((Get-OptionalObjectProperty -InputObject $detail -Name 'scopes' -Default @()))
                $resolved = Resolve-SendGridPersonaFromScopes -Scopes $parentScopes -Scope Parent
                if ($resolved.Persona -eq 'observer') {
                    $compliant = $true
                    if ($resolved.MatchType -eq 'superset') {
                        $driftScopes = @($resolved.ExtraScopes | Where-Object { -not $implicitScopeSet.Contains($_) })
                        if ($driftScopes.Count -gt 0) {
                            Add-SyncWarning "Teammate '$upn' is admin-ro compliant but carries extra scope(s): $($driftScopes -join ', ')"
                        }
                    }
                }
            }

            if ($compliant) {
                Write-Host ("  OK: {0} is parent read-only." -f $upn) -ForegroundColor DarkGray
                continue
            }

            if (-not $ApplyTeammateChanges) {
                Write-Host ("  [WhatIf] Would set {0} to parent read-only (observer)." -f $upn) -ForegroundColor Yellow
                continue
            }

            Invoke-SendGridRequest -Client $sendGridClient -Method PATCH -Path "/v3/sso/teammates/$encoded" -Body @{ is_admin = $false; has_restricted_subuser_access = $false; scopes = @(Get-SendGridPersonaScopes -Persona observer -Scope Parent) } | Out-Null
            $teammateUpdatedCount++
            Write-Host ("  Set {0} to parent read-only (observer)." -f $upn) -ForegroundColor Green
            continue
        }

        # Desired: subuser-scoped access.
        $drift = Get-SubuserAccessDrift -Access $access -DesiredEntries $effective.Entries

        if ($null -eq $drift) {
            Write-Host ("  OK: {0} matches: {1}" -f $upn, $effective.Summary) -ForegroundColor DarkGray
            continue
        }

        if (-not $ApplyTeammateChanges) {
            Write-Host ("  [WhatIf] Would set {0} subuser access to: {1} ({2})" -f $upn, $effective.Summary, $drift) -ForegroundColor Yellow
            continue
        }

        Set-SendGridTeammateSubuserAccess -Client $sendGridClient -TeammateName $teammateName -SubuserAccess $effective.Entries | Out-Null
        $teammateUpdatedCount++

        # Verify by reading back: SendGrid can return 200 yet drop parts of the request.
        $verifyAccess = Get-SendGridTeammateSubuserAccess -Client $sendGridClient -TeammateName $teammateName
        $verifyDrift = Get-SubuserAccessDrift -Access $verifyAccess -DesiredEntries $effective.Entries
        if ($null -eq $verifyDrift) {
            Write-Host ("  Set {0} subuser access to: {1} (verified)" -f $upn, $effective.Summary) -ForegroundColor Green
        }
        else {
            Add-SyncWarning "Teammate '$upn': update did not fully persist — $verifyDrift. SendGrid accepted the PATCH but dropped part of it."
        }

        # Report best-effort scopes SendGrid chose not to keep (informational; never counted as drift).
        $verifyById = New-Object 'System.Collections.Generic.Dictionary[int, object]'
        foreach ($verifyEntry in @($verifyAccess.SubuserAccess)) {
            $verifyId = Get-OptionalObjectProperty -InputObject $verifyEntry -Name 'id' -Default $null
            if ($null -ne $verifyId) { $verifyById[[int]$verifyId] = $verifyEntry }
        }
        foreach ($desiredEntry in @($effective.Entries)) {
            if ([string]$desiredEntry['permission_type'] -ne 'restricted' -or -not $verifyById.ContainsKey([int]$desiredEntry['id'])) {
                continue
            }
            $requestedBestEffort = @(@($desiredEntry['scopes']) | Where-Object { $SendGridBestEffortScopes -contains $_ })
            if ($requestedBestEffort.Count -eq 0) { continue }
            $verifyScopes = @((Get-OptionalObjectProperty -InputObject $verifyById[[int]$desiredEntry['id']] -Name 'scopes' -Default @()))
            $droppedBestEffort = @($requestedBestEffort | Where-Object { $verifyScopes -notcontains $_ })
            if ($droppedBestEffort.Count -gt 0) {
                Add-SyncWarning "Teammate '$upn', subuser id $($desiredEntry['id']): SendGrid did not persist optional scope(s) $($droppedBestEffort -join ', ') — the Marketing tab will not appear there (likely unsupported on this account/subuser)."
            }
        }
    } catch {
        Add-SyncWarning "Failed to reconcile teammate '$upn': $(Get-SendGridErrorMessage -ErrorRecord $_)"
    }
}

# Unmanaged teammates: enabled in Entra but in no role group. Left alone by design.
foreach ($state in $teammateStates) {
    if ($state.Removed -or [string]$state.EntraState.Status -ne 'enabled') {
        continue
    }

    $stateUpn = [string]$state.EntraState.Upn
    if (-not [string]::IsNullOrWhiteSpace($stateUpn) -and -not $effectiveByUpn.ContainsKey($stateUpn)) {
        Add-SyncWarning "Teammate '$($state.LookupKey)' is in no cs-sendgrid-* group (unmanaged); left alone. Add them to a group or remove them manually."
    }
}

# --- 6) SSO access group membership ------------------------------------------
Write-Host ''
Write-Host ("6) SSO access group ({0})" -f $SsoAccessGroupName) -ForegroundColor Cyan

$ssoFilterUri = "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$($SsoAccessGroupName.Replace("'", "''"))'&`$select=id,displayName"
$ssoGroups = @(Get-GraphCollection -Uri $ssoFilterUri)
$ssoGroupId = if ($ssoGroups.Count -gt 0) { [string](Get-OptionalObjectProperty -InputObject $ssoGroups[0] -Name 'id' -Default '') } else { $null }

if ([string]::IsNullOrWhiteSpace($ssoGroupId)) {
    $ssoGroupId = $null
    if ($ApplyEntraGroups) {
        $nickname = ($SsoAccessGroupName -replace '[^A-Za-z0-9-]', '')
        try {
            $newGroup = Invoke-GraphRequest -Uri 'https://graph.microsoft.com/v1.0/groups' -Method Post -Body @{
                displayName     = $SsoAccessGroupName
                mailEnabled     = $false
                mailNickname    = $nickname
                securityEnabled = $true
                groupTypes      = @()
                description     = 'SendGrid SSO access group (managed by SendGrid group sync).'
            }
            $ssoGroupId = [string](Get-OptionalObjectProperty -InputObject $newGroup -Name 'id' -Default '')
            Write-Host ("  Created Entra group: {0}" -f $SsoAccessGroupName) -ForegroundColor Green
        } catch {
            Add-SyncWarning "Failed to create SSO access group '$SsoAccessGroupName': $($_.Exception.Message)"
        }
    } else {
        Write-Host ("  [WhatIf] Would create Entra group: {0}" -f $SsoAccessGroupName) -ForegroundColor Yellow
    }
}

$ssoAdditionCount = 0
if (-not [string]::IsNullOrWhiteSpace($ssoGroupId)) {
    $ssoMemberUpns = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $ssoMembersUri = "https://graph.microsoft.com/v1.0/groups/$ssoGroupId/members/microsoft.graph.user?`$select=id,userPrincipalName"
    foreach ($member in @(Get-GraphCollection -Uri $ssoMembersUri)) {
        $memberUpn = ([string](Get-OptionalObjectProperty -InputObject $member -Name 'userPrincipalName' -Default '')).Trim().ToLowerInvariant()
        if (-not [string]::IsNullOrWhiteSpace($memberUpn)) {
            [void]$ssoMemberUpns.Add($memberUpn)
        }
    }

    foreach ($upn in @($desiredByUpn.Keys | Sort-Object)) {
        if ($ssoMemberUpns.Contains($upn)) {
            continue
        }

        if (-not $ApplySsoAccessGroup) {
            Write-Host ("  [WhatIf] Would add {0} to {1}" -f $upn, $SsoAccessGroupName) -ForegroundColor Yellow
            continue
        }

        $userRecord = if ($userRecordByUpn.ContainsKey($upn)) { $userRecordByUpn[$upn] } else { $null }
        $userId = [string](Get-OptionalObjectProperty -InputObject $userRecord -Name 'id' -Default '')
        if ([string]::IsNullOrWhiteSpace($userId)) {
            Add-SyncWarning "No Entra id for '$upn'; not added to '$SsoAccessGroupName'."
            continue
        }

        try {
            Invoke-GraphRequest -Uri "https://graph.microsoft.com/v1.0/groups/$ssoGroupId/members/`$ref" -Method Post -Body @{
                '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$userId"
            } | Out-Null
            $ssoAdditionCount++
            Write-Host ("  Added {0} to {1}" -f $upn, $SsoAccessGroupName) -ForegroundColor Green
        } catch {
            Add-SyncWarning "Failed to add '$upn' to '$SsoAccessGroupName': $($_.Exception.Message)"
        }
    }
}

# --- 7) State CSV export ------------------------------------------------------
$resolvedExportFolder = $null
if ($ExportStateCsv) {
    $resolvedExportFolder = if (-not [string]::IsNullOrWhiteSpace($ExportFolder)) {
        $ExportFolder
    } else {
        Join-Path ([System.IO.Path]::GetTempPath()) ('SendGridSync_{0}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
    }

    $null = New-Item -ItemType Directory -Path $resolvedExportFolder -Force

    Write-Host ''
    Write-Host ("7) Exporting state CSVs to {0}" -f $resolvedExportFolder) -ForegroundColor Cyan

    Export-SendGridPlanCsv -Path (Join-Path $resolvedExportFolder 'group-memberships.csv') -Columns @('GroupName', 'MemberUpn', 'RoleKind', 'Subuser', 'Role', 'Enabled') -Rows $membershipExportRows.ToArray()

    $subuserNameById = New-Object 'System.Collections.Generic.Dictionary[int, string]'
    foreach ($subuserInfo in $subuserByToken.Values) {
        $subuserNameById[[int]$subuserInfo.Id] = [string]$subuserInfo.Name
    }

    $roleRows = New-Object System.Collections.Generic.List[object]
    foreach ($state in $teammateStates) {
        $baseRow = [ordered]@{
            Teammate    = $state.LookupKey
            EntraUpn    = [string]$state.EntraState.Upn
            EntraStatus = [string]$state.EntraState.Status
        }

        if ($state.Removed) {
            [void]$roleRows.Add([pscustomobject]($baseRow + [ordered]@{ AccessScope = '(removed this run)'; Subuser = ''; Role = ''; ScopeCount = 0 }))
            continue
        }

        if ($state.IsAdmin) {
            [void]$roleRows.Add([pscustomobject]($baseRow + [ordered]@{ AccessScope = 'account-admin'; Subuser = ''; Role = 'admin'; ScopeCount = 0 }))
            continue
        }

        try {
            $access = Get-SendGridTeammateSubuserAccess -Client $sendGridClient -TeammateName $state.LookupKey
        } catch {
            Add-SyncWarning "Could not read access for teammate '$($state.LookupKey)' during export: $(Get-SendGridErrorMessage -ErrorRecord $_)"
            continue
        }

        if ($access.HasRestrictedSubuserAccess) {
            foreach ($entry in @($access.SubuserAccess)) {
                $entryId = Get-OptionalObjectProperty -InputObject $entry -Name 'id' -Default $null
                $entrySubuser = if ($null -ne $entryId -and $subuserNameById.ContainsKey([int]$entryId)) { $subuserNameById[[int]$entryId] } else { "id:$entryId" }
                $permissionType = ([string](Get-OptionalObjectProperty -InputObject $entry -Name 'permission_type' -Default '')).Trim().ToLowerInvariant()
                $entryScopes = @((Get-OptionalObjectProperty -InputObject $entry -Name 'scopes' -Default @()))
                $entryRole = if ($permissionType -eq 'admin') {
                    'admin'
                } else {
                    $resolved = Resolve-SendGridPersonaFromScopes -Scopes $entryScopes -Scope Subuser
                    if ([string]::IsNullOrWhiteSpace([string]$resolved.Persona)) { '(custom)' } else { [string]$resolved.Persona }
                }
                [void]$roleRows.Add([pscustomobject]($baseRow + [ordered]@{ AccessScope = 'subuser'; Subuser = $entrySubuser; Role = $entryRole; ScopeCount = $entryScopes.Count }))
            }
            continue
        }

        try {
            $detail = Get-SendGridTeammate -Client $sendGridClient -TeammateName $state.LookupKey
        } catch {
            Add-SyncWarning "Could not read parent scopes for teammate '$($state.LookupKey)' during export: $(Get-SendGridErrorMessage -ErrorRecord $_)"
            continue
        }

        $parentScopes = @((Get-OptionalObjectProperty -InputObject $detail -Name 'scopes' -Default @()))
        $resolved = Resolve-SendGridPersonaFromScopes -Scopes $parentScopes -Scope Parent
        $parentRole = if ([string]::IsNullOrWhiteSpace([string]$resolved.Persona)) { '(custom)' } else { [string]$resolved.Persona }
        [void]$roleRows.Add([pscustomobject]($baseRow + [ordered]@{ AccessScope = 'parent'; Subuser = ''; Role = $parentRole; ScopeCount = $parentScopes.Count }))
    }

    Export-SendGridPlanCsv -Path (Join-Path $resolvedExportFolder 'sendgrid-roles.csv') -Columns @('Teammate', 'EntraUpn', 'EntraStatus', 'AccessScope', 'Subuser', 'Role', 'ScopeCount') -Rows $roleRows.ToArray()
}

# --- Summary ------------------------------------------------------------------
Write-Host ''
if ($warnings.Count -gt 0) {
    Write-Host 'Warnings:' -ForegroundColor DarkYellow
    $warnings | ForEach-Object { Write-Host ("  - {0}" -f $_) -ForegroundColor DarkYellow }
} else {
    Write-Host 'Completed with no warnings.' -ForegroundColor Green
}

[pscustomobject][ordered]@{
    Region               = $Region
    GroupsCreated        = $createdGroupCount
    TeammatesRemoved     = $removalCount
    TeammatesCreated     = $teammateCreatedCount
    TeammatesUpdated     = $teammateUpdatedCount
    SsoGroupMembersAdded = $ssoAdditionCount
    WarningCount         = $warnings.Count
    Warnings             = $warnings.ToArray()
    ExportFolder         = $resolvedExportFolder
}
