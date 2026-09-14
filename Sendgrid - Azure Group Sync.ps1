# -----------------------------------------------------------------------------
# SendGrid Azure Group Sync (Phase 2/4 reconciler)
# -----------------------------------------------------------------------------
# Purpose
# - Entra groups are the source of truth; SendGrid parent-account SSO Teammates
#   are reconciled to match:
#     cs-sendgrid-admin            -> is_admin = true
#     cs-sendgrid-admin-ro         -> parent Observer scopes
#     cs-sendgrid-<subuser>-<role> -> one subuser_access entry per subuser
#       role admin                  -> permission_type = admin
#       role accountant|developer|marketer|observer -> permission_type =
#       restricted with the subuser persona template
# - Baseline: every SendGrid subuser gets its 5 role groups in Entra (plus the
#   admin / admin-ro core groups); missing groups are created empty.
# - Cleanup: teammates whose Entra user is DISABLED or DELETED are removed.
#   A Graph lookup that fails for any other reason (throttling, outage) marks
#   the user 'unknown' and the teammate is left alone — lookup failure is never
#   treated as permission to delete.
# - Optional CSV export ($ExportStateCsv): group-memberships.csv (which users
#   are in which cs-sendgrid-* group and what role that implies) and
#   sendgrid-roles.csv (what each parent teammate actually holds in SendGrid).
#
# Conflict rules (from the management spec)
# - admin > admin-ro > subuser access. Admin ignores all other groups.
# - admin-ro together with any subuser group: warn, apply nothing for that user.
# - Unknown subuser or unparseable group name: warn and continue.
# - A user in multiple role groups for the SAME subuser: admin wins; multiple
#   personas grant the union of their scope templates (warned).
# - Teammates in no role groups whose Entra user is enabled are reported as
#   unmanaged and left alone (remove/disable is an open design decision).
#
# Idempotency
# - SendGrid injects a small implicit scope set on teammates; those scopes
#   ($SendGridImplicitScopes) are ignored when comparing, so a compliant
#   teammate is not re-patched every run.
#
# Safety
# - Every write category is gated behind its own $Apply* flag (default $false =
#   preview). $ProtectedTeammates lists emails never modified or removed.
#
# Requirements
# - Same layout as the management script: ..\config\LocalVariables.ps1,
#   ..\Delinea\DelineaAuth.ps1, and 'Sendgrid API Functions.ps1' beside this file.
# - Graph app permissions: Group.ReadWrite.All, User.Read.All.
# -----------------------------------------------------------------------------

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Fill-in input block
$Region = 'Global' # Global | EU
$GroupPrefix = 'cs-sendgrid-'
$SsoAccessGroupName = 'cs-ea-Sendgrid-users'
$SendGridSecretIdOverride = $null # Optional int secret ID override
$GraphCertSecretIdOverride = $null # Optional int secret ID override
$ProtectedTeammates = @() # Teammate emails this script must never modify or remove
$ApplyEntraGroups = $false      # Create missing baseline role groups (and the SSO access group)
$ApplySsoAccessGroup = $false   # Add role-group members to the SSO access group
$ApplyTeammateChanges = $false  # Create/update SendGrid teammates from group membership
$ApplyTeammateRemovals = $false # Remove teammates whose Entra user is disabled or deleted
$ExportStateCsv = $true # Export group memberships + current SendGrid roles as CSV files
$ExportFolderOverride = $null # Optional export folder; default: <script folder>\SendGridSync_<timestamp>

# Scopes SendGrid adds on its own; ignored when comparing desired vs current.
$SendGridImplicitScopes = @(
	'2fa_exempt',
	'2fa_required',
	'sender_verification_eligible',
	'sender_verification_legacy',
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

# Resolve this script's folder whether run as a file (F5) or via "Run Selection" (F8)
$ScriptDir = if ($PSScriptRoot) {
	$PSScriptRoot
} elseif (Get-Variable -Name psEditor -ErrorAction SilentlyContinue) {
	Split-Path $psEditor.GetEditorContext().CurrentFile.Path
} else {
	$PWD.Path
}

. "$ScriptDir\..\config\LocalVariables.ps1"
. "$ScriptDir\..\Delinea\DelineaAuth.ps1"
. "$ScriptDir\Sendgrid API Functions.ps1"

Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
Import-Module Microsoft.Graph.Groups -ErrorAction Stop
Import-Module Microsoft.Graph.Users -ErrorAction Stop

function Get-OptionalObjectProperty {
	[CmdletBinding()]
	param(
		[object]$InputObject,
		[string]$Name,
		$Default = $null
	)

	if ($null -eq $InputObject) {
		return $Default
	}

	$prop = $InputObject.PSObject.Properties[$Name]
	if ($null -ne $prop) {
		return $prop.Value
	}

	return $Default
}

$warnings = New-Object System.Collections.Generic.List[string]
function Add-SyncWarning {
	param(
		[Parameter(Mandatory)]
		[string]$Message
	)

	[void]$warnings.Add($Message)
	Write-Host ("  ! {0}" -f $Message) -ForegroundColor DarkYellow
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

	return ,$set
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
	# Hardened: only a genuine 404 counts as not-found; any other Graph failure
	# returns Status 'unknown' so callers never act destructively on flaky lookups.
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[string]$Email
	)

	$normalized = $Email.Trim().ToLowerInvariant()
	$result = [ordered]@{
		Status    = 'unknown'
		Upn       = $null
		UserId    = $null
		GivenName = ''
		Surname   = ''
		DisplayName = ''
		Detail    = ''
	}

	$user = $null
	$notFoundById = $false
	try {
		$user = Get-MgUser -UserId $normalized -Property Id,UserPrincipalName,Mail,AccountEnabled,GivenName,Surname,DisplayName
	}
	catch {
		if ($_.Exception.Message -match 'Request_ResourceNotFound|ResourceNotFound|does not exist') {
			$notFoundById = $true
		}
		else {
			$result.Detail = $_.Exception.Message
			return [pscustomobject]$result
		}
	}

	if ($null -eq $user -and $notFoundById) {
		$escaped = $normalized.Replace("'", "''")
		try {
			$candidates = @(Get-MgUser -Filter "mail eq '$escaped' or userPrincipalName eq '$escaped'" -All -Property Id,UserPrincipalName,Mail,AccountEnabled,GivenName,Surname,DisplayName)
		}
		catch {
			$result.Detail = $_.Exception.Message
			return [pscustomobject]$result
		}

		if ($candidates.Count -gt 0) {
			$exact = $candidates | Where-Object { (([string]$_.UserPrincipalName).Trim().ToLowerInvariant() -eq $normalized) -or (([string]$_.Mail).Trim().ToLowerInvariant() -eq $normalized) } | Select-Object -First 1
			$user = if ($exact) { $exact } else { $candidates[0] }
		}
	}

	if ($null -eq $user) {
		$result.Status = 'not-found'
		return [pscustomobject]$result
	}

	$result.Status = if ([bool](Get-OptionalObjectProperty -InputObject $user -Name 'AccountEnabled' -Default $true)) { 'enabled' } else { 'disabled' }
	$result.Upn = ([string](Get-OptionalObjectProperty -InputObject $user -Name 'UserPrincipalName' -Default $normalized)).Trim().ToLowerInvariant()
	$result.UserId = [string](Get-OptionalObjectProperty -InputObject $user -Name 'Id' -Default '')
	$result.GivenName = [string](Get-OptionalObjectProperty -InputObject $user -Name 'GivenName' -Default '')
	$result.Surname = [string](Get-OptionalObjectProperty -InputObject $user -Name 'Surname' -Default '')
	$result.DisplayName = [string](Get-OptionalObjectProperty -InputObject $user -Name 'DisplayName' -Default '')
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

	$firstName = ([string](Get-OptionalObjectProperty -InputObject $UserRecord -Name 'GivenName' -Default '')).Trim()
	$lastName = ([string](Get-OptionalObjectProperty -InputObject $UserRecord -Name 'Surname' -Default '')).Trim()
	if ([string]::IsNullOrWhiteSpace($firstName) -or [string]::IsNullOrWhiteSpace($lastName)) {
		$displayParts = @((([string](Get-OptionalObjectProperty -InputObject $UserRecord -Name 'DisplayName' -Default '')).Trim() -split '\s+') | Where-Object { $_ })
		if ([string]::IsNullOrWhiteSpace($firstName)) { $firstName = if ($displayParts.Count -gt 0) { $displayParts[0] } else { ($Upn -split '@')[0] } }
		if ([string]::IsNullOrWhiteSpace($lastName)) { $lastName = if ($displayParts.Count -gt 1) { ($displayParts[1..($displayParts.Count - 1)] -join ' ') } else { 'User' } }
	}

	return [pscustomobject]@{ FirstName = $firstName; LastName = $lastName }
}

if ($Region -notin @('Global', 'EU')) {
	throw "Invalid Region '$Region'. Valid values: Global, EU."
}

foreach ($required in @('TenantId', 'EntraClientId', 'SsBase', 'SendGridSecretId', 'SsSecretId')) {
	if (-not (Get-Variable -Name $required -Scope Script -ErrorAction SilentlyContinue) -or
		[string]::IsNullOrWhiteSpace([string](Get-Variable -Name $required -Scope Script).Value)) {
		throw "Required config variable '$required' is missing or empty. Check config/LocalVariables.ps1."
	}
}

if (-not (Get-Command -Name 'Invoke-SendGridRequest' -ErrorAction SilentlyContinue)) {
	throw "Invoke-SendGridRequest was not found. Dot-source 'Sendgrid API Functions.ps1' before running."
}

$protectedSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($protectedEmail in @($ProtectedTeammates)) {
	if (-not [string]::IsNullOrWhiteSpace([string]$protectedEmail)) {
		[void]$protectedSet.Add(([string]$protectedEmail).Trim().ToLowerInvariant())
	}
}

$sendGridSecretId = if ($null -ne $SendGridSecretIdOverride) { [int]$SendGridSecretIdOverride } else { [int]$SendGridSecretId }
$graphSecretId = if ($null -ne $GraphCertSecretIdOverride) { [int]$GraphCertSecretIdOverride } else { [int]$SsSecretId }

Write-Host ''
Write-Host 'Apply flags ($false = preview only):' -ForegroundColor Cyan
Write-Host ("  ApplyEntraGroups      : {0}" -f $ApplyEntraGroups)
Write-Host ("  ApplySsoAccessGroup   : {0}" -f $ApplySsoAccessGroup)
Write-Host ("  ApplyTeammateChanges  : {0}" -f $ApplyTeammateChanges)
Write-Host ("  ApplyTeammateRemovals : {0}" -f $ApplyTeammateRemovals)

Write-Host ''
Write-Host 'Preparing SendGrid + Graph clients...' -ForegroundColor Cyan

$ssToken = Get-DelineaToken -SsBase $SsBase -ProbeSecretId $sendGridSecretId -Comment 'SendGrid group-driven teammate sync.'
$sendGridClient = New-SendGridClientFromDelinea -SsBase $SsBase -SecretId $sendGridSecretId -SsToken $ssToken -Region $Region -DefaultPageSize 500
$pfx = Get-SecretServerPfxObject -SecretId $graphSecretId -SsBase $SsBase -SsToken $ssToken

Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
Connect-MgGraph -TenantId $TenantId -ClientId $EntraClientId -Certificate $pfx.Cert -NoWelcome | Out-Null

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

$prefixGroups = @(Get-MgGroup -Filter "startswith(displayName,'$GroupPrefix')" -All -Property Id,DisplayName)
$groupIdByName = New-Object 'System.Collections.Generic.Dictionary[string, string]' ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($g in $prefixGroups) {
	$gName = ([string](Get-OptionalObjectProperty -InputObject $g -Name 'DisplayName' -Default '')).Trim()
	if (-not [string]::IsNullOrWhiteSpace($gName)) {
		$groupIdByName[$gName] = [string]$g.Id
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
		$newGroup = New-MgGroup -DisplayName $requiredName -MailEnabled:$false -MailNickname $nickname -SecurityEnabled:$true -Description 'SendGrid access group (managed by SendGrid group sync).'
		$groupIdByName[$requiredName] = [string]$newGroup.Id
		$createdGroupCount++
		Write-Host ("  Created Entra group: {0}" -f $requiredName) -ForegroundColor Green
	}
	catch {
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
# upn -> Entra user record from group membership (Id, names, AccountEnabled)
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

	$members = @(Get-MgGroupMemberAsUser -GroupId $groupIdByName[$groupName] -All -Property Id,UserPrincipalName,AccountEnabled,GivenName,Surname,DisplayName)
	foreach ($member in $members) {
		$upn = ([string](Get-OptionalObjectProperty -InputObject $member -Name 'UserPrincipalName' -Default '')).Trim().ToLowerInvariant()
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
			Enabled   = [bool](Get-OptionalObjectProperty -InputObject $member -Name 'AccountEnabled' -Default $true)
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
	}
	else {
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
	}
	elseif (-not [string]::IsNullOrWhiteSpace($email)) {
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
		Write-Host ("  Protected: {0} (Entra status: {1}) — left alone." -f $state.LookupKey, $status) -ForegroundColor DarkGray
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
	}
	catch {
		Add-SyncWarning "Failed to remove teammate '$($state.LookupKey)': $(Get-SendGridErrorMessage -ErrorRecord $_)"
	}
}

if ($removalCount -eq 0 -and $ApplyTeammateRemovals) {
	Write-Host '  Nothing removed.' -ForegroundColor Green
}

# --- 5) Reconcile teammates with group-derived roles -------------------------
Write-Host ''
Write-Host '5) Reconcile teammates (create/update from groups)' -ForegroundColor Cyan

foreach ($upn in @($effectiveByUpn.Keys | Sort-Object)) {
	$effective = $effectiveByUpn[$upn]
	if ($effective.Kind -eq 'conflict') {
		continue
	}

	if ($protectedSet.Contains($upn)) {
		Write-Host ("  Protected: {0} — left alone." -f $upn) -ForegroundColor DarkGray
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

		if (-not [bool](Get-OptionalObjectProperty -InputObject $userRecord -Name 'AccountEnabled' -Default $true)) {
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
			Write-Host ("  Created SSO teammate {0}: {1}" -f $upn, $effective.Summary) -ForegroundColor Green
		}
		catch {
			$createError = Get-SendGridErrorMessage -ErrorRecord $_
			if ($createError -match 'username exists') {
				Add-SyncWarning "Failed to create teammate '$upn': $createError. The username is likely held by a legacy subuser-context teammate — run the legacy migration in 'Sendgrid - Azure Group Apply.ps1' first."
			}
			else {
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
			}
			else {
				Set-SendGridTeammateSubuserAccess -Client $sendGridClient -TeammateName $teammateName -SubuserAccess $effective.Entries | Out-Null
			}
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
	}
	catch {
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

$ssoGroups = @(Get-MgGroup -Filter "displayName eq '$($SsoAccessGroupName.Replace("'", "''"))'" -All -Property Id,DisplayName)
$ssoGroupId = if ($ssoGroups.Count -gt 0) { [string]$ssoGroups[0].Id } else { $null }

if ($null -eq $ssoGroupId) {
	if ($ApplyEntraGroups) {
		$nickname = ($SsoAccessGroupName -replace '[^A-Za-z0-9-]', '')
		try {
			$newGroup = New-MgGroup -DisplayName $SsoAccessGroupName -MailEnabled:$false -MailNickname $nickname -SecurityEnabled:$true -Description 'SendGrid SSO access group (managed by SendGrid group sync).'
			$ssoGroupId = [string]$newGroup.Id
			Write-Host ("  Created Entra group: {0}" -f $SsoAccessGroupName) -ForegroundColor Green
		}
		catch {
			Add-SyncWarning "Failed to create SSO access group '$SsoAccessGroupName': $($_.Exception.Message)"
		}
	}
	else {
		Write-Host ("  [WhatIf] Would create Entra group: {0}" -f $SsoAccessGroupName) -ForegroundColor Yellow
	}
}

if ($null -ne $ssoGroupId) {
	$ssoMemberUpns = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
	foreach ($member in @(Get-MgGroupMemberAsUser -GroupId $ssoGroupId -All -Property Id,UserPrincipalName)) {
		$memberUpn = ([string](Get-OptionalObjectProperty -InputObject $member -Name 'UserPrincipalName' -Default '')).Trim().ToLowerInvariant()
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
		$userId = [string](Get-OptionalObjectProperty -InputObject $userRecord -Name 'Id' -Default '')
		if ([string]::IsNullOrWhiteSpace($userId)) {
			Add-SyncWarning "No Entra id for '$upn'; not added to '$SsoAccessGroupName'."
			continue
		}

		try {
			New-MgGroupMember -GroupId $ssoGroupId -DirectoryObjectId $userId
			Write-Host ("  Added {0} to {1}" -f $upn, $SsoAccessGroupName) -ForegroundColor Green
		}
		catch {
			Add-SyncWarning "Failed to add '$upn' to '$SsoAccessGroupName': $($_.Exception.Message)"
		}
	}
}

# --- 7) State CSV export ------------------------------------------------------
if ($ExportStateCsv) {
	$exportFolder = if (-not [string]::IsNullOrWhiteSpace([string]$ExportFolderOverride)) {
		[string]$ExportFolderOverride
	}
	else {
		Join-Path $ScriptDir ('SendGridSync_{0}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
	}

	$null = New-Item -ItemType Directory -Path $exportFolder -Force

	Write-Host ''
	Write-Host ("7) Exporting state CSVs to {0}" -f $exportFolder) -ForegroundColor Cyan

	Export-SendGridPlanCsv -Path (Join-Path $exportFolder 'group-memberships.csv') -Columns @('GroupName', 'MemberUpn', 'RoleKind', 'Subuser', 'Role', 'Enabled') -Rows $membershipExportRows.ToArray()

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
		}
		catch {
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
				}
				else {
					$resolved = Resolve-SendGridPersonaFromScopes -Scopes $entryScopes -Scope Subuser
					if ([string]::IsNullOrWhiteSpace([string]$resolved.Persona)) { '(custom)' } else { [string]$resolved.Persona }
				}
				[void]$roleRows.Add([pscustomobject]($baseRow + [ordered]@{ AccessScope = 'subuser'; Subuser = $entrySubuser; Role = $entryRole; ScopeCount = $entryScopes.Count }))
			}
			continue
		}

		try {
			$detail = Get-SendGridTeammate -Client $sendGridClient -TeammateName $state.LookupKey
		}
		catch {
			Add-SyncWarning "Could not read parent scopes for teammate '$($state.LookupKey)' during export: $(Get-SendGridErrorMessage -ErrorRecord $_)"
			continue
		}

		$parentScopes = @((Get-OptionalObjectProperty -InputObject $detail -Name 'scopes' -Default @()))
		$resolved = Resolve-SendGridPersonaFromScopes -Scopes $parentScopes -Scope Parent
		$parentRole = if ([string]::IsNullOrWhiteSpace([string]$resolved.Persona)) { '(custom)' } else { [string]$resolved.Persona }
		[void]$roleRows.Add([pscustomobject]($baseRow + [ordered]@{ AccessScope = 'parent'; Subuser = ''; Role = $parentRole; ScopeCount = $parentScopes.Count }))
	}

	Export-SendGridPlanCsv -Path (Join-Path $exportFolder 'sendgrid-roles.csv') -Columns @('Teammate', 'EntraUpn', 'EntraStatus', 'AccessScope', 'Subuser', 'Role', 'ScopeCount') -Rows $roleRows.ToArray()
}

# --- Summary ------------------------------------------------------------------
Write-Host ''
if ($warnings.Count -gt 0) {
	Write-Host 'Warnings:' -ForegroundColor DarkYellow
	$warnings | ForEach-Object { Write-Host ("  - {0}" -f $_) -ForegroundColor DarkYellow }
}
else {
	Write-Host 'Completed with no warnings.' -ForegroundColor Green
}

Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
