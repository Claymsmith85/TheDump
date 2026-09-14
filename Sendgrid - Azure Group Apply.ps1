# -----------------------------------------------------------------------------
# SendGrid Azure Group Apply (consumes the plan CSVs exported by
# 'Sendgrid - Azure Group Management')
# -----------------------------------------------------------------------------
# Purpose
# - Apply a reviewed/hand-tweaked SendGridPlan_<timestamp> folder:
#     01-groups-to-create.csv           -> create Entra security groups
#     02-group-memberships.csv          -> add members to the role groups
#     03-sendgrid-teammate-removals.csv -> delete SendGrid teammates
#     04-legacy-teammate-migrations.csv -> create/patch the parent SSO teammate
#                                          with subuser_access, then delete the
#                                          legacy subuser-context teammate
#     05-sso-group-additions.csv        -> add users to the SSO access group
#   (06-warnings.csv is informational and is not consumed.)
#
# Workflow
# 1) Run 'Sendgrid - Azure Group Management' to export a plan folder.
# 2) Edit the CSVs: delete a row to veto that action; add rows for manual
#    mappings (e.g. copy a warned user into 02 with the group you choose).
# 3) Run this script with all $Apply* flags $false first (pure preview), then
#    set the flags you want to $true and run again.
#
# Safety
# - Every write category has its own $Apply* flag; all default to $false.
# - Rows are re-validated against live state, so already-applied rows are
#   skipped and the script is safe to re-run.
# - A row whose user, group, subuser, or role cannot be resolved is warned and
#   skipped; lookup failures are never treated as permission to act.
# - Migration order per member: create/patch the parent teammate FIRST; legacy
#   subuser-context teammates are deleted only after that succeeds. An existing
#   parent-scoped (non-admin) teammate is a conflict: warned, nothing applied.
#
# Requirements
# - Same layout as the management script: ..\config\LocalVariables.ps1,
#   ..\Delinea\DelineaAuth.ps1, and 'Sendgrid API Functions.ps1' beside this file.
# - Graph app permissions: Group.ReadWrite.All (create groups, add members) and
#   User.Read.All (resolve users). The management script only needs read access.
# -----------------------------------------------------------------------------

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Fill-in input block
$Region = 'Global' # Global | EU
$SsoAccessGroupName = 'cs-ea-Sendgrid-users'
$PlanFolderOverride = $null # Path to a SendGridPlan_* folder; default: newest one beside this script
$SendGridSecretIdOverride = $null # Optional int secret ID override
$GraphCertSecretIdOverride = $null # Optional int secret ID override
$ApplyEntraGroups = $false      # 01 (also creates the SSO access group if 05 needs it)
$ApplyEntraMemberships = $false # 02 + 05
$ApplySendGridRemovals = $false # 03
$ApplyLegacyMigrations = $false # 04

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

function Get-EntraGroupByDisplayNameExact {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[string]$DisplayName
	)

	$escapedName = $DisplayName.Replace("'", "''")
	$groups = @(Get-MgGroup -Filter "displayName eq '$escapedName'" -All -Property Id,DisplayName)
	if ($groups.Count -eq 0) {
		return $null
	}

	if ($groups.Count -gt 1) {
		throw "Multiple Entra groups found with displayName '$DisplayName'."
	}

	return $groups[0]
}

function Get-EntraGroupMemberUpnSet {
	# Direct user members only.
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[string]$GroupId
	)

	$set = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
	$users = @(Get-MgGroupMemberAsUser -GroupId $GroupId -All -Property Id,UserPrincipalName)

	foreach ($user in $users) {
		$upn = [string](Get-OptionalObjectProperty -InputObject $user -Name 'UserPrincipalName' -Default '')
		if (-not [string]::IsNullOrWhiteSpace($upn)) {
			[void]$set.Add($upn.Trim().ToLowerInvariant())
		}
	}

	# Comma prevents PowerShell from unrolling the HashSet into the pipeline on return.
	return ,$set
}

function Resolve-EntraUser {
	# $null on any failure; callers warn and skip (never act on a failed lookup).
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[string]$Upn
	)

	try {
		return Get-MgUser -UserId $Upn.Trim() -Property Id,UserPrincipalName,GivenName,Surname,DisplayName
	}
	catch {
		return $null
	}
}

function Import-PlanCsv {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[string]$Folder,

		[Parameter(Mandatory)]
		[string]$FileName
	)

	$path = Join-Path $Folder $FileName
	if (-not (Test-Path -LiteralPath $path)) {
		Write-Host ("  {0,-40} missing (category skipped)" -f $FileName) -ForegroundColor DarkYellow
		return @()
	}

	$rows = @(Import-Csv -LiteralPath $path)
	Write-Host ("  {0,-40} {1} row(s)" -f $FileName, $rows.Count)
	# Callers wrap in @(); returning the bare array keeps one row per element.
	return $rows
}

$warnings = New-Object System.Collections.Generic.List[string]
function Add-ApplyWarning {
	param(
		[Parameter(Mandatory)]
		[string]$Message
	)

	[void]$warnings.Add($Message)
	Write-Host ("  ! {0}" -f $Message) -ForegroundColor DarkYellow
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

$planFolder = if (-not [string]::IsNullOrWhiteSpace([string]$PlanFolderOverride)) {
	[string]$PlanFolderOverride
}
else {
	$latest = Get-ChildItem -LiteralPath $ScriptDir -Directory -Filter 'SendGridPlan_*' | Sort-Object -Property Name -Descending | Select-Object -First 1
	if ($null -eq $latest) {
		throw "No SendGridPlan_* folder found next to this script. Run 'Sendgrid - Azure Group Management' first, or set `$PlanFolderOverride."
	}
	$latest.FullName
}

if (-not (Test-Path -LiteralPath $planFolder)) {
	throw "Plan folder not found: $planFolder"
}

Write-Host ''
Write-Host ("Loading plan from {0}" -f $planFolder) -ForegroundColor Cyan
$groupRows = @(Import-PlanCsv -Folder $planFolder -FileName '01-groups-to-create.csv')
$membershipRows = @(Import-PlanCsv -Folder $planFolder -FileName '02-group-memberships.csv')
$removalRows = @(Import-PlanCsv -Folder $planFolder -FileName '03-sendgrid-teammate-removals.csv')
$migrationRows = @(Import-PlanCsv -Folder $planFolder -FileName '04-legacy-teammate-migrations.csv')
$ssoRows = @(Import-PlanCsv -Folder $planFolder -FileName '05-sso-group-additions.csv')

Write-Host ''
Write-Host 'Apply flags ($false = preview only):' -ForegroundColor Cyan
Write-Host ("  ApplyEntraGroups      : {0}" -f $ApplyEntraGroups)
Write-Host ("  ApplyEntraMemberships : {0}" -f $ApplyEntraMemberships)
Write-Host ("  ApplySendGridRemovals : {0}" -f $ApplySendGridRemovals)
Write-Host ("  ApplyLegacyMigrations : {0}" -f $ApplyLegacyMigrations)

$sendGridSecretId = if ($null -ne $SendGridSecretIdOverride) { [int]$SendGridSecretIdOverride } else { [int]$SendGridSecretId }
$graphSecretId = if ($null -ne $GraphCertSecretIdOverride) { [int]$GraphCertSecretIdOverride } else { [int]$SsSecretId }

Write-Host ''
Write-Host 'Preparing SendGrid + Graph clients...' -ForegroundColor Cyan

$ssToken = Get-DelineaToken -SsBase $SsBase -ProbeSecretId $sendGridSecretId -Comment 'SendGrid group plan apply.'
$sendGridClient = New-SendGridClientFromDelinea -SsBase $SsBase -SecretId $sendGridSecretId -SsToken $ssToken -Region $Region -DefaultPageSize 500
$pfx = Get-SecretServerPfxObject -SecretId $graphSecretId -SsBase $SsBase -SsToken $ssToken

Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
Connect-MgGraph -TenantId $TenantId -ClientId $EntraClientId -Certificate $pfx.Cert -NoWelcome | Out-Null

# --- 1) Entra groups ---------------------------------------------------------
Write-Host ''
Write-Host '1) Entra groups to create' -ForegroundColor Cyan

$groupIdByName = New-Object 'System.Collections.Generic.Dictionary[string, string]' ([System.StringComparer]::OrdinalIgnoreCase)
$createdGroups = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

function Add-EntraGroupIfMissing {
	param(
		[Parameter(Mandatory)]
		[string]$DisplayName,

		[Parameter(Mandatory)]
		[bool]$Apply
	)

	try {
		$existingGroup = Get-EntraGroupByDisplayNameExact -DisplayName $DisplayName
	}
	catch {
		Add-ApplyWarning "Could not look up group '$DisplayName': $($_.Exception.Message)"
		return
	}

	if ($null -ne $existingGroup) {
		$groupIdByName[$DisplayName] = [string]$existingGroup.Id
		Write-Host ("  Exists: {0}" -f $DisplayName) -ForegroundColor DarkGray
		return
	}

	if (-not $Apply) {
		Write-Host ("  [WhatIf] Would create Entra group: {0}" -f $DisplayName) -ForegroundColor Yellow
		return
	}

	# mailNickname: required, no spaces, limited charset, max 64 chars.
	$nickname = ($DisplayName -replace '[^A-Za-z0-9-]', '')
	if ($nickname.Length -gt 64) { $nickname = $nickname.Substring(0, 64) }
	if ([string]::IsNullOrWhiteSpace($nickname)) { $nickname = 'sendgrid-group' }

	try {
		$newGroup = New-MgGroup -DisplayName $DisplayName -MailEnabled:$false -MailNickname $nickname -SecurityEnabled:$true -Description 'SendGrid access group (managed by SendGrid group sync).'
		$groupIdByName[$DisplayName] = [string]$newGroup.Id
		[void]$createdGroups.Add($DisplayName)
		Write-Host ("  Created Entra group: {0}" -f $DisplayName) -ForegroundColor Green
	}
	catch {
		Add-ApplyWarning "Failed to create group '$DisplayName': $($_.Exception.Message)"
	}
}

$desiredGroupNames = New-Object System.Collections.Generic.List[string]
$desiredGroupSeen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($row in $groupRows) {
	$name = ([string](Get-OptionalObjectProperty -InputObject $row -Name 'GroupName' -Default '')).Trim()
	if (-not [string]::IsNullOrWhiteSpace($name) -and $desiredGroupSeen.Add($name)) {
		[void]$desiredGroupNames.Add($name)
	}
}

# SSO access group(s) referenced by 05 are created here too, if missing.
foreach ($row in $ssoRows) {
	$target = ([string](Get-OptionalObjectProperty -InputObject $row -Name 'TargetGroup' -Default '')).Trim()
	if ([string]::IsNullOrWhiteSpace($target)) { $target = $SsoAccessGroupName }
	if ($desiredGroupSeen.Add($target)) {
		[void]$desiredGroupNames.Add($target)
	}
}

if ($desiredGroupNames.Count -eq 0) {
	Write-Host '  Nothing to do.' -ForegroundColor Green
}
else {
	foreach ($name in $desiredGroupNames) {
		Add-EntraGroupIfMissing -DisplayName $name -Apply $ApplyEntraGroups
	}
}

# --- 2) Entra group memberships (02 + 05) ------------------------------------
Write-Host ''
Write-Host '2) Entra group memberships' -ForegroundColor Cyan

$membersByGroup = New-Object 'System.Collections.Generic.Dictionary[string, object]' ([System.StringComparer]::OrdinalIgnoreCase)
$pairSeen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

function Add-PlannedMember {
	param(
		[string]$GroupName,
		[string]$Upn
	)

	$groupTrimmed = ([string]$GroupName).Trim()
	$upnTrimmed = ([string]$Upn).Trim().ToLowerInvariant()
	if ([string]::IsNullOrWhiteSpace($groupTrimmed) -or [string]::IsNullOrWhiteSpace($upnTrimmed)) {
		return
	}

	if (-not $pairSeen.Add("$groupTrimmed|$upnTrimmed")) {
		return
	}

	if (-not $membersByGroup.ContainsKey($groupTrimmed)) {
		$membersByGroup[$groupTrimmed] = New-Object System.Collections.Generic.List[string]
	}

	[void]$membersByGroup[$groupTrimmed].Add($upnTrimmed)
}

foreach ($row in $membershipRows) {
	Add-PlannedMember -GroupName ([string](Get-OptionalObjectProperty -InputObject $row -Name 'GroupName' -Default '')) -Upn ([string](Get-OptionalObjectProperty -InputObject $row -Name 'MemberUpn' -Default ''))
}

foreach ($row in $ssoRows) {
	$target = ([string](Get-OptionalObjectProperty -InputObject $row -Name 'TargetGroup' -Default '')).Trim()
	if ([string]::IsNullOrWhiteSpace($target)) { $target = $SsoAccessGroupName }
	Add-PlannedMember -GroupName $target -Upn ([string](Get-OptionalObjectProperty -InputObject $row -Name 'MemberUpn' -Default ''))
}

if ($membersByGroup.Count -eq 0) {
	Write-Host '  Nothing to do.' -ForegroundColor Green
}
else {
	foreach ($groupName in @($membersByGroup.Keys | Sort-Object)) {
		$plannedUpns = $membersByGroup[$groupName]

		$groupId = $null
		if ($groupIdByName.ContainsKey($groupName)) {
			$groupId = $groupIdByName[$groupName]
		}
		else {
			try {
				$liveGroup = Get-EntraGroupByDisplayNameExact -DisplayName $groupName
				if ($null -ne $liveGroup) { $groupId = [string]$liveGroup.Id }
			}
			catch {
				Add-ApplyWarning "Could not look up group '$groupName': $($_.Exception.Message)"
				continue
			}
		}

		if ($null -eq $groupId) {
			if (-not $ApplyEntraMemberships) {
				foreach ($upn in $plannedUpns) {
					Write-Host ("  [WhatIf] Would add {0} to {1} (group pending creation)" -f $upn, $groupName) -ForegroundColor Yellow
				}
			}
			else {
				Add-ApplyWarning "Group '$groupName' does not exist; $($plannedUpns.Count) member add(s) skipped. Run with `$ApplyEntraGroups = `$true first."
			}
			continue
		}

		$currentUpns = if ($createdGroups.Contains($groupName)) {
			New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
		}
		else {
			Get-EntraGroupMemberUpnSet -GroupId $groupId
		}

		foreach ($upn in $plannedUpns) {
			if ($currentUpns.Contains($upn)) {
				Write-Host ("  Already member: {0} in {1}" -f $upn, $groupName) -ForegroundColor DarkGray
				continue
			}

			if (-not $ApplyEntraMemberships) {
				Write-Host ("  [WhatIf] Would add {0} to {1}" -f $upn, $groupName) -ForegroundColor Yellow
				continue
			}

			$user = Resolve-EntraUser -Upn $upn
			if ($null -eq $user) {
				Add-ApplyWarning "User '$upn' not found in Entra; not added to '$groupName'."
				continue
			}

			try {
				New-MgGroupMember -GroupId $groupId -DirectoryObjectId ([string]$user.Id)
				Write-Host ("  Added {0} to {1}" -f $upn, $groupName) -ForegroundColor Green
			}
			catch {
				Add-ApplyWarning "Failed to add '$upn' to '$groupName': $($_.Exception.Message)"
			}
		}
	}
}

# --- 3) SendGrid teammate removals -------------------------------------------
Write-Host ''
Write-Host '3) SendGrid teammate removals' -ForegroundColor Cyan

if ($removalRows.Count -eq 0) {
	Write-Host '  Nothing to do.' -ForegroundColor Green
}
else {
	foreach ($row in $removalRows) {
		$teammate = ([string](Get-OptionalObjectProperty -InputObject $row -Name 'Teammate' -Default '')).Trim()
		$scope = ([string](Get-OptionalObjectProperty -InputObject $row -Name 'Scope' -Default '')).Trim().ToLowerInvariant()
		$subuser = ([string](Get-OptionalObjectProperty -InputObject $row -Name 'Subuser' -Default '')).Trim()
		$reason = ([string](Get-OptionalObjectProperty -InputObject $row -Name 'Reason' -Default '')).Trim()

		if ([string]::IsNullOrWhiteSpace($teammate)) {
			Add-ApplyWarning 'Removal row with empty Teammate skipped.'
			continue
		}

		if ($scope -eq 'subuser' -and [string]::IsNullOrWhiteSpace($subuser)) {
			Add-ApplyWarning "Removal row for '$teammate' has Scope=subuser but no Subuser; skipped."
			continue
		}

		$where = if ($scope -eq 'subuser') { "subuser $subuser" } else { 'parent' }

		if (-not $ApplySendGridRemovals) {
			Write-Host ("  [WhatIf] Would remove SendGrid teammate: {0} on {1} ({2})" -f $teammate, $where, $reason) -ForegroundColor Yellow
			continue
		}

		try {
			if ($scope -eq 'subuser') {
				Remove-SendGridTeammate -Client $sendGridClient -TeammateName $teammate -OnBehalfOf $subuser -Confirm:$false
			}
			else {
				Remove-SendGridTeammate -Client $sendGridClient -TeammateName $teammate -Confirm:$false
			}

			Write-Host ("  Removed teammate: {0} on {1}" -f $teammate, $where) -ForegroundColor Green
		}
		catch {
			Add-ApplyWarning "Failed to remove teammate '$teammate' on ${where}: $(Get-SendGridErrorMessage -ErrorRecord $_)"
		}
	}
}

# --- 4) Legacy teammate migrations -------------------------------------------
Write-Host ''
Write-Host '4) Legacy teammate migrations (subuser-context -> parent subuser_access)' -ForegroundColor Cyan

$migrationsByMember = New-Object 'System.Collections.Generic.Dictionary[string, object]' ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($row in $migrationRows) {
	$member = ([string](Get-OptionalObjectProperty -InputObject $row -Name 'Member' -Default '')).Trim().ToLowerInvariant()
	$mappedGroup = ([string](Get-OptionalObjectProperty -InputObject $row -Name 'MappedGroup' -Default '')).Trim()

	if ([string]::IsNullOrWhiteSpace($member)) {
		Add-ApplyWarning "Migration row with empty Member skipped (teammate '$([string](Get-OptionalObjectProperty -InputObject $row -Name 'Teammate' -Default ''))')."
		continue
	}

	if ([string]::IsNullOrWhiteSpace($mappedGroup) -or $mappedGroup -eq '(manual)') {
		Add-ApplyWarning "Migration for '$member' on subuser '$([string](Get-OptionalObjectProperty -InputObject $row -Name 'Subuser' -Default ''))' is unmapped; edit MappedGroup in the CSV to include it."
		continue
	}

	if (-not $migrationsByMember.ContainsKey($member)) {
		$migrationsByMember[$member] = New-Object System.Collections.Generic.List[object]
	}

	[void]$migrationsByMember[$member].Add($row)
}

if ($migrationsByMember.Count -eq 0) {
	Write-Host '  Nothing to do.' -ForegroundColor Green
}
else {
	$subuserIdByName = New-Object 'System.Collections.Generic.Dictionary[string, int]' ([System.StringComparer]::OrdinalIgnoreCase)
	foreach ($subuser in @(Get-SendGridSubusers -Client $sendGridClient -PageSize 500)) {
		$subName = [string](Get-OptionalObjectProperty -InputObject $subuser -Name 'username' -Default '')
		$subId = Get-OptionalObjectProperty -InputObject $subuser -Name 'id' -Default $null
		if (-not [string]::IsNullOrWhiteSpace($subName) -and $null -ne $subId) {
			$subuserIdByName[$subName.Trim()] = [int]$subId
		}
	}

	$parentTeammateByKey = New-Object 'System.Collections.Generic.Dictionary[string, object]' ([System.StringComparer]::OrdinalIgnoreCase)
	foreach ($teammate in @(Get-SendGridTeammates -Client $sendGridClient -PageSize 500)) {
		foreach ($keyName in @('email', 'username')) {
			$key = ([string](Get-OptionalObjectProperty -InputObject $teammate -Name $keyName -Default '')).Trim().ToLowerInvariant()
			if (-not [string]::IsNullOrWhiteSpace($key) -and -not $parentTeammateByKey.ContainsKey($key)) {
				$parentTeammateByKey[$key] = $teammate
			}
		}
	}

	foreach ($member in @($migrationsByMember.Keys | Sort-Object)) {
		$memberRows = $migrationsByMember[$member]
		$entries = New-Object System.Collections.Generic.List[object]
		$legacyToDelete = New-Object System.Collections.Generic.List[object]

		foreach ($row in $memberRows) {
			$subName = ([string](Get-OptionalObjectProperty -InputObject $row -Name 'Subuser' -Default '')).Trim()
			$mappedGroup = ([string](Get-OptionalObjectProperty -InputObject $row -Name 'MappedGroup' -Default '')).Trim().ToLowerInvariant()

			if (-not $subuserIdByName.ContainsKey($subName)) {
				Add-ApplyWarning "Migration for '$member': subuser '$subName' not found in SendGrid; row skipped."
				continue
			}

			if ($mappedGroup -notmatch '-(?<Role>admin|accountant|developer|marketer|observer)$') {
				Add-ApplyWarning "Migration for '$member': cannot read a role from MappedGroup '$mappedGroup'; row skipped."
				continue
			}

			$role = $Matches['Role']
			$entry = if ($role -eq 'admin') {
				@{ id = $subuserIdByName[$subName]; permission_type = 'admin' }
			}
			else {
				@{ id = $subuserIdByName[$subName]; permission_type = 'restricted'; scopes = @(Get-SendGridPersonaScopes -Persona $role -Scope Subuser) }
			}

			[void]$entries.Add([pscustomobject]@{ Subuser = $subName; Role = $role; Payload = $entry })
			[void]$legacyToDelete.Add($row)
		}

		if ($entries.Count -eq 0) {
			continue
		}

		$accessSummary = (@($entries | ForEach-Object { '{0} ({1})' -f $_.Subuser, $_.Role }) -join ', ')

		if (-not $ApplyLegacyMigrations) {
			Write-Host ("  [WhatIf] {0}: would ensure parent subuser_access for {1}, then delete {2} legacy teammate(s)." -f $member, $accessSummary, $legacyToDelete.Count) -ForegroundColor Yellow
			continue
		}

		$existingTeammate = if ($parentTeammateByKey.ContainsKey($member)) { $parentTeammateByKey[$member] } else { $null }
		$parentReady = $false

		try {
			if ($null -ne $existingTeammate -and [bool](Get-OptionalObjectProperty -InputObject $existingTeammate -Name 'is_admin' -Default $false)) {
				Write-Host ("  {0}: parent teammate is admin; already covers {1}." -f $member, $accessSummary) -ForegroundColor Green
				$parentReady = $true
			}
			elseif ($null -ne $existingTeammate) {
				$existingUsername = ([string](Get-OptionalObjectProperty -InputObject $existingTeammate -Name 'username' -Default '')).Trim()
				if ([string]::IsNullOrWhiteSpace($existingUsername)) {
					Add-ApplyWarning "Migration for '$member': existing parent teammate has no username; skipped."
				}
				else {
					$access = Get-SendGridTeammateSubuserAccess -Client $sendGridClient -TeammateName $existingUsername
					if (-not $access.HasRestrictedSubuserAccess) {
						Add-ApplyWarning "Migration for '$member': existing parent teammate is parent-scoped (scopes); cannot add subuser_access. Legacy teammates left in place."
					}
					else {
						$existingIds = New-Object 'System.Collections.Generic.HashSet[int]'
						foreach ($existingEntry in @($access.SubuserAccess)) {
							$existingId = Get-OptionalObjectProperty -InputObject $existingEntry -Name 'id' -Default $null
							if ($null -ne $existingId) { [void]$existingIds.Add([int]$existingId) }
						}

						$newPayloads = @($entries | Where-Object { -not $existingIds.Contains([int]$_.Payload['id']) } | ForEach-Object { $_.Payload })
						if ($newPayloads.Count -eq 0) {
							Write-Host ("  {0}: parent teammate already has subuser_access for {1}." -f $member, $accessSummary) -ForegroundColor Green
						}
						else {
							$fullAccess = @($access.SubuserAccess) + $newPayloads
							Set-SendGridTeammateSubuserAccess -Client $sendGridClient -TeammateName $existingUsername -SubuserAccess $fullAccess | Out-Null
							Write-Host ("  {0}: patched parent teammate with {1} new subuser_access entr(y/ies)." -f $member, $newPayloads.Count) -ForegroundColor Green
						}
						$parentReady = $true
					}
				}
			}
			else {
				$user = Resolve-EntraUser -Upn $member
				if ($null -eq $user) {
					Add-ApplyWarning "Migration for '$member': user not found in Entra; skipped."
				}
				else {
					$firstName = ([string](Get-OptionalObjectProperty -InputObject $user -Name 'GivenName' -Default '')).Trim()
					$lastName = ([string](Get-OptionalObjectProperty -InputObject $user -Name 'Surname' -Default '')).Trim()
					if ([string]::IsNullOrWhiteSpace($firstName) -or [string]::IsNullOrWhiteSpace($lastName)) {
						$displayParts = @((([string](Get-OptionalObjectProperty -InputObject $user -Name 'DisplayName' -Default '')).Trim() -split '\s+') | Where-Object { $_ })
						if ([string]::IsNullOrWhiteSpace($firstName)) { $firstName = if ($displayParts.Count -gt 0) { $displayParts[0] } else { ($member -split '@')[0] } }
						if ([string]::IsNullOrWhiteSpace($lastName)) { $lastName = if ($displayParts.Count -gt 1) { ($displayParts[1..($displayParts.Count - 1)] -join ' ') } else { 'User' } }
					}

					New-SendGridSsoTeammate -Client $sendGridClient -Email $member -FirstName $firstName -LastName $lastName -SubuserAccess @($entries | ForEach-Object { $_.Payload }) | Out-Null
					Write-Host ("  {0}: created parent SSO teammate with subuser_access for {1}." -f $member, $accessSummary) -ForegroundColor Green
					$parentReady = $true
				}
			}
		}
		catch {
			Add-ApplyWarning "Migration for '$member' failed before deleting legacy teammates: $(Get-SendGridErrorMessage -ErrorRecord $_)"
		}

		if (-not $parentReady) {
			continue
		}

		foreach ($row in $legacyToDelete) {
			$legacyName = ([string](Get-OptionalObjectProperty -InputObject $row -Name 'Teammate' -Default '')).Trim()
			$legacySubuser = ([string](Get-OptionalObjectProperty -InputObject $row -Name 'Subuser' -Default '')).Trim()
			if ([string]::IsNullOrWhiteSpace($legacyName) -or [string]::IsNullOrWhiteSpace($legacySubuser)) {
				Add-ApplyWarning "Migration for '$member': legacy row missing Teammate/Subuser; delete skipped."
				continue
			}

			try {
				Remove-SendGridTeammate -Client $sendGridClient -TeammateName $legacyName -OnBehalfOf $legacySubuser -Confirm:$false
				Write-Host ("  {0}: deleted legacy teammate '{1}' on subuser '{2}'." -f $member, $legacyName, $legacySubuser) -ForegroundColor Green
			}
			catch {
				Add-ApplyWarning "Migration for '$member': failed to delete legacy teammate '$legacyName' on subuser '$legacySubuser': $(Get-SendGridErrorMessage -ErrorRecord $_)"
			}
		}
	}
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
