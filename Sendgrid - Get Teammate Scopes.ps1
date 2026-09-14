# -----------------------------------------------------------------------------
# SendGrid teammate scope inspector (read-only diagnostic)
# -----------------------------------------------------------------------------
# Prints exactly what a teammate holds on one subuser (or on the parent, if
# they are parent-scoped): permission_type, the full scope list, whether the
# Marketing-tab pair is present, and which persona template the scopes match.
# -----------------------------------------------------------------------------

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Fill-in input block
$TeammateName = 'bobbisue.morrison@corespecialty.com' # Teammate username/email
$SubuserName = 'Internal Email Communications'        # Subuser to inspect
$Region = 'Global' # Global | EU
$SendGridSecretIdOverride = $null # Optional int secret ID override

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

$sendGridSecretId = if ($null -ne $SendGridSecretIdOverride) { [int]$SendGridSecretIdOverride } else { [int]$SendGridSecretId }
$ssToken = Get-DelineaToken -SsBase $SsBase -ProbeSecretId $sendGridSecretId -Comment 'SendGrid teammate scope inspection (read-only).'
$sendGridClient = New-SendGridClientFromDelinea -SsBase $SsBase -SecretId $sendGridSecretId -SsToken $ssToken -Region $Region -DefaultPageSize 500

function Write-ScopeReport {
	param(
		[Parameter(Mandatory)]
		[string]$Label,

		[string[]]$Scopes,

		[ValidateSet('Parent', 'Subuser')]
		[string]$TemplateScope = 'Subuser'
	)

	$sorted = @($Scopes | Sort-Object -Unique)
	Write-Host ''
	Write-Host ("{0} — {1} scope(s)" -f $Label, $sorted.Count) -ForegroundColor Cyan

	foreach ($probe in @('marketing.read', 'marketing.automation.read', 'marketing_campaigns.read', 'stats.read')) {
		$has = if ($sorted -contains $probe) { 'YES' } else { 'no ' }
		Write-Host ("  [{0}] {1}" -f $has, $probe) -ForegroundColor $(if ($sorted -contains $probe) { 'Green' } else { 'Yellow' })
	}

	$resolved = Resolve-SendGridPersonaFromScopes -Scopes $sorted -Scope $TemplateScope
	$personaText = if ([string]::IsNullOrWhiteSpace([string]$resolved.Persona)) { '(no persona match)' } else { "$($resolved.Persona) ($($resolved.MatchType))" }
	Write-Host ("  Persona match: {0}" -f $personaText)

	Write-Host '  Scopes:'
	$sorted | ForEach-Object { Write-Host ("    {0}" -f $_) }
}

Write-Host ("Teammate: {0}" -f $TeammateName) -ForegroundColor Cyan
Write-Host ("Subuser : {0}" -f $SubuserName) -ForegroundColor Cyan

$detail = Get-SendGridTeammate -Client $sendGridClient -TeammateName $TeammateName
if ([bool](Get-OptionalObjectProperty -InputObject $detail -Name 'is_admin' -Default $false)) {
	Write-Host 'Teammate is an account admin: all scopes on the parent and every subuser.' -ForegroundColor Green
	return
}

$access = Get-SendGridTeammateSubuserAccess -Client $sendGridClient -TeammateName $TeammateName

if (-not $access.HasRestrictedSubuserAccess) {
	Write-Host 'Teammate is PARENT-scoped (no restricted subuser access).' -ForegroundColor Yellow
	Write-ScopeReport -Label 'Parent scopes' -Scopes @((Get-OptionalObjectProperty -InputObject $detail -Name 'scopes' -Default @())) -TemplateScope Parent
	return
}

$subuserId = $null
foreach ($subuser in @(Get-SendGridSubusers -Client $sendGridClient -PageSize 500)) {
	if (([string](Get-OptionalObjectProperty -InputObject $subuser -Name 'username' -Default '')).Trim() -ieq $SubuserName.Trim()) {
		$subuserId = [int](Get-OptionalObjectProperty -InputObject $subuser -Name 'id' -Default 0)
		break
	}
}

if ($null -eq $subuserId) {
	throw "Subuser '$SubuserName' was not found in SendGrid."
}

Write-Host ("Subuser id: {0}; teammate has {1} subuser_access entr(y/ies) total." -f $subuserId, @($access.SubuserAccess).Count)

$entry = @($access.SubuserAccess) | Where-Object { [int](Get-OptionalObjectProperty -InputObject $_ -Name 'id' -Default -1) -eq $subuserId } | Select-Object -First 1
if ($null -eq $entry) {
	Write-Host ("Teammate has NO subuser_access entry for '{0}'. Entries exist for id(s): {1}" -f $SubuserName, (@($access.SubuserAccess | ForEach-Object { Get-OptionalObjectProperty -InputObject $_ -Name 'id' -Default '?' }) -join ', ')) -ForegroundColor Red
	return
}

$permissionType = [string](Get-OptionalObjectProperty -InputObject $entry -Name 'permission_type' -Default '')
Write-Host ("permission_type: {0}" -f $permissionType) -ForegroundColor Cyan

if ($permissionType -ieq 'admin') {
	Write-Host 'Teammate is subuser ADMIN here: all scopes on this subuser, Marketing tab included.' -ForegroundColor Green
	return
}

Write-ScopeReport -Label ("Restricted scopes on '{0}'" -f $SubuserName) -Scopes @((Get-OptionalObjectProperty -InputObject $entry -Name 'scopes' -Default @())) -TemplateScope Subuser
