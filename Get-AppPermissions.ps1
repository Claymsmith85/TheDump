#Requires -Modules Az.Accounts, Az.Resources
<#
.SYNOPSIS
    Reports an app registration's configured API permissions and the roles assigned to its
    related enterprise app (service principal).

.DESCRIPTION
    Sections reported:
      1. Configured API permissions on the app registration (requiredResourceAccess).
      2. Application permissions actually granted (admin-consented app role assignments).
      3. Delegated permissions actually granted (OAuth2 permission grants).
      4. Entra ID directory roles the service principal holds.
      5. Azure RBAC role assignments for the service principal.

.PARAMETER AppId
    Application (client) ID of the app registration.

.PARAMETER DisplayName
    Display name of the app registration. Must match exactly one app.

.PARAMETER AllSubscriptions
    Check Azure RBAC assignments across every subscription visible to the current account.
    Default is the current context subscription only.

.EXAMPLE
    .\Get-AppRegistrationPermissions.ps1 -AppId 00000000-0000-0000-0000-000000000000 -AllSubscriptions

.EXAMPLE
    .\Get-AppRegistrationPermissions.ps1 -DisplayName 'Daily Email Report'
#>
[CmdletBinding(DefaultParameterSetName = 'ByAppId')]
param(
    [Parameter(Mandatory, ParameterSetName = 'ByAppId')]
    [string]$AppId,

    [Parameter(Mandatory, ParameterSetName = 'ByName')]
    [string]$DisplayName,

    [switch]$AllSubscriptions
)

$ErrorActionPreference = 'Stop'

if (-not (Get-AzContext)) {
    Connect-AzAccount | Out-Null
}

# ---------------------------------------------------------------------------
# App registration
# ---------------------------------------------------------------------------
$app = if ($PSCmdlet.ParameterSetName -eq 'ByAppId') {
    Get-AzADApplication -ApplicationId $AppId
}
else {
    Get-AzADApplication -DisplayName $DisplayName
}

if (-not $app)          { throw "App registration not found." }
if (@($app).Count -gt 1) { throw "Multiple app registrations match '$DisplayName'. Use -AppId." }

Write-Host ''
Write-Host "App registration : $($app.DisplayName)"
Write-Host "App (client) ID  : $($app.AppId)"
Write-Host "Object ID        : $($app.Id)"

# Caches so each resource API is only looked up once
$spByAppId = @{}
$spByObjId = @{}

function Get-ResourceSpByAppId ([string]$Id) {
    if (-not $spByAppId.ContainsKey($Id)) { $spByAppId[$Id] = Get-AzADServicePrincipal -ApplicationId $Id }
    $spByAppId[$Id]
}
function Get-ResourceSpByObjectId ([string]$Id) {
    if (-not $spByObjId.ContainsKey($Id)) { $spByObjId[$Id] = Get-AzADServicePrincipal -ObjectId $Id }
    $spByObjId[$Id]
}

# ---------------------------------------------------------------------------
# 1. Configured API permissions (what is declared on the app registration)
# ---------------------------------------------------------------------------
$configured = foreach ($perm in (Get-AzADAppPermission -ObjectId $app.Id)) {
    $resourceSp = Get-ResourceSpByAppId $perm.ApiId

    $name = if ($perm.Type -eq 'Role') {
        ($resourceSp.AppRole | Where-Object Id -eq $perm.Id).Value
    }
    else {
        ($resourceSp.Oauth2PermissionScope | Where-Object Id -eq $perm.Id).Value
    }

    [pscustomobject]@{
        ResourceApp  = $resourceSp.DisplayName
        Permission   = $name
        Type         = if ($perm.Type -eq 'Role') { 'Application' } else { 'Delegated' }
        PermissionId = $perm.Id
    }
}

Write-Host ''
Write-Host '=== 1. Configured API permissions (app registration) ==='
if ($configured) { $configured | Sort-Object ResourceApp, Type, Permission | Format-Table -AutoSize }
else             { Write-Host 'None.' }

# ---------------------------------------------------------------------------
# Enterprise app (service principal)
# ---------------------------------------------------------------------------
$sp = Get-AzADServicePrincipal -ApplicationId $app.AppId
if (-not $sp) {
    Write-Warning 'No enterprise app (service principal) exists for this app registration in this tenant.'
    return
}

Write-Host "Enterprise app   : $($sp.DisplayName)"
Write-Host "SP Object ID     : $($sp.Id)"

# ---------------------------------------------------------------------------
# 2. Granted application permissions (admin-consented app role assignments)
# ---------------------------------------------------------------------------
$grantedAppRoles = foreach ($a in (Get-AzADServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id)) {
    $resourceSp = Get-ResourceSpByObjectId $a.ResourceId
    [pscustomobject]@{
        ResourceApp = $resourceSp.DisplayName
        Permission  = ($resourceSp.AppRole | Where-Object Id -eq $a.AppRoleId).Value
        AppRoleId   = $a.AppRoleId
    }
}

Write-Host ''
Write-Host '=== 2. Granted application permissions (admin consent) ==='
if ($grantedAppRoles) { $grantedAppRoles | Sort-Object ResourceApp, Permission | Format-Table -AutoSize }
else                  { Write-Host 'None.' }

# ---------------------------------------------------------------------------
# 3. Granted delegated permissions (OAuth2 permission grants) - via Graph REST
# ---------------------------------------------------------------------------
$grantUri = "https://graph.microsoft.com/v1.0/servicePrincipals/$($sp.Id)/oauth2PermissionGrants"
$grants   = (Invoke-AzRestMethod -Uri $grantUri -Method GET).Content | ConvertFrom-Json

$delegated = foreach ($g in $grants.value) {
    $resourceSp = Get-ResourceSpByObjectId $g.resourceId
    [pscustomobject]@{
        ResourceApp = $resourceSp.DisplayName
        Scopes      = $g.scope.Trim()
        ConsentType = $g.consentType     # AllPrincipals = admin consent, Principal = single user
        PrincipalId = $g.principalId
    }
}

Write-Host ''
Write-Host '=== 3. Granted delegated permissions ==='
if ($delegated) { $delegated | Format-Table -AutoSize }
else            { Write-Host 'None.' }

# ---------------------------------------------------------------------------
# 4. Entra ID directory roles held by the service principal - via Graph REST
# ---------------------------------------------------------------------------
$memberUri = "https://graph.microsoft.com/v1.0/servicePrincipals/$($sp.Id)/memberOf?`$select=id,displayName"
$memberOf  = (Invoke-AzRestMethod -Uri $memberUri -Method GET).Content | ConvertFrom-Json

$dirRoles = $memberOf.value |
    Where-Object { $_.'@odata.type' -eq '#microsoft.graph.directoryRole' } |
    Select-Object @{n = 'DirectoryRole'; e = { $_.displayName } }, @{n = 'RoleObjectId'; e = { $_.id } }

Write-Host ''
Write-Host '=== 4. Entra ID directory roles ==='
if ($dirRoles) { $dirRoles | Format-Table -AutoSize }
else           { Write-Host 'None.' }

# ---------------------------------------------------------------------------
# 5. Azure RBAC role assignments
# ---------------------------------------------------------------------------
$rbac = if ($AllSubscriptions) {
    $original = Get-AzContext
    try {
        foreach ($sub in Get-AzSubscription) {
            Set-AzContext -SubscriptionId $sub.Id | Out-Null
            Get-AzRoleAssignment -ObjectId $sp.Id |
                Select-Object @{n = 'Subscription'; e = { $sub.Name } }, RoleDefinitionName, Scope
        }
    }
    finally {
        Set-AzContext -Context $original | Out-Null
    }
}
else {
    Get-AzRoleAssignment -ObjectId $sp.Id |
        Select-Object @{n = 'Subscription'; e = { (Get-AzContext).Subscription.Name } }, RoleDefinitionName, Scope
}

Write-Host ''
Write-Host '=== 5. Azure RBAC role assignments ==='
if ($rbac) { $rbac | Sort-Object Subscription, Scope | Format-Table -AutoSize -Wrap }
else       { Write-Host 'None.' }
