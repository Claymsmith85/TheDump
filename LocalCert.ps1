<#
.SYNOPSIS
    Generates a self-signed certificate in the Local Machine store and exports the public key (.cer)
    for upload to an Entra ID app registration / enterprise application.

.DESCRIPTION
    Self-contained: no repo config, Delinea, or module dependencies. Intended for Windows Server,
    Windows PowerShell 5.1, run as Administrator (writing to Cert:\LocalMachine\My requires it).

    1. Creates an RSA 2048 / SHA256 self-signed cert valid for 365 days in Cert:\LocalMachine\My.
    2. Grants Read on the private key to the accounts in -GrantReadTo (default CORESPEC\cs-ad-jml) so
       they can use the cert (e.g. Connect-MgGraph -CertificateThumbprint) without being admins.
    3. Exports the PUBLIC key only (DER .cer) to C:\Temp. The private key never leaves the machine store.
    4. Prints the thumbprint and validity dates.

    SECURITY: anyone granted Read on the private key can authenticate as the app and use ALL of its
    application permissions. Grant only to specific service accounts / small groups.

    Upload the exported .cer in the Entra admin center:
        App registrations > <your app> > Certificates & secrets > Certificates > Upload certificate

.PARAMETER CertName
    Subject CN of the certificate (e.g. "svc-app01.contoso.com"). Also used for the export filename.

.PARAMETER ValidDays
    Validity period in days. Default 365.

.PARAMETER ExportPath
    Folder for the exported public key. Default C:\Temp (created if missing).

.PARAMETER Exportable
    Mark the private key as exportable. Off by default so the key cannot be copied off this server.

.PARAMETER GrantReadTo
    Accounts (DOMAIN\user, DOMAIN\group, or DOMAIN\gmsa$) granted Read on the private key.
    Default CORESPEC\cs-ad-jml. Pass an empty array (-GrantReadTo @()) to skip.

.EXAMPLE
    .\Generate Machine Store Self-Signed Cert.ps1 -CertName "svc-app01.contoso.com"

.EXAMPLE
    .\Generate Machine Store Self-Signed Cert.ps1 -CertName "svc-app01.contoso.com" -GrantReadTo 'CORESPEC\cs-ad-jml', 'CORESPEC\svc-other'
#>
#Requires -Version 5.1
#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$CertName,

    [ValidateRange(1, 3650)]
    [int]$ValidDays = 365,

    [string]$ExportPath = 'C:\Temp',

    [switch]$Exportable,

    [string[]]$GrantReadTo = @('CORESPEC\cs-ad-jml')
)

$ErrorActionPreference = 'Stop'
$storeLocation = 'Cert:\LocalMachine\My'

# ---- Resolve the grant accounts BEFORE creating anything, so a typo fails fast ----
$grantAccounts = @(foreach ($identity in $GrantReadTo) {
    $account = New-Object System.Security.Principal.NTAccount($identity)
    try {
        [void]$account.Translate([System.Security.Principal.SecurityIdentifier])
    } catch {
        throw "Cannot resolve '$identity' to a Windows account. Check the name, or that this server can reach the domain."
    }
    $account
})

# ---- Warn about existing certs with the same subject (rotation leaves the old one in place) ----
$existing = Get-ChildItem -Path $storeLocation | Where-Object { $_.Subject -eq "CN=$CertName" }
if ($existing) {
    Write-Warning "Found $(@($existing).Count) existing cert(s) with subject CN=$CertName in $storeLocation."
    $existing | ForEach-Object {
        Write-Warning ("  {0}  expires {1}" -f $_.Thumbprint, $_.NotAfter.ToString('yyyy-MM-dd'))
    }
    Write-Warning "A new certificate will be created alongside them."
}

# ---- 1) Create the certificate in the machine store ----
$certParams = @{
    Subject           = "CN=$CertName"
    FriendlyName      = "$CertName (Entra app credential)"
    CertStoreLocation = $storeLocation
    KeyAlgorithm      = 'RSA'
    KeyLength         = 2048
    HashAlgorithm     = 'SHA256'
    KeySpec           = 'Signature'
    KeyExportPolicy   = if ($Exportable) { 'Exportable' } else { 'NonExportable' }
    NotBefore         = (Get-Date).AddMinutes(-5)   # small back-date to absorb clock skew
    NotAfter          = (Get-Date).AddDays($ValidDays)
}
$cert = New-SelfSignedCertificate @certParams
Write-Host "Created certificate $($cert.Thumbprint) in $storeLocation." -ForegroundColor Green

# ---- 2) Grant Read on the private key file ----
# Machine-store private keys are ACL'd to SYSTEM + Administrators only. The key file lives under
# ProgramData: Crypto\Keys for CNG (KSP) keys, Crypto\RSA\MachineKeys for legacy CSP keys.
if ($grantAccounts.Count -gt 0) {
    $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($cert)
    if ($rsa -is [System.Security.Cryptography.RSACng]) {
        $keyPath = Join-Path "$env:ProgramData\Microsoft\Crypto\Keys" $rsa.Key.UniqueName
    } else {
        $keyPath = Join-Path "$env:ProgramData\Microsoft\Crypto\RSA\MachineKeys" $rsa.CspKeyContainerInfo.UniqueKeyContainerName
    }

    try {
        $acl = Get-Acl -LiteralPath $keyPath
        foreach ($account in $grantAccounts) {
            $rule = New-Object System.Security.AccessControl.FileSystemAccessRule($account, 'Read', 'Allow')
            $acl.AddAccessRule($rule)
        }
        Set-Acl -LiteralPath $keyPath -AclObject $acl
    } catch {
        Write-Warning "Certificate $($cert.Thumbprint) was created, but granting private key access failed: $($_.Exception.Message)"
        Write-Warning "Grant Read manually: certlm.msc > Personal > Certificates > right-click > All Tasks > Manage Private Keys."
        throw
    }

    foreach ($account in $grantAccounts) {
        Write-Host "Granted Read on private key to $($account.Value)." -ForegroundColor Green
    }
}

# ---- 3) Export the public key only ----
if (-not (Test-Path -LiteralPath $ExportPath)) {
    New-Item -ItemType Directory -Path $ExportPath -Force | Out-Null
}

$safeName = $CertName -replace '[\\/:*?"<>|]', '_'
$cerPath  = Join-Path $ExportPath ("{0}_{1}.cer" -f $safeName, $cert.NotAfter.ToString('yyyyMMdd'))

Export-Certificate -Cert $cert -FilePath $cerPath -Type CERT -Force | Out-Null
Write-Host "Exported public key to $cerPath" -ForegroundColor Green

# ---- 4) Summary ----
Write-Host ''
Write-Host 'Certificate details' -ForegroundColor Cyan
Write-Host ("  Subject     : {0}" -f $cert.Subject)
Write-Host ("  Thumbprint  : {0}" -f $cert.Thumbprint)
Write-Host ("  Not before  : {0}" -f $cert.NotBefore.ToString('yyyy-MM-dd HH:mm'))
Write-Host ("  Not after   : {0}" -f $cert.NotAfter.ToString('yyyy-MM-dd HH:mm'))
Write-Host ("  Store       : {0}" -f $storeLocation)
Write-Host ("  Key export  : {0}" -f $certParams.KeyExportPolicy)
Write-Host ("  Key readers : {0}" -f $(if ($grantAccounts.Count) { ($grantAccounts | ForEach-Object { $_.Value }) -join ', ' } else { '(admins/SYSTEM only)' }))
Write-Host ("  Public key  : {0}" -f $cerPath)
Write-Host ''
Write-Host 'Next: upload the .cer under App registrations > <app> > Certificates & secrets > Certificates,' -ForegroundColor Yellow
Write-Host '      then confirm the thumbprint shown in Entra matches the one above.' -ForegroundColor Yellow

$cert
