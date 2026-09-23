<#
.SYNOPSIS
    Generates a self-signed certificate in the Local Machine store and exports the public key (.cer)
    for upload to an Entra ID app registration / enterprise application.

.DESCRIPTION
    Self-contained: no repo config, Delinea, or module dependencies. Intended for Windows Server,
    Windows PowerShell 5.1, run as Administrator (writing to Cert:\LocalMachine\My requires it).

    1. Creates an RSA 2048 / SHA256 self-signed cert valid for 365 days in Cert:\LocalMachine\My.
    2. Exports the PUBLIC key only (DER .cer) to C:\Temp. The private key never leaves the machine store.
    3. Prints the thumbprint and validity dates.

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

.EXAMPLE
    .\Generate Machine Store Self-Signed Cert.ps1 -CertName "svc-app01.contoso.com"
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

    [switch]$Exportable
)

$ErrorActionPreference = 'Stop'
$storeLocation = 'Cert:\LocalMachine\My'

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

# ---- 2) Export the public key only ----
if (-not (Test-Path -LiteralPath $ExportPath)) {
    New-Item -ItemType Directory -Path $ExportPath -Force | Out-Null
}

$safeName = $CertName -replace '[\\/:*?"<>|]', '_'
$cerPath  = Join-Path $ExportPath ("{0}_{1}.cer" -f $safeName, $cert.NotAfter.ToString('yyyyMMdd'))

Export-Certificate -Cert $cert -FilePath $cerPath -Type CERT -Force | Out-Null
Write-Host "Exported public key to $cerPath" -ForegroundColor Green

# ---- 3) Summary ----
Write-Host ''
Write-Host 'Certificate details' -ForegroundColor Cyan
Write-Host ("  Subject     : {0}" -f $cert.Subject)
Write-Host ("  Thumbprint  : {0}" -f $cert.Thumbprint)
Write-Host ("  Not before  : {0}" -f $cert.NotBefore.ToString('yyyy-MM-dd HH:mm'))
Write-Host ("  Not after   : {0}" -f $cert.NotAfter.ToString('yyyy-MM-dd HH:mm'))
Write-Host ("  Store       : {0}" -f $storeLocation)
Write-Host ("  Key export  : {0}" -f $certParams.KeyExportPolicy)
Write-Host ("  Public key  : {0}" -f $cerPath)
Write-Host ''
Write-Host 'Next: upload the .cer under App registrations > <app> > Certificates & secrets > Certificates,' -ForegroundColor Yellow
Write-Host '      then confirm the thumbprint shown in Entra matches the one above.' -ForegroundColor Yellow

$cert
