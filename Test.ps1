#Requires -Version 7.0
# Minimal EWS read test: signs in as a user and reads the Inbox of each mailbox given.
# Example:
#   .\Test-EwsMailboxAccess.ps1 -TenantId corespecialty.onmicrosoft.com -ClientId <appId> `
#       -UserName CS.BitTitan@corespecialtyins.com -Mailbox apus@corespecialtyins.com, AccountsPayable@corespecialty.com
param(
    [Parameter(Mandatory)][string]$TenantId,
    [Parameter(Mandatory)][string]$ClientId,
    [Parameter(Mandatory)][string]$UserName,
    [Parameter(Mandatory)][string[]]$Mailbox
)
$ErrorActionPreference = 'Stop'

# Sign in (username/password)
$cred = Get-Credential -UserName $UserName -Message 'Password'
$tok = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Body @{
    grant_type = 'password'
    client_id  = $ClientId
    scope      = 'https://outlook.office365.com/EWS.AccessAsUser.All'
    username   = $cred.UserName
    password   = $cred.GetNetworkCredential().Password
}

# Show who the token is for and what it allows
$p = $tok.access_token.Split('.')[1].Replace('-', '+').Replace('_', '/')
$p = $p.PadRight($p.Length + (4 - $p.Length % 4) % 4, '=')
$claims = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($p)) | ConvertFrom-Json
Write-Host "Token user: $($claims.upn)   Scopes: $($claims.scp)"

# Read the Inbox of each mailbox
foreach ($mbx in $Mailbox) {
    $body = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
  <soap:Header><t:RequestServerVersion Version="Exchange2013_SP1"/></soap:Header>
  <soap:Body>
    <m:GetFolder>
      <m:FolderShape><t:BaseShape>Default</t:BaseShape></m:FolderShape>
      <m:FolderIds>
        <t:DistinguishedFolderId Id="inbox"><t:Mailbox><t:EmailAddress>$mbx</t:EmailAddress></t:Mailbox></t:DistinguishedFolderId>
      </m:FolderIds>
    </m:GetFolder>
  </soap:Body>
</soap:Envelope>
"@
    $r = Invoke-WebRequest -Uri 'https://outlook.office365.com/EWS/Exchange.asmx' -Method Post -Body $body `
        -ContentType 'text/xml; charset=utf-8' -SkipHttpErrorCheck `
        -Headers @{ Authorization = "Bearer $($tok.access_token)"; 'X-AnchorMailbox' = $mbx }

    if ([int]$r.StatusCode -ne 200) {
        Write-Host "$mbx : HTTP $($r.StatusCode) $($r.Content)" -ForegroundColor Red
        continue
    }

    $x  = [xml]([string]$r.Content).TrimStart([char]0xFEFF)
    $rm = $x.SelectSingleNode("//*[local-name()='GetFolderResponseMessage']")
    if ($rm.ResponseClass -eq 'Success') {
        $name  = $x.SelectSingleNode("//*[local-name()='DisplayName']").InnerText
        $count = $x.SelectSingleNode("//*[local-name()='TotalCount']").InnerText
        Write-Host "$mbx : OK - $name, $count item(s)" -ForegroundColor Green
    }
    else {
        $code = $rm.SelectSingleNode("*[local-name()='ResponseCode']").InnerText
        $text = $rm.SelectSingleNode("*[local-name()='MessageText']").InnerText
        Write-Host "$mbx : $code - $text" -ForegroundColor Red
    }
}