#Entra tenant and delegated app registration used for Graph sign-in
$TenantId      = "<tenant-id>"
$EntraClientId = "<app-client-id>"

#Install Module
#Install-Module -Name Microsoft.Graph.Users

# NOTE: Graph sign-in happens once up front (below) via a local listener; both the new-account and
#       re-joiner paths use Graph to verify on-prem sync takeover after the delta sync.



#New accounts are created in this staging OU (out of AAD Connect sync scope)
$usersContainer = "OU=AAD Copy - Not Enabled for Sync,DC=CoreSpec,DC=Local"

#Both re-joiners and newly-created accounts are moved into this sync-enabled OU
$SyncOu = "OU=AADConnectSync Enabled Accounts,DC=CoreSpec,DC=Local"

#Browser sign-in (auth code + PKCE) captured on a localhost listener; returns the Graph access token as a SecureString.
#App registration needs redirect URI http://localhost under "Mobile and desktop applications" (Entra ignores the port).
function Get-GraphTokenViaListener {
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ClientId,
        [string[]]$Scopes = @('User.Read.All'),
        [int]$TimeoutSeconds = 180
    )

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    #Pick a free localhost port
    $tcp = New-Object System.Net.Sockets.TcpListener([Net.IPAddress]::Loopback, 0)
    $tcp.Start(); $port = $tcp.LocalEndpoint.Port; $tcp.Stop()
    $redirectUri = "http://localhost:$port/"

    #PKCE verifier/challenge and state
    $bytes = New-Object byte[] 32
    [Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($bytes)
    $verifier  = [Convert]::ToBase64String($bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')
    $hash      = [Security.Cryptography.SHA256]::Create().ComputeHash([Text.Encoding]::ASCII.GetBytes($verifier))
    $challenge = [Convert]::ToBase64String($hash).TrimEnd('=').Replace('+', '-').Replace('/', '_')
    $state     = [guid]::NewGuid().ToString('N')

    $scope   = ($Scopes | ForEach-Object { "https://graph.microsoft.com/$_" }) -join ' '
    $authUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/authorize" +
               "?client_id=$ClientId&response_type=code" +
               "&redirect_uri=$([uri]::EscapeDataString($redirectUri))" +
               "&scope=$([uri]::EscapeDataString($scope))" +
               "&state=$state&code_challenge=$challenge&code_challenge_method=S256&prompt=select_account"

    $listener = New-Object System.Net.HttpListener
    $listener.Prefixes.Add($redirectUri)
    $listener.Start()
    try {
        Start-Process $authUrl
        $params   = @{}
        $deadline = (Get-Date).AddSeconds($TimeoutSeconds)

        #Loop until the redirect carrying code/error arrives (ignores favicon and other stray requests)
        while (-not ($params.code -or $params.error)) {
            $remaining = $deadline - (Get-Date)
            if ($remaining.TotalMilliseconds -le 0) { throw "Timed out waiting for sign-in." }
            $async = $listener.BeginGetContext($null, $null)
            if (-not $async.AsyncWaitHandle.WaitOne($remaining)) { throw "Timed out waiting for sign-in." }
            $ctx = $listener.EndGetContext($async)

            #Request.QueryString can arrive empty; parse RawUrl directly
            $params = @{}
            $query  = ($ctx.Request.RawUrl -split '\?', 2)[1]
            if ($query) {
                foreach ($pair in $query -split '&') {
                    $k, $v = $pair -split '=', 2
                    $params[[uri]::UnescapeDataString($k)] = [uri]::UnescapeDataString(($v -replace '\+', ' '))
                }
            }

            $msg = if ($params.code) { 'Sign-in complete. You can close this tab.' }
                   elseif ($params.error) { "Sign-in failed: $($params.error)" }
                   else { '' }
            $buf = [Text.Encoding]::UTF8.GetBytes("<html><body><h3>$msg</h3></body></html>")
            $ctx.Response.ContentType = 'text/html'
            $ctx.Response.OutputStream.Write($buf, 0, $buf.Length)
            $ctx.Response.Close()
        }
    }
    finally {
        $listener.Stop()
        $listener.Close()
    }

    if ($params.error)            { throw "Sign-in failed: $($params.error) - $($params.error_description)" }
    if ($params.state -ne $state) { throw "State mismatch; discarding sign-in response." }

    #Exchange the auth code for a token
    $token = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
        -ContentType 'application/x-www-form-urlencoded' -Body @{
            client_id     = $ClientId
            grant_type    = 'authorization_code'
            code          = $params.code
            redirect_uri  = $redirectUri
            code_verifier = $verifier
            scope         = $scope
        }

    ConvertTo-SecureString $token.access_token -AsPlainText -Force
}

function Add-LocalADObject {
    ForEach ($object in $input) {
        write-host $object.DisplayName
        write-host $object.UserType
        write-host $object.UserPrincipalName

        #Check if the Azure AD User object is configured with a valid UPN domain
        if ($object.UserPrincipalName -like "*onmicrosoft.com") {
            write-host "Skipping object " $object.DisplayName "because it does not have a custom logon domain"
            continue
        }

        #Create the AD User object
        if ($object.UserType -eq "Member") {
            $userName = $object.UserPrincipalName.split('@')[0]
            Write-Host $userName
            New-ADUser -SamAccountName $userName -UserPrincipalName $object.UserPrincipalName -Name $object.DisplayName -DisplayName $object.DisplayName -Path $usersContainer -AccountPassword (ConvertTo-SecureString $TempPassword -AsPlainText -Force) -Enabled $True -PasswordNeverExpires $True -PassThru 
            $filter = "CN=" + $object.DisplayName
            #$empId = (Get-AzureADUser -ObjectId $object.UserPrincipalName).extensionproperty["employeeId"]
        }


        $localADObject = Get-ADObject -LDAPFilter $filter
        
        ##Update AD user attributes with the same values that are on the Azure AD user object
        if ($object.GivenName -ne $null) { Set-ADObject $localADObject -Add @{givenName = $object.GivenName } }
        if ($object.Surname -ne $null) { Set-ADObject $localADObject -Add @{sn = $object.Surname } }
        if ($object.Mail -ne $null) { Set-ADObject $localADObject -Add @{mail = $object.Mail } }
        if ($object.StreetAddress -ne $null) { Set-ADObject $localADObject -Add @{streetAddress = $object.StreetAddress } }
        if ($object.PostalCode -ne $null) { Set-ADObject $localADObject -Add @{postalCode = $object.PostalCode } }
        if ($object.City -ne $null) { Set-ADObject $localADObject -Add @{l = $object.City } }
        if ($object.State -ne $null) { Set-ADObject $localADObject -Add @{st = $object.State } }        
        if ($object.OfficeLocation -ne $null) { Set-ADObject $localADObject -Add @{physicalDeliveryOfficeName = $object.OfficeLocation } }
        if ($object.BusinessPhones -ne $null) { Set-ADObject $localADObject -Add @{telephoneNumber = $object.BusinessPhones[0] } }
        if ($object.FaxNumber -ne $null) { Set-ADObject $localADObject -Add @{facsimileTelephoneNumber = $object.FaxNumber } }
        if ($object.MobilePhone -ne $null) { Set-ADObject $localADObject -Add @{mobile = $object.MobilePhone } }
        if ($object.JobTitle -ne $null) { Set-ADObject $localADObject -Add @{title = $object.JobTitle } }
        if ($object.Department -ne $null) { Set-ADObject $localADObject -Add @{department = $object.Department } }
        if ($object.CompanyName -ne $null) { Set-ADObject $localADObject -Add @{company = $object.CompanyName } }
        if ($object.employeeid -ne $null) { Set-ADObject $localADObject -Replace @{employeeID = $object.employeeid } }
 
        ##Update AD user ProxyAddresses attribute with the same values that are on the Azure AD user object
        if ($object.ProxyAddresses -ne $null) {
            ForEach ($proxyAddress in $object.ProxyAddresses) {
                if ($proxyAddress -notlike "*onmicrosoft*") {
                    Set-ADObject $localADObject -Add @{ProxyAddresses = $proxyAddress }
                }
            }
        }
         
    }
}

#Poll the cloud user's OnPremisesSyncEnabled until it flips to Yes (true), up to a timeout.
function Wait-OnPremSync {
    param(
        [Parameter(Mandatory)][string]$UserId,
        [int]$TimeoutMinutes = 10,
        [int]$PollSeconds = 30
    )
    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    while ((Get-Date) -lt $deadline) {
        if ((Get-MgUser -UserId $UserId -Property OnPremisesSyncEnabled).OnPremisesSyncEnabled -eq $true) {
            return $true
        }
        Start-Sleep -Seconds $PollSeconds
    }
    return $false
}

#Specify the UPN of the Entra (Azure AD) object you want to sync to / create in Active Directory
$object = $null
$UPN = "eric.weston@corespecialty.com"
$SamAccountName = $UPN.Split('@')[0]
$TempPassword = -join ((33..126) | Get-Random -Count 32 | ForEach-Object { [char]$_ })

#Sign in to Microsoft Graph once, up front, via the local listener -- both paths need it.
#NOTE: User.Read.All covers the user reads and sync polling. The SOA check/reset below also needs
#      User-OnPremisesSyncBehavior.ReadWrite.All (plus Hybrid Administrator); add it to -Scopes if you
#      want that branch to work, otherwise it throws its existing permission error.
$graphToken = Get-GraphTokenViaListener -TenantId $TenantId -ClientId $EntraClientId -Scopes 'User.Read.All'
Connect-MgGraph -AccessToken $graphToken -NoWelcome
$graphToken = $null

#--- Re-joiner detection: is there already an account with this SamAccountName anywhere in the domain? ---
$existing = Get-ADUser -Filter "SamAccountName -eq '$SamAccountName'"

if ($existing) {
    #=== Re-joiner: re-enable the existing account and move it into the sync OU ===
    Write-Host "Found existing account '$SamAccountName' at $($existing.DistinguishedName)" -ForegroundColor Cyan

    #Safety guard: never touch an account that already looks active
    $parentOu = ($existing.DistinguishedName -split ',', 2)[1]
    if ($existing.Enabled -or $parentOu -eq $SyncOu) {
        throw "Account '$SamAccountName' is already enabled and/or already in the sync OU ($parentOu). Refusing to modify what looks like an active account."
    }

    Enable-ADAccount -Identity $existing
    $existing | Move-ADObject -TargetPath $SyncOu
    Write-Host "Re-enabled '$SamAccountName' and moved it to $SyncOu" -ForegroundColor Green
}
else {
    #=== New account: build the on-prem object from the Entra user, then move it into the sync OU ===

    $aws2FA = Read-Host "Add AWS MFA group?  Must do before initial sync. y/n"

    Get-MgUser -UserId $UPN -Property UserPrincipalName, DisplayName, UserType, GivenName, Surname, Mail, StreetAddress, PostalCode, City, State, OfficeLocation, BusinessPhones, FaxNumber, MobilePhone, JobTitle, Department, CompanyName, EmployeeId, ProxyAddresses | Add-LocalADObject

    if ($aws2FA -eq "y") {
        Add-ADGroupMember -Identity "AWS2MFA" -Members $SamAccountName
    }

    Start-Sleep -s 15

    Get-ADUser $SamAccountName | Move-ADObject -TargetPath $SyncOu
}

Start-Sleep -s 15

#=== Trigger sync, then verify on-prem takeover; reset Source of Authority if the cloud still holds it ===
$userId = (Get-MgUser -UserId $UPN -Property Id).Id

Start-ADSyncSyncCycle -PolicyType Delta

if (Wait-OnPremSync -UserId $userId -TimeoutMinutes 10) {
    Write-Host "On-prem sync confirmed: OnPremisesSyncEnabled = Yes for $UPN." -ForegroundColor Green
}
else {
    #Still 'No' after 10 minutes. Is the cloud Source of Authority the blocker?
    try {
        $behavior = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$userId/onPremisesSyncBehavior?`$select=isCloudManaged"
    }
    catch {
        throw "Could not read SOA status (onPremisesSyncBehavior) for $UPN. The token likely lacks 'User-OnPremisesSyncBehavior.Read.All' (or ReadWrite). Underlying error: $($_.Exception.Message)"
    }

    if ($behavior.isCloudManaged -eq $true) {
        Write-Warning "Cloud holds the Source of Authority (isCloudManaged = true). Resetting SOA back to on-premises for $UPN..."
        try {
            Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/users/$userId/onPremisesSyncBehavior" -Body @{ isCloudManaged = $false }
        }
        catch {
            throw "SOA reset failed for $UPN. The token likely lacks 'User-OnPremisesSyncBehavior.ReadWrite.All' / Hybrid Administrator rights. Underlying error: $($_.Exception.Message)"
        }

        Write-Host "SOA reset (isCloudManaged -> false). Re-running delta sync and re-verifying..." -ForegroundColor Cyan
        Start-Sleep -s 15
        Start-ADSyncSyncCycle -PolicyType Delta

        if (Wait-OnPremSync -UserId $userId -TimeoutMinutes 10) {
            Write-Host "On-prem sync confirmed after SOA reset: OnPremisesSyncEnabled = Yes for $UPN." -ForegroundColor Green
        }
        else {
            Write-Warning ("Still not synced after SOA reset + delta sync. The tenant-wide 'blockCloudObjectTakeoverThroughHardMatchEnabled' flag is most likely still enabled, which blocks cloud->on-prem takeover. Disable it (PATCH /beta/directory/onPremisesSynchronization/{id}), run another delta sync, then re-enable it. Docs: https://learn.microsoft.com/en-us/entra/identity/hybrid/how-to-user-source-of-authority-configure")
        }
    }
    else {
        Write-Warning "OnPremisesSyncEnabled is still not Yes, but isCloudManaged = false (not a cloud-SOA object). SOA reset would not help -- likely a different issue (object out of sync scope, anchor/ImmutableID mismatch, or a Connect Sync error). Investigate the Connect Sync run for $UPN."
    }
}
