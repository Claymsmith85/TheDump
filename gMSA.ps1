Import-Module ActiveDirectory

$gmsaName = 'CloudSyncSvc01'
$domain = Get-ADDomain
$agentComputer = Get-ADComputer -Identity $env:COMPUTERNAME

New-ADServiceAccount `
    -Name $gmsaName `
    -SamAccountName "$gmsaName`$" `
    -DNSHostName "$gmsaName.$($domain.DNSRoot)" `
    -PrincipalsAllowedToRetrieveManagedPassword $agentComputer `
    -KerberosEncryptionType AES128,AES256

Install-ADServiceAccount -Identity $gmsaName
Test-ADServiceAccount -Identity $gmsaName

Get-ADServiceAccount -Identity $gmsaName |
    Select-Object Name, SamAccountName
