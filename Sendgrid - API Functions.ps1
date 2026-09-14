Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-SendGridBaseUri {
	[CmdletBinding()]
	param(
		[ValidateSet('Global', 'EU')]
		[string]$Region = 'Global'
	)

	if ($Region -eq 'EU') {
		return 'https://api.eu.sendgrid.com'
	}

	return 'https://api.sendgrid.com'
}

function ConvertTo-SendGridArray {
	[CmdletBinding()]
	param(
		$InputObject
	)

	if ($null -eq $InputObject) {
		return @()
	}

	if ($InputObject -is [string] -or $InputObject -is [System.Collections.IDictionary]) {
		return , $InputObject
	}

	if ($InputObject -is [System.Collections.IEnumerable]) {
		return @($InputObject)
	}

	return , $InputObject
}

function ConvertTo-SendGridQueryString {
	[CmdletBinding()]
	param(
		[hashtable]$Query
	)

	if ($null -eq $Query -or $Query.Count -eq 0) {
		return ''
	}

	$pairs = foreach ($key in $Query.Keys) {
		$value = $Query[$key]
		if ($null -eq $value) {
			continue
		}

		$text = [string]$value
		if ([string]::IsNullOrWhiteSpace($text)) {
			continue
		}

		'{0}={1}' -f [uri]::EscapeDataString([string]$key), [uri]::EscapeDataString($text)
	}

	$pairs = @($pairs)
	if ($pairs.Count -eq 0) {
		return ''
	}

	return '?' + ($pairs -join '&')
}

function Get-SendGridErrorMessage {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[System.Management.Automation.ErrorRecord]$ErrorRecord
	)

	if ($ErrorRecord.ErrorDetails -and -not [string]::IsNullOrWhiteSpace($ErrorRecord.ErrorDetails.Message)) {
		try {
			$parsed = $ErrorRecord.ErrorDetails.Message | ConvertFrom-Json
			if ($parsed.errors) {
				return (($parsed.errors | ForEach-Object { $_.message }) -join '; ')
			}
		}
		catch {
			return $ErrorRecord.ErrorDetails.Message
		}

		return $ErrorRecord.ErrorDetails.Message
	}

	return $ErrorRecord.Exception.Message
}

function New-SendGridClient {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[string]$ApiKey,

		[ValidateSet('Global', 'EU')]
		[string]$Region = 'Global',

		[ValidateRange(1, 500)]
		[int]$DefaultPageSize = 100
	)

	if ([string]::IsNullOrWhiteSpace($ApiKey)) {
		throw 'ApiKey cannot be empty.'
	}

	[pscustomobject]@{
		ApiKey          = $ApiKey
		Region          = $Region
		BaseUri         = Get-SendGridBaseUri -Region $Region
		DefaultPageSize = $DefaultPageSize
	}
}

function New-SendGridClientFromDelinea {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[string]$SsBase,

		[Parameter(Mandatory)]
		[int]$SecretId,

		[Parameter(Mandatory)]
		[string]$SsToken,

		[ValidateSet('Global', 'EU')]
		[string]$Region = 'Global',

		[ValidateRange(1, 500)]
		[int]$DefaultPageSize = 100
	)

	if (-not (Get-Command Get-SendGridApiKeyFromDelinea -ErrorAction SilentlyContinue)) {
		throw 'Get-SendGridApiKeyFromDelinea is not loaded. Dot-source Delinea\DelineaAuth.ps1 first.'
	}

	$apiKey = Get-SendGridApiKeyFromDelinea -SecretId $SecretId -SsBase $SsBase -SsToken $SsToken
	return New-SendGridClient -ApiKey $apiKey -Region $Region -DefaultPageSize $DefaultPageSize
}

function Invoke-SendGridRequest {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[pscustomobject]$Client,

		[Parameter(Mandatory)]
		[ValidateSet('GET', 'POST', 'PUT', 'PATCH', 'DELETE')]
		[string]$Method,

		[Parameter(Mandatory)]
		[string]$Path,

		[hashtable]$Query,

		$Body,

		[string]$OnBehalfOf,

		[string]$BaseUri,

		[switch]$RawResponse
	)

	if ([string]::IsNullOrWhiteSpace($BaseUri)) {
		$BaseUri = [string]$Client.BaseUri
	}

	$headers = @{
		Authorization = "Bearer $($Client.ApiKey)"
		Accept        = 'application/json'
	}

	if (-not [string]::IsNullOrWhiteSpace($OnBehalfOf)) {
		$headers['on-behalf-of'] = $OnBehalfOf
	}

	$uri = '{0}{1}{2}' -f $BaseUri.TrimEnd('/'), $Path, (ConvertTo-SendGridQueryString -Query $Query)

	$invokeParams = @{
		Method      = $Method
		Uri         = $uri
		Headers     = $headers
		ErrorAction = 'Stop'
	}

	if ($PSBoundParameters.ContainsKey('Body')) {
		$invokeParams['ContentType'] = 'application/json'

		if ($Body -is [string]) {
			$invokeParams['Body'] = $Body
		}
		else {
			$invokeParams['Body'] = $Body | ConvertTo-Json -Depth 20 -Compress
		}
	}

	try {
		$response = Invoke-WebRequest @invokeParams
	}
	catch {
		$statusCode = $null
		try {
			if ($_.Exception.Response -and $_.Exception.Response.StatusCode) {
				$statusCode = [int]$_.Exception.Response.StatusCode
			}
		}
		catch {
		}

		$msg = Get-SendGridErrorMessage -ErrorRecord $_
		throw "SendGrid API call failed: $Method $uri. StatusCode=$statusCode. Details=$msg"
	}

	if ($RawResponse) {
		return $response
	}

	if ([string]::IsNullOrWhiteSpace($response.Content)) {
		return $null
	}

	try {
		return $response.Content | ConvertFrom-Json
	}
	catch {
		return $response.Content
	}
}

# Backward-compatible alias; 'Sendgrid - Azure Group Management' expects Invoke-SendGridRequest.
Set-Alias -Name Invoke-SendGridApiRequest -Value Invoke-SendGridRequest

function Invoke-SendGridPagedRequest {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[pscustomobject]$Client,

		[Parameter(Mandatory)]
		[string]$Path,

		[string]$OnBehalfOf,

		[hashtable]$Query,

		[string]$ResponseProperty,

		[ValidateRange(1, 500)]
		[int]$PageSize,

		[string]$BaseUri
	)

	if (-not $PSBoundParameters.ContainsKey('PageSize')) {
		$PageSize = [int]$Client.DefaultPageSize
	}

	$items = [System.Collections.Generic.List[object]]::new()
	$offset = 0

	while ($true) {
		$queryParams = @{}
		if ($null -ne $Query) {
			foreach ($k in $Query.Keys) {
				$queryParams[$k] = $Query[$k]
			}
		}

		$queryParams['limit'] = $PageSize
		$queryParams['offset'] = $offset

		$page = Invoke-SendGridRequest -Client $Client -Method GET -Path $Path -Query $queryParams -OnBehalfOf $OnBehalfOf -BaseUri $BaseUri

		$pageItems = if ([string]::IsNullOrWhiteSpace($ResponseProperty)) {
			ConvertTo-SendGridArray -InputObject $page
		}
		else {
			if ($null -eq $page) {
				@()
			}
			else {
				$prop = $page.PSObject.Properties[$ResponseProperty]
				if ($null -eq $prop) {
					ConvertTo-SendGridArray -InputObject $page
				}
				else {
					ConvertTo-SendGridArray -InputObject $prop.Value
				}
			}
		}

		foreach ($item in $pageItems) {
			[void]$items.Add($item)
		}

		if (@($pageItems).Count -lt $PageSize) {
			break
		}

		$offset += $PageSize
	}

	return $items.ToArray()
}

function Test-SendGridApiKey {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[pscustomobject]$Client
	)

	try {
		$null = Invoke-SendGridRequest -Client $Client -Method GET -Path '/v3/scopes'
		return $true
	}
	catch {
		return $false
	}
}

function Get-SendGridScopes {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[pscustomobject]$Client
	)

	Invoke-SendGridRequest -Client $Client -Method GET -Path '/v3/scopes'
}

function Get-SendGridSubusers {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[pscustomobject]$Client,

		[switch]$IncludeRegion,

		[ValidateSet('all', 'global', 'eu')]
		[string]$RegionFilter = 'all',

		[ValidateRange(1, 500)]
		[int]$PageSize
	)

	$query = @{}
	if ($IncludeRegion) {
		$query['include_region'] = 'true'
		$query['region'] = $RegionFilter
	}

	Invoke-SendGridPagedRequest -Client $Client -Path '/v3/subusers' -Query $query -PageSize $PageSize
}

function Get-SendGridSubuser {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[pscustomobject]$Client,

		[Parameter(Mandatory)]
		[string]$Username
	)

	$encoded = [uri]::EscapeDataString($Username)
	Invoke-SendGridRequest -Client $Client -Method GET -Path "/v3/subusers/$encoded"
}

function New-SendGridSubuser {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[pscustomobject]$Client,

		[Parameter(Mandatory)]
		[string]$Username,

		[Parameter(Mandatory)]
		[string]$Email,

		[Parameter(Mandatory)]
		[string]$Password,

		[string[]]$Ips = @()
	)

	$body = @{
		username = $Username
		email    = $Email
		password = $Password
		ips      = @($Ips)
	}

	Invoke-SendGridRequest -Client $Client -Method POST -Path '/v3/subusers' -Body $body
}

function Remove-SendGridSubuser {
	[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
	param(
		[Parameter(Mandatory)]
		[pscustomobject]$Client,

		[Parameter(Mandatory)]
		[string]$Username
	)

	$encoded = [uri]::EscapeDataString($Username)
	if ($PSCmdlet.ShouldProcess($Username, 'Delete SendGrid subuser')) {
		Invoke-SendGridRequest -Client $Client -Method DELETE -Path "/v3/subusers/$encoded" | Out-Null
	}
}

function Get-SendGridApiKeys {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[pscustomobject]$Client,

		[string]$OnBehalfOf,

		[ValidateRange(1, 500)]
		[int]$PageSize
	)

	Invoke-SendGridPagedRequest -Client $Client -Path '/v3/api_keys' -OnBehalfOf $OnBehalfOf -ResponseProperty 'result' -PageSize $PageSize
}

function Get-SendGridApiKey {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[pscustomobject]$Client,

		[Parameter(Mandatory)]
		[string]$ApiKeyId,

		[string]$OnBehalfOf
	)

	$encoded = [uri]::EscapeDataString($ApiKeyId)
	Invoke-SendGridRequest -Client $Client -Method GET -Path "/v3/api_keys/$encoded" -OnBehalfOf $OnBehalfOf
}

function New-SendGridApiKey {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[pscustomobject]$Client,

		[Parameter(Mandatory)]
		[string]$Name,

		[Parameter(Mandatory)]
		[string[]]$Scopes,

		[string]$OnBehalfOf
	)

	$body = @{
		name   = $Name
		scopes = @($Scopes)
	}

	Invoke-SendGridRequest -Client $Client -Method POST -Path '/v3/api_keys' -Body $body -OnBehalfOf $OnBehalfOf
}

function Remove-SendGridApiKey {
	[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
	param(
		[Parameter(Mandatory)]
		[pscustomobject]$Client,

		[Parameter(Mandatory)]
		[string]$ApiKeyId,

		[string]$OnBehalfOf
	)

	$encoded = [uri]::EscapeDataString($ApiKeyId)
	if ($PSCmdlet.ShouldProcess($ApiKeyId, 'Delete SendGrid API key')) {
		Invoke-SendGridRequest -Client $Client -Method DELETE -Path "/v3/api_keys/$encoded" -OnBehalfOf $OnBehalfOf | Out-Null
	}
}

function Get-SendGridTeammates {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[pscustomobject]$Client,

		[string]$OnBehalfOf,

		[ValidateRange(1, 500)]
		[int]$PageSize
	)

	Invoke-SendGridPagedRequest -Client $Client -Path '/v3/teammates' -OnBehalfOf $OnBehalfOf -ResponseProperty 'result' -PageSize $PageSize
}

function Get-SendGridTeammate {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[pscustomobject]$Client,

		[Parameter(Mandatory)]
		[string]$TeammateName,

		[string]$OnBehalfOf
	)

	$encoded = [uri]::EscapeDataString($TeammateName)
	Invoke-SendGridRequest -Client $Client -Method GET -Path "/v3/teammates/$encoded" -OnBehalfOf $OnBehalfOf
}

function New-SendGridSsoTeammate {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[pscustomobject]$Client,

		[Parameter(Mandatory)]
		[string]$Email,

		[Parameter(Mandatory)]
		[string]$FirstName,

		[Parameter(Mandatory)]
		[string]$LastName,

		[string[]]$Scopes = @('user.profile.read', 'user.profile.update'),

		[bool]$IsAdmin = $false,

		# One entry per subuser: @{ id = <subuser id>; permission_type = 'admin' } or
		# @{ id = <subuser id>; permission_type = 'restricted'; scopes = @(...) }
		[object[]]$SubuserAccess
	)

	# SendGrid constraint: a non-admin teammate is EITHER parent-scoped (scopes)
	# OR subuser-scoped (subuser_access); the two cannot be combined.
	$hasSubuserAccess = $PSBoundParameters.ContainsKey('SubuserAccess') -and @($SubuserAccess).Count -gt 0
	if ($hasSubuserAccess -and $PSBoundParameters.ContainsKey('Scopes')) {
		throw 'A teammate cannot have both parent scopes and subuser_access. Pass -Scopes or -SubuserAccess, not both.'
	}

	$body = @{
		email      = $Email
		first_name = $FirstName
		last_name  = $LastName
		is_sso     = $true
		is_admin   = $IsAdmin
	}

	if ($hasSubuserAccess) {
		$body['has_restricted_subuser_access'] = $true
		$body['subuser_access'] = @($SubuserAccess)
	}
	elseif (-not $IsAdmin) {
		$body['scopes'] = @($Scopes)
	}

	Invoke-SendGridRequest -Client $Client -Method POST -Path '/v3/sso/teammates' -Body $body
}

function Set-SendGridTeammateSubuserAccess {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[pscustomobject]$Client,

		[Parameter(Mandatory)]
		[string]$TeammateName,

		[Parameter(Mandatory)]
		[object[]]$SubuserAccess,

		[bool]$IsAdmin = $false,

		[bool]$HasRestrictedSubuserAccess = $true
	)

	$encoded = [uri]::EscapeDataString($TeammateName)
	$body = @{
		is_admin                      = $IsAdmin
		has_restricted_subuser_access = $HasRestrictedSubuserAccess
		subuser_access                = @($SubuserAccess)
	}

	Invoke-SendGridRequest -Client $Client -Method PATCH -Path "/v3/sso/teammates/$encoded" -Body $body
}

function Get-SendGridTeammateSubuserAccess {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[pscustomobject]$Client,

		[Parameter(Mandatory)]
		[string]$TeammateName,

		[ValidateRange(1, 500)]
		[int]$Limit = 100
	)

	function Get-OptionalNestedPropertyValue {
		param(
			[object]$InputObject,
			[string]$PropertyName,
			$Default = $null
		)

		if ($null -eq $InputObject) {
			return $Default
		}

		$prop = $InputObject.PSObject.Properties[$PropertyName]
		if ($null -eq $prop) {
			return $Default
		}

		return $prop.Value
	}

	$encoded = [uri]::EscapeDataString($TeammateName)
	$result = [System.Collections.Generic.List[object]]::new()
	$afterSubuserId = $null
	$hasRestrictedSubuserAccess = $false

	do {
		$query = @{ limit = $Limit }
		if ($afterSubuserId) {
			$query['after_subuser_id'] = $afterSubuserId
		}

		$response = Invoke-SendGridRequest -Client $Client -Method GET -Path "/v3/teammates/$encoded/subuser_access" -Query $query
		$hasRestrictedSubuserAccess = [bool](Get-OptionalNestedPropertyValue -InputObject $response -PropertyName 'has_restricted_subuser_access' -Default $false)
		foreach ($row in @(ConvertTo-SendGridArray -InputObject (Get-OptionalNestedPropertyValue -InputObject $response -PropertyName 'subuser_access'))) {
			[void]$result.Add($row)
		}

		$afterSubuserId = $null
		$metadata = Get-OptionalNestedPropertyValue -InputObject $response -PropertyName '_metadata'
		$nextParams = Get-OptionalNestedPropertyValue -InputObject $metadata -PropertyName 'next_params'
		$nextAfterSubuserId = Get-OptionalNestedPropertyValue -InputObject $nextParams -PropertyName 'after_subuser_id'
		if (-not [string]::IsNullOrWhiteSpace([string]$nextAfterSubuserId)) {
			$afterSubuserId = [string]$nextAfterSubuserId
		}
	} while ($afterSubuserId)

	[pscustomobject]@{
		HasRestrictedSubuserAccess = $hasRestrictedSubuserAccess
		SubuserAccess              = $result.ToArray()
	}
}

function Remove-SendGridTeammate {
	[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
	param(
		[Parameter(Mandatory)]
		[pscustomobject]$Client,

		[Parameter(Mandatory)]
		[string]$TeammateName,

		[string]$OnBehalfOf
	)

	$encoded = [uri]::EscapeDataString($TeammateName)
	$target = if ([string]::IsNullOrWhiteSpace($OnBehalfOf)) { $TeammateName } else { "$TeammateName (subuser: $OnBehalfOf)" }
	if ($PSCmdlet.ShouldProcess($target, 'Delete SendGrid teammate')) {
		Invoke-SendGridRequest -Client $Client -Method DELETE -Path "/v3/teammates/$encoded" -OnBehalfOf $OnBehalfOf | Out-Null
	}
}

function Get-SendGridVerifiedSenders {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[pscustomobject]$Client,

		[string]$OnBehalfOf
	)

	Invoke-SendGridRequest -Client $Client -Method GET -Path '/v3/verified_senders' -OnBehalfOf $OnBehalfOf
}

function New-SendGridVerifiedSender {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[pscustomobject]$Client,

		[Parameter(Mandatory)]
		[hashtable]$Sender,

		[string]$OnBehalfOf
	)

	Invoke-SendGridRequest -Client $Client -Method POST -Path '/v3/verified_senders' -Body $Sender -OnBehalfOf $OnBehalfOf
}

function Remove-SendGridVerifiedSender {
	[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
	param(
		[Parameter(Mandatory)]
		[pscustomobject]$Client,

		[Parameter(Mandatory)]
		[string]$SenderId,

		[string]$OnBehalfOf
	)

	$encoded = [uri]::EscapeDataString($SenderId)
	if ($PSCmdlet.ShouldProcess($SenderId, 'Delete SendGrid verified sender')) {
		Invoke-SendGridRequest -Client $Client -Method DELETE -Path "/v3/verified_senders/$encoded" -OnBehalfOf $OnBehalfOf | Out-Null
	}
}

function Get-SendGridInboundParseSettings {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[pscustomobject]$Client,

		[string]$OnBehalfOf,

		[ValidateSet('Global', 'EU')]
		[string]$Region
	)

	$baseUri = if ($PSBoundParameters.ContainsKey('Region')) { Get-SendGridBaseUri -Region $Region } else { $null }
	Invoke-SendGridRequest -Client $Client -Method GET -Path '/v3/user/webhooks/parse/settings' -OnBehalfOf $OnBehalfOf -BaseUri $baseUri
}

function New-SendGridInboundParseSetting {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[pscustomobject]$Client,

		[Parameter(Mandatory)]
		[string]$Hostname,

		[Parameter(Mandatory)]
		[string]$Url,

		[bool]$SpamCheck = $true,

		[bool]$SendRaw = $false,

		[string]$OnBehalfOf,

		[ValidateSet('Global', 'EU')]
		[string]$Region
	)

	$baseUri = if ($PSBoundParameters.ContainsKey('Region')) { Get-SendGridBaseUri -Region $Region } else { $null }
	$body = @{
		hostname   = $Hostname
		url        = $Url
		spam_check = $SpamCheck
		send_raw   = $SendRaw
	}

	Invoke-SendGridRequest -Client $Client -Method POST -Path '/v3/user/webhooks/parse/settings' -Body $body -OnBehalfOf $OnBehalfOf -BaseUri $baseUri
}

function Remove-SendGridInboundParseSetting {
	[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
	param(
		[Parameter(Mandatory)]
		[pscustomobject]$Client,

		[Parameter(Mandatory)]
		[string]$Hostname,

		[string]$OnBehalfOf,

		[ValidateSet('Global', 'EU')]
		[string]$Region
	)

	$baseUri = if ($PSBoundParameters.ContainsKey('Region')) { Get-SendGridBaseUri -Region $Region } else { $null }
	$encoded = [uri]::EscapeDataString($Hostname)
	if ($PSCmdlet.ShouldProcess($Hostname, 'Delete SendGrid inbound parse setting')) {
		Invoke-SendGridRequest -Client $Client -Method DELETE -Path "/v3/user/webhooks/parse/settings/$encoded" -OnBehalfOf $OnBehalfOf -BaseUri $baseUri | Out-Null
	}
}

function Send-SendGridMailMessage {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[pscustomobject]$Client,

		[Parameter(Mandatory)]
		[string]$From,

		[Parameter(Mandatory)]
		[string[]]$To,

		[Parameter(Mandatory)]
		[string]$Subject,

		[string]$PlainTextContent,

		[string]$HtmlContent,

		[string[]]$Cc,

		[string[]]$Bcc,

		[string]$OnBehalfOf
	)

	if ([string]::IsNullOrWhiteSpace($PlainTextContent) -and [string]::IsNullOrWhiteSpace($HtmlContent)) {
		throw 'Provide PlainTextContent, HtmlContent, or both.'
	}

	$personalization = @{
		to = @($To | ForEach-Object { @{ email = $_ } })
	}

	if ($Cc -and $Cc.Count -gt 0) {
		$personalization['cc'] = @($Cc | ForEach-Object { @{ email = $_ } })
	}

	if ($Bcc -and $Bcc.Count -gt 0) {
		$personalization['bcc'] = @($Bcc | ForEach-Object { @{ email = $_ } })
	}

	$content = [System.Collections.Generic.List[hashtable]]::new()
	if (-not [string]::IsNullOrWhiteSpace($PlainTextContent)) {
		$content.Add(@{ type = 'text/plain'; value = $PlainTextContent })
	}
	if (-not [string]::IsNullOrWhiteSpace($HtmlContent)) {
		$content.Add(@{ type = 'text/html'; value = $HtmlContent })
	}

	$body = @{
		from             = @{ email = $From }
		subject          = $Subject
		personalizations = @($personalization)
		content          = $content.ToArray()
	}

	Invoke-SendGridRequest -Client $Client -Method POST -Path '/v3/mail/send' -Body $body -OnBehalfOf $OnBehalfOf -RawResponse
}

