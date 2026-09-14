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


# -----------------------------------------------------------------------------
# SendGrid role/persona catalog (shared by 'Sendgrid - Azure Group Management'
# and 'Sendgrid - Azure Group Apply.ps1').
# -----------------------------------------------------------------------------
# Role/spec catalog. Reconciliation code (Phase 2+) consumes these definitions.
$SendGridRoleSpec = [ordered]@{
	GroupPrefix = 'cs-sendgrid-'

	Admin = [ordered]@{
		Key          = 'admin'
		GroupName    = 'cs-sendgrid-admin'
		RoleKind     = 'administrator'
		AccessTarget = 'main-and-all-subusers'
		Precedence   = 300
		Notes        = 'is_admin = true on the parent SSO teammate (all scopes).'
	}

	AdminReadOnly = [ordered]@{
		Key          = 'admin-ro'
		GroupName    = 'cs-sendgrid-admin-ro'
		RoleKind     = 'persona'
		PersonaSlug  = 'observer'
		AccessTarget = 'parent'
		Precedence   = 200
		Notes        = 'Observer persona scopes on the parent. Cannot be combined with subuser_access.'
	}

	SubuserAccess = [ordered]@{
		Key          = 'subuser-access'
		RoleKind     = 'subuser_access'
		AccessTarget = 'single-subuser'
		Precedence   = 100
		# Format: cs-sendgrid-<subuser>-<role>
		# role 'admin' -> permission_type admin; persona slugs -> permission_type restricted.
		GroupPattern = '^(?i)cs-sendgrid-(?<Subuser>[a-z0-9][a-z0-9-]*)-(?<Role>admin|accountant|developer|marketer|observer)$'
	}
}

# Twilio SendGrid persona scopes (parent account) source:
# https://www.twilio.com/docs/sendgrid/ui/account-and-settings/teammate-permissions#persona-scopes
$SendGridPersonaScopes = [ordered]@{
	accountant = @(
		'billing.create',
		'billing.read',
		'billing.update',
		'billing.delete',
		'mail_settings.read',
		'partner_settings.read',
		'tracking_settings.read',
		'stats.read',
		'stats.global.read',
		'categories.stats.read',
		'categories.stats.sums.read',
		'devices.stats.read',
		'clients.stats.read',
		'clients.phone.stats.read',
		'clients.tablet.stats.read',
		'clients.webmail.stats.read',
		'clients.desktop.stats.read',
		'geo.stats.read',
		'mailbox_providers.stats.read',
		'browsers.stats.read',
		'subusers.stats.read',
		'subusers.stats.sums.read',
		'subusers.stats.monthly.read',
		'user.webhooks.parse.stats.read',
		'user.account.read',
		'user.credits.read',
		'user.email.read',
		'user.profile.read',
		'user.profile.update',
		'user.timezone.read',
		'user.username.read',
		'user.settings.enforced_tls.read',
		'categories.read',
		'sender_verification_eligible',
		'sender_verification_legacy',
		'2fa_exempt',
		'2fa_required'
	)

	developer = @(
		'alerts.create',
		'alerts.read',
		'alerts.update',
		'alerts.delete',
		'asm.groups.create',
		'asm.groups.read',
		'asm.groups.update',
		'asm.groups.delete',
		'ips.warmup.create',
		'ips.warmup.read',
		'ips.warmup.update',
		'ips.warmup.delete',
		'ips.pools.create',
		'ips.pools.read',
		'ips.pools.update',
		'ips.pools.delete',
		'ips.pools.ips.create',
		'ips.pools.ips.read',
		'ips.pools.ips.update',
		'ips.pools.ips.delete',
		'ips.assigned.read',
		'ips.create',
		'ips.read',
		'ips.update',
		'ips.delete',
		'mail.send',
		'mail_settings.read',
		'mail_settings.bcc.read',
		'mail_settings.bcc.update',
		'mail_settings.address_whitelist.read',
		'mail_settings.address_whitelist.update',
		'mail_settings.footer.read',
		'mail_settings.footer.update',
		'mail_settings.forward_spam.read',
		'mail_settings.forward_spam.update',
		'mail_settings.plain_content.read',
		'mail_settings.plain_content.update',
		'mail_settings.spam_check.read',
		'mail_settings.spam_check.update',
		'mail_settings.bounce_purge.read',
		'mail_settings.bounce_purge.update',
		'mail_settings.forward_bounce.read',
		'mail_settings.forward_bounce.update',
		'partner_settings.read',
		'tracking_settings.read',
		'tracking_settings.click.read',
		'tracking_settings.click.update',
		'tracking_settings.subscription.read',
		'tracking_settings.subscription.update',
		'tracking_settings.open.read',
		'tracking_settings.open.update',
		'tracking_settings.google_analytics.read',
		'tracking_settings.google_analytics.update',
		'user.webhooks.event.settings.read',
		'user.webhooks.event.settings.update',
		'user.webhooks.event.test.create',
		'user.webhooks.event.test.read',
		'user.webhooks.event.test.update',
		'user.webhooks.parse.settings.create',
		'user.webhooks.parse.settings.read',
		'user.webhooks.parse.settings.update',
		'user.webhooks.parse.settings.delete',
		'stats.read',
		'stats.global.read',
		'categories.stats.read',
		'categories.stats.sums.read',
		'devices.stats.read',
		'clients.stats.read',
		'clients.phone.stats.read',
		'clients.tablet.stats.read',
		'clients.webmail.stats.read',
		'clients.desktop.stats.read',
		'geo.stats.read',
		'mailbox_providers.stats.read',
		'browsers.stats.read',
		'subusers.stats.read',
		'subusers.stats.sums.read',
		'subusers.stats.monthly.read',
		'user.webhooks.parse.stats.read',
		'templates.create',
		'templates.read',
		'templates.update',
		'templates.delete',
		'templates.versions.create',
		'templates.versions.read',
		'templates.versions.update',
		'templates.versions.delete',
		'templates.versions.activate.create',
		'user.account.read',
		'user.credits.read',
		'user.email.read',
		'user.profile.read',
		'user.profile.update',
		'user.timezone.read',
		'user.username.read',
		'user.settings.enforced_tls.read',
		'api_keys.create',
		'api_keys.read',
		'api_keys.update',
		'api_keys.delete',
		'categories.create',
		'categories.read',
		'categories.update',
		'categories.delete',
		'mail_settings.template.read',
		'mail_settings.template.update',
		'marketing_campaigns.create',
		'marketing_campaigns.read',
		'marketing_campaigns.update',
		'marketing_campaigns.delete',
		'mail.batch.create',
		'mail.batch.read',
		'mail.batch.update',
		'mail.batch.delete',
		'user.scheduled_sends.create',
		'user.scheduled_sends.read',
		'user.scheduled_sends.update',
		'user.scheduled_sends.delete',
		'access_settings.whitelist.create',
		'access_settings.whitelist.read',
		'access_settings.whitelist.update',
		'access_settings.whitelist.delete',
		'access_settings.activity.read',
		'suppression.create',
		'suppression.read',
		'suppression.update',
		'suppression.delete',
		'email_testing.read',
		'email_testing.write',
		'sender_verification_eligible',
		'sender_verification_legacy',
		'2fa_exempt',
		'2fa_required'
	)

	marketer = @(
		'alerts.create',
		'alerts.read',
		'alerts.update',
		'alerts.delete',
		'asm.groups.create',
		'asm.groups.read',
		'asm.groups.update',
		'asm.groups.delete',
		'mail_settings.read',
		'mail_settings.spam_check.read',
		'mail_settings.spam_check.update',
		'partner_settings.read',
		'tracking_settings.read',
		'tracking_settings.click.read',
		'tracking_settings.click.update',
		'tracking_settings.subscription.read',
		'tracking_settings.subscription.update',
		'tracking_settings.open.read',
		'tracking_settings.open.update',
		'tracking_settings.google_analytics.read',
		'tracking_settings.google_analytics.update',
		'stats.global.read',
		'categories.stats.read',
		'categories.stats.sums.read',
		'devices.stats.read',
		'clients.stats.read',
		'clients.phone.stats.read',
		'clients.tablet.stats.read',
		'clients.webmail.stats.read',
		'clients.desktop.stats.read',
		'geo.stats.read',
		'mailbox_providers.stats.read',
		'browsers.stats.read',
		'subusers.stats.read',
		'subusers.stats.sums.read',
		'subusers.stats.monthly.read',
		'user.webhooks.parse.stats.read',
		'templates.create',
		'templates.read',
		'templates.update',
		'templates.delete',
		'templates.versions.create',
		'templates.versions.read',
		'templates.versions.update',
		'templates.versions.delete',
		'templates.versions.activate.create',
		'user.account.read',
		'user.credits.read',
		'user.email.read',
		'user.profile.read',
		'user.profile.update',
		'user.timezone.read',
		'user.username.read',
		'user.settings.enforced_tls.read',
		'categories.read',
		'marketing_campaigns.create',
		'marketing_campaigns.read',
		'marketing_campaigns.update',
		'marketing_campaigns.delete',
		'mail.batch.read',
		'user.scheduled_sends.read',
		'suppression.create',
		'suppression.read',
		'suppression.update',
		'suppression.delete',
		'email_testing.read',
		'email_testing.write',
		'sender_verification_eligible',
		'sender_verification_legacy',
		'2fa_exempt',
		'2fa_required'
	)

	observer = @(
		'alerts.read',
		'asm.groups.read',
		'billing.read',
		'ips.warmup.read',
		'ips.pools.read',
		'ips.pools.ips.read',
		'ips.assigned.read',
		'ips.read',
		'mail_settings.read',
		'mail_settings.bcc.read',
		'mail_settings.address_whitelist.read',
		'mail_settings.footer.read',
		'mail_settings.forward_spam.read',
		'mail_settings.plain_content.read',
		'mail_settings.spam_check.read',
		'mail_settings.bounce_purge.update',
		'mail_settings.forward_bounce.read',
		'partner_settings.read',
		'partner_settings.new_relic.read',
		'partner_settings.sendwithus.read',
		'tracking_settings.read',
		'tracking_settings.click.read',
		'tracking_settings.subscription.read',
		'tracking_settings.open.read',
		'tracking_settings.google_analytics.read',
		'user.webhooks.event.settings.read',
		'user.webhooks.event.test.read',
		'user.webhooks.parse.settings.read',
		'stats.read',
		'stats.global.read',
		'categories.stats.read',
		'categories.stats.sums.read',
		'devices.stats.read',
		'clients.stats.read',
		'clients.phone.stats.read',
		'clients.tablet.stats.read',
		'clients.webmail.stats.read',
		'clients.desktop.stats.read',
		'geo.stats.read',
		'mailbox_providers.stats.read',
		'browsers.stats.read',
		'subusers.stats.read',
		'subusers.stats.sums.read',
		'subusers.stats.monthly.read',
		'user.webhooks.parse.stats.read',
		'subusers.read',
		'subusers.monitor.read',
		'subusers.credits.read',
		'subusers.credits.remaining.read',
		'subusers.reputations.read',
		'subusers.summary.read',
		'templates.read',
		'templates.versions.read',
		'user.account.read',
		'user.credits.read',
		'user.email.read',
		'user.profile.read',
		'user.profile.update',
		'user.timezone.read',
		'user.username.read',
		'user.settings.enforced_tls.read',
		'api_keys.read',
		'categories.read',
		'mail_settings.template.read',
		'mail.batch.read',
		'user.scheduled_sends.read',
		'access_settings.whitelist.read',
		'access_settings.activity.read',
		'suppression.read',
		'messages.read',
		'email_testing.read',
		'sender_verification_eligible',
		'sender_verification_legacy',
		'2fa_exempt',
		'2fa_required'
	)
}

# Scopes SendGrid accepts on a subuser_access entry with permission_type = restricted.
# This is a different (smaller) list than the parent persona scopes. Source:
# https://support.sendgrid.com/hc/en-us/articles/27274820796059-Twilio-SendGrid-SSO-Teammate-Permissions-for-a-Subuser
$SendGridSubuserRestrictedScopes = @(
	'access_settings.activity.read',
	'access_settings.whitelist.create',
	'access_settings.whitelist.delete',
	'access_settings.whitelist.read',
	'access_settings.whitelist.update',
	'alerts.create',
	'alerts.delete',
	'alerts.read',
	'alerts.update',
	'api_keys.create',
	'api_keys.delete',
	'api_keys.read',
	'api_keys.update',
	'asm.groups.create',
	'asm.groups.delete',
	'asm.groups.read',
	'asm.groups.suppressions.create',
	'asm.groups.suppressions.delete',
	'asm.groups.suppressions.read',
	'asm.groups.suppressions.update',
	'asm.groups.update',
	'asm.suppressions.global.create',
	'asm.suppressions.global.delete',
	'asm.suppressions.global.read',
	'asm.suppressions.global.update',
	'browsers.stats.read',
	'categories.create',
	'categories.delete',
	'categories.read',
	'categories.stats.read',
	'categories.stats.sums.read',
	'categories.update',
	'clients.desktop.stats.read',
	'clients.phone.stats.read',
	'clients.stats.read',
	'clients.tablet.stats.read',
	'clients.webmail.stats.read',
	'credentials.create',
	'credentials.delete',
	'credentials.read',
	'credentials.update',
	'design_library.create',
	'design_library.delete',
	'design_library.read',
	'design_library.update',
	'devices.stats.read',
	'di.bounce_block_classification.read',
	'email_testing.read',
	'email_testing.write',
	'geo.stats.read',
	'ips.assigned.read',
	'ips.pools.create',
	'ips.pools.delete',
	'ips.pools.ips.create',
	'ips.pools.ips.delete',
	'ips.pools.ips.read',
	'ips.pools.ips.update',
	'ips.pools.read',
	'ips.pools.update',
	'ips.warmup.create',
	'ips.warmup.delete',
	'ips.warmup.read',
	'ips.warmup.update',
	'mail.batch.create',
	'mail.batch.delete',
	'mail.batch.read',
	'mail.batch.update',
	'mail.send',
	'mail_settings.address_whitelist.create',
	'mail_settings.address_whitelist.delete',
	'mail_settings.address_whitelist.read',
	'mail_settings.address_whitelist.update',
	'mail_settings.bcc.create',
	'mail_settings.bcc.delete',
	'mail_settings.bcc.read',
	'mail_settings.bcc.update',
	'mail_settings.bounce_purge.create',
	'mail_settings.bounce_purge.delete',
	'mail_settings.bounce_purge.read',
	'mail_settings.bounce_purge.update',
	'mail_settings.footer.create',
	'mail_settings.footer.delete',
	'mail_settings.footer.read',
	'mail_settings.footer.update',
	'mail_settings.forward_bounce.create',
	'mail_settings.forward_bounce.delete',
	'mail_settings.forward_bounce.read',
	'mail_settings.forward_bounce.update',
	'mail_settings.forward_spam.create',
	'mail_settings.forward_spam.delete',
	'mail_settings.forward_spam.read',
	'mail_settings.forward_spam.update',
	'mail_settings.plain_content.create',
	'mail_settings.plain_content.delete',
	'mail_settings.plain_content.read',
	'mail_settings.plain_content.update',
	'mail_settings.read',
	'mail_settings.spam_check.create',
	'mail_settings.spam_check.delete',
	'mail_settings.spam_check.read',
	'mail_settings.spam_check.update',
	'mail_settings.template.create',
	'mail_settings.template.delete',
	'mail_settings.template.read',
	'mail_settings.template.update',
	'mailbox_providers.stats.read',
	'marketing_campaigns.create',
	'marketing_campaigns.delete',
	'marketing_campaigns.read',
	'marketing_campaigns.update',
	'marketing.read',
	'marketing.automation.read',
	'messages.read',
	'partner_settings.new_relic.create',
	'partner_settings.new_relic.delete',
	'partner_settings.new_relic.read',
	'partner_settings.new_relic.update',
	'partner_settings.read',
	'partner_settings.sendwithus.create',
	'partner_settings.sendwithus.delete',
	'partner_settings.sendwithus.read',
	'partner_settings.sendwithus.update',
	'recipients.erasejob.create',
	'recipients.erasejob.read',
	'stats.global.read',
	'stats.read',
	'suppression.blocks.create',
	'suppression.blocks.delete',
	'suppression.blocks.read',
	'suppression.blocks.update',
	'suppression.bounces.create',
	'suppression.bounces.delete',
	'suppression.bounces.read',
	'suppression.bounces.update',
	'suppression.create',
	'suppression.delete',
	'suppression.invalid_emails.create',
	'suppression.invalid_emails.delete',
	'suppression.invalid_emails.read',
	'suppression.invalid_emails.update',
	'suppression.read',
	'suppression.spam_reports.create',
	'suppression.spam_reports.delete',
	'suppression.spam_reports.read',
	'suppression.spam_reports.update',
	'suppression.unsubscribes.create',
	'suppression.unsubscribes.delete',
	'suppression.unsubscribes.read',
	'suppression.unsubscribes.update',
	'suppression.update',
	'templates.create',
	'templates.delete',
	'templates.read',
	'templates.update',
	'templates.versions.activate.create',
	'templates.versions.activate.delete',
	'templates.versions.activate.read',
	'templates.versions.activate.update',
	'templates.versions.create',
	'templates.versions.delete',
	'templates.versions.read',
	'templates.versions.update',
	'tracking_settings.click.create',
	'tracking_settings.click.delete',
	'tracking_settings.click.read',
	'tracking_settings.click.update',
	'tracking_settings.google_analytics.create',
	'tracking_settings.google_analytics.delete',
	'tracking_settings.google_analytics.read',
	'tracking_settings.google_analytics.update',
	'tracking_settings.open.create',
	'tracking_settings.open.delete',
	'tracking_settings.open.read',
	'tracking_settings.open.update',
	'tracking_settings.read',
	'tracking_settings.subscription.create',
	'tracking_settings.subscription.delete',
	'tracking_settings.subscription.read',
	'tracking_settings.subscription.update',
	'user.account.read',
	'user.credits.read',
	'user.email.read',
	'user.scheduled_sends.create',
	'user.scheduled_sends.delete',
	'user.scheduled_sends.read',
	'user.scheduled_sends.update',
	'user.settings.enforced_tls.read',
	'user.settings.enforced_tls.update',
	'user.timezone.create',
	'user.timezone.delete',
	'user.timezone.read',
	'user.timezone.update',
	'user.username.read',
	'user.webhooks.event.settings.create',
	'user.webhooks.event.settings.delete',
	'user.webhooks.event.settings.read',
	'user.webhooks.event.settings.update',
	'user.webhooks.event.test.create',
	'user.webhooks.event.test.delete',
	'user.webhooks.event.test.read',
	'user.webhooks.event.test.update',
	'user.webhooks.parse.settings.create',
	'user.webhooks.parse.settings.delete',
	'user.webhooks.parse.settings.read',
	'user.webhooks.parse.settings.update',
	'user.webhooks.parse.stats.read',
	'whitelabel.create',
	'whitelabel.delete',
	'whitelabel.read',
	'whitelabel.update'
)

# Subuser persona templates: persona scopes limited to what restricted subuser_access accepts.
# These are the scope payloads used for permission_type = restricted entries.
$SendGridSubuserPersonaScopes = [ordered]@{}
$SendGridSubuserPersonaDroppedScopes = [ordered]@{}
$subuserAllowedScopeSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($allowedScope in $SendGridSubuserRestrictedScopes) {
	[void]$subuserAllowedScopeSet.Add($allowedScope)
}
foreach ($personaKey in @($SendGridPersonaScopes.Keys)) {
	$SendGridSubuserPersonaScopes[$personaKey] = @($SendGridPersonaScopes[$personaKey] | Where-Object { $subuserAllowedScopeSet.Contains($_) })
	$SendGridSubuserPersonaDroppedScopes[$personaKey] = @($SendGridPersonaScopes[$personaKey] | Where-Object { -not $subuserAllowedScopeSet.Contains($_) })
}

# The SendGrid UI gates the subuser "Marketing" tab behind the newer marketing.*
# read scopes, not the legacy marketing_campaigns.* scopes the documented
# persona lists carry. Personas that grant campaigns access (marketer,
# developer) get them added so the tab actually shows for those teammates.
$SendGridSubuserMarketingUiScopes = @('marketing.read', 'marketing.automation.read')
foreach ($personaKey in @($SendGridSubuserPersonaScopes.Keys)) {
	if (@($SendGridSubuserPersonaScopes[$personaKey]) -contains 'marketing_campaigns.read') {
		$SendGridSubuserPersonaScopes[$personaKey] = @(@($SendGridSubuserPersonaScopes[$personaKey]) + $SendGridSubuserMarketingUiScopes | Select-Object -Unique)
	}
}

function Get-SendGridAdminRoleSpec {
	[CmdletBinding()]
	param()

	return [pscustomobject]$SendGridRoleSpec.Admin
}

function Get-SendGridAdminReadOnlyRoleSpec {
	[CmdletBinding()]
	param()

	return [pscustomobject]$SendGridRoleSpec.AdminReadOnly
}

function Get-SendGridPersonaScopes {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[ValidateSet('accountant', 'developer', 'marketer', 'observer')]
		[string]$Persona,

		# Parent = documented persona list. Subuser = payload for a restricted subuser_access entry.
		[ValidateSet('Parent', 'Subuser')]
		[string]$Scope = 'Parent'
	)

	$key = $Persona.ToLowerInvariant()
	if ($Scope -eq 'Subuser') {
		return @($SendGridSubuserPersonaScopes[$key])
	}

	return @($SendGridPersonaScopes[$key])
}

function Get-SupportedPersonaSlugs {
	[CmdletBinding()]
	param()

	return @($SendGridPersonaScopes.Keys | Sort-Object)
}

function Resolve-SendGridRoleFromGroupName {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[string]$GroupName
	)

	$groupNameNormalized = $GroupName.Trim().ToLowerInvariant()

	if ($groupNameNormalized -eq $SendGridRoleSpec.Admin.GroupName) {
		return [pscustomobject]@{
			RoleKey        = $SendGridRoleSpec.Admin.Key
			GroupName      = $groupNameNormalized
			AccessTarget   = $SendGridRoleSpec.Admin.AccessTarget
			Subuser        = $null
			Role           = 'admin'
			PermissionType = $null
			Persona        = $null
		}
	}

	if ($groupNameNormalized -eq $SendGridRoleSpec.AdminReadOnly.GroupName) {
		return [pscustomobject]@{
			RoleKey        = $SendGridRoleSpec.AdminReadOnly.Key
			GroupName      = $groupNameNormalized
			AccessTarget   = $SendGridRoleSpec.AdminReadOnly.AccessTarget
			Subuser        = $null
			Role           = $SendGridRoleSpec.AdminReadOnly.PersonaSlug
			PermissionType = $null
			Persona        = $SendGridRoleSpec.AdminReadOnly.PersonaSlug
		}
	}

	$groupPattern = [string]$SendGridRoleSpec.SubuserAccess.GroupPattern
	$match = [regex]::Match($groupNameNormalized, $groupPattern)

	if (-not $match.Success) {
		return $null
	}

	$role = $match.Groups['Role'].Value
	$isSubuserAdmin = ($role -eq 'admin')

	return [pscustomobject]@{
		RoleKey        = $SendGridRoleSpec.SubuserAccess.Key
		GroupName      = $groupNameNormalized
		AccessTarget   = $SendGridRoleSpec.SubuserAccess.AccessTarget
		Subuser        = $match.Groups['Subuser'].Value
		Role           = $role
		PermissionType = if ($isSubuserAdmin) { 'admin' } else { 'restricted' }
		Persona        = if ($isSubuserAdmin) { $null } else { $role }
	}
}


function ConvertTo-GroupToken {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[string]$Value
	)

	$normalized = $Value.Trim().ToLowerInvariant()
	$normalized = [regex]::Replace($normalized, '[^a-z0-9]+', '-')
	$normalized = [regex]::Replace($normalized, '-{2,}', '-')
	$normalized = $normalized.Trim('-')

	if ([string]::IsNullOrWhiteSpace($normalized)) {
		throw "Could not build a valid group token from '$Value'."
	}

	return $normalized
}

function Resolve-SendGridPersonaFromScopes {
	# Matches an actual scope list to a persona template.
	# Exact match wins. Otherwise the persona whose template is fully contained in
	# the actual scopes with the fewest extras wins (SendGrid adds a minimum scope
	# set of its own, so exact matches are not guaranteed). No match -> Persona = $null.
	[CmdletBinding()]
	param(
		[string[]]$Scopes,

		[ValidateSet('Parent', 'Subuser')]
		[string]$Scope = 'Parent'
	)

	$templates = if ($Scope -eq 'Subuser') { $SendGridSubuserPersonaScopes } else { $SendGridPersonaScopes }

	$actualSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
	foreach ($scopeName in @($Scopes)) {
		if (-not [string]::IsNullOrWhiteSpace([string]$scopeName)) {
			[void]$actualSet.Add(([string]$scopeName).Trim())
		}
	}

	$noMatch = [pscustomobject]@{
		Persona     = $null
		MatchType   = 'none'
		ExtraScopes = @()
	}

	if ($actualSet.Count -eq 0) {
		return $noMatch
	}

	$best = $null
	foreach ($persona in @($templates.Keys)) {
		$template = @($templates[$persona])
		if ($template.Count -eq 0) {
			continue
		}

		# Best-effort scopes (the Marketing-tab pair) may be silently dropped by
		# SendGrid; their absence must not break persona matching.
		$missing = @($template | Where-Object { -not $actualSet.Contains($_) -and $SendGridSubuserMarketingUiScopes -notcontains $_ })
		if ($missing.Count -gt 0) {
			continue
		}

		$templateSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
		foreach ($templateScope in $template) {
			[void]$templateSet.Add($templateScope)
		}

		$extra = @($actualSet | Where-Object { -not $templateSet.Contains($_) } | Sort-Object)

		if ($null -eq $best -or $extra.Count -lt @($best.ExtraScopes).Count) {
			$best = [pscustomobject]@{
				Persona     = $persona
				MatchType   = if ($extra.Count -eq 0) { 'exact' } else { 'superset' }
				ExtraScopes = $extra
			}
		}
	}

	if ($null -ne $best) {
		return $best
	}

	return $noMatch
}

function Export-SendGridPlanCsv {
	# Writes a header-only file when there are no rows so every run produces the full file set.
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[string]$Path,

		[Parameter(Mandatory)]
		[string[]]$Columns,

		[object[]]$Rows
	)

	$rowsArray = @($Rows | Where-Object { $null -ne $_ })
	if ($rowsArray.Count -gt 0) {
		$rowsArray | Select-Object -Property $Columns | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
	}
	else {
		Set-Content -LiteralPath $Path -Value ('"{0}"' -f ($Columns -join '","')) -Encoding UTF8
	}

	Write-Host ("  {0,-40} {1} row(s)" -f (Split-Path -Path $Path -Leaf), $rowsArray.Count)
}
