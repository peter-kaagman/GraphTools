# Private functions


function Invoke-GraphRequest {
<#
.SYNOPSIS
Executes a Microsoft Graph REST request.

.DESCRIPTION
Internal helper function that wraps Invoke-RestMethod for Microsoft Graph API calls.
Supports GET, POST, PATCH and DELETE operations and automatically serializes request bodies to JSON.

.PARAMETER Headers
Graph authentication headers.

.PARAMETER Uri
Microsoft Graph endpoint URI.

.PARAMETER Method
HTTP method to execute.

.PARAMETER Body
Optional request body. Automatically converted to JSON when supplied.

.OUTPUTS
System.Object

.EXAMPLE
Invoke-GraphRequest `
    -Headers $headers `
    -Uri "https://graph.microsoft.com/v1.0/groups" `
    -Method GET

.EXAMPLE
Invoke-GraphRequest `
    -Headers $headers `
    -Uri "https://graph.microsoft.com/v1.0/teams/$TeamId/archive" `
    -Method POST
#>	
	[CmdletBinding()]
	param (
		[Parameter(Mandatory)] [System.Collections.Generic.Dictionary[[String], [String]]] $Headers,
		[Parameter(Mandatory)] [string] $Uri,
		[ValidateSet("GET","POST","PATCH","DELETE")] [string] $Method = "GET",
		[object] $Body
	)
	$requestSplat = @{
		Uri         = $Uri
		Headers     = $Headers
		Method      = $Method
		ErrorAction = "Stop"
	}
	if ($PSBoundParameters.ContainsKey('Body')) {
		$requestSplat.Body = $Body | ConvertTo-Json -Depth 10
		$requestSplat.ContentType = 'application/json'
	}
	$response = Invoke-RestMethod @requestSplat
	return $response
}

# Public functions for GraphTools

# Group collection functions

function Get-GraphCollection {
<#
.SYNOPSIS
Retrieve a paged collection from Microsoft Graph.

.DESCRIPTION
Retrieves all items from a Microsoft Graph collection endpoint and automatically follows @odata.nextLink pages.

When -HashTableKeyProperty is specified, results are returned as a dictionary indexed by the specified property. Otherwise an ArrayList is returned.

.PARAMETER Headers
Authentication headers used for Graph requests.

.PARAMETER Uri
Microsoft Graph collection endpoint URI.

.PARAMETER Filter
Optional OData filter expression.

.PARAMETER Select
Optional OData select expression.

.PARAMETER HashTableKeyProperty
Property used as dictionary key. When specified, results are returned as a Dictionary[string,object].

.EXAMPLE
$groups = Get-GraphCollection `
    -Headers $Headers `
    -Uri "https://graph.microsoft.com/v1.0/groups"

.EXAMPLE
$groups = Get-GraphCollection `
    -Headers $Headers `
    -Uri "https://graph.microsoft.com/v1.0/groups" `
    -HashTableKeyProperty "mailNickname"

.NOTES
Designed as an internal helper function for retrieving Graph collections.
#>
	[CmdletBinding()]
	param (
		[Parameter(Mandatory)] [System.Collections.Generic.Dictionary[[String], [String]]] $Headers,
		[Parameter(Mandatory)] [string] $Uri,
		[string] $Filter,
		[string] $Select,
		[string] $HashTableKeyProperty
	)
	# Add filter if present
	$separator = if ($Uri -match '\?') { '&' } else { '?' }
	if ($Filter) {
		$encodedFilter = [System.Uri]::EscapeDataString($Filter)
		$Uri += "$separator`$filter=$encodedFilter"
		$separator = '&'
	}

	if ($Select) {
		$encodedSelect = [System.Uri]::EscapeDataString($Select)
		$Uri += "$separator`$select=$encodedSelect"
	}

	$requestSplat = @{
		Uri         = $Uri
		Headers     = $Headers
		Method      = "GET"
		Verbose     = $false
		ErrorAction = "Stop"
	}

	if ($HashTableKeyProperty) {
		$result = @{}
	} else {
		$result = [System.Collections.ArrayList]::new()
	}

	$response = Invoke-RestMethod @requestSplat
	while ($null -ne $response) {
		if ($response.Value) {
			foreach ($item in $response.Value) {
				if ($HashTableKeyProperty) {
					$key = [string]$item.$HashTableKeyProperty
					if ([string]::IsNullOrEmpty($key) ) {
						throw "Item is missing property '$HashTableKeyProperty'."
					}
					if ($result.ContainsKey($key)) {
						throw "Duplicate key '$key' found in results."
					}
					$result[$key] = $item
				} else {
					[void]$result.Add($item)
				}
			}				
		}
		if ($response.'@odata.nextlink') {
			# If there is a next link, we need to call the API again
			$requestSplat.Uri = $response.'@odata.nextlink'
			$response = Invoke-RestMethod @requestSplat
		} else {
			# No next link, we are done
			$response = $null
		}
	}	
	return $result
}

function Get-GraphHeaders {
<#
.SYNOPSIS
Create authentication headers for Microsoft Graph.

.DESCRIPTION
Obtains an application access token using the OAuth2 client credentials flow and returns a header dictionary that can be used with Graph requests.

.PARAMETER TenantId
Microsoft Entra tenant ID.

.PARAMETER ClientId
Application (client) ID.

.PARAMETER ClientSecret
Application client secret value.

.EXAMPLE
$Headers = Get-GraphHeaders `
    -TenantId $TenantId `
    -ClientId $ClientId `
    -ClientSecret $ClientSecret

.NOTES
Uses the Microsoft Graph v2.0 token endpoint.
#>
	[CmdletBinding()]
    param (
	[Parameter(Mandatory)][string] $TenantId,
	[Parameter(Mandatory)][string] $ClientId,
	[Parameter(Mandatory)][string] $ClientSecret,
	[string] $Scope = "https://graph.microsoft.com/.default" # default scope for Microsoft Graph
    )
    try {
        # Write-Verbose "Creating Access Token"
        $baseUri = "https://login.microsoftonline.com"
        $authUri = $baseUri + "/$TenantId/oauth2/v2.0/token"
    
        $body = @{
            grant_type		= "client_credentials"
            client_id		= $ClientId
            client_secret	= $ClientSecret
            scope		= $Scope #"https://graph.microsoft.com/.default"
        }
    
	Write-Verbose "Requesting Access Token for AppId [$ClientId] in tenant [$TenantId]."
        $Response = Invoke-RestMethod -Method POST -Uri $authUri -Body $body -ContentType 'application/x-www-form-urlencoded'
	Write-Verbose "Access Token acquired for AppId [$ClientId] in tenant [$TenantId]."
	Write-Verbose "Response: $($Response | ConvertTo-Json -Depth 10)"
        $accessToken = $Response.access_token
    
        #Add the authorization header to the request
        # Write-Verbose 'Adding Authorization headers'

        $headers = [System.Collections.Generic.Dictionary[[String], [String]]]::new()
        $headers.Add('Authorization', "Bearer $accessToken")
        $headers.Add('Accept', 'application/json')
        # Needed to filter on specific attributes (https://docs.microsoft.com/en-us/graph/aad-advanced-queries)
	# caller specifiek laten toevoegen indien nodig.
        # $headers.Add('ConsistencyLevel', 'eventual')

        return $headers
    }
    catch {
        $errorMessage = $_.Exception.Message
        if ($errorMessage -match "AADSTS7000215|invalid_client") {
            $clientSecretHint = if ($ClientSecret -match '^[0-9a-fA-F-]{36}$') { "The configured secret looks like a secret ID GUID instead of the secret value." } else { "The configured secret was rejected by Microsoft Entra ID." }
            throw "Unable to acquire a Graph access token for AppId [$ClientId] in tenant [$TenantId]. $clientSecretHint Check whether the secret is current, belongs to this app registration, and is stored as the secret value rather than the secret ID."
        }

        throw $_
    }

}	

function Get-Groups {
<#
.SYNOPSIS
Retrieve Microsoft 365 Groups.

.DESCRIPTION
Retrieves Groups from the Microsoft Graph /groups endpoint.

Returns only Group properties available from the Groups endpoint, such as Id,
DisplayName, Mail, MailNickname, GroupTypes and ResourceProvisioningOptions.

Use Get-Teams to retrieve only Groups that have Team capabilities.

.PARAMETER Headers
Authentication headers used for Graph requests.

.PARAMETER Select
Optional OData select expression.

.PARAMETER Filter
Optional OData filter expression.

.PARAMETER HashTableKeyProperty
Property used as dictionary key.

.EXAMPLE
$groups = Get-Groups -Headers $Headers

.EXAMPLE
$groups = Get-Groups `
    -Headers $Headers `
    -HashTableKeyProperty mailNickname

.NOTES
Uses the Microsoft Graph /groups endpoint.
#>	
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)] [System.Collections.Generic.Dictionary[[String], [String]]] $Headers,
        [string] $Select,
        [string] $Filter,
        [string] $HashTableKeyProperty
    )

	# Define URI for reading groups
	$uri = "https://graph.microsoft.com/v1.0/groups"
	$groupsSplat = @{
		Headers = $Headers
		Uri = $uri
	}
	if ($Select) {
		$groupsSplat.Select = $Select
	}
	if ($Filter) {
		$groupsSplat.Filter = $Filter
	}
	if ($HashTableKeyProperty) {
		$groupsSplat.HashTableKeyProperty = $HashTableKeyProperty
	}
	$groups = Get-GraphCollection @groupsSplat
	return $groups
}

function Get-Teams {
<#
.SYNOPSIS
Retrieve Microsoft 365 Groups that have Team capabilities.

.DESCRIPTION
Retrieves Teams by querying the Microsoft Graph Groups endpoint and filtering on
resourceProvisioningOptions.

Because the Groups endpoint is used, only Group properties are returned.
Team-specific properties such as IsArchived, MemberSettings, GuestSettings and
MessagingSettings are not included.

.PARAMETER Headers
Authentication headers used for Graph requests.

.PARAMETER Select
Optional OData select expression.

.PARAMETER Filter
Optional OData filter expression. Combined with the Team filter.

.PARAMETER HashTableKeyProperty
Property used as dictionary key.

.EXAMPLE
$teams = Get-Teams -Headers $Headers

.EXAMPLE
$teams = Get-Teams `
    -Headers $Headers `
    -HashTableKeyProperty mailNickname

.NOTES
Uses the /groups endpoint rather than the /teams endpoint.
Use team-specific functions when Team properties are required.
#>	
	[CmdletBinding()]
	param (
		[Parameter(Mandatory)] [System.Collections.Generic.Dictionary[[String], [String]]] $Headers,
		[string] $Select,
		[string] $Filter,
		[string] $HashTableKeyProperty
	)

	$teamFilter = "resourceProvisioningOptions/Any(x:x eq 'Team')"

	if ($Filter) {
		$teamFilter = "$teamFilter and ($Filter)"
	}

	$groupsSplat = @{
		Headers = $Headers
		Filter = $teamFilter
	}
	if ($Select) {
		$groupsSplat.Select = $Select
	}
	if ($HashTableKeyProperty) {
		$groupsSplat.HashTableKeyProperty = $HashTableKeyProperty
	}

	$teams = Get-Groups @groupsSplat
	return $teams
}

# Functions for Group/Team-specific operations
#
function Get-GroupByMailNickname {
<#
.SYNOPSIS
Get a group by mailNickname.

.DESCRIPTION
Returns the first matching Microsoft 365 group.

.PARAMETER Headers
Authentication headers for Microsoft Graph.

.PARAMETER MailNickname
The group's mailNickname.

.PARAMETER Select
Optional comma-separated list of properties to return.
#>
	[CmdletBinding()]
	param (
		[Parameter(Mandatory)] [System.Collections.Generic.Dictionary[[String], [String]]] $Headers,
		[Parameter(Mandatory)][string] $MailNickname,
		[string] $Select
	)
	$groupSplat = @{
		Headers = $Headers
		Filter = "mailNickname eq '$MailNickname'"
	}
	if ($Select) {
		$groupSplat.Select = $Select
	}

	return Get-Groups @groupSplat| Select-Object -First 1
}

function Get-GroupById {
<#
.SYNOPSIS
Get a group by id.

.DESCRIPTION
Retrieves a Microsoft 365 group directly by its Graph id.

.PARAMETER Headers
Authentication headers for Microsoft Graph.

.PARAMETER Id
The group's unique identifier.

.PARAMETER Select
Optional comma-separated list of properties to return.
#>	[CmdletBinding()]
	param (
		[Parameter(Mandatory)] [System.Collections.Generic.Dictionary[[String], [String]]] $Headers,
		[Parameter(Mandatory)][string] $Id,
		[string] $Select
	)

	$uri = "https://graph.microsoft.com/v1.0/groups/$Id"
	if ($Select) {
		$uri += "?`$select=$Select"
	}
	# Write-Host "Getting group wiht URI: $($uri)"
	$groupSplat = @{
		Headers = $Headers
		Uri = $uri
		Method = "GET"
	}

	$result = Invoke-GraphRequest @groupSplat
	return $result
}

# Tests
#
function Test-TeamArchived{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory)] [System.Collections.Generic.Dictionary[[String], [String]]] $Headers,
		[Parameter(Mandatory)][string] $Id
	)
	$testSplat = @{
		Headers = $Headers
		Uri = "https://graph.microsoft.com/v1.0/teams/$($Id)?`$select=isArchived"
	}
	$response = Invoke-GraphRequest @testSplat
	return $response.isArchived
}

function Test-IsTeam{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory)] [System.Collections.Generic.Dictionary[[String], [String]]] $Headers,
		[Parameter(Mandatory)][string] $Id
	)
	$testSplat = @{
		Headers = $Headers
		Uri = "https://graph.microsoft.com/v1.0/groups/$($Id)?`$select=id,resourceProvisioningOptions"
	}
	$response = Invoke-GraphRequest @testSplat

	return $response.resourceProvisioningOptions -contains "Team"
}

# Archiving functions
#

function Set-TeamArchived{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory)] [System.Collections.Generic.Dictionary[[String], [String]]] $Headers,
		[Parameter(Mandatory)][string] $Id
	)
	# Get some info about the team
	$teamInfo = Get-GroupById -Headers $Headers -Id $Id -Select "displayName,description,mailNickname,resourceProvisioningOptions"
	if ($teamInfo.resourceProvisioningOptions -notcontains "Team") {
		Throw "Group $Id does not have Team capabilities and cannot be archived."
	}
	# Lets see if the team is already archived
	$isArchived = Test-TeamArchived -Headers $Headers -Id $Id
	if ($isArchived) {
		Throw "Team $Id is already archived."
	}
	# Nothing to block us from archiving the team, lets do it
	$archiveSplat = @{
		Headers = $Headers
		Uri = "https://graph.microsoft.com/v1.0/teams/$($Id)/archive"
		Method = "POST"
	}
	$null = Invoke-GraphRequest @archiveSplat
	# Rename the team to indicate it is archived
	if (-not $teamInfo.displayName.StartsWith('Archived_') ) {
		$newDisplayName = "Archived_$($teamInfo.displayName)"
	}
	$now = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
	$newDescription = "[Archived team at $now] $($teamInfo.description)"
	$renameSplat = @{
		Headers = $Headers
		Uri = "https://graph.microsoft.com/v1.0/groups/$($Id)"
		Method = "PATCH"
		Body = @{
			displayName = $newDisplayName
			description = $newDescription
		}
	}
	$null = Invoke-GraphRequest @renameSplat
}

function Clear-TeamArchived{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory)] [System.Collections.Generic.Dictionary[[String], [String]]] $Headers,
		[Parameter(Mandatory)][string] $Id
	)
	# Get some info about the team
	$teamInfo = Get-GroupById -Headers $Headers -Id $Id -Select "displayName,description,mailNickname,resourceProvisioningOptions"
	if ($teamInfo.resourceProvisioningOptions -notcontains "Team") {
		Throw "Group $Id does not have Team capabilities and cannot be un-archived."
	}
	# Lets see if the team is archived
	$isArchived = Test-TeamArchived -Headers $Headers -Id $Id
	if (-not $isArchived) {
		Throw "Team $Id is not archived."
	}
	# Nothing to block us from archiving the team, lets do it
	$archiveSplat = @{
		Headers = $Headers
		Uri = "https://graph.microsoft.com/v1.0/teams/$($Id)/unarchive"
		Method = "POST"
	}
	$null = Invoke-GraphRequest @archiveSplat

	# Remove the "Archived_" prefix from the team name if it exists
	$newDisplayName = $teamInfo.displayName -replace '^Archived_',''
	$newDescription = $teamInfo.description -replace '^\[Archived team at [^\]]+\]\s*', ''
	$renameSplat = @{
		Headers = $Headers
		Uri = "https://graph.microsoft.com/v1.0/groups/$($Id)"
		Method = "PATCH"
		Body = @{
			displayName = $newDisplayName
			description = $newDescription
		}
	}
	$null = Invoke-GraphRequest @renameSplat
}

# Life cycle functions
#
function Remove-Group {
	[CmdletBinding( SupportsShouldProcess)]
	param (
		[Parameter(Mandatory)] [System.Collections.Generic.Dictionary[[String], [String]]] $Headers,
		[Parameter(Mandatory)][string] $Id,
		[switch] $Force

	)

	$isTeam = Test-IsTeam -Headers $Headers -Id $Id

	if ($isTeam  -and (-not $Force) ) {
		$archived = Test-TeamArchived -Headers $Headers -Id $Id
		if (-not $archived) {
			Throw "The group $Id is in fact a team and is not archived. Use -Force to remove a non-archived team."
		}
	}

	$removeSplat = @{
		Headers = $Headers
		Uri = "https://graph.microsoft.com/v1.0/groups/$($Id)"
		Method = "DELETE"
	}
	if ($PSCmdlet.ShouldProcess("Group $Id", "Remove group")) {
		$null = Invoke-GraphRequest @removeSplat
	}
}

#
# Public API
Export-ModuleMember -Function @(
	'Get-GraphHeaders',
	'Get-GraphCollection',
	'Get-Groups',
	'Get-Teams',
	'Get-GroupByMailNickname',
	'Get-GroupById',
	'Test-TeamArchived',
	'Set-TeamArchived',
	'Clear-TeamArchived',
	'Remove-Group'
)
