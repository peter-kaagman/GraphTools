# GraphTools

GraphTools is a small PowerShell module containing helper functions for working with Microsoft Graph through its REST API.

The module was created for reuse in provisioning and lifecycle scripts, including scripts executed by HelloID. It uses application authentication with the OAuth 2.0 client credentials flow and does not depend on the Microsoft Graph PowerShell SDK.

## Requirements

- Windows PowerShell 5.1 or PowerShell 7+
- A Microsoft Entra app registration
- Microsoft Graph application permissions appropriate for the operations being performed
- Tenant ID, application ID and client secret

Grant only the permissions required by the scripts using this module.

## Installation

Place the module directory in a location included in `$env:PSModulePath`, or import it directly:

```powershell
Import-Module "C:\HelloID\Modules\GraphTools\GraphTools.psd1"
```

If no module manifest is present:

```powershell
Import-Module "C:\HelloID\Modules\GraphTools\GraphTools.psm1"
```

## Authentication

Create the HTTP headers used by the other functions:

```powershell
$headers = Get-GraphHeaders `
    -TenantId $tenantId `
    -ClientId $clientId `
    -ClientSecret $clientSecret
```

Advanced Graph query headers can be added by the caller when required:

```powershell
$headers['ConsistencyLevel'] = 'eventual'
```

## Functions

### Authentication

#### `Get-GraphHeaders`

Obtains an application access token using the OAuth 2.0 client credentials flow and returns a header dictionary for Microsoft Graph requests.

### Group and Team retrieval

#### `Get-GraphCollection`

Retrieves all objects from a paged Microsoft Graph collection endpoint and follows `@odata.nextLink` automatically.

Results are returned as an `ArrayList` by default. When `-HashTableKeyProperty` is provided, the results are indexed by that property.

#### `Get-Groups`

Retrieves Microsoft 365 groups from the Microsoft Graph `/groups` endpoint.

Supports OData `$filter` and `$select` expressions and can return the results indexed by a selected property.

#### `Get-Teams`

Retrieves groups with Microsoft Teams capabilities.

This function uses the `/groups` endpoint and filters on `resourceProvisioningOptions`. It therefore returns group properties rather than team-specific settings.

#### `Get-GroupById`

Retrieves a single Microsoft 365 group by its Graph object ID.

#### `Get-GroupByMailNickname`

Retrieves the first Microsoft 365 group matching the supplied `mailNickname`.

A `mailNickname` is not guaranteed to be unique. This function should therefore only be used where returning the first matching object is acceptable.

### Team state

#### `Test-TeamArchived`

Returns whether a Team is archived.

#### `Set-TeamArchived`

Archives a Team and updates its display name and description to indicate when it was archived.

#### `Clear-TeamArchived`

Unarchives a Team and removes the archive marker from its display name and description.

### Group lifecycle

#### `Remove-Group`

Removes a Microsoft 365 group.

When the group has Teams capabilities, it must be archived before removal unless `-Force` is specified. The function supports `-WhatIf` and `-Confirm` through PowerShell's `ShouldProcess` mechanism.

## Examples

Retrieve all groups:

```powershell
$groups = Get-Groups -Headers $headers
```

Retrieve selected group properties:

```powershell
$groups = Get-Groups `
    -Headers $headers `
    -Select 'id,displayName,mailNickname'
```

Retrieve Teams indexed by object ID:

```powershell
$teams = Get-Teams `
    -Headers $headers `
    -Select 'id,displayName,mailNickname' `
    -HashTableKeyProperty 'id'
```

Archive a Team:

```powershell
Set-TeamArchived `
    -Headers $headers `
    -Id $teamId
```

Remove an archived Team or regular group:

```powershell
Remove-Group `
    -Headers $headers `
    -Id $groupId
```

Preview removal without making changes:

```powershell
Remove-Group `
    -Headers $headers `
    -Id $groupId `
    -WhatIf
```

## Notes

GraphTools is intended as a lightweight helper module. It does not try to replace the Microsoft Graph PowerShell SDK or provide complete coverage of the Microsoft Graph API.

## License

This project is licensed under the MIT License. See the LICENSE file for details.


