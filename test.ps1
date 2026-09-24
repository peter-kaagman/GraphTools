<#
.SYNOPSIS
Smoke test for GraphTools.

.DESCRIPTION
Validates the most important GraphTools functions against a live Microsoft 365 tenant.

The script uses an existing test group and test team. Team archive and
unarchive operations are exercised and the original archive state is restored
after testing.

ASSUMPTIONS
    - A test group with mailNickname 'test-peter' exists.
    - At least one Team with mailNickname starting with 'Test' exists.
    - The configured application has sufficient Graph permissions.

SIDE EFFECTS
    - The selected Team is archived and unarchived during the test.
    - The original archive state is restored before the script exits.

EXIT CODES
    0 = Success
    1 = One or more tests failed
#>

Import-Module "$PSScriptRoot\GraphTools.psm1" -Force


# ----------------------------------------------------------------------
# Test state
# ----------------------------------------------------------------------

$script:TestFailures = 0


# ----------------------------------------------------------------------
# Assertions
# ----------------------------------------------------------------------

function Assert-True {
    param(
        [Parameter(Mandatory)]
        [bool] $Value,

        [Parameter(Mandatory)]
        [string] $Message
    )

    if (-not $Value) {
        Write-Host "FAIL: $Message" -ForegroundColor Red
        $script:TestFailures++
        return
    }

    Write-Host "PASS: $Message" -ForegroundColor Green
}


function Assert-Equal {
    param(
        [Parameter(Mandatory)]
        [AllowNull()]
        [object] $Expected,

        [Parameter(Mandatory)]
        [AllowNull()]
        [object] $Actual,

        [Parameter(Mandatory)]
        [string] $Message
    )

    if ($Expected -ne $Actual) {
        Write-Host `
            "FAIL: $Message (expected '$Expected', actual '$Actual')" `
            -ForegroundColor Red

        $script:TestFailures++
        return
    }

    Write-Host "PASS: $Message" -ForegroundColor Green
}


function Assert-NotNull {
    param(
        [Parameter(Mandatory)]
        [AllowNull()]
        [object] $Value,

        [Parameter(Mandatory)]
        [string] $Message
    )

    if ($null -eq $Value) {
        Write-Host "FAIL: $Message" -ForegroundColor Red
        $script:TestFailures++
        return
    }

    Write-Host "PASS: $Message" -ForegroundColor Green
}


function Assert-Throws {
    param(
        [Parameter(Mandatory)]
        [scriptblock] $ScriptBlock,

        [Parameter(Mandatory)]
        [string] $Message,

        [string] $ExpectedMessage
    )

    try {
        & $ScriptBlock
    }
    catch {
        if (
            $ExpectedMessage -and
            $_.Exception.Message -notmatch $ExpectedMessage
        ) {
            Write-Host `
                "FAIL: $Message (expected exception matching '$ExpectedMessage', got '$($_.Exception.Message)')" `
                -ForegroundColor Red

            $script:TestFailures++
            return
        }

        Write-Host "PASS: $Message" -ForegroundColor Green
        return
    }

    Write-Host `
        "FAIL: $Message (expected exception was not thrown)" `
        -ForegroundColor Red

    $script:TestFailures++
}


# ----------------------------------------------------------------------
# Wait helpers
# ----------------------------------------------------------------------

function Wait-TeamArchivedState {
    param(
        [Parameter(Mandatory)]
        [hashtable] $Parameters,

        [Parameter(Mandatory)]
        [bool] $ExpectedState,

        [int] $Retries = 10,

        [int] $DelaySeconds = 2
    )

    for ($attempt = 1; $attempt -le $Retries; $attempt++) {
        $currentState = Test-TeamArchived @Parameters

        if ($currentState -eq $ExpectedState) {
            return $true
        }

        if ($attempt -lt $Retries) {
            Start-Sleep -Seconds $DelaySeconds
        }
    }

    return $false
}


function Wait-TeamMetadataState {
    param(
        [Parameter(Mandatory)]
        [hashtable] $Parameters,

        [Parameter(Mandatory)]
        [bool] $ExpectedArchivedState,

        [int] $Retries = 10,

        [int] $DelaySeconds = 2
    )

    for ($attempt = 1; $attempt -le $Retries; $attempt++) {
        $teamInfo = Get-GroupById @Parameters

        $hasArchivedPrefix =
            $teamInfo.displayName -match '^Archived_'

        $hasArchiveTimestamp =
            $teamInfo.description -match '^\[Archived team at '

        if (
            $hasArchivedPrefix -eq $ExpectedArchivedState -and
            $hasArchiveTimestamp -eq $ExpectedArchivedState
        ) {
            return $teamInfo
        }

        if ($attempt -lt $Retries) {
            Start-Sleep -Seconds $DelaySeconds
        }
    }

    return $null
}


# ----------------------------------------------------------------------
# Configuration
# ----------------------------------------------------------------------

$config = Get-Content `
    -Path "$PSScriptRoot\..\JSON\.json" `
    -Raw |
    ConvertFrom-Json

$prefix = "Test"
$group  = "test-peter"


# ----------------------------------------------------------------------
# Authentication
# ----------------------------------------------------------------------

Write-Host "Creating Graph connection..." -ForegroundColor Cyan

$headers = Get-GraphHeaders `
    -TenantId $config.GRAPH_TENANT_ID `
    -ClientId $config.GRAPH_CLIENT_ID `
    -ClientSecret $config.GRAPH_CLIENT_SECRET

Assert-NotNull `
    -Value $headers `
    -Message "Graph headers created"


# ----------------------------------------------------------------------
# Shared group query
# ----------------------------------------------------------------------

$groupsSplat = @{
    Headers = $headers
    Filter  = "startswith(mailNickname,'$prefix')"
    Select  = "id,displayName,mail,mailNickname"
}


# ----------------------------------------------------------------------
# Get-Groups
# ----------------------------------------------------------------------

#region Get-Groups

$groups = @(Get-Groups @groupsSplat)

Assert-True `
    -Value ($groups.Count -gt 0) `
    -Message "Get-Groups returns one or more groups"

#endregion


# ----------------------------------------------------------------------
# Get-Teams
# ----------------------------------------------------------------------

#region Get-Teams

$teams = @(Get-Teams @groupsSplat)

Assert-True `
    -Value ($teams.Count -gt 0) `
    -Message "Get-Teams returns one or more teams"

if ($teams.Count -eq 0) {
    Write-Host `
        "Cannot continue without a test team." `
        -ForegroundColor Red

    exit 1
}

$teamId = $teams[0].id

#endregion


# ----------------------------------------------------------------------
# Get-GroupByMailNickname
# ----------------------------------------------------------------------

#region Get-GroupByMailNickname

$getGroupByMailNicknameSplat = @{
    Headers      = $headers
    MailNickname = $group
    Select       = "id,mailNickname"
}

$groupByMailNickname =
    Get-GroupByMailNickname @getGroupByMailNicknameSplat

Assert-NotNull `
    -Value $groupByMailNickname `
    -Message "Group retrieved by mailNickname"

Assert-Equal `
    -Expected $group `
    -Actual $groupByMailNickname.mailNickname `
    -Message "Retrieved group has expected mailNickname"

$groupId = $groupByMailNickname.id

#endregion


# ----------------------------------------------------------------------
# Get-GroupById
# ----------------------------------------------------------------------

#region Get-GroupById

$getGroupByIdSplat = @{
    Headers = $headers
    Id      = $groupId
    Select  = "id,displayName,mailNickname"
}

$groupById = Get-GroupById @getGroupByIdSplat

Assert-NotNull `
    -Value $groupById `
    -Message "Group retrieved by id"

Assert-Equal `
    -Expected $groupId `
    -Actual $groupById.id `
    -Message "Retrieved group has expected id"

#endregion


# ----------------------------------------------------------------------
# Test-TeamArchived
# ----------------------------------------------------------------------

#region Test-TeamArchived

$nonTeamSplat = @{
    Headers = $headers
    Id      = $groupId
}

Assert-Throws `
    -Message "Test-TeamArchived throws for a group without a team" `
    -ScriptBlock {
        Test-TeamArchived @nonTeamSplat
    }

$teamArchiveSplat = @{
    Headers = $headers
    Id      = $teamId
}

$initialArchivedState =
    Test-TeamArchived @teamArchiveSplat

Assert-True `
    -Value ($initialArchivedState -is [bool]) `
    -Message "Test-TeamArchived returns a Boolean for a team"

Write-Host `
    "Initial archive state of team ${teamId}: $initialArchivedState" `
    -ForegroundColor Cyan

#endregion


# ----------------------------------------------------------------------
# Set-TeamArchived on non-team group
# ----------------------------------------------------------------------

$nonTeamArchiveSplat = @{
    Headers = $headers
    Id      = $groupId
}

Assert-Throws `
    -Message "Set-TeamArchived throws for a group without a team" `
    -ScriptBlock {
        Set-TeamArchived @nonTeamArchiveSplat
    }


# ----------------------------------------------------------------------
# Shared metadata query
# ----------------------------------------------------------------------

$getTeamByIdSplat = @{
    Headers = $headers
    Id      = $teamId
    Select  = "id,displayName,description"
}


# ----------------------------------------------------------------------
# Archive transition
# ----------------------------------------------------------------------

function Test-ArchiveTransition {
    Write-Host
    Write-Host "Testing archive transition..." -ForegroundColor Cyan

    Set-TeamArchived @teamArchiveSplat

    $archiveStateChanged = Wait-TeamArchivedState `
        -Parameters $teamArchiveSplat `
        -ExpectedState $true `
        -Retries 10 `
        -DelaySeconds 2

    Assert-True `
        -Value $archiveStateChanged `
        -Message "Team is archived"

    $teamInfo = Wait-TeamMetadataState `
        -Parameters $getTeamByIdSplat `
        -ExpectedArchivedState $true `
        -Retries 10 `
        -DelaySeconds 2

    Assert-NotNull `
        -Value $teamInfo `
        -Message "Archived team metadata became available"

    if ($null -ne $teamInfo) {
        Assert-True `
            -Value ($teamInfo.displayName -match '^Archived_') `
            -Message "Archived prefix added"

        Assert-True `
            -Value ($teamInfo.description -match '^\[Archived team at ') `
            -Message "Archive timestamp added"
    }
}


# ----------------------------------------------------------------------
# De-archive transition
# ----------------------------------------------------------------------

function Test-DeArchiveTransition {
    Write-Host
    Write-Host "Testing de-archive transition..." -ForegroundColor Cyan

    Clear-TeamArchived @teamArchiveSplat

    $deArchiveStateChanged = Wait-TeamArchivedState `
        -Parameters $teamArchiveSplat `
        -ExpectedState $false `
        -Retries 10 `
        -DelaySeconds 2

    Assert-True `
        -Value $deArchiveStateChanged `
        -Message "Team is no longer archived"

    $teamInfo = Wait-TeamMetadataState `
        -Parameters $getTeamByIdSplat `
        -ExpectedArchivedState $false `
        -Retries 10 `
        -DelaySeconds 2

    Assert-NotNull `
        -Value $teamInfo `
        -Message "Active team metadata became available"

    if ($null -ne $teamInfo) {
        Assert-True `
            -Value ($teamInfo.displayName -notmatch '^Archived_') `
            -Message "Archived prefix removed"

        Assert-True `
            -Value ($teamInfo.description -notmatch '^\[Archived team at ') `
            -Message "Archive timestamp removed"
    }
}


# ----------------------------------------------------------------------
# Run transitions in an order based on the initial state
# ----------------------------------------------------------------------

if ($initialArchivedState) {
    Test-DeArchiveTransition
    Test-ArchiveTransition
}
else {
    Test-ArchiveTransition
    Test-DeArchiveTransition
}


# ----------------------------------------------------------------------
# Verify original state was restored
# ----------------------------------------------------------------------

$finalArchivedState =
    Test-TeamArchived @teamArchiveSplat

Assert-Equal `
    -Expected $initialArchivedState `
    -Actual $finalArchivedState `
    -Message "Team restored to its original archive state"


# ----------------------------------------------------------------------
# Test summary
# ----------------------------------------------------------------------

Write-Host

if ($script:TestFailures -eq 0) {
    Write-Host "All smoke tests passed." -ForegroundColor Green
    exit 0
}

Write-Host `
    "$script:TestFailures smoke test(s) failed." `
    -ForegroundColor Red

exit 1
