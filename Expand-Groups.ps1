#Requires -Version 5.1
<#
.SYNOPSIS
Exports user members of Microsoft Entra groups and Dataverse teams in a CSV.

.DESCRIPTION
Uses PrincipalId, PrincipalType, PrincipalDisplayName, and environmentId.
Group rows are processed by default (case-insensitive). Team rows are included
only when ExpandTeams is $true; it defaults to $false. Group IDs are
Microsoft Entra object IDs, queried once per distinct ID through
Get-MgGroupTransitiveMember -All, including nested members. Only user objects are
exported; UPN and display name come from AdditionalProperties.

Team IDs are Dataverse team IDs. Team rows must have an environmentId, used by
Get-AdminPowerAppEnvironment to find the Dataverse URL. Each team/environment
pair is inspected for teamtype and azureactivedirectoryobjectid.
Security Group (2) and Office/Microsoft 365 Group (3) teams use
Get-MgGroupTransitiveMember with their backing Entra group object ID. All
transitive users are included, not just members materialized in Dataverse.
Entra memberships are fetched once per backing group ID, even when shared by
several teams or also present as a Group row.

Owner (0) and Access (1) teams use teams(<id>)/teammembership_association,
following all pages. UserId is azureactivedirectoryobjectid (not systemuserid),
UPN is domainname, and UserDisplayName is fullname.

The first matching row supplies each group's/team's display name. For a Group,
EnvironmentId is taken from that first row, if present; Graph membership itself
is tenant-wide. A user in several groups/teams has a separate row for each.

IncludePause defaults to $true and adds a random 0-2000 millisecond delay before
each Graph or Dataverse Web API request, including team inspection and subsequent
pages. This reduces request frequency
but does not guarantee that Microsoft Graph will not throttle requests.

Install prerequisites once:
  Install-Module Microsoft.Graph.Groups -Scope CurrentUser

Before running, connect to the correct tenant with consented read permissions:
  Connect-MgGraph -Scopes 'GroupMember.Read.All', 'User.ReadBasic.All'
Hidden memberships additionally require Member.Read.Hidden.
Group rows and group-backed Team rows require a Graph connection. For team-only
input, Graph is initialized only if a group-backed team is encountered.

When ExpandTeams is $true, Team rows also require read access to each environment:
  Install-Module Microsoft.PowerApps.Administration.PowerShell -Scope CurrentUser
  Install-Module Az.Accounts -Scope CurrentUser
  Add-PowerAppsAccount
  Connect-AzAccount -Tenant '<tenant-id>'
Power Apps administrative access alone may not grant Dataverse access.
Dataverse-Helpers.ps1 must be alongside this script. No memberships are changed.

Output columns: UserId, UPN, UserDisplayName, GroupId, GroupDisplayName,
PrincipalType, EnvironmentId. For Team rows, GroupId/GroupDisplayName contain
the Dataverse team ID/name.
Lookup failures or missing user details produce warnings and processing
continues. Failed groups/teams contribute no partial rows, and the final export
is explicitly marked incomplete in a warning if any failed.
No matching groups or no user members produces a header-only CSV.

.PARAMETER InputCsvPath
Required path to the UTF-8, comma-delimited source CSV.
.PARAMETER OutputCsvPath
Required path to the output CSV. Existing output is overwritten on success.
Missing parent directories are created. Must differ from the input path.
.PARAMETER IncludePause
Optional Boolean, default $true. Set to $false to disable request delays.
.PARAMETER ExpandTeams
Optional Boolean, default $false. Set to $true to include Dataverse Team rows.
When $false, Team rows are ignored before team validation or authentication.
.EXAMPLE
.\Expand-Groups.ps1 -InputCsvPath .\t.csv -OutputCsvPath .\GroupUsers.csv
.EXAMPLE
.\Expand-Groups.ps1 -InputCsvPath .\t.csv -OutputCsvPath .\GroupUsers.csv -IncludePause $false
.EXAMPLE
.\Expand-Groups.ps1 -InputCsvPath .\t.csv -OutputCsvPath .\GroupUsers.csv -ExpandTeams $true
.LINK
https://learn.microsoft.com/powershell/module/microsoft.graph.groups/get-mggrouptransitivemember
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$InputCsvPath,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputCsvPath,

    [bool]$IncludePause = $true,

    [bool]$ExpandTeams = $false
)

$ErrorActionPreference = 'Stop'

function Assert-GraphConnection {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Import-Module Microsoft.Graph.Groups -ErrorAction Stop
    if ($null -eq (Get-MgContext -ErrorAction Stop)) {
        throw "Connect to the correct tenant using Connect-MgGraph -Scopes 'GroupMember.Read.All', 'User.ReadBasic.All' before running this script."
    }
}

function Get-EntraGroupUsers {
    param(
        [string]$GroupId,
        [hashtable]$Cache,
        [hashtable]$RequestOptions
    )

    if (-not $Cache.ContainsKey($GroupId)) {
        try {
            $Cache[$GroupId] = @(Get-MgGroupTransitiveMember -GroupId $GroupId @RequestOptions |
                Where-Object { $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.user' } |
                ForEach-Object {
                    [pscustomobject]@{
                        UserId          = [string]$_.Id
                        UPN             = [string]$_.AdditionalProperties['userPrincipalName']
                        UserDisplayName = [string]$_.AdditionalProperties['displayName']
                    }
                })
        }
        catch {
            # Do not retry the same failed group for every team that references it.
            $Cache[$GroupId] = $_
        }
    }
    if ($Cache[$GroupId] -is [System.Management.Automation.ErrorRecord]) {
        throw $Cache[$GroupId]
    }
    $Cache[$GroupId]
}

$inputFile = Get-Item -LiteralPath $InputCsvPath -ErrorAction Stop
if ($inputFile -isnot [System.IO.FileInfo]) {
    throw 'InputCsvPath must identify a file.'
}
$outputFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputCsvPath)
if ($inputFile.FullName -eq $outputFile) {
    throw 'InputCsvPath and OutputCsvPath must be different files.'
}

$requiredColumns = @('PrincipalId', 'PrincipalType', 'PrincipalDisplayName')
$principals = @(Import-Csv -LiteralPath $inputFile.FullName -Encoding UTF8 -ErrorAction Stop)
if ($principals.Count -gt 0) {
    $missingColumns = @($requiredColumns | Where-Object { $_ -notin $principals[0].PSObject.Properties.Name })
    if ($missingColumns.Count -gt 0) {
        throw "Input CSV is missing required columns: $($missingColumns -join ', ')."
    }
}

$includedTypes = @('Group')
if ($ExpandTeams) {
    $includedTypes += 'Team'
}
Write-Host "Found $($principals.Count) rows in CSV file, now filtering unique $($includedTypes -join '/') rows."

$groups = [System.Collections.Generic.List[object]]::new()
$seenGroupIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$processedRows = 0
$lastFilterPercent = -1
try {
    foreach ($principal in ($principals | Select-Object -Property ($requiredColumns + @('environmentId')))) {
        $processedRows++
        $filterPercent = [int][Math]::Floor(100.0 * $processedRows / $principals.Count)
        if ($filterPercent -ne $lastFilterPercent) {
            Write-Progress -Activity 'Filtering CSV rows' -Status "$processedRows of $($principals.Count) rows ($filterPercent%)" -PercentComplete $filterPercent
            $lastFilterPercent = $filterPercent
        }
        $principalType = ([string]$principal.PrincipalType).Trim()
        if ($principalType -notin $includedTypes) {
            continue
        }
        $groupId = ([string]$principal.PrincipalId).Trim()
        if ([string]::IsNullOrWhiteSpace($groupId)) {
            throw "A Group or Team row ('$($principal.PrincipalDisplayName)') has an empty PrincipalId."
        }
        $environmentId = ([string]$principal.environmentId).Trim()
        if ($principalType -eq 'Team') {
            $principalType = 'Team'
            if ([string]::IsNullOrWhiteSpace($environmentId)) {
                throw "Team '$($principal.PrincipalDisplayName)' ($groupId) requires an environmentId column with a non-empty value."
            }
            $teamId = [guid]::Empty
            if (-not [guid]::TryParse($groupId, [ref]$teamId) -or $teamId -eq [guid]::Empty) {
                throw "Team '$($principal.PrincipalDisplayName)' has an invalid Dataverse team PrincipalId: '$groupId'."
            }
            $groupId = $teamId.ToString()
            $groupKey = "Team|$environmentId|$groupId"
        }
        else {
            $principalType = 'Group'
            $groupKey = "Group|$groupId"
        }
        if ($seenGroupIds.Add($groupKey)) {
            $groups.Add([pscustomobject]@{
                Id            = $groupId
                DisplayName   = $principal.PrincipalDisplayName
                PrincipalType = $principalType
                EnvironmentId = $environmentId
            })
        }
    }
}
finally {
    Write-Progress -Activity 'Filtering CSV rows' -Completed
}

Write-Host "Found $($groups.Count) unique groups to expand."

$graphReady = $false
if (@($groups | Where-Object { $_.PrincipalType -eq 'Group' }).Count -gt 0) {
    Assert-GraphConnection
    $graphReady = $true
}
if (@($groups | Where-Object { $_.PrincipalType -eq 'Team' }).Count -gt 0) {
    Import-Module Microsoft.PowerApps.Administration.PowerShell -ErrorAction Stop
    Import-Module Az.Accounts -ErrorAction Stop
    if ($null -eq (Get-AzContext -ErrorAction Stop)) {
        throw "Sign in to the correct tenant using Add-PowerAppsAccount and Connect-AzAccount -Tenant '<tenant-id>' before expanding Dataverse teams."
    }
    . (Join-Path $PSScriptRoot 'Dataverse-Helpers.ps1')
}
if ($groups.Count -eq 0) {
    Write-Warning "No matching $($includedTypes -join '/') rows found. Writing a header-only CSV."
}

$requestOptions = @{
    All         = $true
    Property    = @('id', 'displayName', 'userPrincipalName')
    ErrorAction = 'Stop'
}
if ($IncludePause) {
    # An HTTP pipeline step also runs for the extra pages requested by -All.
    $requestOptions.HttpPipelinePrepend = @({
        param($request, $callback, $next)
        Start-Sleep -Milliseconds (Get-Random -Minimum 0 -Maximum 2001)
        $next.SendAsync($request, $callback)
    })
}

$results = [System.Collections.Generic.List[object]]::new()
$dataverseContexts = @{}
$graphMemberCache = @{}
$failedCount = 0
$index = 0
try {
    foreach ($group in $groups) {
        $index++
        Write-Progress -Activity 'Expanding groups' -Status "$index of $($groups.Count): $($group.DisplayName)" -PercentComplete (($index / $groups.Count) * 100)
        try {
            $usesGraph = $group.PrincipalType -eq 'Group'
            $entraGroupId = $group.Id
            if (-not $usesGraph) {
                if (-not $dataverseContexts.ContainsKey($group.EnvironmentId)) {
                    $dataverseContexts[$group.EnvironmentId] = @{
                        Url = Get-DataverseEnvironmentUrl -Id $group.EnvironmentId
                    }
                }
                $context = $dataverseContexts[$group.EnvironmentId]
                $team = Invoke-DataverseRequest -Context $context -IncludePause $IncludePause -RequestUri "$($context.Url)/api/data/v9.2/teams($($group.Id))?`$select=teamtype,azureactivedirectoryobjectid"
                if ($null -eq $team -or $null -eq $team.teamtype -or $team.teamtype -notin @(0, 1, 2, 3)) {
                    throw 'Dataverse did not return a supported teamtype (0, 1, 2, or 3).'
                }
                if ($team.teamtype -in @(2, 3)) {
                    $backingGroupId = [guid]::Empty
                    if (-not [guid]::TryParse(([string]$team.azureactivedirectoryobjectid).Trim(), [ref]$backingGroupId) -or $backingGroupId -eq [guid]::Empty) {
                        throw 'The group-backed team has no valid backing Entra group object ID (azureactivedirectoryobjectid).'
                    }
                    $usesGraph = $true
                    $entraGroupId = $backingGroupId.ToString()
                }
                else {
                    $members = @(Invoke-DataverseQuery -Context $context -IncludePause $IncludePause -Query "teams($($group.Id))/teammembership_association?`$select=azureactivedirectoryobjectid,domainname,fullname" |
                        ForEach-Object {
                            [pscustomobject]@{
                                UserId          = [string]$_.azureactivedirectoryobjectid
                                UPN             = [string]$_.domainname
                                UserDisplayName = [string]$_.fullname
                            }
                        })
                }
            }
            if ($usesGraph) {
                if (-not $graphReady) {
                    Assert-GraphConnection
                    $graphReady = $true
                }
                $members = @(Get-EntraGroupUsers -GroupId $entraGroupId -Cache $graphMemberCache -RequestOptions $requestOptions)
                $missingDetailsMessage = 'Check user read permissions, including User.ReadBasic.All.'
            }
            else {
                $missingDetailsMessage = 'Dataverse users must have azureactivedirectoryobjectid, domainname, and fullname; systemuserid and email are not substitutes.'
            }
            $groupResults = @(foreach ($member in $members) {
                if ([string]::IsNullOrWhiteSpace($member.UserId) -or
                    $member.UserId -eq [guid]::Empty.ToString() -or
                    [string]::IsNullOrWhiteSpace($member.UPN) -or
                    [string]::IsNullOrWhiteSpace($member.UserDisplayName)) {
                    throw "A member ('$($member.UserId)') has no Entra ID, UPN, or display name. $missingDetailsMessage"
                }
                [pscustomobject][ordered]@{
                        UserId           = $member.UserId
                        UPN              = $member.UPN
                        UserDisplayName  = $member.UserDisplayName
                        GroupId          = $group.Id
                        GroupDisplayName = $group.DisplayName
                        PrincipalType    = $group.PrincipalType
                        EnvironmentId    = $group.EnvironmentId
                }
            })
            $results.AddRange($groupResults)
        }
        catch {
            $failedCount++
            Write-Warning "Failed to expand $($group.PrincipalType) '$($group.DisplayName)' ($($group.Id)), environment '$($group.EnvironmentId)': $($_.Exception.Message)"
        }
    }
}
finally {
    Write-Progress -Activity 'Expanding groups' -Completed
}

$outputDirectory = Split-Path -Path $outputFile -Parent
if (-not (Test-Path -LiteralPath $outputDirectory -PathType Container)) {
    $null = New-Item -Path $outputDirectory -ItemType Directory -Force -ErrorAction Stop
}
if ($results.Count -gt 0) {
    $results | Export-Csv -LiteralPath $outputFile -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
}
else {
    '"UserId","UPN","UserDisplayName","GroupId","GroupDisplayName","PrincipalType","EnvironmentId"' |
        Set-Content -LiteralPath $outputFile -Encoding UTF8 -ErrorAction Stop
}
if ($failedCount -gt 0) {
    Write-Warning "INCOMPLETE EXPORT: $failedCount of $($groups.Count) groups/teams failed. Exported $($results.Count) memberships from successful lookups to '$outputFile'."
}
else {
    Write-Host "Exported $($results.Count) user/group memberships from $($groups.Count) groups to '$outputFile'."
}