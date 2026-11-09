#Requires -Version 5.1
<#
.SYNOPSIS
Exports one row per app and sharing identity from an app inventory CSV.

.DESCRIPTION
Reads appId, appDisplayName, appType, and environmentId from the input CSV.
Processes apps in environmentId order, optionally restricted to one environment.
CSV headers are checked, but row values are trusted and app IDs are used as supplied.
All results are held in memory and exported only after every selected app has
been processed.

Canvas, code, and generic apps use Get-AdminPowerAppRoleAssignment. Model-driven
apps are excluded unless IncludeModelApps is $true. When included, they use the
Dataverse Web API to resolve associated security roles to directly assigned
users and teams, including role copies in child business units.
Groups and teams are not expanded. Multiple roles for an identity are combined.
PrincipalId is the Power Apps principal ID or Dataverse user/team ID.
EntraObjectId is provided separately where available; tenant-wide shares use
the tenant ID in PrincipalId, not EntraObjectId. Roles and RoleIds are
semicolon-separated. Duplicate inventory rows for the same app/environment
are processed once.

IncludePause defaults to $true, adding a random 0-2000 millisecond delay before
each Power Apps administrative lookup and Dataverse Web API request, including
subsequent pages. Set it to $false to disable these delays. This reduces request
frequency but does not guarantee that the APIs will not throttle requests.

This is an assignment report, not an effective-access audit: model-driven
implicit administrator access, licensing, table privileges, and team membership
are not evaluated. Apps without associated roles produce a warning, not a claim
that nobody can access them. Disabled users with role assignments are included.

Install prerequisites once:
  Install-Module Microsoft.PowerApps.Administration.PowerShell -Scope CurrentUser

Before running, sign in with an appropriately privileged account in the tenant:
  Add-PowerAppsAccount

To include model-driven apps, also install Az.Accounts and sign in:
  Install-Module Az.Accounts -Scope CurrentUser
  Connect-AzAccount -Tenant '<tenant-id>'
Az.Accounts is needed only when selected rows include model-driven apps. That
account needs read access to apps, roles, users, teams, and role associations in
each Dataverse environment. Power Apps administrative access alone may not grant
Dataverse access. The script does not change any apps or permissions.
Keep Dataverse-Helpers.ps1 in the same directory as this script.

Lookup failures are warned about individually and processing continues. The
final CSV contains only successfully processed apps; a terminating error after
export explicitly marks it as incomplete. Unsupported app types also fail rather
than silently disappearing. The output path must differ from the input path.
An empty result creates a CSV containing the column headers.

.PARAMETER InputCsvPath
Required path to the app inventory CSV (UTF-8, comma-delimited).
.PARAMETER OutputCsvPath
Required path to the output CSV. An existing output file is overwritten.
.PARAMETER EnvironmentId
Optional environment ID. Only matching input rows are processed.
.PARAMETER IncludeModelApps
Optional Boolean. Defaults to $false, excluding model-driven apps from lookups
and output. Set to $true to include their Dataverse security-role assignments.
.PARAMETER IncludeManagedApps
Optional Boolean. Defaults to $false, excluding managed apps from lookups
and output. Managed apps currently not supported, so recommend exclude for now.
.PARAMETER IncludePause
Optional Boolean, default $true. Adds a random 0-2 second delay before each
Power Apps lookup and Dataverse page request. Set to $false to disable delays.
.EXAMPLE
.\Get-App-Shared-List.ps1 -InputCsvPath .\ListAllPowerApps.csv -OutputCsvPath .\AppSharing.csv
.EXAMPLE
.\Get-App-Shared-List.ps1 -InputCsvPath .\ListAllPowerApps.csv -OutputCsvPath .\AppSharing.csv -EnvironmentId '00000000-0000-0000-0000-000000000000'
.EXAMPLE
.\Get-App-Shared-List.ps1 -InputCsvPath .\ListAllPowerApps.csv -OutputCsvPath .\AppSharing.csv -IncludeModelApps $true
.EXAMPLE
.\Get-App-Shared-List.ps1 -InputCsvPath .\ListAllPowerApps.csv -OutputCsvPath .\AppSharing.csv -IncludePause $false
.LINK
https://learn.microsoft.com/powershell/module/microsoft.powerapps.administration.powershell/get-adminpowerapproleassignment
.LINK
https://learn.microsoft.com/power-apps/developer/data-platform/webapi/quick-start-ps
.LINK
https://learn.microsoft.com/power-apps/maker/model-driven-apps/share-model-driven-app
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$InputCsvPath,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputCsvPath,

    [ValidateNotNullOrEmpty()]
    [string]$EnvironmentId,

    [bool]$IncludeModelApps = $false,

    
    [bool]$IncludeManagedApps = $false,

    [bool]$IncludePause = $true
    
)

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'Dataverse-Helpers.ps1')

function Get-DataverseContext {
    param([string]$Id)

    $context = @{
        Url            = Get-DataverseEnvironmentUrl -Id $Id -IncludePause $IncludePause
        Roles          = @()
        RolePrincipals = @{}
    }
    $context.Roles = @(Invoke-DataverseQuery -Context $context -IncludePause $IncludePause -Query 'roles?$select=roleid,name,_parentrootroleid_value')
    return $context
}

function Get-ModelDrivenAssignments {
    param(
        [hashtable]$Context,
        [string]$AppId
    )

    $appRoles = @(Invoke-DataverseQuery -Context $Context -IncludePause $IncludePause -Query "appmodules($AppId)/appmoduleroles_association?`$select=roleid,name,_parentrootroleid_value")
    if ($appRoles.Count -eq 0) {
        Write-Warning "Model-driven app '$AppId' has no associated security roles. No role-derived identities will be exported; implicit access is not evaluated."
        return
    }

    $roleIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($appRole in $appRoles) {
        $rootId = [string]$appRole._parentrootroleid_value
        if ([string]::IsNullOrWhiteSpace($rootId) -or $rootId -eq [guid]::Empty.ToString()) {
            $rootId = [string]$appRole.roleid
        }
        $matchingRoles = @($Context.Roles | Where-Object {
            $_.roleid -eq $appRole.roleid -or $_.roleid -eq $rootId -or $_._parentrootroleid_value -eq $rootId
        })
        if ($matchingRoles.Count -eq 0) {
            throw "Could not resolve app security role '$($appRole.roleid)' in the environment."
        }
        foreach ($role in $matchingRoles) {
            if (-not $roleIds.Add([string]$role.roleid)) {
                continue
            }
            if (-not $Context.RolePrincipals.ContainsKey([string]$role.roleid)) {
                $principals = [System.Collections.Generic.List[object]]::new()
                $users = @(Invoke-DataverseQuery -Context $Context -IncludePause $IncludePause -Query "roles($($role.roleid))/systemuserroles_association?`$select=systemuserid,fullname,internalemailaddress,azureactivedirectoryobjectid")
                foreach ($user in $users) {
                    $principals.Add([pscustomobject]@{
                        PrincipalId          = [string]$user.systemuserid
                        PrincipalType        = 'User'
                        PrincipalDisplayName = $user.fullname
                        PrincipalEmail       = $user.internalemailaddress
                        EntraObjectId        = $user.azureactivedirectoryobjectid
                        DataversePrincipalId = $user.systemuserid
                        RoleName             = $role.name
                        RoleId               = $role.roleid
                        SharingSource        = 'DataverseSecurityRole'
                    })
                }
                $teams = @(Invoke-DataverseQuery -Context $Context -IncludePause $IncludePause -Query "roles($($role.roleid))/teamroles_association?`$select=teamid,name,emailaddress,azureactivedirectoryobjectid")
                foreach ($team in $teams) {
                    $principals.Add([pscustomobject]@{
                        PrincipalId          = [string]$team.teamid
                        PrincipalType        = 'Team'
                        PrincipalDisplayName = $team.name
                        PrincipalEmail       = $team.emailaddress
                        EntraObjectId        = $team.azureactivedirectoryobjectid
                        DataversePrincipalId = $team.teamid
                        RoleName             = $role.name
                        RoleId               = $role.roleid
                        SharingSource        = 'DataverseSecurityRole'
                    })
                }
                $Context.RolePrincipals[[string]$role.roleid] = $principals.ToArray()
            }
            $Context.RolePrincipals[[string]$role.roleid]
        }
    }
}

function Get-PowerAppsAssignments {
    param([object]$App)

    $lookupErrors = @()
    if ($IncludePause) {
        Start-Sleep -Milliseconds (Get-Random -Minimum 0 -Maximum 2001)
    }
    $assignments = @(Get-AdminPowerAppRoleAssignment -EnvironmentName $App.environmentId -AppName $App.appId -ErrorAction Stop -ErrorVariable lookupErrors)
    if ($lookupErrors.Count -gt 0) {
        throw "Power Apps sharing lookup failed: $($lookupErrors[0])"
    }
    foreach ($assignment in $assignments) {
        [pscustomobject]@{
            PrincipalId          = [string]$assignment.PrincipalObjectId
            PrincipalType        = $assignment.PrincipalType
            PrincipalDisplayName = $assignment.PrincipalDisplayName
            PrincipalEmail       = $assignment.PrincipalEmail
            EntraObjectId        = $(if ($assignment.PrincipalType -ne 'Tenant') { $assignment.PrincipalObjectId } else { '' })
            DataversePrincipalId = ''
            RoleName             = $assignment.RoleType
            RoleId               = $assignment.RoleId
            SharingSource        = 'PowerAppsRoleAssignment'
        }
    }
}

#this is the main code execution
$inputFile = Get-Item -LiteralPath $InputCsvPath -ErrorAction Stop
if ($inputFile -isnot [System.IO.FileInfo]) {
    throw 'InputCsvPath must identify a file.'
}
$outputFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputCsvPath)
if ($inputFile.FullName -eq $outputFile) {
    throw 'InputCsvPath and OutputCsvPath must be different files.'
}

$apps = @(Import-Csv -LiteralPath $inputFile.FullName -Encoding UTF8 -ErrorAction Stop)
$requiredColumns = @('appId', 'appDisplayName', 'appType', 'environmentId','modelDrivenAppId')
if ($apps.Count -gt 0) {
    $missingColumns = @($requiredColumns | Where-Object { $_ -notin $apps[0].PSObject.Properties.Name })
    if ($missingColumns.Count -gt 0) {
        throw "Input CSV is missing required columns: $($missingColumns -join ', ')."
    }
}

# Filter apps based on the provided parameters
if ($EnvironmentId) {
    $apps = @($apps | Where-Object { $_.environmentId -eq $EnvironmentId })
}
if (-not $IncludeModelApps) {
    $apps = @($apps | Where-Object { $_.appType -ne 'microsoft.powerapps/modeldrivenapps' })
}
if (-not $IncludeManagedApps) {
    $apps = @($apps | Where-Object { $_.appType -ne 'microsoft.powerapps/apps' })
}

#sort apps list by environmentId and appId, ensuring uniqueness
$apps = @($apps | Sort-Object environmentId, appId -Unique)

if ($apps.Count -gt 0) {
    Import-Module Microsoft.PowerApps.Administration.PowerShell -ErrorAction Stop
    if (@($apps | Where-Object { $_.appType -eq 'microsoft.powerapps/modeldrivenapps' }).Count -gt 0) {
        Import-Module Az.Accounts -ErrorAction Stop
        if ($null -eq (Get-AzContext -ErrorAction Stop)) {
            throw "Sign in to the correct tenant using Connect-AzAccount -Tenant '<tenant-id>' before running this script."
        }
    }
}
else {
    Write-Warning 'No input apps match the requested scope. Writing a header-only CSV.'
}

$outputDirectory = Split-Path -Path $outputFile -Parent
if (-not (Test-Path -LiteralPath $outputDirectory -PathType Container)) {
    $null = New-Item -Path $outputDirectory -ItemType Directory -Force -ErrorAction Stop
}

$results = [System.Collections.Generic.List[object]]::new()
$failures = [System.Collections.Generic.List[string]]::new()
$contexts = @{}
$index = 0

foreach ($app in $apps) {
    $index++
    Write-Progress -Activity 'Retrieving app sharing identities' -Status "$index of $($apps.Count): $($app.appDisplayName)" -PercentComplete (($index / $apps.Count) * 100)
    try {
        Write-Verbose "Processing environment '$($app.environmentId)', app '$($app.appId)' ($($app.appType))."
        switch ($app.appType) {
            'microsoft.powerapps/modeldrivenapps' {
                if (-not $contexts.ContainsKey($app.environmentId)) {
                    $contexts[$app.environmentId] = Get-DataverseContext -Id $app.environmentId
                }
                $assignments = @(Get-ModelDrivenAssignments -Context $contexts[$app.environmentId] -AppId $app.modelDrivenAppId)
            }
            #- 'microsoft.powerapps/apps' is managed apps, and we leave this out because Get-AdminPowerAppRoleAssignment doesnt support managed apps
            { $_ -in @('microsoft.powerapps/canvasapps', 'microsoft.powerapps/codeapps','microsoft.powerapps/apps') } {
                $assignments = @(Get-PowerAppsAssignments -App $app)
            }
            default {
                write-warning "Unsupported appType '$($app.appType)'."
            }
        }
        foreach ($assignment in $assignments) {
            if ([string]::IsNullOrWhiteSpace($assignment.PrincipalId) -or [string]::IsNullOrWhiteSpace($assignment.PrincipalType)) {
                throw 'A sharing assignment is missing its identity ID or type.'
            }
        }
        foreach ($identity in ($assignments | Group-Object -Property PrincipalType, PrincipalId | Sort-Object Name)) {
            $principal = $identity.Group[0]
            $results.Add([pscustomobject][ordered]@{
                appId                = $app.appId
                appDisplayName       = $app.appDisplayName
                appType              = $app.appType
                environmentId        = $app.environmentId
                PrincipalId          = $principal.PrincipalId
                PrincipalType        = $principal.PrincipalType
                PrincipalDisplayName = $principal.PrincipalDisplayName
                PrincipalEmail       = $principal.PrincipalEmail
                EntraObjectId        = $principal.EntraObjectId
                DataversePrincipalId = $principal.DataversePrincipalId
                Roles                = ($identity.Group.RoleName | Sort-Object -Unique) -join '; '
                RoleIds              = ($identity.Group.RoleId | Sort-Object -Unique) -join '; '
                SharingSource        = $principal.SharingSource
            })
        }
    }
    catch {
        $message = "App '$($app.appDisplayName)' ($($app.appId)), environment '$($app.environmentId)': $($_.Exception.Message)"
        $failures.Add($message)
        Write-Warning $message
    }
}
Write-Progress -Activity 'Retrieving app sharing identities' -Completed

$columns = @(
    'appId', 'appDisplayName', 'appType', 'environmentId',
    'PrincipalId', 'PrincipalType', 'PrincipalDisplayName', 'PrincipalEmail',
    'EntraObjectId', 'DataversePrincipalId', 'Roles', 'RoleIds', 'SharingSource'
)
if ($results.Count -gt 0) {
    $results | Select-Object -Property $columns | Export-Csv -LiteralPath $outputFile -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
}
else {
    ('"' + ($columns -join '","') + '"') | Set-Content -LiteralPath $outputFile -Encoding UTF8 -ErrorAction Stop
}
if ($failures.Count -gt 0) {
    Write-Error  "INCOMPLETE EXPORT: $($failures.Count) of $($apps.Count) apps failed. Exported $($results.Count) identity rows for successful apps to '$outputFile'. See the warnings above."
}
Write-Host "Exported $($results.Count) identity rows from $($apps.Count) apps to '$outputFile'."