param(
    [Parameter(Mandatory = $true, HelpMessage = "Provide path to the CSV file containing the list of cost centre AD groups.")]
    [string]$InputGroupsCsvFilePath,    # E.g. ".\Verified_AuditStream_CostCentre_AdGroups.csv" or ".\TestGroups.csv"

    [Parameter(Mandatory = $false, HelpMessage = "Provide the Azure AD tenant ID where groups are located. Default is '70e99cc0-effd-4446-8505-773f8af647fe' (FSP Futures).")]
    [string]$TenantId = "70e99cc0-effd-4446-8505-773f8af647fe",     #FSP Futures

    [Parameter(Mandatory = $false, HelpMessage = "If specified, any pre-existing output files from previous runs will be removed before processing.")]
    [switch]$CleanOutputFiles = $false
)


Update-TypeData -TypeName 'AzurePowerCommands.User' -DefaultDisplayPropertySet @('ObjectType', 'ObjectId', 'DisplayName', 'UserPrincipalName', 'AccountEnabled') -Force
Update-TypeData -TypeName 'AzurePowerCommands.Group' -DefaultDisplayPropertySet @('ObjectType', 'ObjectId', 'DisplayName', 'Mail', 'SecurityEnabled', 'IsAssignableToRole') -Force
Update-TypeData -TypeName 'AzurePowerCommands.ServicePrincipal' -DefaultDisplayPropertySet @('ObjectType', 'ObjectId', 'DisplayName', 'AppId', 'AccountEnabled') -Force
Update-TypeData -TypeName 'AzurePowerCommands.Application' -DefaultDisplayPropertySet @('ObjectType', 'ObjectId', 'DisplayName', 'AppId') -Force
Update-TypeData -TypeName 'AzurePowerCommands.DirectoryObject' -DefaultDisplayPropertySet @('ObjectType', 'ObjectId', 'DisplayName') -Force

$script:AzurePowerDirectoryObjectCache = @{}
$script:AzurePowerDirectoryObjectCacheTenantId = $null


####################################################################################
# START OF FUNCTION DEFINITIONS
####################################################################################

function Assert-AzurePowerGraphConnection {
    [CmdletBinding()]
    param()

    if (-not (Get-Command Get-MgContext -ErrorAction SilentlyContinue)) {
        throw 'Microsoft Graph PowerShell is not installed or imported.'
    }

    $Context = Get-MgContext
    if (-not $Context) {
        throw "You're not connected with Microsoft Graph. Connect with Connect-MgGraph."
    }

    if ($script:AzurePowerDirectoryObjectCacheTenantId -ne $Context.TenantId) {
        $script:AzurePowerDirectoryObjectCache = @{}
        $script:AzurePowerDirectoryObjectCacheTenantId = $Context.TenantId
    }
}

function Get-AzurePowerPropertyValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $InputObject,

        [Parameter(Mandatory = $true)]
        [string[]]$Name
    )

    foreach ($PropertyName in $Name) {
        if ($null -ne $InputObject.PSObject.Methods['ContainsKey']) {
            if ($InputObject.ContainsKey($PropertyName)) {
                return $InputObject[$PropertyName]
            }
        }
        elseif ($InputObject -is [System.Collections.IDictionary]) {
            $Dictionary = [System.Collections.IDictionary]$InputObject
            if ($Dictionary.Contains([object]$PropertyName)) {
                return $Dictionary[$PropertyName]
            }
        }

        $Property = $InputObject.PSObject.Properties[$PropertyName]
        if ($null -ne $Property) {
            return $Property.Value
        }

        $AdditionalPropertiesProperty = $InputObject.PSObject.Properties['AdditionalProperties']
        if ($null -ne $AdditionalPropertiesProperty) {
            $AdditionalProperties = $AdditionalPropertiesProperty.Value

            if ($null -ne $AdditionalProperties -and $null -ne $AdditionalProperties.PSObject.Methods['ContainsKey']) {
                if ($AdditionalProperties.ContainsKey($PropertyName)) {
                    return $AdditionalProperties[$PropertyName]
                }
            }
            elseif ($AdditionalProperties -is [System.Collections.IDictionary]) {
                $Dictionary = [System.Collections.IDictionary]$AdditionalProperties
                if ($Dictionary.Contains([object]$PropertyName)) {
                    return $Dictionary[$PropertyName]
                }
            }
        }
    }

    return $null
}

function Get-AzurePowerObjectId {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $InputObject
    )

    if ($InputObject -is [string]) {
        return [string]$InputObject
    }

    $Id = Get-AzurePowerPropertyValue -InputObject $InputObject -Name @('Id', 'ObjectId', 'id', 'objectId')
    if ([string]::IsNullOrWhiteSpace([string]$Id)) {
        throw 'The supplied object does not contain an Id or ObjectId property.'
    }

    return [string]$Id
}

function Get-AzurePowerObjectType {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $InputObject
    )

    $ObjectType = Get-AzurePowerPropertyValue -InputObject $InputObject -Name @('ObjectType', 'objectType')
    if ($ObjectType) {
        switch -Regex ([string]$ObjectType) {
            '^user$'             { return 'User' }
            '^group$'            { return 'Group' }
            '^serviceprincipal$' { return 'ServicePrincipal' }
            '^application$'      { return 'Application' }
        }
    }

    $ODataType = Get-AzurePowerPropertyValue -InputObject $InputObject -Name @('@odata.type', 'ODataType')
    if ($ODataType) {
        switch -Regex ([string]$ODataType) {
            'microsoft\.graph\.user$'             { return 'User' }
            'microsoft\.graph\.group$'            { return 'Group' }
            'microsoft\.graph\.servicePrincipal$' { return 'ServicePrincipal' }
            'microsoft\.graph\.application$'      { return 'Application' }
        }
    }

    foreach ($TypeName in $InputObject.PSObject.TypeNames) {
        switch -Regex ($TypeName) {
            'MicrosoftGraphUser$'             { return 'User' }
            'MicrosoftGraphGroup$'            { return 'Group' }
            'MicrosoftGraphServicePrincipal$' { return 'ServicePrincipal' }
            'MicrosoftGraphApplication$'      { return 'Application' }
            'AzurePowerCommands\.User$'             { return 'User' }
            'AzurePowerCommands\.Group$'            { return 'Group' }
            'AzurePowerCommands\.ServicePrincipal$' { return 'ServicePrincipal' }
            'AzurePowerCommands\.Application$'      { return 'Application' }
        }
    }

    if (Get-AzurePowerPropertyValue -InputObject $InputObject -Name @('UserPrincipalName', 'userPrincipalName')) {
        return 'User'
    }

    if (Get-AzurePowerPropertyValue -InputObject $InputObject -Name @('AppId', 'appId')) {
        return 'ServicePrincipal'
    }

    if ($null -ne (Get-AzurePowerPropertyValue -InputObject $InputObject -Name @('SecurityEnabled', 'securityEnabled'))) {
        return 'Group'
    }

    return 'DirectoryObject'
}

function ConvertTo-AzurePowerDirectoryObject {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        $InputObject,

        [Parameter(Mandatory = $false)]
        [ValidateSet('User', 'Group', 'ServicePrincipal', 'Application', 'DirectoryObject')]
        [string]$ObjectType
    )

    process {
        $ResolvedObjectType = $ObjectType
        if (-not $ResolvedObjectType) {
            $ResolvedObjectType = Get-AzurePowerObjectType -InputObject $InputObject
        }

        $Id = Get-AzurePowerObjectId -InputObject $InputObject
        $ODataType = Get-AzurePowerPropertyValue -InputObject $InputObject -Name @('@odata.type', 'ODataType')

        [PSCustomObject]@{
            PSTypeName           = "AzurePowerCommands.$ResolvedObjectType"
            ObjectType           = $ResolvedObjectType
            ObjectId             = $Id
            Id                   = $Id
            DisplayName          = Get-AzurePowerPropertyValue -InputObject $InputObject -Name @('DisplayName', 'displayName')
            UserPrincipalName    = Get-AzurePowerPropertyValue -InputObject $InputObject -Name @('UserPrincipalName', 'userPrincipalName')
            AppId                = Get-AzurePowerPropertyValue -InputObject $InputObject -Name @('AppId', 'appId')
            AccountEnabled       = Get-AzurePowerPropertyValue -InputObject $InputObject -Name @('AccountEnabled', 'accountEnabled')
            Mail                 = Get-AzurePowerPropertyValue -InputObject $InputObject -Name @('Mail', 'mail')
            SecurityEnabled      = Get-AzurePowerPropertyValue -InputObject $InputObject -Name @('SecurityEnabled', 'securityEnabled')
            IsAssignableToRole   = Get-AzurePowerPropertyValue -InputObject $InputObject -Name @('IsAssignableToRole', 'isAssignableToRole')
            ServicePrincipalType = Get-AzurePowerPropertyValue -InputObject $InputObject -Name @('ServicePrincipalType', 'servicePrincipalType')
            ODataType            = $ODataType
        }
    }
}

function Resolve-AzurePowerDirectoryObject {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        $DirectoryObject
    )

    process {
        $Id = Get-AzurePowerObjectId -InputObject $DirectoryObject
        $ObjectType = Get-AzurePowerObjectType -InputObject $DirectoryObject
        $CacheKey = "$ObjectType|$Id"

        if ($script:AzurePowerDirectoryObjectCache.ContainsKey($CacheKey)) {
            $script:AzurePowerDirectoryObjectCache[$CacheKey]
            return
        }

        $ResolvedObject = $null

        try {
            switch ($ObjectType) {
                'User' {
                    $ResolvedObject = Get-MgUser -UserId $Id -Property @('Id', 'DisplayName', 'UserPrincipalName', 'AccountEnabled', 'Mail') -ErrorAction Stop |
                        ConvertTo-AzurePowerDirectoryObject -ObjectType User
                }
                'Group' {
                    $ResolvedObject = Get-MgGroup -GroupId $Id -Property @('Id', 'DisplayName', 'Mail', 'SecurityEnabled', 'IsAssignableToRole') -ErrorAction Stop |
                        ConvertTo-AzurePowerDirectoryObject -ObjectType Group
                }
                'ServicePrincipal' {
                    $ResolvedObject = Get-MgServicePrincipal -ServicePrincipalId $Id -Property @('Id', 'DisplayName', 'AppId', 'AccountEnabled', 'ServicePrincipalType') -ErrorAction Stop |
                        ConvertTo-AzurePowerDirectoryObject -ObjectType ServicePrincipal
                }
                'Application' {
                    $ResolvedObject = Get-MgApplication -ApplicationId $Id -Property @('Id', 'DisplayName', 'AppId') -ErrorAction Stop |
                        ConvertTo-AzurePowerDirectoryObject -ObjectType Application
                }
                default {
                    $ResolvedObject = ConvertTo-AzurePowerDirectoryObject -InputObject $DirectoryObject -ObjectType DirectoryObject
                }
            }
        }
        catch {
            Write-Verbose "Could not resolve $ObjectType object ${Id}: $($_.Exception.Message)"
            $ResolvedObject = ConvertTo-AzurePowerDirectoryObject -InputObject $DirectoryObject -ObjectType $ObjectType
        }

        $script:AzurePowerDirectoryObjectCache[$CacheKey] = $ResolvedObject
        $ResolvedObject
    }
}

function Invoke-AzurePowerGraphCollectionRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri
    )

    $NextLink = $Uri
    while ($NextLink) {
        $Response = Invoke-MgGraphRequest -Method GET -Uri $NextLink -ErrorAction Stop
        $Values = Get-AzurePowerPropertyValue -InputObject $Response -Name @('value')

        if ($null -ne $Values) {
            foreach ($Value in @($Values)) {
                $Value
            }
        }

        $NextLink = Get-AzurePowerPropertyValue -InputObject $Response -Name @('@odata.nextLink')
    }
}

function Get-AzurePowerGroupDirectMembers {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupId,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeServicePrincipals
    )

    if ($IncludeServicePrincipals) {
        # The v1.0 group members endpoint has a documented limitation where service
        # principals can be omitted. Beta is used only for this compatibility path.
        $EscapedGroupId = [uri]::EscapeDataString($GroupId)
        $Members = Invoke-AzurePowerGraphCollectionRequest -Uri "https://graph.microsoft.com/beta/groups/$EscapedGroupId/members"
    }
    else {
        $Members = Get-MgGroupMember -GroupId $GroupId -All -ErrorAction Stop
    }

    foreach ($Member in @($Members)) {
        Resolve-AzurePowerDirectoryObject -DirectoryObject $Member
    }
}

function Get-GroupMembersRecursive {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupId,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Users', 'Groups', 'ServicePrincipals')]
        [string]$Mode,

        [Parameter(Mandatory = $true)]
        [hashtable]$VisitedGroups,

        [Parameter(Mandatory = $true)]
        [hashtable]$DistinctObjectIds
    )

    if ($VisitedGroups.ContainsKey($GroupId)) {
        return
    }
    $VisitedGroups[$GroupId] = $true

    $Members = @(Get-AzurePowerGroupDirectMembers -GroupId $GroupId -IncludeServicePrincipals:($Mode -eq 'ServicePrincipals')) | Where-Object { -not $DistinctObjectIds.ContainsKey($_.Id) }
    $Members | Where-Object { (-not $DistinctObjectIds.ContainsKey($_.Id)) -and ($_.ObjectType -eq 'User') } | Select-Object Id | ForEach-Object { $DistinctObjectIds[$_.Id] = $true }
    $GroupMembers = @($Members | Where-Object { $_.ObjectType -eq 'Group' })

    switch ($Mode) {
        'Users' {
            $Members | Where-Object { $_.ObjectType -eq 'User' }
        }
        'Groups' {
            $GroupMembers
        }
        'ServicePrincipals' {
            $Members | Where-Object { $_.ObjectType -eq 'ServicePrincipal' }
        }
    }

    foreach ($GroupMember in $GroupMembers) {
        Get-GroupMembersRecursive -GroupId $GroupMember.ObjectId -Mode $Mode -VisitedGroups $VisitedGroups -DistinctObjectIds $DistinctObjectIds
    }
}

####################################################################################
# END OF FUNCTION DEFINITIONS
####################################################################################


New-Variable -Name USER_DETAILS_OUTPUT_FILE_PATH -Value ".\Audit_Users_AllDetails_For_MindBridge.csv" -Option Constant
New-Variable -Name USER_EMAILS_ONLY_OUTPUT_FILE_PATH -Value ".\Audit_Users_EmailsOnly_For_MindBridge.csv" -Option Constant
New-Variable -Name USER_UPNS_ONLY_OUTPUT_FILE_PATH -Value ".\Audit_Users_UPNsOnly_For_MindBridge.csv" -Option Constant
New-Variable -Name ERROR_GROUPS_OUTPUT_FILE_PATH -Value ".\ERRORS_CostCentre_AdGroups_NotFound.csv" -Option Constant

if ($CleanOutputFiles) {
    # Remove any existing output files to start fresh
    foreach ($filePath in @($USER_DETAILS_OUTPUT_FILE_PATH, $USER_EMAILS_ONLY_OUTPUT_FILE_PATH, $USER_UPNS_ONLY_OUTPUT_FILE_PATH, $ERROR_GROUPS_OUTPUT_FILE_PATH)) {
        if (Test-Path -Path $filePath) {
            Write-Host "Removing existing output file: $filePath"
            Remove-Item -Path $filePath -Force
        }
    }
} 
else {
    foreach ($filePath in @($USER_DETAILS_OUTPUT_FILE_PATH, $USER_EMAILS_ONLY_OUTPUT_FILE_PATH, $USER_UPNS_ONLY_OUTPUT_FILE_PATH, $ERROR_GROUPS_OUTPUT_FILE_PATH)) {
        if (Test-Path -Path $filePath) {
            Write-Warning "Delete all output files from previous runs: $filePath"
            exit
        }
    }
}  

# Read the list of cost centre AD groups from the CSV file
$costCentreAdGroups = Import-Csv -Path $InputGroupsCsvFilePath

if (-not $costCentreAdGroups) {
    Write-Error "No cost centre AD groups found in the CSV file."
    exit
}

try {
    # Import the required module for Entra ID (Azure AD) operations
    Import-Module Microsoft.Entra -ErrorAction Stop
}
catch {
    Write-Error "Failed to import Microsoft.Entra module. Please ensure it is installed."
    exit
}

try {

    # Identified users will be stored in this array for later reporting
    $outputUsers = @()
    # Groups not found in Entra ID will be stored in this array for later reporting
    $errorGroups = @()
    # Visited Groups (i.e. those already processed) to avoid infinite loops
    $visitedGroups = @{}
    # User IDs to avoid duplicates
    $userIds = @{}

    Write-Host "Connecting to Entra ID...`n"
    # Connect to Entra ID (requires Graph API permissions)
    Connect-Entra -Scopes 'GroupMember.Read.All' -TenantId $TenantId -NoWelcome

    # Loop around each cost centre AD group and retrieve its members recursively
    foreach ($adGroup in $costCentreAdGroups) {
        #Debug:
        #Write-Host "Got from csv: '$($adGroup)'"

        $rootGroup = Get-EntraGroup -Filter "DisplayName eq '$($adGroup.GroupName)'"
        if (-not $rootGroup) {
            Write-Warning "Group '$($adGroup.GroupName)' not found in Entra ID (Tenant: $TenantId).`n"
            $errorGroups += $adGroup.GroupName
        }
        else {
            Write-Host "Processing group '$($adGroup.GroupName)' (ID: $($rootGroup.Id))..."
            $outputUsers = Get-GroupMembersRecursive -GroupId $rootGroup.Id -Mode 'Users' -VisitedGroups $visitedGroups -DistinctObjectIds $userIds | Select-Object Id, DisplayName, ObjectType, mail, userPrincipalName

            #Debug:
            Write-Host "Found $($outputUsers.Count) distinct users in group '$($adGroup.GroupName)' and all nested groups that we haven't processed yet."
            write-Host "$($userIds.Count) distinct user Id's already processed."

            # Write users to csv output file
            Write-Host "Writing output to CSV...`n"
            $outputUsers | Select-Object Id, DisplayName, ObjectType, mail, userPrincipalName | Export-Csv -Path $USER_DETAILS_OUTPUT_FILE_PATH -NoTypeInformation -Append
            $outputUsers | Where-Object { -not [string]::IsNullOrWhiteSpace($_.mail) } | Select-Object mail | Export-Csv -Path $USER_EMAILS_ONLY_OUTPUT_FILE_PATH -NoTypeInformation -Append
            #$outputUsers | Where-Object { -not [string]::IsNullOrWhiteSpace($_.userPrincipalName) } | Select-Object userPrincipalName | Export-Csv -Path $USER_UPNS_ONLY_OUTPUT_FILE_PATH -NoTypeInformation -Append
        }
    }
    
    # Write any groups that were not found to a separate CSV file
    Write-Host "Writing $($errorGroups.Count) error groups (i.e. groups not found in Entra ID) to CSV...`n"
    $errorGroups | Select-Object @{Name="GroupName"; Expression={$_.ToString()}} | Export-Csv -Path $ERROR_GROUPS_OUTPUT_FILE_PATH -NoTypeInformation -Append

    Write-Host "Completed processing. Output files generated:"
    Write-Host "- User details: $USER_DETAILS_OUTPUT_FILE_PATH"
    Write-Host "- User emails only: $USER_EMAILS_ONLY_OUTPUT_FILE_PATH"
    #Write-Host "- User UPNs only: $USER_UPNS_ONLY_OUTPUT_FILE_PATH"
    Write-Host "- Error groups: $ERROR_GROUPS_OUTPUT_FILE_PATH"
    Write-Host "`n"
}
catch {
    Write-Error "Failed to get Audit Stream users from Entra ID. Error: $_`n"
    exit
}
finally {
    Write-Host "Disconnecting from Entra ID..."
    Disconnect-Entra
}
