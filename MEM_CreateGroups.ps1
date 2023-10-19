<#

.DESCRIPTION
    Create Dynamic AzureAD groups from CSV file. Can use "GroupsFile" as parameter for the name of CSV file with the list of groups to create, file should be on same directory. 
    It can operate on a specified tenant or the default tenant associated with the provided credentials.

.NOTES
    - The script now accepts an optional "TenantId" parameter to specify the Azure AD tenant where the groups will be created.
    - Import filename and path should be provided as a parameter. The default path is the execution path, and the default filename is "MEM_CreateGroups.csv".
    - The import file should contain the following headers: "GroupType", "GroupDisplayName", "GroupDescription", "GroupMembershipType", "GroupMembershipRule", "GroupOwner".
    - Valid strings for "GroupMembershipType" are "AA" for assigned groups, "DD" for Dynamic Device groups, and "DU" for Dynamic User groups.
    - The script checks for local administrative privileges and also verifies the AzureADPreview PowerShell module is installed.

    Script created or based on Alex Durante's (tw:@ADurrante) Blog:
    Source: https://letsconfigmgr.com/bulk-create-intune-groups-script/#The_Script

.EXAMPLE
    .\MEM_CreateGroups.ps1 -GroupsFile "MEM_CreateGroups.csv"
    Creates all groups listed in "MEM_CreateGroups.csv" file in the default tenant.

    .\MEM_CreateGroups.ps1 -GroupsFile "MEM_CreateGroups.csv" -GroupsConfirm "Y"
    Creates groups in "MEM_CreateGroups.csv" file in the default tenant, does not request confirmation before creating groups.

    .\MEM_CreateGroups.ps1 -GroupsFile "MEM_CreateGroups.csv" -TenantId "your-tenant-id"
    Creates all groups listed in "MEM_CreateGroups.csv" file in the specified tenant with ID "your-tenant-id".

    .\MEM_CreateGroups.ps1 -GroupsFile "MEM_CreateGroups.csv" -GroupsConfirm "Y" -TenantId "your-tenant-id"
    Creates groups in "MEM_CreateGroups.csv" file in the specified tenant with ID "your-tenant-id", does not request confirmation before creating groups.

#>


#region Settings

param (
        [Parameter()]
    [string]$GroupsFile = "MEM_CreateGroups.csv",
    # Confirm has to be "N" to skip confirmation for each Group to be created
    [string]$GroupsConfirm = "Y",
    [string]$TenantId = $null  # Add this parameter, default to $null
)

$Error.Clear()
$errMessage = ""
$t = Get-Date
$ImportPath = ".\"
$ImportFilename = $GroupsFile
$GroupsObj = New-Object PSObject

#Give me some space, please
Write-Host "`n`n"

#endregion Settings


#region Functions

# Verify if running as Local Administrator
function Test-IsAdmin {

    If (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {

        # Does not have Admin privileges
        Write-Host "Script needs to run with Administrative privileges"
        Return $false

    }
    else {

        # Yes, has Admin rights
        Write-Host "Adminitrator rights have been confirmed"
        Return $true
    
    }
    
}


# Install Azure AD Preview PS Module and connect to Tenant.
function ConnectToAAD {
    param (
        [string]$TenantId = $null  # Initialize TenantId to $null by default
    )

    # Check if AzureADPreview module is installed
    if (-not (Get-Module -ListAvailable -Name AzureADPreview)) {
        Write-Host "Installing AzureAD PowerShell Module" -ForegroundColor Green
        try {
            Install-Module -Name AzureADPreview -AllowClobber -Force -Scope CurrentUser
        } catch {
            Write-Host "Failed to install AzureADPreview Module. Please install it manually and rerun the script." -ForegroundColor Red
            return $false
        }
    }

    # Import Azure AD Preview Module
    Write-Host "Importing AzureADPreview Module" -ForegroundColor Green
    try {
        Import-Module AzureADPreview -Force
    } catch {
        Write-Host "Failed to import AzureADPreview Module. Please resolve the issue and rerun the script." -ForegroundColor Red
        return $false
    }

    # Sign into Azure AD
    Write-Host "Please log into AzureAD" -ForegroundColor Green
    try {
        if ($null -eq $TenantId -or $TenantId -eq "") {
            # Connect without TenantId if it's null or empty
            Connect-AzureAD
        } else {
            # Connect with TenantId
            Connect-AzureAD -TenantId $TenantId
        }
    } catch {
        Write-Host "Failed to connect to Azure AD. Please check your credentials or TenantId and try again." -ForegroundColor Red
        return $false
    }

    return $true
}



# Find user in AzureAD

function Find-AzureADUser {

    param
    (
        [Parameter(Mandatory=$true)]
        $aadUser
    )

    if (-not (($null -eq $aadUser) -or ($aadUser -eq ""))) {
        try {
            # Find user in Azure AD. If error, return $null.
            $aadUserObj = Get-AzureADUser -Filter "userPrincipalName eq '$aadUser'"
        }
        catch {
            # Error finding user, notify error, return null and keep going.
            Write-Error "Error finding user $aadUSer in Azure AD`n`t$($error.Exception.Message)"
            return $null
        }

        # Verify we have ID of Azure AD user
        if (($null -eq $aadUserObj.ObjectId) -or ($aadUserObj.ObjectId -eq "")) {
            # If we don't find owner ID, notify and return null.
            Write-Warning "Didn't find user $aadUser, owner property will be left blank"
            return $null
        }

        return $aadUserObj
    }

    else {

        #Blank o null query returns null result.
        return $null
    }


    
}



#endregion Functions


#######################################################################

#region Main

#Verify if running as Admin, exit if not.
if (-not(Test-IsAdmin)) {
    Exit 1
}


#Import file with groups to be created. End if can't import
try {

    $GroupsObj = Import-Csv -Path "$ImportPath$ImportFilename"
    
}
catch {
    Write-Host "Error importing $ImportFileName.`nPlease verify File name and/or extension.`nYou can use parameter ""GroupsFile""  to specify file name. `n`n" -ForegroundColor Red
    Write-Error $error.Exception.Message
    Exit 1
}

# Connect to Azure AD
$connectionResult = ConnectToAAD -TenantId $TenantId  # Pass the TenantId parameter

# Check if connection was successful
if ($connectionResult -eq $false) {
    Write-Host "Failed to connect to Azure AD. Exiting script." -ForegroundColor Red
    exit 1  # Exit the script with an error code
}


$GroupsObj | Select-Object GroupType, GroupDisplayName, GroupDescription, GroupMembershipType, GroupMembershipRule, GroupOwner | Format-Table

#Create Groups

foreach ($Group in $GroupsObj) {

    [string]$confirmGroup = $null
    [string]$GroupTypes = $null
    [bool]$GroupRoleAssign = $false

    if (($Group.GroupMembershipType -eq "DU") -or ($Group.GroupMembershipType -eq "DD") -or ($Group.GroupMembershipType -eq "AA")) {

        Write-Host "Creating Azure AD Group: $($Group.GroupDisplayName)" -ForegroundColor Green

        # Get group information to variables
        if (-not (($null -eq $Group.GroupDisplayName))) {
            
            $Groupname = $Group.GroupDisplayName
            $GroupDesc = $Group.GroupDescription
            $GroupOwn = Find-AzureADUser ($Group.GroupOwner)
            $confirmGroup = "N"

            if ($Group.GroupAadRoles -eq "YES") {
                $GroupRoleAssign = $true
            }
        
        }
        else {
            
            Write-Host "Can not create Group. `nGroup definition in file should specify a name. Please verify informations and headers in file" -ForegroundColor Red
            continue
        }

        # In case of a Dynamic group, get the query rules to a variable and define GroupType variable as DynamicMembership
        if (($Group.GroupMembershipType -eq "DU") -or ($Group.GroupMembershipType -eq "DD")) {

            # Validate that query exists, assign variable needed or Dynamic Group
            if (-not ($null -eq $Group.GroupMembershipRule)) {

                $GroupQuery = $Group.GroupMembershipRule
                $GroupTypes = "DynamicMembership"

            }
            else {
            
                Write-Host "Can not create Dynamic Group. `nDynamic Group definition should specify query rules. Please verify information and headers in file" -ForegroundColor Red
                continue
            }
    
        }

        #Get confirmation before creating group

        if ($Confirm -eq "N") {
            $confirmGroup = "Y"
        }
        else {
            $confirmGroup = $(Write-Host "`tPlease confirm that you want to create Azure AD Group ""$Groupname"" (Y/N)?: " -ForegroundColor Green -NoNewline; Read-Host)
        }

        if ($confirmGroup -eq "Y") {

            try {

                # Keep it simple, 2 different complete commands, depending if group is Dynamic or Assigned
                if ($Group.GroupMembershipType -eq "AA") {

                    $AzureGroup = New-AzureADMSGroup `
                    -DisplayName "$Groupname" `
                    -Description "$GroupDesc" `
                    -MailEnabled $false `
                    -SecurityEnabled $true `
                    -IsAssignableToRole $GroupRoleAssign `
                    -MailNickname "$($Groupname.replace(' ',''))" `
                    -ErrorAction Stop

                }
                else {

                    $AzureGroup = New-AzureADMSGroup `
                    -DisplayName "$Groupname" `
                    -Description "$GroupDesc" `
                    -MailEnabled $false `
                    -SecurityEnabled $true `
                    -IsAssignableToRole $GroupRoleAssign `
                    -MailNickname "$($Groupname.replace(' ',''))" `
                    -GroupTypes $GroupTypes `
                    -MembershipRule "$GroupQuery" `
                    -MembershipRuleProcessingState 'On' `
                    -ErrorAction Stop

                }
                
            }
            catch {
                
                # If error, notify and continue.
                $errMessage = $_.Exception.ErrorContent.Message
                Write-Host "`tUnable to create $Groupname. `n`tERROR: $errMessage" -ForegroundColor Red
    
                continue
            }
    
    
            # Define Owner for the new Group
            if ($null -ne $GroupOwn) {
                Add-AzureADGroupOwner -ObjectId "$($AzureGroup.Id)" -RefObjectId "$($GroupOwn.ObjectId)"
            }

            Write-Host "...Successfully created Azure AD Group $Groupname"

        }

        else {
            Write-Host "`tAzure AD Group $Groupname was not created." -ForegroundColor Yellow
        }

        
    }

    else {
        Write-Host "`tAzure AD Group $Groupname was not created. You must specify Group Membership Type." -ForegroundColor Yellow
        Write-Host "`tVerify Group file.`n`n" -ForegroundColor Yellow
        Write-Host "`t   Group Type can be AA for assgined group, DD for Dynamic Device, DU for Dynamic User. `n`t   Please verify files and header." -ForegroundColor Yellow
    }


}

# The end.
Write-Host "`nFinished creating Groups!"

# Give me some space, please.
Write-Host "`n`n"

#endregion Main"