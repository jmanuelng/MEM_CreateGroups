<#

.Description
    Create Dynamic AzureAD groups from CSV file.

.Notes

    Import filename and path should be provided as a parameter. Default path is the execution path, default filename "MEM_CreateGroups.csv"
    Import file should, at least, contain the following headers (names should be exactly like that): "GroupType", "GroupDisplayName", "GroupsDescription", "GroupMembershipType", "GroupsMembershipRule", "GroupOwner" 

    Script created or based on Alex Durante's (tw:@ADurrante) Blog:
    Source: https://letsconfigmgr.com/bulk-create-intune-groups-script/#The_Script


To do:
    - Add parameters for filepath
    - Add parameters to skip verification

#>

#region Settings

$Error.Clear()
$errMessage = ""
$t = Get-Date
$ImportPath = ".\"
$ImportFilename = "MEM_CreateGroups.csv"
$GroupsObj = New-Object PSObject

#Give me some space, please
Write-Host "`n`n"

#endregion Settings


#region Functions

# Verify if running as Local Administrator
Function Test-IsAdmin {

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


# Install Azure AD Preview PS Module
function ConnectToAAD {

    if (Get-Module -ListAvailable -Name AzureADPreview) {
    } 
    else {
        Write-Host "Installing AzureAD PowerShell Module" -ForegroundColor Green       
        Install-Module -Name AzureADPreview -AllowClobber -Force
    }
    
    # Import Azure AD Preview Module
    Write-Host "Importing AzureADPreview Module" -ForegroundColor Green
    Import-Module AzureADPreview -Force
    
    # Sign into Azure AD
    Write-Host "Please log into AzureAD" -ForegroundColo Green
    Connect-AzureAD
    
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
    Write-Error $error.Exception.Message
    Exit 1
}

#Connect to Azure AD
ConnectToAAD

Write-Warning "Script only creates Dynamic Groups, as defined by GroupMemberhip ""DD"" or ""DU"""

$GroupsObj | Select-Object GroupType, GroupDisplayName, GroupDescription, GroupMembershipType, GroupMembershipRule, GroupOwner | Format-Table

#Create Groups

foreach ($Group in $GroupsObj) {

    if (($Group.GroupMembershipType -eq "DU") -or ($Group.GroupMembershipType -eq "DD")) {

        Write-Host "Creating Azure AD Dynamic Group: $($Group.GroupDisplayName)" -ForegroundColor Green

        if (-not (($null -eq $Group.GroupDisplayName) -and ($null -eq $Group.GroupMembershipRules))) {
            
            $Groupname = $Group.GroupDisplayName
            $GroupDesc = $Group.GroupDescription
            $GroupQuery = $Group.GroupMembershipRule
            $GroupOwn = Find-AzureADUser ($Group.GroupOwner)
            $confirmGroup = "N"
    
        }
        else {
            
            Write-Host "Can not create Group. `nGroup definition in file should specify rules and name. Please verify files and headers in file" -ForegroundColor Red
            continue
        }
    
        $confirmGroup = $(Write-Host "`tPlease confirm that you want to create Azure AD Group ""$Groupname"" (Y/N)?: " -ForegroundColor Green -NoNewline; Read-Host)
       
        if ($confirmGroup -eq "Y") {

            try {

                $AzureGroup = New-AzureADMSGroup `
                -DisplayName "$Groupname" `
                -Description "$GroupDesc" `
                -MailEnabled $false `
                -SecurityEnabled $true `
                -MailNickname "$($Groupname.replace(' ',''))" `
                -GroupTypes "DynamicMembership" `
                -MembershipRule "$GroupQuery" `
                -MembershipRuleProcessingState 'On' `
                -ErrorAction SilentlyContinue
                
            }
            catch {
                
                # If error, notify and continue.
                $errMessage = $_.Exception.ErrorContent.Message
                Write-Host "`tUnable to create $Groupname. `n`tERROR: $errMessage" -ForegroundColor Red
    
                continue
            }
    
    
            # Define Owner for the new Dynamic Group
            if ($null -ne $GroupOwn) {
                Add-AzureADGroupOwner -ObjectId "$($AzureGroup.Id)" -RefObjectId "$($GroupOwn.ObjectId)"
            }

            Write-Host "...Successfully created Azure AD Group $Groupname"

        }

        else {
            Write-Host "`tAzure AD Group $Groupname was not created." -ForegroundColor Yellow
        }

        
    }


}

# The end.
Write-Host "`nFinished creating Groups!"

# Give me some space, please.
Write-Host "`n`n"

#endregion Main"