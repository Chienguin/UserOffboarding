<# 
This is a general offboarding script. It will disable the user account in AD and move that user to a specified "Disabled
Users" OU, convert the O365 mailbox to a shared mailbox, remove from all distribution groups, hide from global address
list, and remove all licenses.

PREREQUISITES:
- Must be running Server 2012, 2012 R2, or 2016
- Have the Active Directory module for Windows Powershell installed
- Have the MSOnline module installed (for MFA-enabled accounts)
- Have the ExchangeOnlineManagement module installed

CHANGELOG:
Version: WIP-0.1 - 10/07/2021
- Initial commit
- Added framework for ITGlue API integration
- Added Active Directory integration
- Added O365 integration (no-MFA)

Version: WIP-0.2 - 10/08/2021
- Upgraded Active Directory integration to work without an implicit module
- Moves user account to Disabled Objects
- Disables user account
- Upgraded O365 integration to work with MFA enabled accounts
- Remove user from distribuion groups
- Converts user mailbox to shared mailbox
- Hides user from global address list
- Removes all user licenses
- Started integration for ITGlue - needing API Key

Version: WIP-0.3 - 10/11/2021
- Added editing Contact Type in ITGlue to show "GONE no longer with the company" for the disabled user

Version: WIP-0.4 - 05/26/2022
- Removed all ITGlue integration due to Kaseya mishap. Will possibly revisit in the future
- Added removal of Teams and Unified Groups in M365 offboarding
- Prep for GUI integration

Version WIP-0.5 - 05/27/2022
- Created basic GUI
- Integrated GUI into existing script

Author: Chien Nguyen
Version: WIP-0.5
#>
$DISABLED_OU = 'PATH_TO_DISABLED_USERS_OU'

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$frm_UserOffboarding             = New-Object system.Windows.Forms.Form
$frm_UserOffboarding.ClientSize  = New-Object System.Drawing.Point(407,461)
$frm_UserOffboarding.text        = "Courant User Offboarding"
$frm_UserOffboarding.TopMost     = $false

$btn_Offboard                    = New-Object system.Windows.Forms.Button
$btn_Offboard.text               = "Offboard"
$btn_Offboard.width              = 106
$btn_Offboard.height             = 47
$btn_Offboard.location           = New-Object System.Drawing.Point(138,97)
$btn_Offboard.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',12,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$lbl_AD                          = New-Object system.Windows.Forms.Label
$lbl_AD.text                     = "AD Username"
$lbl_AD.AutoSize                 = $true
$lbl_AD.width                    = 25
$lbl_AD.height                   = 10
$lbl_AD.location                 = New-Object System.Drawing.Point(9,23)
$lbl_AD.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$lbl_Email                       = New-Object system.Windows.Forms.Label
$lbl_Email.text                  = "M365 Email Address"
$lbl_Email.AutoSize              = $true
$lbl_Email.width                 = 25
$lbl_Email.height                = 10
$lbl_Email.location              = New-Object System.Drawing.Point(9,55)
$lbl_Email.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$lbl_Output                      = New-Object system.Windows.Forms.Label
$lbl_Output.AutoSize             = $false
$lbl_Output.width                = 397
$lbl_Output.height               = 285
$lbl_Output.location             = New-Object System.Drawing.Point(5,172)
$lbl_Output.Font                 = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$lbl_Output.ForeColor            = [System.Drawing.ColorTranslator]::FromHtml("#000000")
$lbl_Output.BackColor            = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")

$txt_Username                    = New-Object system.Windows.Forms.TextBox
$txt_Username.multiline          = $false
$txt_Username.width              = 261
$txt_Username.height             = 20
$txt_Username.location           = New-Object System.Drawing.Point(138,19)
$txt_Username.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$txt_Email                       = New-Object system.Windows.Forms.TextBox
$txt_Email.multiline             = $false
$txt_Email.width                 = 261
$txt_Email.height                = 20
$txt_Email.location              = New-Object System.Drawing.Point(138,51)
$txt_Email.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$frm_UserOffboarding.controls.AddRange(@($btn_Offboard,$lbl_AD,$lbl_Email,$lbl_Output,$txt_Username,$txt_Email))

#Logic functions
function OffboardUser {
    $RESULTS = ''
    $USER = txt_Username.text
    $UPN = txt_Email.text

    # Move user to Disabled Objects and disables the user account
    $RESULTS = 'Disabling user in AD and moving to Disabled Objects OU'
    $lbl_Output.text = $RESULTS
    Get-ADUser $USER | Move-ADObject -TargetPath $DISABLED_OU
    Disable-ADAccount -Identity $USER

    # Connect to Azure AD
    $RESULTS = $RESULTS + '`n' + 'Logging into Azure AD.'
    $lbl_Output.text = $RESULTS
    Connect-AzureAD

    # Connect to M365
    $RESULTS = $RESULTS + '`n' + 'Connecting to Exchange Online'
    $lbl_Output.text = $RESULTS
    Connect-ExchangeOnline
    $OffboardDN = (Get-Mailbox -Identity $UPN -IncludeInactiveMailbox).DistinguishedName

    # Removed user from all distribution groups
    $RESULTS = $RESULTS + '`n' + 'Removing user from distribution groups'
    $lbl_Output.text = $RESULTS
    Get-Recipient -Filter "Members -eq '$OffboardDN'" | foreach-object { 
    $RESULTS = $RESULTS + '`n' + "Removing user from $($_.name)"
    $lbl_Output.text = $RESULTS
    Remove-DistributionGroupMember -Identity $_.ExternalDirectoryObjectId -Member $OffboardDN -BypassSecurityGroupManagerCheck -Confirm:$false 
    }

    # Remove user from all Teams and Unified Groups
    $RESULTS = $RESULTS + '`n' + "Removing user from all Teams and Unified Groups"
    $lbl_Output.text = $RESULTS
    Get-Recipient -Filter "Members -eq '$OffboardDN'" -RecipientTypeDetails 'GroupMailbox' | foreach-object {
    $RESULTS = $RESULTS + '`n' + "Removing user from $($_.name)"
    $lbl_Output.text = $RESULTS
    Remove-UnifiedGroupLinks -Identity $_.ExternalDirectoryObjectId -Links $UPN -LinkType Member -Confirm:$false
    }

    # Converts user mailbox to a shared mailbox
    $RESULTS = $RESULTS + '`n' + "Converting user's mailbox to shared mailbox"
    $lbl_Output.text = $RESULTS
    Set-Mailbox $UPN -Type Shared

    # Hide user from global address list
    $RESULTS = $RESULTS + '`n' + "Hiding user from Global Address List"
    $lbl_Output.text = $RESULTS
    Set-Mailbox $UPN -HiddenFromAddressListsEnabled $true

    # Removing all of user's licenses
    $RESULTS = $RESULTS + '`n' + "Removing user licenses"
    $lbl_Output.text = $RESULTS

    $AssignedLicensesTable = Get-AzureADUser -ObjectId $UPN | Get-AzureADUserLicenseDetail | Select-Object @{n = "License"; e = { $_.SkuPartNumber } }, skuid 
    if ($AssignedLicensesTable) {
        $body = @{
            addLicenses    = @()
            removeLicenses = @($AssignedLicensesTable.skuid)
        }
        Set-AzureADUserLicense -ObjectId $UPN -AssignedLicenses $body
    }
 
    $RESULTS = $RESULTS + '`n' + "Removed licenses:" + '`n' + $AssignedLicensesTable
    $lbl_Output.text = $RESULTS
}

$btn_Offboard.Add_Click({OffboardUser})

[void]$frm_UserOffboarding.ShowDialog()
