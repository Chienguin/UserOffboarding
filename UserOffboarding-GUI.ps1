<# 
.SYNOPSIS
    General offboarding script. Has functions to disable a user account in localAD, move the user to
    a specified Disabled Users OU or to a specified Shared Mailbox OU for ADSync upkeep. In M365, it
    has the functions to convert the mailbox to a shared mailbox, remove user from all distribution 
    groups, remove user from all Teams, hide user from the Global Address List, and remove all 
    licenses from user.

.DESCRIPTION
    Pre-requisites:
    - Must be running Server 2012 or higher
    - Have the Active Directory module for Windows Powershell installed
    - Have the MSOnline module installed (for MFA-enabled accounts)
    - Have the ExchangeOnlineManagement module installed

    Prior to usage, please provide the fully qualified path to a Disabled Users OU and a Shared
    Mailbox OU. If there is no Shared Mailbox OU, the script will auto-default it to Disabled Users
    OU path. If there is no Disabled Users OU, the script will auto-default it to the OU of the user 
    in the textbox.

.NOTES
    File Name           : UserOffboarding-GUI.ps1
    Author              : Chien Nguyen (chien@gocourant.com)
    Current Version     : WIP-0.7 (testing needed)
    Copyright           : GNU General Public License v3.0

.LINK
    Script posted over:
    https://github.com/Chienguin/UserOffboarding
#>

# Global Variables
$DISABLED_OU = 'PATH_TO_DISABLED_USERS_OU'
$SHARED_MAILBOX_OU = 'PATH_TO_SHARED_MAILBOX_OU'

$OU_CHECKED = $false
$AD_CHECKED = $false
$M365_CHECKED = $false
$RESULTS = ''


#region FormGUI
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$frm_UserOffboarding             = New-Object system.Windows.Forms.Form
$frm_UserOffboarding.ClientSize  = New-Object System.Drawing.Point(830,427)
$frm_UserOffboarding.text        = "Courant User Offboarding"
$frm_UserOffboarding.TopMost     = $false

$lbl_AD                          = New-Object system.Windows.Forms.Label
$lbl_AD.text                     = "Username"
$lbl_AD.AutoSize                 = $true
$lbl_AD.width                    = 25
$lbl_AD.height                   = 10
$lbl_AD.location                 = New-Object System.Drawing.Point(15,38)
$lbl_AD.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$lbl_Email                       = New-Object system.Windows.Forms.Label
$lbl_Email.text                  = "Email Address"
$lbl_Email.AutoSize              = $true
$lbl_Email.width                 = 25
$lbl_Email.height                = 10
$lbl_Email.location              = New-Object System.Drawing.Point(15,38)
$lbl_Email.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$txt_Username                    = New-Object system.Windows.Forms.TextBox
$txt_Username.multiline          = $false
$txt_Username.width              = 261
$txt_Username.height             = 20
$txt_Username.location           = New-Object System.Drawing.Point(122,35)
$txt_Username.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$txt_Email                       = New-Object system.Windows.Forms.TextBox
$txt_Email.multiline             = $false
$txt_Email.width                 = 261
$txt_Email.height                = 20
$txt_Email.location              = New-Object System.Drawing.Point(122,35)
$txt_Email.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$lbl_Output                      = New-Object system.Windows.Forms.Label
$lbl_Output.AutoSize             = $false
$lbl_Output.width                = 402
$lbl_Output.height               = 419
$lbl_Output.location             = New-Object System.Drawing.Point(424,4)
$lbl_Output.Font                 = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$lbl_Output.ForeColor            = [System.Drawing.ColorTranslator]::FromHtml("#000000")
$lbl_Output.BackColor            = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")

$pnl_ActiveDirectory             = New-Object system.Windows.Forms.Panel
$pnl_ActiveDirectory.height      = 127
$pnl_ActiveDirectory.width       = 400
$pnl_ActiveDirectory.location    = New-Object System.Drawing.Point(11,22)

$lbl_ActiveDirectory             = New-Object system.Windows.Forms.Label
$lbl_ActiveDirectory.text        = "Active Directory"
$lbl_ActiveDirectory.AutoSize    = $true
$lbl_ActiveDirectory.width       = 25
$lbl_ActiveDirectory.height      = 10
$lbl_ActiveDirectory.location    = New-Object System.Drawing.Point(135,7)
$lbl_ActiveDirectory.Font        = New-Object System.Drawing.Font('Microsoft Sans Serif',14,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$lbl_ADDescription               = New-Object system.Windows.Forms.Label
$lbl_ADDescription.text          = "Description"
$lbl_ADDescription.AutoSize      = $true
$lbl_ADDescription.width         = 25
$lbl_ADDescription.height        = 10
$lbl_ADDescription.location      = New-Object System.Drawing.Point(15,64)
$lbl_ADDescription.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$txt_Description                 = New-Object system.Windows.Forms.TextBox
$txt_Description.multiline       = $false
$txt_Description.width           = 261
$txt_Description.height          = 20
$txt_Description.location        = New-Object System.Drawing.Point(122,61)
$txt_Description.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$chk_DisableUser                 = New-Object system.Windows.Forms.CheckBox
$chk_DisableUser.text            = "Disable User"
$chk_DisableUser.AutoSize        = $true
$chk_DisableUser.width           = 95
$chk_DisableUser.height          = 20
$chk_DisableUser.location        = New-Object System.Drawing.Point(15,95)
$chk_DisableUser.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$chk_SharedOU                    = New-Object system.Windows.Forms.CheckBox
$chk_SharedOU.text               = "Shared Mailbox OU"
$chk_SharedOU.AutoSize           = $true
$chk_SharedOU.width              = 95
$chk_SharedOU.height             = 20
$chk_SharedOU.location           = New-Object System.Drawing.Point(245,95)
$chk_SharedOU.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$pnl_Microsoft365                = New-Object system.Windows.Forms.Panel
$pnl_Microsoft365.height         = 190
$pnl_Microsoft365.width          = 400
$pnl_Microsoft365.location       = New-Object System.Drawing.Point(11,161)

$lbl_Microsoft365                = New-Object system.Windows.Forms.Label
$lbl_Microsoft365.text           = "Microsoft 365"
$lbl_Microsoft365.AutoSize       = $true
$lbl_Microsoft365.width          = 25
$lbl_Microsoft365.height         = 10
$lbl_Microsoft365.location       = New-Object System.Drawing.Point(137,7)
$lbl_Microsoft365.Font           = New-Object System.Drawing.Font('Microsoft Sans Serif',14,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$chk_DistroGroups                = New-Object system.Windows.Forms.CheckBox
$chk_DistroGroups.text           = "Remove from Distribution Groups"
$chk_DistroGroups.AutoSize       = $true
$chk_DistroGroups.width          = 95
$chk_DistroGroups.height         = 20
$chk_DistroGroups.location       = New-Object System.Drawing.Point(82,158)
$chk_DistroGroups.Font           = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$chk_Teams                       = New-Object system.Windows.Forms.CheckBox
$chk_Teams.text                  = "Remove from Teams"
$chk_Teams.AutoSize              = $true
$chk_Teams.width                 = 95
$chk_Teams.height                = 20
$chk_Teams.location              = New-Object System.Drawing.Point(15,81)
$chk_Teams.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$chk_SharedMailbox               = New-Object system.Windows.Forms.CheckBox
$chk_SharedMailbox.text          = "Convert to Shared Mailbox"
$chk_SharedMailbox.AutoSize      = $true
$chk_SharedMailbox.width         = 95
$chk_SharedMailbox.height        = 20
$chk_SharedMailbox.location      = New-Object System.Drawing.Point(197,81)
$chk_SharedMailbox.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$chk_AddressBook                 = New-Object system.Windows.Forms.CheckBox
$chk_AddressBook.text            = "Hide from Address Book"
$chk_AddressBook.AutoSize        = $true
$chk_AddressBook.width           = 95
$chk_AddressBook.height          = 20
$chk_AddressBook.location        = New-Object System.Drawing.Point(197,119)
$chk_AddressBook.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$chk_Licenses                    = New-Object system.Windows.Forms.CheckBox
$chk_Licenses.text               = "Remove licenses"
$chk_Licenses.AutoSize           = $true
$chk_Licenses.width              = 95
$chk_Licenses.height             = 20
$chk_Licenses.location           = New-Object System.Drawing.Point(15,119)
$chk_Licenses.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$btn_Offboard                    = New-Object system.Windows.Forms.Button
$btn_Offboard.text               = "Offboard User"
$btn_Offboard.width              = 228
$btn_Offboard.height             = 30
$btn_Offboard.location           = New-Object System.Drawing.Point(90,375)
$btn_Offboard.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$pnl_ActiveDirectory.controls.AddRange(@($lbl_AD,$txt_Username,$lbl_ActiveDirectory,$lbl_ADDescription,$txt_Description,$chk_DisableUser,$chk_SharedOU))
$pnl_Microsoft365.controls.AddRange(@($lbl_Email,$txt_Email,$lbl_Microsoft365,$chk_DistroGroups,$chk_Teams,$chk_SharedMailbox,$chk_AddressBook,$chk_Licenses))
$frm_UserOffboarding.controls.AddRange(@($lbl_Output,$pnl_ActiveDirectory,$pnl_Microsoft365,$btn_Offboard))
#endregion FormGUI

#region functions
function isConnected {
    try {
        $var = Get-AzureADTenantDetail
    } 
    catch [Microsoft.Open.Azure.AD.CommonLibrary.AadNeedAuthenticationException] {
        $RESULTS = "Logging into Azure AD"
        $lbl_Output.text = $RESULTS
        Connect-AzureAD
    }

    if (!(Get-PSSession | Where-Object {$_.Name -match 'ExchangeOnline' -and $_.Availability -eq 'Available'})) {
        $RESULTS = $RESULTS + '`n' + "Connecting to Exchange Online"
        $lbl_Output.text = $RESULTS
        Connect-ExchangeOnline
    }

    return
}

function CheckInput {
    if ($AD_CHECKED -eq $false) {
        $inputUN = $txt_Username.text

        if ($inputUN = Get-ADUser -Filter "SamAccountName -eq '$inputUN'") {
            $AD_CHECKED = $true
        } else {
            $RESULTS = "$inputUN is not a valid user!"
            $lbl_Output.text = $RESULTS
            return
        }
    }

    if ($OU_CHECKED -eq $false) {
        if (($DISABLED_OU -eq 'PATH_TO_DISABLED_USERS_OU') -OR ($DISABLED_OU -eq "")){
            $DN = (Get-ADUser -Identity $User -Properties DistinguishedName).DistinguishedName
            $DISABLED_OU = $DN.Substring($DN.IndexOf('OU='))
        }
    
        if (($SHARED_MAILBOX_OU -eq 'PATH_TO_SHARED_MAILBOX_OU') -OR ($SHARED_MAILBOX_OU -eq "")){
            $SHARED_MAILBOX_OU = $DISABLED_OU
        }

        $OU_CHECKED = $true
    }

    if ($M365_CHECKED -eq $false) {
        {isConnected}

        $inputEmail = $txt_Email.text

        if ($inputEmail = Get-EXOMailbox -Filter "EmailAddresses -eq '$inputEmail'") {
            $M365_CHECKED = $true
        } else {
            $RESULTS = "$inputEmail is not a valid email address!"
            $lbl_Output.text = $RESULTS
            return
        }
    }

    return
}

function OffboardUser {
    $USER = txt_Username.text
    $UPN = txt_Email.text
    $DESCRIPTION = $txt_Description.text

    {CheckInput}

    if (($AD_CHECKED -eq $false) -OR ($OU_CHECKED -eq $false) -OR ($M365_CHECKED -eq $false)) {
        return
    }

    # Disables User Account if box is checked
    if ($chk_DisableUser.Checked -eq $true) {
        $RESULTS = "Disabling $USER"
        $lbl_Output.text = $RESULTS
        Disable-ADAccount -Identity $USER
    }
    
    # Checks for which OU user account should be moved to
    if ($chk_SharedOU.Checked -eq $true) {
        $RESULTS = $RESULTS + '`n' + "Moving $USER to the Shared Mailbox OU"
        $lbl_Output.text = $RESULTS
        Get-ADUser $USER | Move-ADObject -TargetPath $SHARED_MAILBOX_OU
    } else {
        $RESULTS = $RESULTS + '`n' + "Moving $USER to the Disabled Users OU"
        $lbl_Output.text = $RESULTS
        Get-ADUser $USER | Move-ADObject -TargetPath $DISABLED_OU
    }

    # Changes User Account's Description
    $previousDesc = (Get-ADUser $USER -Properties Description).Description
    $newDesc = $previousDesc + " - " + $DESCRIPTION
    $RESULTS = $RESULTS + '`n' + "Updating description for $USER"
    $lbl_Output.text = $RESULTS
    Set-ADUser $USER -Replace @{Description = $newDesc}
    

    # Begin M365 block
    $OffboardDN = (Get-EXOMailbox -Identity $UPN -IncludeInactiveMailbox).DistinguishedName

    # Remove user from all distribution groups
    if ($chk_DistroGroups.Checked -eq $true) {
        $RESULTS = $RESULTS + '`n' + "Removing $UPN from distribution groups"
        $lbl_Output.text = $RESULTS
        Get-EXORecipient -Filter "Members -eq '$OffboardDN'" | foreach-object { 
            $RESULTS = $RESULTS + '`n' + "Removing $UPN from $($_.name)"
            $lbl_Output.text = $RESULTS
            Remove-DistributionGroupMember -Identity $_.ExternalDirectoryObjectId -Member $OffboardDN -BypassSecurityGroupManagerCheck -Confirm:$false 
        }
    }
    

    # Remove user from all Teams and Unified Groups
    if ($chk_Teams.Checked -eq $true){
        $RESULTS = $RESULTS + '`n' + "Removing $UPN from all Teams and Unified Groups"
        $lbl_Output.text = $RESULTS
        Get-EXORecipient -Filter "Members -eq '$OffboardDN'" -RecipientTypeDetails 'GroupMailbox' | foreach-object {
            $RESULTS = $RESULTS + '`n' + "Removing $UPN from $($_.name)"
            $lbl_Output.text = $RESULTS
            Remove-UnifiedGroupLinks -Identity $_.ExternalDirectoryObjectId -Links $UPN -LinkType Member -Confirm:$false
        }
    }

    # Convert user mailbox to a shared mailbox
    if ($chk_SharedMailbox.Checked -eq $true) {
        $RESULTS = $RESULTS + '`n' + "Converting $UPN to a Shared Mailbox"
        $lbl_Output.text = $RESULTS
        Set-Mailbox $UPN -Type Shared
    }

    # Hide user from global address list
    if ($chk_AddressBook.Checked -eq $true) {
        $RESULTS = $RESULTS + '`n' + "Hiding $UPN from Global Address List"
        $lbl_Output.text = $RESULTS
        Set-Mailbox $UPN -HiddenFromAddressListsEnabled $true
    }

    # Removing all of user's licenses
    if ($chk_Licenses.Checked -eq $true) {
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

    return
}
#endregion functions

$btn_Offboard.Add_Click({OffboardUser})

[void]$frm_UserOffboarding.ShowDialog()
