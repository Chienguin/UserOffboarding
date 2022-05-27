<# 
.NAME
    User Offboarding
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
    $Results = ''
    $AD = txt_Username.text
    $UPN = txt_Email.text


}
#End functions

$btn_Offboard.Add_Click({OffboardUser})

[void]$frm_UserOffboarding.ShowDialog()