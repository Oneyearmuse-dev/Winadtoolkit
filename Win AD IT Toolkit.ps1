# IT Toolkit GUI for Scripts written Oneyearmuse (Created 28/03/2025 Modiefied 16/05/2025)
# Version 1.0 Open Source - You are free to use and modify this script as you see fit - just give credit where credit is due.

$ErrorActionPreference = "SilentlyContinue"
Set-ExecutionPolicy Bypass -Scope CurrentUser -Force

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#[System.Windows.Forms.Application]::EnableVisualStyles();  

Remove-Module -Name Start-Svc,Remove-TeamFolders,Get-CDriveSpace,Get-PCType,Get-Model,Get-AveragePing,Get-ServicesStatus,Get-Network,Get-PCInfo,Get-Boottime,Remove-TeamFolders,Get-LastLoggedOnUser,Disable-Hibernate,Remove-OldUserProfiles,Remove-SWDFolder,Remove-WV2,Add-PCGroup,Add-FilesDesktop,Set-RemoteRegistrySvc,Set-StopStartService,Get-ADUsername,Add-OutputText -ErrorAction SilentlyContinue

#$samusername = ""

$pathExplorerPlus = "C:\Win IT Toolkit\Exp++"
$pathToolkit = "C:\Win IT Toolkit"

# Import functions for the scripts

    Import-Module -Name ActiveDirectory

    #PC Info Modules
    Import-Module -Name "$pathToolkit\PCInfoModules.psm1" -Force
    #C Drive Space Modules
    Import-Module -Name "$pathToolkit\CDriveSpaceModules.psm1" -Force
    #Misc Modules
    Import-Module -Name "$pathToolkit\MiscModules.psm1" -Force
    #User Modules
    Import-Module -Name "$pathToolkit\UserModules.psm1" -Force

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Win AD IT Toolkit V1 (Oneyearmuse)"
$form.Size = New-Object System.Drawing.Size(1000, 900)
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

# Create the font style
$fontStyle = [System.Drawing.FontStyle]::Bold
$fontStyleN = [System.Drawing.FontStyle]::Regular
$fontStyleUL = [System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Underline

#Create Tooltip Object
$toolTip = New-Object System.Windows.Forms.ToolTip

# Create the Users info group
$usersGroup = New-Object System.Windows.Forms.GroupBox
$usersGroup.Text = "User Info"
$usersGroup.Location = New-Object System.Drawing.Point(15, 30)
$usersGroup.Size = New-Object System.Drawing.Size(340, 475)

# Create the Users actions group
$usersactionsGroup = New-Object System.Windows.Forms.GroupBox
$usersactionsGroup.Text = "User Actions"
$usersactionsGroup.Location = New-Object System.Drawing.Point(15, 510)
$usersactionsGroup.Size = New-Object System.Drawing.Size(340, 150)

$form.Controls.Add($usersactionsGroup)

$form.Controls.Add($usersGroup)

#Users Panel Start
$usernameLabel = New-Object System.Windows.Forms.Label
$usernameLabel.Text = "Username"
$usernameLabel.Location = New-Object System.Drawing.Point(10, 20)
$usernameLabel.AutoSize = $true
$usernameLabel.Font = New-Object System.Drawing.Font("Arial", 10, $fontStyle)
$usersGroup.Controls.Add($usernameLabel)

$usernameTextBox = New-Object System.Windows.Forms.TextBox
$usernameTextBox.Location = New-Object System.Drawing.Point(10, 40)
$usernameTextBox.Size = New-Object System.Drawing.Size(120, 20)
$usernameTextBox.Add_KeyDown({
    param($sender, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        if ([string]::IsNullOrWhiteSpace($usernameTextBox.Text)) {
            Add-OutputText "Username field cannot be empty. Please enter a username."
        } else {
            Add-OutputText "Extracting User info..."
            # Add your Username search script here
            Get-ADUsername -param $usernameTextBox.Text
        }
    }
})
$usersGroup.Controls.Add($usernameTextBox)

#Wildcard Search Button
$wildcardsrcButton = New-Object System.Windows.Forms.Button
$wildcardsrcButton.Text = "Wildcard Search"
$wildcardsrcButton.Location = New-Object System.Drawing.Point(135, 13)
$wildcardsrcButton.AutoSize = $true

$wildcardsrcButton.Add_Click({
    if ([string]::IsNullOrWhiteSpace($usernameTextBox.Text)) {
        Add-OutputText "Username field cannot be empty. Please enter a partial username."
        return
    }
    Add-OutputText "Searching for AD Username..."
    Get-WildcardADUsername -param $usernameTextBox.Text
})
$usersGroup.Controls.Add($wildcardsrcButton)
$toolTip.SetToolTip($wildcardsrcButton, "Enter partial username to populate box below.")

$wildbutton = New-Object System.Windows.Forms.Button
$wildbutton.Text = "*"
$wildbutton.Location = New-Object System.Drawing.Point(233,14)
$wildButton.AutoSize = $fale
$wildbutton.Size = New-Object System.Drawing.Size(17,23)
$wildbutton.Add_Click({
    $selectedUser = $wildcardsrcDropDownBox.SelectedItem
    if ($selectedUser) {
        Add-OutputText "Extracting User info..."
        # Add your Username search script here
        Get-ADUsername -param $selectedUser
        $usernameTextBox.Text = $selectedUser
    } else {
        [System.Windows.Forms.MessageBox]::Show("Please select a user.")
    }
})

$usersGroup.Controls.Add($wildButton)
$toolTip.SetToolTip($wildButton, "Search for AD Username on selected user.")

$emailsrcButton = New-Object System.Windows.Forms.Button
$emailsrcButton.Text = "Email Search"
$emailsrcButton.Location = New-Object System.Drawing.Point(252, 13)
$emailsrcButton.AutoSize = $true

$emailsrcButton.Add_Click({
    if ([string]::IsNullOrWhiteSpace($emailTextBox.Text)) {
        Add-OutputText "Email field cannot be empty. Please enter an email address."
        return
    }
    Add-OutputText "Searching for AD Username using email..."
    Get-EmailADUsername -param $emailTextBox.Text
})
$usersGroup.Controls.Add($emailsrcButton)
$toolTip.SetToolTip($emailsrcButton, "Enter Email Addrress to populate box below.")

$wildcardsrcDropDownBox = New-Object System.Windows.Forms.ComboBox
$wildcardsrcDropDownBox.Location = New-Object System.Drawing.Point(135, 40)
$wildcardsrcDropDownBox.Size = New-Object System.Drawing.Size(200, 20)
$wildcardsrcDropDownBox.DropDownHeight = 200
$usersGroup.Controls.Add($wildcardsrcDropDownBox)

#User Info Labels and Text Boxes
#Display Name
$displaynameLabel = New-Object System.Windows.Forms.Label
$displaynameLabel.Location = New-Object System.Drawing.Point(5,85)
$displaynameLabel.Size = New-Object System.Drawing.Size(130,20)
$displaynameLabel.Text = "Display Name"
$displaynameLabel.Font = New-Object System.Drawing.Font("Arial", 9, $fontStyle)
$usersGroup.Controls.Add($displaynameLabel)

$displaynameTextBox = New-Object System.Windows.Forms.TextBox
$displaynameTextBox.Location = New-Object System.Drawing.Point(135,80)
$displaynameTextBox.Size = New-Object System.Drawing.Size(200,20)
$displaynameTextBox.ReadOnly = $true;
$displaynameTextBox.Text = ""
$displaynameTextBox.Font = New-Object System.Drawing.Font("Arial", 10, $fontStyleN)
$usersGroup.Controls.Add($displaynameTextBox)

#Job Title
$jobtitleLabel = New-Object System.Windows.Forms.Label
$jobtitleLabel.Location = New-Object System.Drawing.Point(5,109)
$jobtitleLabel.Size = New-Object System.Drawing.Size(130,20)
$jobtitleLabel.Text = "Job Title"
$jobtitleLabel.Font = New-Object System.Drawing.Font("Arial", 9, $fontStyle)
$usersGroup.Controls.Add($jobtitleLabel)

$jobtitleTextBox = New-Object System.Windows.Forms.TextBox
$jobtitleTextBox.Location = New-Object System.Drawing.Point(135,104)
$jobtitleTextBox.Size = New-Object System.Drawing.Size(200,20)
$jobtitleTextBox.Text = ""
$jobtitleTextBox.ReadOnly = $true;
$jobtitleTextBox.Font = New-Object System.Drawing.Font("Arial", 10, $fontStyleN)
$usersGroup.Controls.Add($jobtitleTextBox)

#Department
$deptLabel = New-Object System.Windows.Forms.Label
$deptLabel.Location = New-Object System.Drawing.Point(5,132)
$deptLabel.Size = New-Object System.Drawing.Size(130,20)
$deptLabel.Text = "Department"
$deptLabel.Font = New-Object System.Drawing.Font("Arial", 9, $fontStyle)
$usersGroup.Controls.Add($deptLabel)

$deptTextBox = New-Object System.Windows.Forms.TextBox
$deptTextBox.Location = New-Object System.Drawing.Point(135,128)
$deptTextBox.Size = New-Object System.Drawing.Size(200,20)
$deptTextBox.ReadOnly = $true;
$deptTextBox.Text = ""
$deptTextBox.Font = New-Object System.Drawing.Font("Arial", 10, $fontStyleN)
$usersGroup.Controls.Add($deptTextBox)

#Email Address
$emailLabel = New-Object System.Windows.Forms.Label
$emailLabel.Location = New-Object System.Drawing.Point(5,156)
$emailLabel.Size = New-Object System.Drawing.Size(130,20)
$emailLabel.Text = "Email Address"
$emailLabel.Font = New-Object System.Drawing.Font("Arial", 9, $fontStyle)
$usersGroup.Controls.Add($emailLabel)

$emailTextBox = New-Object System.Windows.Forms.TextBox
$emailTextBox.Location = New-Object System.Drawing.Point(135,152)
$emailTextBox.Size = New-Object System.Drawing.Size(200,20)
$emailTextBox.ReadOnly = $false;
$emailTextBox.Text = ""  
$emailTextBox.Font = New-Object System.Drawing.Font("Arial", 10, $fontStyleN)
$usersGroup.Controls.Add($emailTextBox)

#Phone Number
$phoneLabel = New-Object System.Windows.Forms.Label
$phoneLabel.Location = New-Object System.Drawing.Point(5,180)
$phoneLabel.Size = New-Object System.Drawing.Size(130,20)
$phoneLabel.Text = "Telephone Number"
$phoneLabel.Font = New-Object System.Drawing.Font("Arial", 9, $fontStyle)
$usersGroup.Controls.Add($phoneLabel)

$phoneTextBox = New-Object System.Windows.Forms.TextBox
$phoneTextBox.Location = New-Object System.Drawing.Point(135,176)
$phoneTextBox.Size = New-Object System.Drawing.Size(200,20)
$phoneTextBox.ReadOnly = $true;
$phoneTextBox.Text = ""
$phoneTextBox.Font = New-Object System.Drawing.Font("Arial", 10, $fontStyleN)
$usersGroup.Controls.Add($phoneTextBox)

#Description
$descrLabel = New-Object System.Windows.Forms.Label
$descrLabel.Location = New-Object System.Drawing.Point(5,205)
$descrLabel.Size = New-Object System.Drawing.Size(130,20)
$descrLabel.Text = "Description"
$descrLabel.Font = New-Object System.Drawing.Font("Arial", 9, $fontStyle)
$usersGroup.Controls.Add($descrLabel)

$descrTextBox = New-Object System.Windows.Forms.TextBox
$descrTextBox.Location = New-Object System.Drawing.Point(135,200)
$descrTextBox.Size = New-Object System.Drawing.Size(200,20)
$descrTextBox.ReadOnly = $true;
$descrTextBox.Text = ""
$descrTextBox.Font = New-Object System.Drawing.Font("Arial", 10, $fontStyleN)
$usersGroup.Controls.Add($descrTextBox)

#Date Created
$datecrLabel = New-Object System.Windows.Forms.Label
$datecrLabel.Location = New-Object System.Drawing.Point(5,230)
$datecrLabel.Size = New-Object System.Drawing.Size(130,20)
$datecrLabel.Text = "Date Created"
$datecrLabel.Font = New-Object System.Drawing.Font("Arial", 9, $fontStyle)
$usersGroup.Controls.Add($datecrLabel)

$datecrTextBox = New-Object System.Windows.Forms.TextBox
$datecrTextBox.Location = New-Object System.Drawing.Point(135,223)
$datecrTextBox.Size = New-Object System.Drawing.Size(200,20)
$datecrTextBox.ReadOnly = $true;
$datecrTextBox.Text = ""
$datecrTextBox.Font = New-Object System.Drawing.Font("Arial", 10, $fontStyleN)
$usersGroup.Controls.Add($datecrTextBox)

#Date Modified
$datemodLabel = New-Object System.Windows.Forms.Label
$datemodLabel.Location = New-Object System.Drawing.Point(5,255)
$datemodLabel.Size = New-Object System.Drawing.Size(130,20)
$datemodLabel.Text = "Date Modified"
$datemodLabel.Font = New-Object System.Drawing.Font("Arial", 9, $fontStyle)
$usersGroup.Controls.Add($datemodLabel)

$datemodTextBox = New-Object System.Windows.Forms.TextBox
$datemodTextBox.Location = New-Object System.Drawing.Point(135,247)
$datemodTextBox.Size = New-Object System.Drawing.Size(200,20)
$datemodTextBox.ReadOnly = $true;
$datemodTextBox.Text = ""
$datemodTextBox.Font = New-Object System.Drawing.Font("Arial", 10, $fontStyleN)
$usersGroup.Controls.Add($datemodTextBox)

#Last Logon
$lastlogondateLabel = New-Object System.Windows.Forms.Label
$lastlogondateLabel.Location = New-Object System.Drawing.Point(5,280)
$lastlogondateLabel.Size = New-Object System.Drawing.Size(130,20)
$lastlogondateLabel.Text = "Last Logon Date"
$lastlogondateLabel.Font = New-Object System.Drawing.Font("Arial", 9, $fontStyle)
$usersGroup.Controls.Add($lastlogondateLabel)

$lastlogondateTextBox = New-Object System.Windows.Forms.TextBox
$lastlogondateTextBox.Location = New-Object System.Drawing.Point(135,271)
$lastlogondateTextBox.Size = New-Object System.Drawing.Size(200,20)
$lastlogondateTextBox.ReadOnly = $true;
$lastlogondateTextBox.Text = ""
$lastlogondateTextBox.Font = New-Object System.Drawing.Font("Arial", 10, $fontStyleN)
$usersGroup.Controls.Add($lastlogondateTextBox)

#Password Last Set
$pwddateLabel = New-Object System.Windows.Forms.Label
$pwddateLabel.Location = New-Object System.Drawing.Point(5,301)
$pwddateLabel.Size = New-Object System.Drawing.Size(130,20)
$pwddateLabel.Text = "Password Last Set"
$pwddateLabel.Font = New-Object System.Drawing.Font("Arial", 9, $fontStyle)
$usersGroup.Controls.Add($pwddateLabel)

$pwddateTextBox = New-Object System.Windows.Forms.TextBox
$pwddateTextBox.Location = New-Object System.Drawing.Point(135,295)
$pwddateTextBox.Size = New-Object System.Drawing.Size(200,20)
$pwddateTextBox.ReadOnly = $true;
$pwddateTextBox.Text = ""
$pwddateTextBox.Font = New-Object System.Drawing.Font("Arial", 10, $fontStyleN)
$usersGroup.Controls.Add($pwddateTextBox)

#Password Expired
$pwdexpLabel = New-Object System.Windows.Forms.Label
$pwdexpLabel.Location = New-Object System.Drawing.Point(5,324)
$pwdexpLabel.Size = New-Object System.Drawing.Size(130,20)
$pwdexpLabel.Text = "Password Expired"
$pwdexpLabel.Font = New-Object System.Drawing.Font("Arial", 9, $fontStyle)
$usersGroup.Controls.Add($pwdexpLabel)

$pwdexpTextBox = New-Object System.Windows.Forms.TextBox
$pwdexpTextBox.Location = New-Object System.Drawing.Point(135,319)
$pwdexpTextBox.Size = New-Object System.Drawing.Size(200,20)
$pwdexpTextBox.ReadOnly = $true;
$pwdexpTextBox.Text = ""
$pwdexpTextBox.Font = New-Object System.Drawing.Font("Arial", 10, $fontStyleN)
$usersGroup.Controls.Add($pwdexpTextBox)

#Never Expires
$pwdnexpLabel = New-Object System.Windows.Forms.Label
$pwdnexpLabel.Location = New-Object System.Drawing.Point(5,349)
$pwdnexpLabel.Size = New-Object System.Drawing.Size(130,20)
$pwdnexpLabel.Text = "Never Expires"
$pwdnexpLabel.Font = New-Object System.Drawing.Font("Arial", 9, $fontStyle)
$usersGroup.Controls.Add($pwdnexpLabel)

$pwdnexpTextBox = New-Object System.Windows.Forms.TextBox
$pwdnexpTextBox.Location = New-Object System.Drawing.Point(135,343)
$pwdnexpTextBox.Size = New-Object System.Drawing.Size(200,20)
$pwdnexpTextBox.ReadOnly = $true;
$pwdnexpTextBox.Text = ""
$pwdnexpTextBox.Font = New-Object System.Drawing.Font("Arial", 10, $fontStyleN)
$usersGroup.Controls.Add($pwdnexpTextBox)

#Last Bad Password
$lastbadpwdLabel = New-Object System.Windows.Forms.Label
$lastbadpwdLabel.Location = New-Object System.Drawing.Point(5,372)
$lastbadpwdLabel.Size = New-Object System.Drawing.Size(130,20)
$lastbadpwdLabel.Text = "Last Bad Password"
$lastbadpwdLabel.Font = New-Object System.Drawing.Font("Arial", 9, $fontStyle)
$usersGroup.Controls.Add($lastbadpwdLabel)

$lastbadpwdTextBox = New-Object System.Windows.Forms.TextBox
$lastbadpwdTextBox.Location = New-Object System.Drawing.Point(135,367)
$lastbadpwdTextBox.Size = New-Object System.Drawing.Size(200,20)
$lastbadpwdTextBox.ReadOnly = $true;
$lastbadpwdTextBox.Text = ""
$lastbadpwdTextBox.Font = New-Object System.Drawing.Font("Arial", 10, $fontStyleN)
$usersGroup.Controls.Add($lastbadpwdTextBox)

#Locked Out
$lockedoutLabel = New-Object System.Windows.Forms.Label
$lockedoutLabel.Location = New-Object System.Drawing.Point(5,396)
$lockedoutLabel.Size = New-Object System.Drawing.Size(130,20)
$lockedoutLabel.Text = "Locked Out"
$lockedoutLabel.Font = New-Object System.Drawing.Font("Arial", 9, $fontStyle)
$usersGroup.Controls.Add($lockedoutLabel)

$lockedoutTextBox = New-Object System.Windows.Forms.TextBox
$lockedoutTextBox.Location = New-Object System.Drawing.Point(135,391)
$lockedoutTextBox.Size = New-Object System.Drawing.Size(200,20)
$lockedoutTextBox.ReadOnly = $true;
$lockedoutTextBox.Text = ""
$lockedoutTextBox.Font = New-Object System.Drawing.Font("Arial", 10, $fontStyleN)
$usersGroup.Controls.Add($lockedoutTextBox)

#Lock Out Time
$lockedoutdateLabel = New-Object System.Windows.Forms.Label
$lockedoutdateLabel.Location = New-Object System.Drawing.Point(5,420)
$lockedoutdateLabel.Size = New-Object System.Drawing.Size(130,20)
$lockedoutdateLabel.Text = "Locked Out Date"
$lockedoutdateLabel.Font = New-Object System.Drawing.Font("Arial", 9, $fontStyle)
$usersGroup.Controls.Add($lockedoutdateLabel)

$lockedoutdateTextBox = New-Object System.Windows.Forms.TextBox
$lockedoutdateTextBox.Location = New-Object System.Drawing.Point(135,415)
$lockedoutdateTextBox.Size = New-Object System.Drawing.Size(200,20)
$lockedoutdateTextBox.ReadOnly = $true;
$lockedoutdateTextBox.Text = ""
$lockedoutdateTextBox.Font = New-Object System.Drawing.Font("Arial", 10, $fontStyleN)
$usersGroup.Controls.Add($lockedoutdateTextBox)

#Account Enabled
$accountenabledLabel = New-Object System.Windows.Forms.Label
$accountenabledLabel.Location = New-Object System.Drawing.Point(5,444)
$accountenabledLabel.Size = New-Object System.Drawing.Size(130,20)
$accountenabledLabel.Text = "Account Enabled"
$accountenabledLabel.Font = New-Object System.Drawing.Font("Arial", 9, $fontStyle)
$usersGroup.Controls.Add($accountenabledLabel)

$accountenabledTextBox = New-Object System.Windows.Forms.TextBox
$accountenabledTextBox.Location = New-Object System.Drawing.Point(135,439)
$accountenabledTextBox.Size = New-Object System.Drawing.Size(200,20)
$accountenabledTextBox.ReadOnly = $true;
$accountenabledTextBox.Text = ""
$accountenabledTextBox.Font = New-Object System.Drawing.Font("Arial", 10, $fontStyleN)
$usersGroup.Controls.Add($accountenabledTextBox)

#Butttons at the bottom: PW Reset, VPN Password Reset, AD Unlock, Get User Groups, Get Object Info
#$usersactionsGroup

#PW Reset Button
$pwresetButton = New-Object System.Windows.Forms.Button
$pwresetButton.Text = "Password Reset"
$pwresetButton.Location = New-Object System.Drawing.Point(5, 30)
$pwresetButton.AutoSize = $true
$pwresetButton.Add_Click({
    Add-OutputText "Resetting AD password (pre-expired)..."
    Add-Type -AssemblyName Microsoft.VisualBasic
    $Password = [Microsoft.VisualBasic.Interaction]::InputBox('Enter new password:', 'Reset Password', "Enter New Password")
    Set-ADAccountPassword -Identity $usernameTextBox.Text -Reset -NewPassword (ConvertTo-SecureString -AsPlainText "$Password" -Force)
    Set-ADUser -Identity $usernameTextBox.Text -ChangePasswordAtLogon $true
    Add-OutputText "Password reset successfully and set to change at next logon."
    #Refresh the user info
    $pwdexpTextBox.Text = ""
    $UserInfo = Get-ADUsername -param $usernameTextBox.Text
})
$usersactionsGroup.Controls.Add($pwresetButton)
$toolTip.SetToolTip($pwresetButton, "Reset AD Password, pre-expired.")

#VPN Password Reset Button
$vpnpwresetButton = New-Object System.Windows.Forms.Button
$vpnpwresetButton.Text = "VPN Password Reset"
$vpnpwresetButton.Location = New-Object System.Drawing.Point(105, 30)
$vpnpwresetButton.AutoSize = $true
$vpnpwresetButton.Add_Click({
    Add-OutputText "Resetting AD password for VPN (not expired)..."
    Add-Type -AssemblyName Microsoft.VisualBasic
    $Password = [Microsoft.VisualBasic.Interaction]::InputBox('Enter new password:', 'Reset Password', "Enter New Password")
    Set-ADAccountPassword -Identity $usernameTextBox.Text -Reset -NewPassword (ConvertTo-SecureString -AsPlainText "$Password" -Force)
    Set-ADuser -Identity $usernameTextBox.Text -ChangePasswordAtLogon $false
    #Refresh the user info
    $pwdexpTextBox.Text = ""
    $UserInfo = Get-ADUsername -param $usernameTextBox.Text
})
$usersactionsGroup.Controls.Add($vpnpwresetButton)
$toolTip.SetToolTip($vpnpwresetButton, "Reset AD Password for VPN not-expired.")

#AD Unlock
$adunlockButton = New-Object System.Windows.Forms.Button
$adunlockButton.Text = "AD Unlock"
$adunlockButton.Location = New-Object System.Drawing.Point(230, 30)
$adunlockButton.AutoSize = $true
$adunlockButton.Add_Click({
    Add-OutputText "Unlocking the AD Account..."
    Unlock-ADAccount -Identity $usernameTextBox.Text
    #Refresh the user info
    $lockedoutTextBox.Text = ""
    #2nd Unlock-ADAccount to refresh the status and wait 2 seconds - just to be annoying
    Start-sleep -Seconds 2
    Unlock-ADAccount -Identity $usernameTextBox.Text
    Get-ADUsername -param $usernameTextBox.Text
}) 
$usersactionsGroup.Controls.Add($adunlockButton)
$toolTip.SetToolTip($adunlockButton, "Unlocks the AD Account.")

#Get User Groups
$getusergroupsButton = New-Object System.Windows.Forms.Button
$getusergroupsButton.Text = "Get User Groups"
$getusergroupsButton.Location = New-Object System.Drawing.Point(5, 70)
$getusergroupsButton.AutoSize = $true
$getusergroupsButton.Add_Click({
    Add-OutputText "Getting User Groups for AD Account..."
    $GetADUserGroups = Get-ADPrincipalGroupMembership -Identity $usernameTextBox.Text | Select-Object -ExpandProperty Name | Sort-Object
    foreach ($group in $GetADUserGroups) {
        Add-OutputText "$group"
    }
})

$usersactionsGroup.Controls.Add($getusergroupsButton)
$toolTip.SetToolTip($getusergroupsButton, "Gets Users Groups for AD Account.")

#Get Object Info
$getobjectButton = New-Object System.Windows.Forms.Button
$getobjectButton.Text = "Get Object OU Info"
$getobjectButton.Location = New-Object System.Drawing.Point(107, 70)
$getobjectButton.AutoSize = $true
$getobjectButton.Add_Click({
    Add-OutputText "Getting Object OU for AD Account..."
    $UserOU = Get-ADUser -Identity $usernameTextBox.Text -Properties * | Select-Object CanonicalName 
    Add-OutputText "Object OU is $UserOU"
})
$usersactionsGroup.Controls.Add($getobjectButton)
$toolTip.SetToolTip($getobjectButton, "Gets Object OU for AD Account.")

#Add Group
$addgroupButton = New-Object System.Windows.Forms.Button
$addgroupButton.Text = "Add Group"
$addgroupButton.Location = New-Object System.Drawing.Point(5, 110)
$addgroupButton.AutoSize = $true
$addgroupButton.Add_Click({
    Add-OutputText "Adding AD Group to user..."
    try {
        $groupName = $addgroupTextBox.Text
        if ([string]::IsNullOrWhiteSpace($groupName)) {
            Add-OutputText "Group name cannot be empty."
            return
        }

        # Check if the group exists in Active Directory
        $groupExists = Get-ADGroup -Filter { Name -eq $groupName } -ErrorAction SilentlyContinue
        if ($null -eq $groupExists) {
            Add-OutputText "Group '$groupName' does not exist in Active Directory."
            return
        }

        # Add the user to the group
        Add-ADGroupMember -Identity $groupName -Members $usernameTextBox.Text
        Add-OutputText "User '$($usernameTextBox.Text)' successfully added to group '$groupName'."
    } catch {
        Add-OutputText "Failed to add user to group."
    }
})
$usersactionsGroup.Controls.Add($addgroupButton)
$toolTip.SetToolTip($addgroupButton, "Adds an AD Security Group to a user.")

#Add Group Text Box
$addgroupTextBox = New-Object System.Windows.Forms.TextBox
$addgroupTextBox.Location = New-Object System.Drawing.Point(85, 112)
$addgroupTextBox.Size = New-Object System.Drawing.Size(150,20)
$addgroupTextBox.ReadOnly = $false;
$addgroupTextBox.Text = ""
$addgroupTextBox.Font = New-Object System.Drawing.Font("Arial", 10, $fontStyleN)
$usersactionsGroup.Controls.Add($addgroupTextBox)

#Users Panel End

#Create the PC info group
$pcsGroup = New-Object System.Windows.Forms.GroupBox
$pcsGroup.Text = "PC Info"
$pcsGroup.Location = New-Object System.Drawing.Point(360, 30)
$pcsGroup.Size = New-Object System.Drawing.Size(610, 197)

#Create the PC actions group
$pcactionsGroup = New-Object System.Windows.Forms.GroupBox
$pcactionsGroup.Text = "PC Actions"
$pcactionsGroup.Location = New-Object System.Drawing.Point(360, 235)
$pcactionsGroup.Size = New-Object System.Drawing.Size(610, 320)
$form.Controls.Add($pcactionsGroup)

$form.Controls.Add($pcsGroup)

#Create Hostname Label and text box

$hostnameLabel = New-Object System.Windows.Forms.Label
$hostnameLabel.Text = "Hostname/PC Number"
$hostnameLabel.Location = New-Object System.Drawing.Point(10, 20)
$hostnameLabel.AutoSize = $true
$hostnameLabel.Font = New-Object System.Drawing.Font("Arial", 10, $fontStyle)
$pcsGroup.Controls.Add($hostnameLabel)

$hostnameTextBox = New-Object System.Windows.Forms.TextBox
$hostnameTextBox.Location = New-Object System.Drawing.Point(160, 20)
$hostnameTextBox.Size = New-Object System.Drawing.Size(120, 20)
$hostnameTextBox.Add_KeyDown({
    param($sender, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        Add-OutputText "Extracting PC Info..."
        # Add your PC Info script here

        # Test network connectivity
        $pingResult = Test-Connection -ComputerName $hostnameTextBox.Text -Count 1 -Quiet
        if ($pingResult) {  
            #Get PC Info
            $PCInfo = Get-PCInfo -param $hostnameTextBox.Text
            #Apply Label Text Changes  
                        
            $modelTextBoxLabel.Text = $PCInfo.Model
            $cDriveTextBoxLabel.Text = $PCInfo.CDriveSpace
            $laptopPcTextBoxLabel.Text = $PCInfo.LapDesk	
            $networkTextBoxLabel.Text = $PCInfo.Network
            $pingSpeedTextBox.Text = $PCInfo.PingSpd
            $printSpoolerStatusLabel.Text = $PCInfo.PSSvr
            $remoteRegStatusLabel.Text = $PCInfo.RRSvr
            $wusStatusLabel.Text = $PCInfo.wusSvr
            $mqStatusLabel.Text = $PCInfo.mqSvr
            $LastUserStatusLabel.Text = $PCInfo.LastUser          
            $boottimeStatusLabel.Text = $PCInfo.Boottime

            Add-OutputText "PC Info updated."
            
            $os = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $hostnameTextBox.Text
            $osInfo = $os.Caption + " " + $os.Version + " - " + $modelTextBoxLabel.Text
            Add-OutputText "OS Info: $osInfo"

        } else {
            
            Add-OutputText "Ping failed. Please check the hostname or network connection."
            
            $networkTextBoxLabel.Text = "None"
        }
        $e.Handled = $true
    }
})
$pcsGroup.Controls.Add($hostnameTextBox)

$LastUserLabel = New-Object System.Windows.Forms.Label
$LastUserLabel.Text = "Last User"
$LastUserLabel.Location = New-Object System.Drawing.Point(300, 20)
$LastUserLabel.AutoSize = $true
$LastUserLabel.Font = New-Object System.Drawing.Font("Arial", 10, $fontStyle)
$pcsGroup.Controls.Add($LastUserLabel)

$LastUserStatusLabel = New-Object System.Windows.Forms.Label
$LastUserStatusLabel.Text = "..."
$LastUserStatusLabel.Location = New-Object System.Drawing.Point(450, 20)
$LastUserStatusLabel.AutoSize = $true
$LastUserStatusLabel.Font = New-Object System.Drawing.Font("Arial", 9, $fontStyleN)
$pcsGroup.Controls.Add($LastUserStatusLabel)

#Create Labels and Text boxes for PC Info
$modelLabel = New-Object System.Windows.Forms.Label
$modelLabel.Text = "Model"
$modelLabel.Location = New-Object System.Drawing.Point(10, 50)
$modelLabel.AutoSize = $true
$modelLabel.Font = New-Object System.Drawing.Font("Arial", 10, $fontStyle)
$pcsGroup.Controls.Add($modelLabel)

$modelTextBoxLabel = New-Object System.Windows.Forms.Label
$modelTextBoxLabel.Text = "..."
$modelTextBoxLabel.Location = New-Object System.Drawing.Point(160, 50)
$modelTextBoxLabel.AutoSize = $true
$modelTextBoxLabel.Font = New-Object System.Drawing.Font("Arial", 9, $fontStyleN)
$pcsGroup.Controls.Add($modelTextBoxLabel)

$cDriveLabel = New-Object System.Windows.Forms.Label
$cDriveLabel.Text = "C Drive Free/Total"
$cDriveLabel.Location = New-Object System.Drawing.Point(10, 80)
$cDriveLabel.AutoSize = $true
$cDriveLabel.Font = New-Object System.Drawing.Font("Arial", 10, $fontStyle)
$pcsGroup.Controls.Add($cDriveLabel)

$cDriveTextBoxLabel = New-Object System.Windows.Forms.Label
$cDriveTextBoxLabel.Text = "..."
$cDriveTextBoxLabel.Location = New-Object System.Drawing.Point(160, 80)
$cDriveTextBoxLabel.AutoSize = $true
$cDriveTextBoxLabel.Font = New-Object System.Drawing.Font("Arial", 9, $fontStyleN)
$pcsGroup.Controls.Add($cDriveTextBoxLabel)

$laptopPcLabel = New-Object System.Windows.Forms.Label
$laptopPcLabel.Text = "Laptop/Desktop"
$laptopPcLabel.Location = New-Object System.Drawing.Point(10, 110)
$laptopPcLabel.AutoSize = $true
$laptopPcLabel.Font = New-Object System.Drawing.Font("Arial", 10, $fontStyle)
$pcsGroup.Controls.Add($laptopPcLabel)

$laptopPcTextBoxLabel = New-Object System.Windows.Forms.Label
$laptopPcTextBoxLabel.Text = "..."
$laptopPcTextBoxLabel.Location = New-Object System.Drawing.Point(160, 110)
$laptopPcTextBoxLabel.AutoSize = $true
$laptopPcTextBoxLabel.Font = New-Object System.Drawing.Font("Arial", 9, $fontStyleN)
$pcsGroup.Controls.Add($laptopPcTextBoxLabel)

$networkLabel = New-Object System.Windows.Forms.Label
$networkLabel.Text = "Network"
$networkLabel.Location = New-Object System.Drawing.Point(10, 140)
$networkLabel.AutoSize = $true
$networkLabel.Font = New-Object System.Drawing.Font("Arial", 10, $fontStyle)
$pcsGroup.Controls.Add($networkLabel)

$networkTextBoxLabel = New-Object System.Windows.Forms.Label
$networkTextBoxLabel.Text = "..."
$networkTextBoxLabel.Location = New-Object System.Drawing.Point(160, 140)
$networkTextBoxLabel.AutoSize = $true
$networkTextBoxLabel.Font = New-Object System.Drawing.Font("Arial", 10, $fontStyleN)
$pcsGroup.Controls.Add($networkTextBoxLabel)

$pingSpeedLabel = New-Object System.Windows.Forms.Label
$pingSpeedLabel.Text = "Ping Speed (m/s)"
$pingSpeedLabel.Location = New-Object System.Drawing.Point(10, 170)
$pingSpeedLabel.AutoSize = $true
$pingSpeedLabel.Font = New-Object System.Drawing.Font("Arial", 10, $fontStyle)
$pcsGroup.Controls.Add($pingSpeedLabel)

$pingSpeedTextBox = New-Object System.Windows.Forms.Label
$pingSpeedTextBox.Text = "..."
$pingSpeedTextBox.Location = New-Object System.Drawing.Point(160, 170)
$pingSpeedTextBox.AutoSize = $true
$pingSpeedTextBox.Font = New-Object System.Drawing.Font("Arial", 9)
$pcsGroup.Controls.Add($pingSpeedTextBox)

$printSpoolerLabel = New-Object System.Windows.Forms.Label
$printSpoolerLabel.Text = "Print Spooler"
$printSpoolerLabel.Location = New-Object System.Drawing.Point(300, 50)
$printSpoolerLabel.AutoSize = $true
$printSpoolerLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$pcsGroup.Controls.Add($printSpoolerLabel)

$printSpoolerStatusLabel = New-Object System.Windows.Forms.Label
$printSpoolerStatusLabel.Text = "..."
$printSpoolerStatusLabel.Location = New-Object System.Drawing.Point(450, 50)
$printSpoolerStatusLabel.AutoSize = $true
$printSpoolerStatusLabel.Font = New-Object System.Drawing.Font("Arial", 10, $fontStyleN)
$pcsGroup.Controls.Add($printSpoolerStatusLabel)

$remoteRegLabel = New-Object System.Windows.Forms.Label
$remoteRegLabel.Text = "Remote Registry"
$remoteRegLabel.Location = New-Object System.Drawing.Point(300, 80)
$remoteRegLabel.AutoSize = $true
$remoteRegLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$pcsGroup.Controls.Add($remoteRegLabel)

$remoteRegStatusLabel = New-Object System.Windows.Forms.Label
$remoteRegStatusLabel.Text = "..."
$remoteRegStatusLabel.Location = New-Object System.Drawing.Point(450, 80)
$remoteRegStatusLabel.AutoSize = $true
$remoteRegStatusLabel.Font = New-Object System.Drawing.Font("Arial", 10, $fontStyleN)
$pcsGroup.Controls.Add($remoteRegStatusLabel)

$wusLabel = New-Object System.Windows.Forms.Label
$wusLabel.Text = "Windows Update"
$wusLabel.Location = New-Object System.Drawing.Point(300, 110)
$wusLabel.AutoSize = $true
$wusLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$pcsGroup.Controls.Add($wusLabel)

$wusStatusLabel = New-Object System.Windows.Forms.Label
$wusStatusLabel.Text = "..."
$wusStatusLabel.Location = New-Object System.Drawing.Point(450, 110)
$wusStatusLabel.AutoSize = $true
$wusStatusLabel.Font = New-Object System.Drawing.Font("Arial", 10, $fontStyleN)
$pcsGroup.Controls.Add($wusStatusLabel)

$mqLabel = New-Object System.Windows.Forms.Label
$mqLabel.Text = "Message Queuing"
$mqLabel.Location = New-Object System.Drawing.Point(300, 140)
$mqLabel.AutoSize = $true
$mqLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$pcsGroup.Controls.Add($mqLabel)

$mqStatusLabel = New-Object System.Windows.Forms.Label
$mqStatusLabel.Text = "..."
$mqStatusLabel.Location = New-Object System.Drawing.Point(450, 140)
$mqStatusLabel.AutoSize = $true
$mqStatusLabel.Font = New-Object System.Drawing.Font("Arial", 10, $fontStyleN)
$pcsGroup.Controls.Add($mqStatusLabel)

$boottimeLabel = New-Object System.Windows.Forms.Label
$boottimeLabel.Text = "PC Boot Time"
$boottimeLabel.Location = New-Object System.Drawing.Point(300, 170)
$boottimeLabel.AutoSize = $true
$boottimeLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$pcsGroup.Controls.Add($boottimeLabel)

$boottimeStatusLabel = New-Object System.Windows.Forms.Label
$boottimeStatusLabel.Text = "..."
$boottimeStatusLabel.Location = New-Object System.Drawing.Point(450, 170)
$boottimeStatusLabel.AutoSize = $true
$boottimeStatusLabel.Font = New-Object System.Drawing.Font("Arial", 10, $fontStyleN)
$pcsGroup.Controls.Add($boottimeStatusLabel)

#Create scripts buttons and labels
#C Drive Space Scripts
$clearcdriveLabel = New-Object System.Windows.Forms.Label
$clearcdriveLabel.Text = "Clear C Drive Space"
$clearcdriveLabel.Location = New-Object System.Drawing.Point(10, 20)
$clearcdriveLabel.AutoSize = $true
$clearcdriveLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$pcactionsGroup.Controls.Add($clearcdriveLabel)

$turnOffHibernateButton = New-Object System.Windows.Forms.Button
$turnOffHibernateButton.Text = "Turn Off Hibernate"
$turnOffHibernateButton.Location = New-Object System.Drawing.Point(10, 50)
$turnOffHibernateButton.AutoSize = $true
$turnOffHibernateButton.Add_Click({
    Add-OutputText "Checking if PC is a Laptop or Desktop..."
    # Add your Turn Off Hibernate script here
    $pctype = Get-PCType -param1 $hostnameTextBox.Text
    if ($pctype -eq "Desktop") {
        Add-OutputText "Turning Off Hibernate on Desktop..."
        Disable-Hibernate -param1 $hostnameTextBox.Text
        
        #Refresh Disk Space Data
        $diskInfo = Get-CDriveSpace -param1 $hostnameTextBox.Text
        $freeSpace = $diskInfo.FreeSpace
        $totalSpace = $diskInfo.TotalSpace
        # Prepare the output text
        $diskInfoText = "$freeSpace GB / $totalSpace GB"
        $cDriveTextBoxLabel.Text = $diskInfoText  

    } else {
        
        Add-OutputText "This is a Laptop PC. Hibernate should stay enabled. "
        
    }
       
})
$pcactionsGroup.Controls.Add($turnOffHibernateButton)
$toolTip.SetToolTip($turnOffHibernateButton, "Turns off Hibernate on a PC.")

$swdWipeButton = New-Object System.Windows.Forms.Button
$swdWipeButton.Text = "Clear SW Distribution"
$swdWipeButton.Location = New-Object System.Drawing.Point(120, 50)
$swdWipeButton.AutoSize = $true
$swdWipeButton.Add_Click({
    Add-OutputText "Deleting Software Distribution Folder."
    # Add your Clear SW Distribution script here
    Remove-SWDFolder -param1 $hostnameTextBox.Text
    #Refresh Disk Space Data
    $diskInfo = Get-CDriveSpace -param1 $hostnameTextBox.Text
    $freeSpace = $diskInfo.FreeSpace
    $totalSpace = $diskInfo.TotalSpace
    # Prepare the output text
    $diskInfoText = "$freeSpace GB / $totalSpace GB"
    $cDriveTextBoxLabel.Text = $diskInfoText
})
$pcactionsGroup.Controls.Add($swdWipeButton)
$toolTip.SetToolTip($swdWipeButton, "This deletes the Software Distribution Folder.")

$wv2FolderWipeButton = New-Object System.Windows.Forms.Button
$wv2FolderWipeButton.Text = "WV2"
$wv2FolderWipeButton.Location = New-Object System.Drawing.Point(244, 50)
$wv2FolderWipeButton.AutoSize = $true

$wv2FolderWipeButton.Add_Click({
    Add-OutputText "Executing WV2 Folder Wipe script..."
    # Add your 30 Day Teams Folder Wipe script here
    Remove-WV2 -param1 $hostnameTextBox.Text
    #Refresh Disk Space Data
    $diskInfo = Get-CDriveSpace -param1 $hostnameTextBox.Text
    $freeSpace = $diskInfo.FreeSpace
    $totalSpace = $diskInfo.TotalSpace
    # Prepare the output text
    $diskInfoText = "$freeSpace GB / $totalSpace GB"
    $cDriveTextBoxLabel.Text = $diskInfoText
})
$pcactionsGroup.Controls.Add($wv2FolderWipeButton)
$toolTip.SetToolTip($wv2FolderWipeButton, "Deletes WV2 Folders.  This can take a while.")

$teamsFolderWipeButton = New-Object System.Windows.Forms.Button
$teamsFolderWipeButton.Text = "Teams Folder Wipe"
$teamsFolderWipeButton.Location = New-Object System.Drawing.Point(321, 50)
$teamsFolderWipeButton.AutoSize = $true

$teamsFolderWipeButton.Add_Click({
    Add-OutputText "Executing 30 Day Teams Folder Wipe script..."
    # Add your 30 Day Teams Folder Wipe script here
    Remove-TeamFolders -param1 $hostnameTextBox.Text
    #Refresh Disk Space Data
    $diskInfo = Get-CDriveSpace -param1 $hostnameTextBox.Text
    $freeSpace = $diskInfo.FreeSpace
    $totalSpace = $diskInfo.TotalSpace
    # Prepare the output text
    $diskInfoText = "$freeSpace GB / $totalSpace GB"
    $cDriveTextBoxLabel.Text = $diskInfoText
})
$pcactionsGroup.Controls.Add($teamsFolderWipeButton)
$toolTip.SetToolTip($teamsFolderWipeButton, "30 Day Teams Folder Wipe.  This can take a while.")

$profileWipeButton = New-Object System.Windows.Forms.Button
$profileWipeButton.Text = "User Profile Wipe"
$profileWipeButton.Location = New-Object System.Drawing.Point(436, 50)
$profileWipeButton.AutoSize = $true
$profileWipeButton.Add_Click({
    Add-OutputText "Executing 30 Day Profile Wipe script..."
    # Add your 30 Day Profile Wipe script here
    Remove-OldUserProfiles -param1 $hostnameTextBox.Text -days 30
    #Refresh Disk Space Data
    $diskInfo = Get-CDriveSpace -param1 $hostnameTextBox.Text
    $freeSpace = $diskInfo.FreeSpace
    $totalSpace = $diskInfo.TotalSpace
    # Prepare the output text
    $diskInfoText = "$freeSpace GB / $totalSpace GB"
    $cDriveTextBoxLabel.Text = $diskInfoText  
})
$pcactionsGroup.Controls.Add($profileWipeButton)
$toolTip.SetToolTip($profileWipeButton, "30 Day User Profile Wipe.  This can take a while.")

$printersLabel = New-Object System.Windows.Forms.Label
$printersLabel.Text = "Printers"
$printersLabel.Location = New-Object System.Drawing.Point(10, 80)
$printersLabel.AutoSize = $true
$printersLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$pcactionsGroup.Controls.Add($printersLabel)

$pcGroupLabel = New-Object System.Windows.Forms.Label
$pcGroupLabel.Text = "PC Groups"
$pcGroupLabel.Location = New-Object System.Drawing.Point(247, 80)
$pcGroupLabel.AutoSize = $true
$pcGroupLabel.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$pcactionsGroup.Controls.Add($pcGroupLabel)

$printSpoolRestartButton = New-Object System.Windows.Forms.Button
$printSpoolRestartButton.Text = "Print Spooler Restart"
$printSpoolRestartButton.Location = New-Object System.Drawing.Point(10, 105)
$printSpoolRestartButton.AutoSize = $true
$printSpoolRestartButton.Add_Click({
    Add-OutputText "Executing Print Spool Restart script..."
    # Add your Print Spool Restart script here
    try {
        Invoke-Command -ComputerName $hostnameTextBox.Text -ScriptBlock {
        Stop-Service -Name "Spooler" -Force
        Start-Sleep -Seconds 3
        Start-Service -Name "Spooler"
        }
        Add-OutputText "Print Spooler restarted..."
        }
        catch {
            $errorDetails = $_
            Add-OutputText "Print Spooler failed to restart. Error: $errorDetails"
        }          
})
$pcactionsGroup.Controls.Add($printSpoolRestartButton)

$deleteSpoolerFilesButton = New-Object System.Windows.Forms.Button
$deleteSpoolerFilesButton.Text = "Delete Spooler Files"
$deleteSpoolerFilesButton.Location = New-Object System.Drawing.Point(131, 105)
$deleteSpoolerFilesButton.AutoSize = $true
$deleteSpoolerFilesButton.Add_Click({
    Add-OutputText "Executing Delete Spooler Files script..."
    # Add your Delete Spooler Files script here
    try {
        $delSpooler = Invoke-Command -ComputerName $hostnameTextBox.Text -ScriptBlock {
            $SpoolerPath = "C:\Windows\System32\spool\PRINTERS"

            #Stop Spooler Service
            Stop-Service -Name "Spooler" -Force
            Start-Sleep -Seconds 3

            if (Test-Path -Path $SpoolerPath) {
                Get-ChildItem -Path $SpoolerPath -Force | Remove-Item -Force -Recurse
                Start-Service -Name "Spooler"
                Start-Sleep -Seconds 5
                return "Spooler files deleted successfully on $env:COMPUTERNAME."
            } else {
                
                Add-OutputText "Spooler path does not exist on $env:COMPUTERNAME."
                
            }
        }
    } catch {
        Add-OutputText "Failed to delete spooler files on $hostnameTextBox.Text Error: $_"
    }
    Add-OutputText "$delSpooler"
    $dpsservice = Get-Service -ComputerName $hostnameTextBox.Text -Name "Spooler" -ErrorAction Stop
})
$pcactionsGroup.Controls.Add($deleteSpoolerFilesButton)

$addpcGroupButton = New-Object System.Windows.Forms.Button
$addpcGroupButton.Text = "Add Group"
$addpcGroupButton.Location = New-Object System.Drawing.Point(249, 105)
$addpcGroupButton.AutoSize = $true
$addpcGroupButton.Add_Click({
    if ([string]::IsNullOrWhiteSpace($hostnameTextBox.Text)) {
        Add-OutputText "Hostname/PC Number field cannot be empty. Please enter a PC No."
        return
    }
    if ([string]::IsNullOrWhiteSpace($addPCGroupTextBox.Text)) {
        Add-OutputText "Group field cannot be empty. Please security group Name."
        return
    }
    Add-OutputText "Adding PC to correct security group..."
    # Add the PC to the correct security group
    Add-PCGroup -param1 $hostnameTextBox.Text -param2 $addPCGroupTextBox.Text
})
$pcactionsGroup.Controls.Add($addpcGroupButton)
$toolTip.SetToolTip($addpcGroupButton, "Add PC security groups.")

$addPCGroupTextBox = New-Object System.Windows.Forms.TextBox
$addPCGroupTextBox.Location = New-Object System.Drawing.Point(330, 105)
$addPCGroupTextBox.Size = New-Object System.Drawing.Size(180, 20)
$addPCGroupTextBox.Font = New-Object System.Drawing.Font("Arial", 10, $fontStyleN)
$addPCGroupTextBox.Text = ""
$pcactionsGroup.Controls.Add($addPCGroupTextBox)

$pcGroupsButton = New-Object System.Windows.Forms.Button
$pcGroupsButton.Text = "Get PC Groups"
$pcGroupsButton.Location = New-Object System.Drawing.Point(515, 105)
$pcGroupsButton.AutoSize = $true
$pcGroupsButton.Add_Click({
    # Add your PC groups script here
    Add-OutputText "Fetching Active Directory group memberships for the PC..."
    try {
        # Get the hostname from the textbox
        $hostname = $hostnameTextBox.Text

        # Ensure the hostname is not empty
        if ([string]::IsNullOrWhiteSpace($hostname)) {
            
            Add-OutputText "PC No. box cannot be empty."
            
            return
        }

        # Fetch the group memberships
        $groups = Get-ADComputer -Identity $hostname -Properties MemberOf | Select-Object -ExpandProperty MemberOf        
        $groupNames = $groups | ForEach-Object { Get-ADGroup -Identity $_ | Select-Object -ExpandProperty Name }

        # Check if any groups were found
        if ($groupNames) {
            Add-OutputText "Groups for $hostname "
            foreach ($group in $groupNames) {
                Add-OutputText "$group"  # Append each group on a new line
            }
        } else {
            
            Add-OutputText "No groups found for $hostname."
           
        } 
    } catch {
        Add-OutputText "Error fetching group memberships: $_"
    } finally {
        Add-OutputText "Group membership fetch operation completed."
    }
})
$pcactionsGroup.Controls.Add($pcGroupsButton)
$toolTip.SetToolTip($pcGroupsButton, "Use Hostname/PC Number box to get the groups for the PC.")

$miscLabel = New-Object System.Windows.Forms.Label
$miscLabel.Text = "Misc"
$miscLabel.Location = New-Object System.Drawing.Point(10, 135)
$miscLabel.AutoSize = $true
$miscLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$pcactionsGroup.Controls.Add($miscLabel)

$pingTestButton = New-Object System.Windows.Forms.Button
$pingTestButton.Text = "Ping Test"
$pingTestButton.Location = New-Object System.Drawing.Point(10, 160)
$pingTestButton.AutoSize = $true
$pingTestButton.Add_Click({
    Add-OutputText "Executing Ping test..."
    $ipAddress = (Resolve-DnsName -Name $hostnameTextBox.Text | Where-Object { $_.QueryType -eq 'A' }).IPAddress
    Add-OutputText "IP Address: $ipAddress"
    try {
        $pingResults = Test-Connection -ComputerName $hostnameTextBox.Text -Count 4 -ErrorAction Stop
        $totalResponseTime = 0
        $pingCount = 0
        foreach ($ping in $pingResults) {
            Add-OutputText "Reply from $($ping.Address): time=$($ping.ResponseTime)ms"
            $totalResponseTime += $ping.ResponseTime
            $pingCount++
        }
        if ($pingCount -gt 0) {
            $averagePing = $totalResponseTime / $pingCount
            Add-OutputText "Average Ping Speed: $averagePing ms"
        }
    } catch {
        Add-OutputText "Ping test failed for $hostnameTextBox.Text"
    }
})
$pcactionsGroup.Controls.Add($pingTestButton)
$toolTip.SetToolTip($pingTestButton, "Runs a ping test.")

$osButton = New-Object System.Windows.Forms.Button
$osButton.Text = "OS"
$osButton.Location = New-Object System.Drawing.Point(87, 160)
$osButton.AutoSize = $false
$osButton.Size = New-Object System.Drawing.Size(40, 24)
$osButton.Add_Click({
    Add-OutputText "Showing Operating System..."
    $os = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $hostnameTextBox.Text
    $osInfo = $os.Caption + " " + $os.Version
    Add-OutputText "OS Info: $osInfo"
    
})
$pcactionsGroup.Controls.Add($osButton)
$toolTip.SetToolTip($osButton, "Shows the Operating System.")

#Uses Explorer++ to open the C$ share to get round the Windows Explorer issues
$cdollarButton = New-Object System.Windows.Forms.Button
$cdollarButton.Text = "C$"
$cdollarButton.Location = New-Object System.Drawing.Point(129, 160)
$cdollarButton.AutoSize = $false
#$cdollarButton.Enabled = $false
$cdollarButton.Size = New-Object System.Drawing.Size(40, 24)
$cdollarButton.Add_Click({
   Add-OutputText "Opening C$ Share via File Explorer..."
    # Define the remote computer and credentials
    $remoteComputer = $hostnameTextBox.Text
    Start-Process -FilePath "$pathExplorerPlus\Explorer++.exe" -ArgumentList "\\$remoteComputer\c$"

})
$pcactionsGroup.Controls.Add($cdollarButton)
$toolTip.SetToolTip($cdollarButton, "Opens the C$ Share using Explorer++.")

$dblbackslashButton = New-Object System.Windows.Forms.Button
$dblbackslashButton.Text = "Shares"
$dblbackslashButton.Location = New-Object System.Drawing.Point(171, 160)
$dblbackslashButton.AutoSize = $True
$dblbackslashButton.Size = New-Object System.Drawing.Size(40, 24)
$dblbackslashButton.Add_Click({
   Add-OutputText "Opening PC Shares via File Explorer..."
    # Define the remote computer and credentials
    $remoteComputer = $hostnameTextBox.Text
    Start-Process -FilePath "$pathExplorerPlus\Explorer++.exe" -ArgumentList "\\$remoteComputer"

})
$pcactionsGroup.Controls.Add($dblbackslashButton)
$toolTip.SetToolTip($dblbackslashButton, "Opens the PC Shares using Explorer++.")

$filesPDesktopButton = New-Object System.Windows.Forms.Button
$filesPDesktopButton.Text = "Copy Files to Desktop"
$filesPDesktopButton.Location = New-Object System.Drawing.Point(223, 160)
$filesPDesktopButton.AutoSize = $true
$filesPDesktopButton.Add_Click({
    Add-OutputText "Executing Copy File to Public Desktop script..."
    # Add your Copy File to Public Desktop script here
    $filesPDesktopStatusText = Add-FilesDesktop -param1 $hostnameTextBox.Text
    Add-OutputText "$filesPDesktopStatusText"
})
$pcactionsGroup.Controls.Add($filesPDesktopButton)
$toolTip.SetToolTip($filesPDesktopButton, "Copy Files to Public Desktop.")

$remoteRegistryButton = New-Object System.Windows.Forms.Button
$remoteRegistryButton.Text = "Remote Registry Enable"
$remoteRegistryButton.Location = New-Object System.Drawing.Point(350, 160)
$remoteRegistryButton.AutoSize = $true
$remoteRegistryButton.Add_Click({
    if ($remoteRegistryButton.Text -eq "Remote Registry Enable") {
        $remoteRegistryButton.Text = "Remote Registry Disable"
        Add-OutputText "Executing Remote Registry Enable script..."
        # Add your Remote Registry Enable script here
        try {
            Set-RemoteRegistrySvc -ComputerName $hostnameTextBox.Text -svcname "RemoteRegistry" -Action "start"
        
        #Update Remote Registry Status Label
        $servicesStatus = Get-ServicesStatus -param1 $hostnameTextBox.Text
        $remoteRegStatusLabel.Text = $servicesStatus[1]      
        }
        catch {
            Add-OutputText "Remote Registry failed to start..."
        }      
            } else {
        $remoteRegistryButton.Text = "Remote Registry Enable"
        Add-OutputText "Executing Remote Registry Disable script..."
        # Add your Remote Registry Disable script here      
        try {
            Set-RemoteRegistrySvc -ComputerName $hostnameTextBox.Text -svcname "RemoteRegistry" -Action "stop"
    
            #Update Remote Registry Status Label
            $servicesStatus = Get-ServicesStatus -param1 $hostnameTextBox.Text
            $remoteRegStatusLabel.Text = $servicesStatus[1]      
            }
            catch {
                Add-OutputText "Remote Registry failed to stop..."
            }      
    }
})
$pcactionsGroup.Controls.Add($remoteRegistryButton)
$toolTip.SetToolTip($remoteRegistryButton, "Always Remember to DISABLE this Service after use.")

$wusStartStopButton = New-Object System.Windows.Forms.Button
$wusStartStopButton.Text = "Stop Windows Update Svc"
$wusStartStopButton.Location = New-Object System.Drawing.Point(10, 190)
$wusStartStopButton.AutoSize = $true
$wusStartStopButton.Add_Click({
    if ($wusStartStopButton.Text -eq "Stop Windows Update Svc") {
        $wusStartStopButton.Text = "Start Windows Update Svc"
        Add-OutputText "Executing Windows Update Service script..."
        # Add your Windows Update Stop script here
        try {
             Set-StopStartService -ComputerName $hostnameTextBox.Text -svcname "wuauserv" -Action "stop"
        
        #Update Windows Update Status Label
        $servicesStatus = Get-ServicesStatus -param1 $hostnameTextBox.Text
        $wusStatusLabel.Text = $servicesStatus[2]      
        }
        catch {
            Add-OutputText "Windows Update Service failed to start..."
        }      
            } else {
        $wusStartStopButton.Text = "Start Windows Update Svc"
        Add-OutputText "Executing Windows Update Service script..."
        # Add your Windows Update Service Start script here      
        try {
            Set-StopStartService -ComputerName $hostnameTextBox.Text -svcname "wuauserv" -Action "start"
    
            #Update Remote Registry Status Label
            $servicesStatus = Get-ServicesStatus -param1 $hostnameTextBox.Text
            $remoteRegStatusLabel.Text = $servicesStatus[2] 
            $wusStartStopButton.Text = "Stop Windows Update Svc"     
            }
            catch {
                Add-OutputText "Windows Update Svc failed to stop..."
            }      
    }
})
$pcactionsGroup.Controls.Add($wusStartStopButton)
$toolTip.SetToolTip($wusStartStopButton, "Stop/Start Windows Update Service.")

$mqStartButton = New-Object System.Windows.Forms.Button
$mqStartButton.Text = "Start Msg Queuing Svc"
$mqStartButton.Location = New-Object System.Drawing.Point(160, 190)
$mqStartButton.AutoSize = $true
$mqStartButton.Add_Click({
        Add-OutputText "Executing Message Queuing Service script..."
        # Add your Message Queuing Stop script here
        try {
             Set-StopStartService -ComputerName $hostnameTextBox.Text -svcname "MSMQ" -Action "start"
             $service = Get-Service -ComputerName $remoteComputer -Name "MSMQTriggers" -ErrorAction Stop
        if ($service.Status -eq 'Stopped') {
             Set-StopStartService -ComputerName $hostnameTextBox.Text -svcname "MSMQTriggers" -Action "start"
        }
        #Update Message Queuing Status Label
        $servicesStatus = Get-ServicesStatus -param1 $hostnameTextBox.Text
        $mqStatusLabel.Text = $servicesStatus[3]      
        }
        catch {
            Add-OutputText "Message Queuing Service failed to start..."
        }                    
})
$pcactionsGroup.Controls.Add($mqStartButton)
$toolTip.SetToolTip($mqStartButton, "Start Message Queuing.")

$pcrestartrsnLabel = New-Object System.Windows.Forms.Label
$pcrestartrsnLabel.Text = "Enter Reason for Restart"
$pcrestartrsnLabel.Location = New-Object System.Drawing.Point(100, 230)
$pcrestartrsnLabel.AutoSize = $true
$pcrestartrsnLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)

$pcactionsGroup.Controls.Add($pcrestartrsnLabel)

$pcrestartdlyLabel = New-Object System.Windows.Forms.Label
$pcrestartdlyLabel.Text = "Delay in Seconds"
$pcrestartdlyLabel.Location = New-Object System.Drawing.Point(425, 230)
$pcrestartdlyLabel.AutoSize = $true
$pcrestartdlyLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$pcactionsGroup.Controls.Add($pcrestartdlyLabel)

$pcRestartButton = New-Object System.Windows.Forms.Button
$pcRestartButton.Text = "PC Restart"
$pcRestartButton.Location = New-Object System.Drawing.Point(10, 260)
$pcRestartButton.AutoSize = $true
$pcRestartButton.Add_Click({
    # Add your PC Restart script here
    try {
        if ([string]::IsNullOrWhiteSpace($hostnameTextBox.Text)) {
            Add-OutputText "PC No. field cannot be empty. Please enter a PC No."
            return
        }
        if ([string]::IsNullOrWhiteSpace($pcrestartrsnTextBox.Text) -or [string]::IsNullOrWhiteSpace($pcrestartdlyTextBox.Text)) {
            Add-OutputText "Please enter a reason and delay for the restart."
            return
        } else {
            $dly = $pcrestartdlyTextBox.Text
            $rsn = $pcrestartrsnTextBox.Text
            Add-OutputText "Executing PC Restart script..."
            Add-OutputText "Restarting PC in $dly seconds with reason: $rsn"
        }
        # Use Invoke-Command to restart the remote computer
        Invoke-Command -ComputerName $hostnameTextBox.Text -ScriptBlock {
            param($delay, $reason)
            C:\Windows\system32\shutdown.exe /r /t $delay /c "$reason"
        } -ArgumentList $dly, $rsn
    } catch {
        Add-OutputText "Failed to restart PC: $_"
    }
})
$pcactionsGroup.Controls.Add($pcRestartButton)

$pcrestartrsnTextBox = New-Object System.Windows.Forms.TextBox
$pcrestartrsnTextBox.Location = New-Object System.Drawing.Point(100, 262)
$pcrestartrsnTextBox.Size = New-Object System.Drawing.Size(300, 20)
$pcrestartrsnTextBox.Font = New-Object System.Drawing.Font("Arial", 10, $fontStyleN)
$pcrestartrsnTextBox.Text = ""
$pcactionsGroup.Controls.Add($pcrestartrsnTextBox)

$pcrestartdlyTextBox = New-Object System.Windows.Forms.TextBox
$pcrestartdlyTextBox.Location = New-Object System.Drawing.Point(425, 262)
$pcrestartdlyTextBox.Font = New-Object System.Drawing.Font("Arial", 10, $fontStyleN)
$pcrestartdlyTextBox.Size = New-Object System.Drawing.Size(50, 20)

$pcrestartdlyTextBox.Text = 30
$pcactionsGroup.Controls.Add($pcrestartdlyTextBox)

#*********************************************
# Create the Admin Tools group
$adminToolsGroup = New-Object System.Windows.Forms.GroupBox
$adminToolsGroup.Text = "Admin Tools"
$adminToolsGroup.Location = New-Object System.Drawing.Point(360, 560)
$adminToolsGroup.Size = New-Object System.Drawing.Size(610, 100)
$form.Controls.Add($adminToolsGroup)

#Active Directory Button
$adButton = New-Object System.Windows.Forms.Button
$adButton.Text = "Active Directory"
$adButton.Location = New-Object System.Drawing.Point(10, 20)
$adButton.AutoSize = $true
$adButton.Add_Click({
    Add-OutputText "Opening Active Directory..."
    # Add Active Directory script here
    try {
        C:\Windows\system32\dsa.msc        
    }
    catch {
        Add-OutputText "Error Opening Active Directory. Please try again."
    }
})
$adminToolsGroup.Controls.Add($adButton)

#DHCP Button
$dhcpButton = New-Object System.Windows.Forms.Button
$dhcpButton.Text = "DHCP"
$dhcpButton.Location = New-Object System.Drawing.Point(105, 20)
$dhcpButton.AutoSize = $true
$dhcpButton.Add_Click({
    Add-OutputText "Opening DHCP..."
    # Add DHCP script here
    try {
        C:\Windows\system32\dhcpmgmt.msc        
    }
    catch {
        Add-OutputText "Error Opening DHCP. Please try again."
    }
})
$adminToolsGroup.Controls.Add($dhcpButton)

#Group Policy Button
$gpButton = New-Object System.Windows.Forms.Button
$gpButton.Text = "Group Policy"
$gpButton.Location = New-Object System.Drawing.Point(182, 20)
$gpButton.AutoSize = $true
$gpButton.Add_Click({
    Add-OutputText "Opening Group Policy..."
    # Add Group Policy script here
    try {
        C:\Windows\system32\gpmc.msc        
    }
    catch {
        Add-OutputText "Error Opening Group Policy. Please try again."
    }
})
$adminToolsGroup.Controls.Add($gpButton)

#Print Mgmt Button
$pmButton = New-Object System.Windows.Forms.Button
$pmButton.Text = "Print Mgmt"
$pmButton.Location = New-Object System.Drawing.Point(263, 20)
$pmButton.AutoSize = $true
$pmButton.Add_Click({
    Add-OutputText "Opening Printer Management..."
    # Add Print Management script here
    try {
        C:\Windows\system32\printmanagement.msc        
    }
    catch {
        Add-OutputText "Error Opening Printer Management. Please try again."
    }
})
$adminToolsGroup.Controls.Add($pmButton)

#7-Zip Manager Button
$zpmButton = New-Object System.Windows.Forms.Button
$zpmButton.Text = "7-Zip Manager"
$zpmButton.Location = New-Object System.Drawing.Point(340, 20)
$zpmButton.AutoSize = $true
$zpmButton.Add_Click({
    Add-OutputText "Opening 7-Zip Manager..."
    # Add 7-Zip script here
    try {
        & "C:\Program Files\7-Zip\7zFM.exe"
    }
    catch {
        Add-OutputText "Error Opening 7-Zip Manager. Please try again."
    }
})
$adminToolsGroup.Controls.Add($zpmButton)

#AD Powershell Button
$adpsButton = New-Object System.Windows.Forms.Button
$adpsButton.Text = "AD Powershell"
$adpsButton.Location = New-Object System.Drawing.Point(430, 20)
$adpsButton.AutoSize = $true
$adpsButton.Add_Click({
    Add-OutputText "AD Powershell..."
    # Add AD Powershell script here
    try {
        & Start-Process -FilePath "$env:windir\system32\WindowsPowerShell\v1.0\powershell.exe" -ArgumentList '-noexit -command "import-module ActiveDirectory"' -WorkingDirectory "$env:USERPROFILE"             
    }
    catch {
        Add-OutputText "Error Opening AD Powershell. Please try again."
    }
})
$adminToolsGroup.Controls.Add($adpsButton)

#CMD Button
$cmdButton = New-Object System.Windows.Forms.Button
$cmdButton.Text = "CMD"
$cmdButton.Location = New-Object System.Drawing.Point(520, 20)
$cmdButton.AutoSize = $true
$cmdButton.Add_Click({
    Add-OutputText "Opening Command Prompt..."
    # Add Command Prompt script here
    try {
        & Start-Process -FilePath "$env:windir\system32\cmd.exe" -ArgumentList '/k' -WorkingDirectory "$env:USERPROFILE"             
    }
    catch {
        Add-OutputText "Error Opening Command Prompt. Please try again."
    }
})
$adminToolsGroup.Controls.Add($cmdButton)

#Regedit Button
$regButton = New-Object System.Windows.Forms.Button
$regButton.Text = "Regedit"
$regButton.Location = New-Object System.Drawing.Point(10, 60)
$regButton.AutoSize = $true
$regButton.Add_Click({
    Add-OutputText "Opening Registry Editor..."
    # Add Registry Editor script here
    try {
        & Start-Process -FilePath "$env:windir\regedit.exe" -ArgumentList '/m' -WorkingDirectory "$env:USERPROFILE"             
    }
    catch {
        Add-OutputText "Error Opening Registry Editor. Please try again."
    }
})
$adminToolsGroup.Controls.Add($regButton)

#Shared Drive App Button
$sdaButton = New-Object System.Windows.Forms.Button
$sdaButton.Text = "Shared Drive App"
$sdaButton.Location = New-Object System.Drawing.Point(87, 60)
$sdaButton.AutoSize = $true
$sdaButton.Add_Click({
    Add-OutputText "Opening Shared Drive App..."
    # Add Shared Drive App script here
    try {
        & Start-Process -FilePath "C:\Program Files (x86)\RemotePackages\Shared Drive Admin.rdp" -WorkingDirectory "C:\Program Files (x86)\RemotePackages" 
    }
    catch {
        Add-OutputText "Error Opening Shared Drive App. Please try again."
    }
})
$adminToolsGroup.Controls.Add($sdaButton)

#Explorer++ Button
$ExpPlusButton = New-Object System.Windows.Forms.Button
$ExpPlusButton.Text = "Explorer++"
$ExpPlusButton.Location = New-Object System.Drawing.Point(192, 60)
$ExpPlusButton.AutoSize = $true
$ExpPlusButton.Add_Click({
    Add-OutputText "Opening Explorer++..."
    # Add Explorer++ script here
    try {
        & Start-Process -FilePath "$pathExplorerPlus\Explorer++.exe" -ArgumentList "C:\Temp" 
    }
    catch {
        Add-OutputText "Error Opening Explorer++. Please try again."
    }
})
$adminToolsGroup.Controls.Add($ExpPlusButton)

#**********************************************
#Externa Buttons to the Group Panels
#Always Button on/off at the top of the form
$AlwaysButton = New-Object System.Windows.Forms.Button
$AlwaysButton.Size = New-Object System.Drawing.Size(100, 25)
$AlwaysButton.Location = New-Object System.Drawing.Point(870, 1)
$AlwaysButton.AutoSize = $false
$AlwaysButton.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$AlwaysButton.Text = "Always Off"

# Always on Top/Off Button
$form.Controls.Add($AlwaysButton)
$toolTip.SetToolTip($AlwaysButton, "Window Always on top or not.")
$form.TopMost = $false

# Define the button click event
$AlwaysButton.Add_Click({
    if ($form.TopMost) {
        $form.TopMost = $false
        $AlwaysButton.Text = "Always Off"
    } else {
        $form.TopMost = $true
        $AlwaysButton.Text = "Always on Top"
    }
})

#Create Reset buttons for PCs/Users
$pcsresetButton = New-Object System.Windows.Forms.Button
$pcsresetButton.Text = "Reset PC Boxes"
$pcsresetButton.Size = New-Object System.Drawing.Size(110, 25)
$pcsresetButton.Location = New-Object System.Drawing.Point(360, 1)
$pcsresetButton.AutoSize = $false
$pcsresetButton.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$pcsresetButton.Add_Click({
    
    # Add your Reset boxes script here
    $hostnameTextBox.Text = ""
    $modelTextBoxLabel.Text = "..."
    $cDriveTextBoxLabel.Text = "..."	
    $laptopPcTextBoxLabel.Text = "..."
    $networkTextBoxLabel.Text = "..."
    $printSpoolerStatusLabel.Text = "..."
    $remoteRegStatusLabel.Text = "..."
    $eprGWStatusLabel.Text = "..."
    $eprQMStatusLabel.Text = "..."
    $pingSpeedTextBox.Text = "..."
    $pcrestartrsnTextBox.Text = "Enter Reason for Restart"
    $pcrestartdlyTextBox.Text = "30"
    $addPCGroupTextBox.Text = ""
    $boottimeStatusLabel.Text = "..."
    $addPrinterPCTextBox.Text = ""
    $LastUserStatusLabel.Text = "..."
})
$form.Controls.Add($pcsresetButton)

$allresetButton = New-Object System.Windows.Forms.Button
$allresetButton.Text = "Reset All Boxes"
$allresetButton.Size = New-Object System.Drawing.Size(120, 25)
$allresetButton.Location = New-Object System.Drawing.Point(749, 1)
$allresetButton.AutoSize = $false
$allresetButton.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$allresetButton.Add_Click({
    
    # Add your Reset boxes script here
    $hostnameTextBox.Text = ""
    $modelTextBoxLabel.Text = "..."
    $cDriveTextBoxLabel.Text = "..."	
    $laptopPcTextBoxLabel.Text = "..."
    $networkTextBoxLabel.Text = "..."
    $printSpoolerStatusLabel.Text = "..."
    $remoteRegStatusLabel.Text = "..."
    $wusStatusLabel.Text = "..."
    $mqStatusLabel.Text = "..."
    $pingSpeedTextBox.Text = "..."
    $pcrestartrsnTextBox.Text = "Enter Reason for Restart"
    $pcrestartdlyTextBox.Text = "30"
    $addPCGroupTextBox.Text = ""
    $boottimeStatusLabel.Text = "..."
    $addPrinterPCTextBox.Text = ""
    $LastUserStatusLabel.Text = "..."
    $usernameTextBox.Text = ""
    $wildcardsrcDropDownBox.Items.Clear()
    $wildcardsrcDropDownBox.Text = ""
    $displaynameTextBox.Text = ""
    $jobtitleTextBox.Text = ""
    $deptTextBox.Text = ""
    $emailTextBox.Text = ""  
    $phoneTextBox.Text = ""
    $descrTextBox.Text = ""
    $datecrTextBox.Text = ""
    $datemodTextBox.Text = ""
    $lastlogondateTextBox.Text = ""
    $pwddateTextBox.Text = ""
    $pwdexpTextBox.Text = ""
    $pwdnexpTextBox.Text = ""
    $lastbadpwdTextBox.Text = ""
    $lockedoutTextBox.Text = ""
    $lockedoutdateTextBox.Text = ""
    $accountenabledTextBox.Text = ""
    $addgroupTextBox.Text = ""
})
$form.Controls.Add($allresetButton)



#Reset Output Panel Button
$resetOutputButton = New-Object System.Windows.Forms.Button
$resetOutputButton.Text = "Clear Output Panel"
$resetOutputButton.Size = New-Object System.Drawing.Size(130, 25)
$resetOutputButton.Location = New-Object System.Drawing.Point(360, 665)
$resetOutputButton.AutoSize = $true
$resetOutputButton.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$resetOutputButton.Add_Click({
    
    # Add your Clear Output Panel script here
    $outputTextBox.Text = ""
})
$form.Controls.Add($resetOutputButton)

$userresetButton = New-Object System.Windows.Forms.Button
$userresetButton.Text = "Reset User Boxes"
$userresetButton.Size = New-Object System.Drawing.Size(120, 25)
$userresetButton.Location = New-Object System.Drawing.Point(15, 1)
$userresetButton.AutoSize = $false
$userresetButton.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$userresetButton.Add_Click({   
    # Add your Reset boxes script here
    $usernameTextBox.Text = ""
    $wildcardsrcDropDownBox.Items.Clear()
    $wildcardsrcDropDownBox.Text = ""
    $displaynameTextBox.Text = ""
    $jobtitleTextBox.Text = ""
    $deptTextBox.Text = ""
    $emailTextBox.Text = ""  
    $phoneTextBox.Text = ""
    $descrTextBox.Text = ""
    $datecrTextBox.Text = ""
    $datemodTextBox.Text = ""
    $lastlogondateTextBox.Text = ""
    $pwddateTextBox.Text = ""
    $pwdexpTextBox.Text = ""
    $pwdnexpTextBox.Text = ""
    $lastbadpwdTextBox.Text = ""
    $lockedoutTextBox.Text = ""
    $lockedoutdateTextBox.Text = ""
    $accountenabledTextBox.Text = ""
    $addgroupTextBox.Text = ""
    
})
$form.Controls.Add($userresetButton)

# Create a XKCD button
$xkcdbutton = New-Object System.Windows.Forms.Button
$xkcdbutton.Text = "Do Not Click Me!"
$xkcdbutton.Size = New-Object System.Drawing.Size(120, 25)
$xkcdbutton.Location = New-Object System.Drawing.Point(848, 665)
$xkcdButton.AutoSize = $false
$xkcdButton.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$xkcdbutton.Add_Click({
    # Generate a random XKCD URL
    #$randomNumber = Get-Random -Minimum 1 -Maximum 4500
    $url = "https://c.xkcd.com/random/comic/"
    
    # Open Microsoft Edge with the random XKCD URL
    Start-Process "msedge.exe" -ArgumentList $url
})
$form.Controls.Add($xkcdbutton)

# Create a Linktree button
$linktreebutton = New-Object System.Windows.Forms.Button
$linktreebutton.Text = "OYM"
$linktreebutton.Size = New-Object System.Drawing.Size(50, 25)
$linktreebutton.Location = New-Object System.Drawing.Point(795, 665)
$linktreeButton.AutoSize = $false
$linktreeButton.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$linktreebutton.Add_Click({
    # Opens my linktr.ee page in Microsoft Edge
    $url = "https://linktr.ee/seadevil4"
    
    # Open Microsoft Edge with the random Linktree URL
    Start-Process "msedge.exe" -ArgumentList $url
})
$form.Controls.Add($linktreebutton)

# Create output preview pane
$outputLabel = New-Object System.Windows.Forms.Label
$outputLabel.Text = "Output Panel"
$outputLabel.Location = New-Object System.Drawing.Point(10, 670)
$outputLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($outputLabel)

$outputTextBox = New-Object System.Windows.Forms.RichTextBox
$outputTextBox.Location = New-Object System.Drawing.Point(10, 700)
$outputTextBox.Size = New-Object System.Drawing.Size(960, 150)
$outputTextBox.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Regular)
$outputTextBox.Multiline = $true
$outputTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$form.Controls.Add($outputTextBox)

# Show the form
[void] $form.ShowDialog()