function Get-ADUsername {
    param (
        [Parameter(Mandatory = $true)]
        [string]$param1
    )

    $UserName = $param1
    $user = Get-ADUser -Identity $UserName -Properties *
    if ($null -eq $user) {
        Add-OutputText "User not found"
        return "User not found`r`n"
    }
    $displaynameTextBox.Text = $user.DisplayName
    $jobtitleTextBox.Text = $user.Title
    $deptTextBox.Text = $user.Department
    $emailTextBox.Text = $user.EmailAddress
    $phoneTextBox.Text = $user.TelephoneNumber
    $descrTextBox.Text = $user.Description
    $datecrTextBox.Text = $user.whenCreated.ToString("dd/MM/yyyy HH:mm:ss")
    $datemodTextBox.Text = $user.whenChanged.ToString("dd/MM/yyyy HH:mm:ss")
    $lastlogondateTextBox.Text = $user.LastLogonDate.ToString("dd/MM/yyyy HH:mm:ss")
    $pwddateTextBox.Text = $user.PasswordLastSet.ToString("dd/MM/yyyy HH:mm:ss")
    $pwdexpTextBox.Text = $user.PasswordExpired
    $pwdnexpTextBox.Text = $user.PasswordNeverExpires
    $BadPasswordCountTextBox = $user.BadPasswordCount
    $lastbadpwdTextBox.Text = $user.LastBadPasswordAttempt.ToString("dd/MM/yyyy HH:mm:ss") + " BPC - $BadPasswordCountTextBox"
    $lockedoutTextBox.Text = $user.LockedOut
    
    # Convert the lockout time from file time to DateTime
    if ($user.lockoutTime -eq 0) {
        $lockedoutdateTextBox.Text = "False"
    }else {
        $lockoutDateTime = [DateTime]::FromFileTime($user.lockoutTime)
        $lockedoutdateTextBox.Text = $lockoutDateTime.ToString("dd/MM/yyyy HH:mm:ss")    
    }
        
    $accountenabledTextBox.Text = $user.Enabled

}
function Get-WildcardADUsername {
    param (
        [Parameter(Mandatory = $true)]
        [string]$param1
    )

    $UserName = $param1
    $wildcardsrcDropDownBox.Items.Clear()
    try {
        $SAMAccountNames = Get-ADUser -Filter "Name -like '*$username*'" -Properties SAMAccountName | Sort-Object SAMAccountName
        if ($SAMAccountNames) {
            foreach ($SAMAccountName in $SAMAccountNames) {
                $wildcardsrcDropDownBox.Items.Add($SAMAccountName.SAMAccountName)
            }
            Add-OutputText "Select from the drop down menu and press * button."
        } else {
            Add-OutputText "No Users"
            return "No Users"
        }
    }
    catch {
        Add-OutputText "User not found"
        return "User not found"
    }
}
function Get-EmailADUsername {
    param (
        [Parameter(Mandatory = $true)]
        [string]$param1
    )

    $email = $param1
    $user = Get-ADUser -Filter {EmailAddress -eq $email} -Properties *
    if ($null -eq $user) {
        Add-OutputText "User not found"
        return "User not found`r`n"
    }
    $usernameTextBox.Text = $user.SamAccountName
    $displaynameTextBox.Text = $user.DisplayName
    $jobtitleTextBox.Text = $user.Title
    $deptTextBox.Text = $user.Department
    $emailTextBox.Text = $user.EmailAddress
    $phoneTextBox.Text = $user.TelephoneNumber
    $descrTextBox.Text = $user.Description
    $datecrTextBox.Text = $user.whenCreated.ToString("dd/MM/yyyy HH:mm:ss")
    $datemodTextBox.Text = $user.whenChanged.ToString("dd/MM/yyyy HH:mm:ss")
    $lastlogondateTextBox.Text = $user.LastLogonDate.ToString("dd/MM/yyyy HH:mm:ss")
    $pwddateTextBox.Text = $user.PasswordLastSet.ToString("dd/MM/yyyy HH:mm:ss")
    $pwdexpTextBox.Text = $user.PasswordNeverExpires
    $pwdnexpTextBox.Text = $user.PasswordExpired
    $BadPasswordCountTextBox = $user.BadPasswordCount
    $lastbadpwdTextBox.Text = $user.LastBadPasswordAttempt.ToString("dd/MM/yyyy HH:mm:ss") + " BPC - $BadPasswordCountTextBox"
    $lockedoutTextBox.Text = $user.LockedOut
    # Convert the lockout time from file time to DateTime
    if ($user.lockoutTime -eq 0) {
        $lockedoutdateTextBox.Text = "False"
    }else {
        $lockoutDateTime = [DateTime]::FromFileTime($user.lockoutTime)
        $lockedoutdateTextBox.Text = $lockoutDateTime.ToString("dd/MM/yyyy HH:mm:ss")    
    }
    $accountenabledTextBox.Text = $user.Enabled
}