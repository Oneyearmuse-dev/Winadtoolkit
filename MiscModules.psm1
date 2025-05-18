function Add-FilesDesktop {
    param (
        [Parameter(Mandatory = $true)]
        [string]$param1
    )
    $remoteComputer = $param1

    # Define Public Desktop path on the remote computer
    $publicDesktopPath = "\\$remoteComputer\C$\Users\Public\Desktop"   
    # Open a file dialog to select files
    Add-Type -AssemblyName System.Windows.Forms
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Multiselect = $true
    $fileDialog.Title = "Select Files to Copy to Remote Public Desktop"
    $fileDialog.Filter = "All Files (*.*)|*.*"

    if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $selectedFiles = $fileDialog.FileNames

        # Copy each selected file to the remote public desktop folder
        foreach ($file in $selectedFiles) {
            $destinationPath = Join-Path -Path $publicDesktopPath -ChildPath (Split-Path -Leaf $file)
            Copy-Item -Path $file -Destination $destinationPath -Force
            Add-OutputText "Copied $file to $destinationPath"
        }
    } else {
        Add-OutputText "No files were selected."
    }
}
function Set-RemoteRegistrySvc {
    param (
        [string]$ComputerName,
        [string]$svcname,
        [ValidateSet("start", "stop")]
        [string]$Action
    )

    # Define the service name
    $serviceName = $svcname

    # Check the action and perform the corresponding operation
    if ($Action -eq "start") {
        Invoke-Command -ComputerName $ComputerName -ScriptBlock {
            Set-Service -Name $using:serviceName -StartupType Manual
            Start-Service -Name $using:serviceName
        }
        Add-OutputText "$serviceName service started on $ComputerName"
        return "$serviceName service started on $ComputerName"
    } elseif ($Action -eq "stop") {
        Invoke-Command -ComputerName $ComputerName -ScriptBlock {
            Stop-Service -Name $using:serviceName
            Start-Sleep -Seconds 5
            Set-Service -Name $using:serviceName -StartupType Disabled
        }
        Add-OutputText "$serviceName service stopped on $ComputerName"
        return "$serviceName service stopped on $ComputerName"
    } else {
        return "Invalid action specified. Use 'start' or 'stop'."
    }
}

# Example usage:
# Set-RemoteRegistrySvc -ComputerName "RemotePCName" -svcname "RemoteRegistry" -Action "start"
# Set-RemoteRegistrySvc -ComputerName "RemotePCName" -svcname "RemoteRegistry" -Action "stop"

function Set-StopStartService {
    param (
        [string]$ComputerName,
        [string]$svcname,
        [ValidateSet("start", "stop")]
        [string]$Action
    )

    # Define the service name
    $serviceName = $svcname

    # Check the action and perform the corresponding operation
    if ($Action -eq "start") {
        Invoke-Command -ComputerName $ComputerName -ScriptBlock {
            Start-Service -Name $using:serviceName
        }
        Add-OutputText "$serviceName service started on $ComputerName"
        return "$serviceName service started on $ComputerName"
    } elseif ($Action -eq "stop") {
        Invoke-Command -ComputerName $ComputerName -ScriptBlock {
            Stop-Service -Name $using:serviceName
            Start-Sleep -Seconds 5
        }
        Add-OutputText "$serviceName service stopped on $ComputerName"
        return "$serviceName service stopped on $ComputerName"
    } else {
        return "Invalid action specified. Use 'start' or 'stop'."
    }
}

# Example usage:
# Set-StopStartService -ComputerName "RemotePCName" -svcname "RemoteRegistry" -Action "start"
# Set-StopStartService -ComputerName "RemotePCName" -svcname "RemoteRegistry" -Action "stop"
function Add-PCGroup {
    param (
        [Parameter(Mandatory = $true)]
        [string]$param1, # Hostname
        [string]$param2  # Security group name
    )

    # Set Variables & Prefix the security group name and hostname
    $remoteComputer = $param1
    $securityGroupName = "$param2"

    # Check if the security group exists in Active Directory
    $group = Get-ADGroup -Filter { Name -eq $securityGroupName } -ErrorAction SilentlyContinue

    if ($null -ne $group) {
        #return "Security group '$securityGroupName' exists. Proceeding to add it to the hostname '$remoteComputer'."

        # Get the computer object in Active Directory
        $computer = Get-ADComputer -Filter { Name -eq $remoteComputer } -ErrorAction SilentlyContinue

        if ($null -ne $computer) {
            # Add the security group to the computer object   
            Add-Adgroupmember -id  $securityGroupName -Members "$remoteComputer$"
            Add-OutputText "Successfully added security group $securityGroupName to the hostname $remoteComputer"
            return "Successfully added security group '$securityGroupName' to the hostname '$remoteComputer'."
        } else {
            
            Add-OutputText "Error: Hostname '$remoteComputer' does not exist in Active Directory."
            return "Error: Hostname '$remoteComputer' does not exist in Active Directory."
            
        }
    } else {
        
        Add-OutputText "Error: Security group '$securityGroupName' does not exist in Active Directory."
        return "Error: Security group '$securityGroupName' does not exist in Active Directory."
        
    }
}
function Add-OutputText {
        param (
        [string]$text
        )
    $outputTextBox.AppendText("$text`n")
    $outputTextBox.SelectionStart = $outputTextBox.Text.Length
    $outputTextBox.ScrollToCaret()
    }
    
    # Example usage
#    Add-OutputText "This is a sample output text."
#    Add-OutputText "Another line of output."
    