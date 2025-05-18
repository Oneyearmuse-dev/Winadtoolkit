function Disable-Hibernate {
    param (
        [string]$param1
    )
    $remoteComputer = $param1

    # Define the script block to turn off hibernate
    $scriptBlock = {
        $hibStatus = ""
        # Disable hibernate
        powercfg.exe /hibernate off       
    }
    try {
        Invoke-Command -ComputerName $remoteComputer -ScriptBlock $scriptBlock -ErrorAction Stop
        Add-OutputText "Hibernate disabled successfully on $remoteComputer"
    } catch {
                 
        Add-OutputText "Failed to disable hibernate on $remoteComputer. Error: $_"
        
    }
    
}

# Example usage of the function
#Disable-Hibernate -param1 "RemotePCName"  # Replace "RemotePCName" with the actual remote computer name
#Remove Software Distribution Folder - stops wuauserv first - clear folder - restart wuauserv
function Remove-SWDFolder {
    param (
        [Parameter(Mandatory = $true)]
        [string]$param1  # Remote computer name
    )
    $remoteComputer = $param1
    try {
        Invoke-Command -ComputerName $remoteComputer -ScriptBlock {
            try {
                # Stop the Windows Update service
                Stop-Service -Name wuauserv -Force

                # Clear the SoftwareDistribution folder
                $swdFolder = "C:\Windows\SoftwareDistribution"
                Remove-Item -Path "$swdFolder\*" -Recurse -Force

                # Start the Windows Update service
                Start-Service -Name wuauserv
            } catch {
                         
                Add-OutputText "Error during operation on remote computer: $_"
                throw "Error during operation on remote computer: $_"
                
            }
        }

        $wuservStatus = Get-Service -Name wuauserv -ComputerName $remoteComputer
        if ($wuservStatus.Status -eq 'Running') {
            Add-OutputText "Software Distribution folder cleared and wuauserv service restarted on $remoteComputer"
        } else {
                     
            Add-OutputText "Failed to restart wuauserv service on $remoteComputer"
            
        }
    } catch {
                 
        Add-OutputText "An error occurred while processing $remoteComputer. Error: $_"
        
    }
}
function Remove-WV2 {
    param (
        [Parameter(Mandatory = $true)]
        [string]$param1
    )
    # Created using Microsoft Copilot AI - Testing by Warren Order 20th Dec 2024
    # Clear WV2 folders
    # Remove all .WV2 folders in all profiles
    $remoteComputer = $param1

    # Define the base path to the Users directory on the remote PC
    $basePath = "\\$remoteComputer\C$\Users"

    # Get all user profiles in the Users directory
    $userProfiles = Get-ChildItem -Path $basePath -Directory

    foreach ($profile in $userProfiles) {
        try {
            # Define the path to the WV2 folder within the Local AppData\Local directory for each user profile
            $wv2Path = Join-Path -Path $profile.FullName -ChildPath "AppData\Local\*.WV2"

            # Get all .WV2 folders
            $wv2Folders = Get-ChildItem -Path $wv2Path -Directory -ErrorAction SilentlyContinue

            foreach ($wv2Folder in $wv2Folders) {
                try {
                    # Delete each .WV2 folder
                    Remove-Item -Path $wv2Folder.FullName -Recurse -Force
                    Add-OutputText "Deleted folder: $($wv2Folder.FullName) on $remoteComputer"
                } catch {
                             
                    Add-OutputText "Failed to delete folder: $($wv2Folder.FullName) on $remoteComputer. Error: $_"
                    
                }
            }
        } catch {
            
            Add-OutputText "Failed to process profile: $($profile.Name) on $remoteComputer. Error: $_"
            
        }
    }
    }
    function Remove-TeamFolders {
        param (
            [string]$param1
        )
        # Created using Microsoft Copilot AI - Testing by Warren Order 20th Dec 2024
        # Clear Teams folders older than 30 Days
        # Uses PC Number from main form
        $hostname = $param1
    
        # Define the base path to the Users directory on the remote PC
        $basePath = "\\$hostname\C$\Users"
    
        # Get the current date
        $currentDate = Get-Date
    
        # Define the age threshold (30 days)
        $thresholdDate = $currentDate.AddDays(-30)
    
        # Get all user profiles in the Users directory
        $userProfiles = Get-ChildItem -Path $basePath -Directory
    
        foreach ($profile in $userProfiles) {
            try {
            # Define the path to the Teams folder within the Local AppData\Local\Microsoft directory for each user profile
            $teamsPath = Join-Path -Path $profile.FullName -ChildPath "AppData\Local\Microsoft\Teams"
        
            # Check if the Teams folder exists
            if (Test-Path -Path $teamsPath) {
                try {
                # Get the Teams folder
                $teamsFolder = Get-Item -Path $teamsPath
        
                # Check the last write time of the Teams folder
                if ($teamsFolder.LastWriteTime -lt $thresholdDate) {
                    try {
                    # If the folder is older than the threshold, delete it
                    Remove-Item -Path $teamsFolder.FullName -Recurse -Force
                    Add-OutputText "Deleted folder: $($teamsFolder.FullName) on $hostname"
                    } catch {
                    
                    Add-OutputText "Failed to delete folder: $($teamsFolder.FullName) on $hostname. Error: $_"
                    
                    }
                } else {
                    Add-OutputText "The Teams folder is not older than 30 days in profile $($profile.Name) on $hostname"
                }
                } catch {
                    
                    Add-OutputText "Failed to process Teams folder in profile $($profile.Name) on $hostname. Error: $_"
                    
                }
            } else {
                Add-OutputText "Teams folder not found in profile $($profile.Name) on $hostname"
            }
            } catch {
            
                Add-OutputText "An error occurred while processing profile $($profile.Name) on $hostname. Error: $_"
                
            }
        }
        
        Add-OutputText "Team Folders Cleaned for $param1"
    } 
#Clean the User Profiles for 30 Days and older using Powershell instead of Delprof2
# Define the number of days 
function Remove-OldUserProfiles {
    param (   
        [string]$param1,  # Remote computer name
        [int]$days = 30   # Number of days
    )
    $remoteComputer = $param1
    Invoke-Command -ComputerName $remoteComputer -ScriptBlock {
        param ($days)
        
        # Get the current date
        $currentDate = Get-Date

        # Get all user profiles
        $userProfiles = Get-WmiObject -Class Win32_UserProfile | Where-Object { $_.Special -eq $false }

        foreach ($profile in $userProfiles) {
            # Calculate the age of the profile
            $profileAge = ($currentDate - $profile.LastUseTime).Days

            if ($profileAge -ge $days) {
                try {
                    # Remove the profile directory
                    Remove-Item -Path $profile.LocalPath -Recurse -Force

                    # Remove the registry entry
                    $sid = $profile.SID
                    Remove-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid" -Recurse -Force

                    Add-OutputText "Removed profile: $($profile.LocalPath)"
                } catch {
                    
                    Add-OutputText "Failed to remove profile: $($profile.LocalPath). Error: $_"
                    
                }
            }
        }
    } -ArgumentList $days
}

# Example usage
#Remove-OldUserProfiles -param1 $hostnameTextBox.Text -days 30  
    