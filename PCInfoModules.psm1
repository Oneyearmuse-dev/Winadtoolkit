function Get-CDriveSpace {
    param (
        [string]$param1
    )
    $remoteComputer = $param1  
    $driveLetter = "C:"

    try {
        $diskSpace = Invoke-Command -ComputerName $remoteComputer -ScriptBlock {
            param($driveLetter)
            $drive = Get-WmiObject -Class Win32_LogicalDisk -Filter "DeviceID='$driveLetter'"
            [PSCustomObject]@{
                ComputerName = $env:COMPUTERNAME
                DriveLetter  = $drive.DeviceID
                FreeSpace    = [math]::Round($drive.FreeSpace / 1GB, 2)
                TotalSpace   = [math]::Round($drive.Size / 1GB, 2)
            }
        } -ArgumentList $driveLetter

        return $diskSpace
    } catch {
        $errorDetails = $_
        Add-OutputText "Failed to retrieve disk space information for $remoteComputer. Error: $errorDetails"
        return $null
    }
}

function Get-PCType {
    param (
        [string]$param1
    )
#Is this a Laptop / PC
#
# Uses PC Number from main form
$remoteComputer = $param1

try {
    $battery = Invoke-Command -ComputerName $remoteComputer -ScriptBlock {Get-WmiObject -Class Win32_Battery}
    
    if ($battery) {
        $pcType = "Laptop"
    } else {
        $pcType = "Desktop"
    }
    return $pcType
} catch {
    $errorDetails = $_
    Add-OutputText "Failed to determine PC type for $remoteComputer. Error: $errorDetails"
    return "Error"
    }
}
function Get-Model {
    param (
        [Parameter(Mandatory = $true)]
        [string]$param1
    )
    # Get the computer name from the parameter
    $remoteComputer = $param1
    # Get the model information using WMI
    try {
        $model = Invoke-Command -ComputerName $remoteComputer -ScriptBlock {
            Get-WmiObject -Class Win32_ComputerSystem | Select-Object -ExpandProperty SystemFamily
        }
        return $model
    } catch {
        $errorDetails = $_
        Add-OutputText "Failed to retrieve model information for $remoteComputer. Error: $errorDetails"
        return "Error"
    }
}
function Get-AveragePing {
    param (
        [string]$param1,
        [int]$pingCount = 4
    )
    # Get the remote computer name from the parameter
    $remoteComputer = $param1

    # Initialize total time
    $totalTime = 0

    # Perform the ping test
    $pingResults = Test-Connection -ComputerName $remoteComputer -Count $pingCount

    # Calculate the total time
    foreach ($result in $pingResults) {
        $totalTime += $result.ResponseTime
    }

    # Calculate the average time
    $averageTime = $totalTime / $pingCount
    return $averageTime
}
function Get-ServicesStatus {
    param (
        [string]$param1
    )
    $remoteComputer = $param1

    # Initialize Services Array
    $ServicesList = @()

    # Spooler Status
    try {
        $service = Get-Service -ComputerName $remoteComputer -Name "Spooler" -ErrorAction Stop
        if ($service.Status -eq 'Running') {
            $spoolerStatus = "Running"
        } elseif ($service.Status -eq 'Stopped') {
            $spoolerStatus = "Stopped"
        } else {
            $spoolerStatus = "Unknown"
        }
    } catch {
        Add-OutputText "Failed to retrieve Print Spooler service status on $remoteComputer $_"
        $spoolerStatus = "Error"
    }
    $ServicesList += $spoolerStatus

    # Remote Registry Status
    try {
        $service = Get-Service -ComputerName $remoteComputer -Name "RemoteRegistry" -ErrorAction Stop
        if ($service.Status -eq 'Running') {
            $remoteRegStatus = "Running"
        } elseif ($service.Status -eq 'Stopped') {
            $remoteRegStatus = "Stopped"
        } else {
            $remoteRegStatus = "Unknown"
        }
    } catch {
        Add-OutputText "Failed to retrieve Remote Registry service status on $remoteComputer $_"
        $remoteRegStatus = "Error"
    }
    $ServicesList += $remoteRegStatus

    # Windows Update Status
    try {
        $service = Get-Service -ComputerName $remoteComputer -Name "wuauserv" -ErrorAction Stop
        if ($service.Status -eq 'Running') {
            $wusStatus = "Running"
        } elseif ($service.Status -eq 'Stopped') {
            $wusStatus = "Stopped"
        } else {
            $wusStatus = "Unknown"
        }
    } catch {
        Add-OutputText "Failed to retrieve Windows Update service status on $remoteComputer $_"
        $wusStatus = "Error"
    }
    $ServicesList += $wusStatus

    # Message Queue Status
    try {
        $service = Get-Service -ComputerName $remoteComputer -Name "MSMQ" -ErrorAction Stop
        if ($service.Status -eq 'Running') {
            $mqStatus = "Running"
        } elseif ($service.Status -eq 'Stopped') {
            $mqStatus = "Stopped"
        } else {
            $mqStatus = "Unknown"
        }
    } catch {
        Add-OutputText "Failed to retrieve Message Queueing service status on $remoteComputer $_"
        $mqStatus = "Error"
    }
    $ServicesList += $mqStatus

    # Return the array of service statuses
    return $ServicesList
}
# Function to get connection type from a remote PC
function Get-Network {
    param (
        [string]$param1
    )
    $remoteComputer = $param1

    if (-not $remoteComputer) {
        Write-Error "ComputerName parameter cannot be null or empty."
        return
    }

    $WirelessConnected = $null
    $WiredConnected = $null
    $VPNConnected = $null

    $scriptBlock = {
        # Detecting PowerShell version, and call the best cmdlets
        if ($PSVersionTable.PSVersion.Major -gt 2) {
            # Using Get-CimInstance for PowerShell version 3.0 and higher
            $WirelessAdapters = Get-CimInstance -Namespace "root\WMI" -Class MSNdis_PhysicalMediumType -Filter 'NdisPhysicalMediumType = 9'
            $WiredAdapters = Get-CimInstance -Namespace "root\WMI" -Class MSNdis_PhysicalMediumType -Filter "NdisPhysicalMediumType = 0 and (NOT InstanceName like '%pangp%') and (NOT InstanceName like '%cisco%') and (NOT InstanceName like '%juniper%') and (NOT InstanceName like '%vpn%') and (NOT InstanceName like 'Hyper-V%') and (NOT InstanceName like 'VMware%') and (NOT InstanceName like 'VirtualBox Host-Only%')"
            $ConnectedAdapters = Get-CimInstance -Class win32_NetworkAdapter -Filter 'NetConnectionStatus = 2'
            $VPNAdapters = Get-CimInstance -Class Win32_NetworkAdapterConfiguration -Filter "Description like '%pangp%' or Description like '%cisco%' or Description like '%juniper%' or Description like '%vpn%'"
        } else {
            # Needed this script to work on PowerShell 2.0 (don't ask)
            $WirelessAdapters = Get-WmiObject -Namespace "root\WMI" -Class MSNdis_PhysicalMediumType -Filter 'NdisPhysicalMediumType = 9'
            $WiredAdapters = Get-WmiObject -Namespace "root\WMI" -Class MSNdis_PhysicalMediumType -Filter "NdisPhysicalMediumType = 0 and (NOT InstanceName like '%pangp%') and (NOT InstanceName like '%cisco%') and (NOT InstanceName like '%juniper%') and (NOT InstanceName like '%vpn%') and (NOT InstanceName like 'Hyper-V%') and (NOT InstanceName like 'VMware%') and (NOT InstanceName like 'VirtualBox Host-Only%')"
            $ConnectedAdapters = Get-WmiObject -Class win32_NetworkAdapter -Filter 'NetConnectionStatus = 2'
            $VPNAdapters = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter "Description like '%pangp%' or Description like '%cisco%' or Description like '%juniper%' or Description like '%vpn%'"
        }

        Foreach ($Adapter in $ConnectedAdapters) {
            If ($WirelessAdapters.InstanceName -contains $Adapter.Name) {
                $WirelessConnected = $true
            }
        }

        Foreach ($Adapter in $ConnectedAdapters) {
            If ($WiredAdapters.InstanceName -contains $Adapter.Name) {
                $WiredConnected = $true
            }
        }

        Foreach ($Adapter in $ConnectedAdapters) {
            If ($VPNAdapters.Index -contains $Adapter.DeviceID) {
                $VPNConnected = $true
            }
        }

        If (($WirelessConnected -ne $true) -and ($WiredConnected -eq $true)) { $ConnectionType = "Wired" }
        #If (($WirelessConnected -eq $true) -and ($WiredConnected -eq $true)) { $ConnectionType = "WIRED AND WIRELESS" }
        If (($WirelessConnected -eq $true) -and ($WiredConnected -ne $true)) { $ConnectionType = "Wi-Fi" }
        If ($VPNConnected -eq $true) { $ConnectionType = "VPN" }

        return $ConnectionType
    }
    
    try {
        Invoke-Command -ComputerName $remoteComputer -ScriptBlock $scriptBlock
    } catch {
        $errorDetails = $_
        Add-OutputText "Failed to retrieve network information for $remoteComputer. Error: $errorDetails"
        return $null
    }
}

# Call the function with $param1
#Get-Network -ComputerName $param1
#$remoteComputer = $param1
function Get-BootTime {
    param (
        [string]$param1  # Remote computer name
    )
    $remoteComputer = $param1
    try {
        $currentTime = Get-Date
        $boottime = Invoke-Command -ComputerName $remoteComputer -Command { ([wmi]'').ConvertToDateTime((gwmi win32_operatingsystem).LastBootUpTime) }
    
        $formattedDate = $boottime.ToString("dd/MM/yy")
        $daysSinceBoot = ($currentTime - $boottime).Days
        $dateDays = $formattedDate + " (" + $daysSinceBoot + " days)"

        return $dateDays
    }
    catch {
        $errorDetails = $_
        Add-OutputText "Failed to retrieve boot time for $remoteComputer. Error: $errorDetails"
        return "Error"
    }
    

    
}
function Get-LastLoggedOnUser {
    param (
        [string]$param1  # Remote computer name
    )
    $remoteComputer = $param1
    try {
        # Query the Win32_ComputerSystem class on the remote computer
        $computerSystem = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $remoteComputer -ErrorAction Stop

        # Extract the username of the last logged-on user
        $lastUser = $computerSystem.UserName

        if ($lastUser) {
            # Remove the domain from the username (e.g., DOMAIN\Username -> Username)
            $lastUser = $lastUser -replace '.*\\', ''  # Regex to remove everything before and including the backslash
            return $lastUser
        } else {
            return "No Data"
        }
    } catch {
        $errorDetails = $_
        Add-OutputText "Failed to retrieve last logged-on user for $remoteComputer. Error: $errorDetails"
        return "Error"
    }
}

# Example usage
#$lastUser = Get-LastLoggedOnUser -ComputerName "RemotePCName"

# Import functions for the scripts

    Import-Module -Name ActiveDirectory
    
function Get-PCInfo {
    param (
        [Parameter(Mandatory = $true)]
        [string]$param1
    )

    $remoteComputer = $param1
            
    #Get Model type of PC
    $model = Get-Model -param1 $remoteComputer
    
    #Get PC Type (Laptop or Desktop)
    $pctype = Get-PCType -param1 $remoteComputer
    
    #Get Disk Space Data
    $diskInfo = Get-CDriveSpace -param1 $remoteComputer
    $freeSpace = $diskInfo.FreeSpace
    $totalSpace = $diskInfo.TotalSpace

    # Prepare the output text
    $diskInfoText = "$freeSpace GB / $totalSpace GB"
    
    # Get Services Status for Print Spooler, Remote Registry, EPR Gateway, and EPR Queue Monitor
    $servicesStatus = Get-ServicesStatus -param1 $remoteComputer
    
    #Get Network Info
    $networkInfo = Get-Network -param $remoteComputer

    #Ping Speed Test
    $pingSpeed = Get-AveragePing -param1 $remoteComputer -pingCount 4

    #Boot Time Info
    $pcboottime = Get-BootTime -param1 $remoteComputer

    #Last Logged On User
    $lluser = Get-LastLoggedOnUser -param1 $remoteComputer

$PCInfoData = [PSCustomObject]@{
    Model        = $model
    CDriveSpace  = $diskInfoText
    LapDesk      = $pctype
    Network      = $networkInfo
    PSSvr		 = $servicesStatus[0]
    RRSvr        = $servicesStatus[1]
    wusSvr       = $servicesStatus[2]
    mqSvr        = $servicesStatus[3]
    PingSpd      = $pingSpeed
    Boottime     = $pcboottime
    LastUser    = $lluser
}
    return $PCInfoData
  
}