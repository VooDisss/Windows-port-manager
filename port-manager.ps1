# Port Reservation Automation Script
# This script helps diagnose port conflicts and reserve port ranges on Windows

param(
    [switch]$Elevated
)

# Check if running as Administrator
function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Relaunch script as Administrator
function Invoke-Elevated {
    $scriptPath = $MyInvocation.MyCommand.Path
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "pwsh.exe" # Use pwsh for PowerShell 7+, or "powershell.exe" for Windows PowerShell 5.1
    $psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -Elevated"
    $psi.Verb = "RunAs"
    $psi.WorkingDirectory = Get-Location
    
    try {
        $process = [System.Diagnostics.Process]::Start($psi)
        # Do not wait for exit or exit the current shell, just let the new one take over.
        # The user can close the non-admin window manually.
        $process.WaitForExit()
        exit
    } catch {
        Write-Host "Failed to elevate privileges. Please run this script as Administrator manually." -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit
    }
}

# Check for admin rights and elevate if needed
if (-not (Test-Administrator) -and -not $Elevated) {
    Clear-Host
    Write-Host "============================================" -ForegroundColor Red
    Write-Host "   Administrator Privileges Required" -ForegroundColor Red
    Write-Host "============================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "This script requires administrator privileges to manage port reservations." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Please select an option:" -ForegroundColor Yellow
    Write-Host "1. Relaunch as Administrator (Recommended)"
    Write-Host "2. Exit"
    Write-Host ""
    
    $choice = Read-Host "Enter your choice (1-2)"
    
    if ($choice -eq "1") {
        Write-Host "Relaunching with administrator privileges..." -ForegroundColor Yellow
        Invoke-Elevated
    } else {
        Write-Host "Exiting..." -ForegroundColor Yellow
        exit
    }
}

# Main script continues here if running as Administrator
function Show-Header {
    Clear-Host
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host "   Windows Port Reservation Automation Tool" -ForegroundColor Cyan
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host "Running with Administrator privileges" -ForegroundColor Green
    Write-Host ""
}

function Show-Menu {
    Write-Host "Please select an option:" -ForegroundColor Yellow
    Write-Host "1. Show current port status (reserved ranges & dynamic range)"
    Write-Host "2. Check if specific port is in use"
    Write-Host "3. Reserve a port range"
    Write-Host "4. Delete a specific reserved port range"
    Write-Host "5. Change dynamic port range"
    Write-Host "6. Restart Windows NAT service (can resolve some conflicts)"
    Write-Host "7. Full diagnostics for a port range"
    Write-Host "8. Exit"
    Write-Host ""
    $choice = Read-Host "Enter your choice (1-8)"
    return $choice
}

function Show-PortStatus {
    Write-Host "Current Reserved Port Ranges:" -ForegroundColor Green
    $result = netsh int ipv4 show excludedportrange protocol=tcp
    $result
    
    Write-Host ""
    Write-Host "Current Dynamic Port Range:" -ForegroundColor Green
    $result = netsh int ipv4 show dynamicportrange tcp
    $result
}

# OPTIMIZED: Uses native Get-NetTCPConnection/Get-NetUDPEndpoint
function Check-PortUsage {
    param(
        [int]$port
    )
    
    Write-Host "Checking if port $port is in use (FAST Check)..." -ForegroundColor Yellow
    
    # Use native PowerShell cmdlets for better performance and object output
    $tcpResult = Get-NetTCPConnection -ErrorAction SilentlyContinue | Where-Object { $_.LocalPort -eq $port -and ($_.State -ne 'Closed' -and $_.State -ne 'TimeWait') }
    $udpResult = Get-NetUDPEndpoint -ErrorAction SilentlyContinue | Where-Object { $_.LocalPort -eq $port }
    
    $results = @($tcpResult) + @($udpResult)

    if ($results.Count -gt 0) {
        Write-Host "Port $port is in use by the following connections/endpoints:" -ForegroundColor Red
        
        $results | ForEach-Object {
            $processId = $_.OwningProcess
            $protocol = if ($_) { $_.PSComputerName } else { if ($_.LocalPort) { "UDP" } else { "TCP" } } # Simple protocol check
            $state = if ($_.State) {$_.State} else {"Bound"}
            $foreignPort = if ($_.RemotePort) {$_.RemotePort} else {"N/A"}

            try {
                $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
                $processName = if ($process) { $process.ProcessName } else { "Could not identify" }
                
                Write-Host ("- {0} Port {1} State: {2} (Foreign Port: {3}) by {4} (PID: {5})" -f $protocol, $_.LocalPort, $state, $foreignPort, $processName, $processId) -ForegroundColor Red
            } catch {
                Write-Host ("- {0} Port {1} State: {2} by Unknown Process (PID: {3})" -f $protocol, $_.LocalPort, $state, $processId) -ForegroundColor Red
            }
        }
        return $true
    } else {
        Write-Host "Port $port is not currently in use." -ForegroundColor Green
        return $false
    }
}

# OPTIMIZED: Uses native Get-NetTCPConnection/Get-NetUDPEndpoint and single Get-Process batch
function Check-RangeUsage {
    param(
        [int]$startPort,
        [int]$numberOfPorts
    )
    
    $endPort = $startPort + $numberOfPorts - 1
    Write-Host ("Checking ports {0} to {1} for usage (FAST Check)..." -f $startPort, $endPort) -ForegroundColor Yellow
    
    # Use native cmdlets to get all connections/endpoints in one go, then filter.
    # This is *significantly* faster than looping with netstat/findstr.
    $allTcp = Get-NetTCPConnection -ErrorAction SilentlyContinue
    $allUdp = Get-NetUDPEndpoint -ErrorAction SilentlyContinue
    
    $tcpResult = $allTcp | Where-Object { 
        $_.LocalPort -ge $startPort -and $_.LocalPort -le $endPort -and ($_.State -ne 'Closed' -and $_.State -ne 'TimeWait')
    }
    $udpResult = $allUdp | Where-Object { 
        $_.LocalPort -ge $startPort -and $_.LocalPort -le $endPort 
    }
    
    $results = @($tcpResult) + @($udpResult)
    
    $usedPortsInfo = @{}
    $PIDsToQuery = @{}
    
    if ($results.Count -gt 0) {
        
        $results | ForEach-Object {
            $processId = $_.OwningProcess
            $port = $_.LocalPort
            
            $PIDsToQuery[$processId] = $null # Collect unique PID
            
            # Store the port usage info
            if (-not $usedPortsInfo.ContainsKey($processId)) {
                $usedPortsInfo[$processId] = @{ ProcessName = ""; Ports = @() }
            }
            $usedPortsInfo[$processId].Ports += $port
        }
        
        # Query all unique process names in one batch
        $validPIDs = $PIDsToQuery.Keys | Where-Object { $_ -gt 0 }
        $processes = if ($validPIDs) { Get-Process -Id ($validPIDs -as [int[]]) -ErrorAction SilentlyContinue } else { @() }
        $processMap = @{}
        $processes | ForEach-Object { $processMap[$_.Id] = $_.ProcessName }
        
        # Fill in ProcessName and display summary
        $usedPortsInfo.GetEnumerator() | ForEach-Object {
            $processId = $_.Name
            $info = $_.Value
            $processName = if ($processMap.ContainsKey($processId)) { $processMap[$processId] } else { "Could not identify" }
            $info.ProcessName = $processName
        }
    }
    
    if ($usedPortsInfo.Count -gt 0) {
        Write-Host ("Found {0} processes using ports in the range {1}-{2}." -f $usedPortsInfo.Count, $startPort, $endPort) -ForegroundColor Red
        return $usedPortsInfo
    } else {
        Write-Host ("No ports in use in the range {0}-{1}." -f $startPort, $endPort) -ForegroundColor Green
        return $null
    }
}

# New helper function to check for overlap with existing reserved ranges
function Test-ReservedRangeConflict {
    param(
        [int]$startPort,
        [int]$endPort
    )
    
    $reservedRanges = netsh int ipv4 show excludedportrange protocol=tcp
    $lines = $reservedRanges -split '[\r\n]+'
    
    # Detect netsh output format by checking the header
    $headerLine = $lines | Where-Object { $_ -match "Start Port" }
    $isEndPortFormat = $headerLine -like "*End Port*"
    
    foreach ($line in $lines) {
        # Pattern matches start port and number of ports
        if ($line -match "^\s*(\d+)\s+(\d+)") {
            $rangeStart = [int]$matches[1]
            $secondValue = [int]$matches[2]
            
            # Calculate End Port based on detected format
            $rangeEnd = if ($isEndPortFormat) { $secondValue } else { $rangeStart + $secondValue - 1 }

            # Check for overlap: (Range A ends AFTER Range B starts) AND (Range B ends AFTER Range A starts)
            if ($endPort -ge $rangeStart -and $startPort -le $rangeEnd) {
                Write-Host ("CONFLICT: Desired range overlaps with existing reserved range {0}-{1}!" -f $rangeStart, $rangeEnd) -ForegroundColor Red
                return $true # Conflict found
            }
        }
    }
    return $false # No conflict found
}

function Get-ReservedRanges {
    [CmdletBinding()]
    param()

    $reservedRangesOutput = netsh int ipv4 show excludedportrange protocol=tcp
    $lines = $reservedRangesOutput -split '[\r\n]+'
    
    # Detect netsh output format by checking the header
    $headerLine = $lines | Where-Object { $_ -match "Start Port" }
    $isEndPortFormat = $headerLine -like "*End Port*"
    
    $ranges = @()
    foreach ($line in $lines) {
        # Pattern matches start port and the second value
        if ($line -match "^\s*(\d+)\s+(\d+)") {
            $startPort = [int]$matches[1]
            $secondValue = [int]$matches[2]
            
            $endPort = 0
            $numberOfPorts = 0

            if ($isEndPortFormat) {
                $endPort = $secondValue
                $numberOfPorts = $endPort - $startPort + 1
            } else { # Number of Ports format
                $numberOfPorts = $secondValue
                $endPort = $startPort + $numberOfPorts - 1
            }
            
            $ranges += [pscustomobject]@{ StartPort = $startPort; EndPort = $endPort; NumberOfPorts = $numberOfPorts }
        }
    }
    return $ranges
}

function Reserve-PortRange {
    try {
        $startPort = Read-Host "Enter the starting port number"
        $numberOfPorts = Read-Host "Enter the number of ports to reserve"
        
        # Validate input
        if (-not ($startPort -match '^\d+$') -or -not ($numberOfPorts -match '^\d+$')) {
            Write-Host "Invalid input. Please enter valid numbers." -ForegroundColor Red
            return
        }
        
        $startPort = [int]$startPort
        $numberOfPorts = [int]$numberOfPorts
        $endPort = $startPort + $numberOfPorts - 1
        
        Write-Host ("Attempting to reserve ports {0} to {1}..." -f $startPort, $endPort) -ForegroundColor Yellow
        
        # Check if ports are in use
        $usedPortsInfo = Check-RangeUsage -startPort $startPort -numberOfPorts $numberOfPorts -ErrorAction SilentlyContinue
        
        if ($usedPortsInfo) {
            Write-Host "`nThe following processes are using ports in the desired range:" -ForegroundColor Yellow
            $usedPortsInfo.GetEnumerator() | ForEach-Object {
                Write-Host ("- PID: {0}, Process: {1}, Ports: {2}" -f $_.Name, $_.Value.ProcessName, ($_.Value.Ports -join ', ')) -ForegroundColor Yellow
            }

            $response = Read-Host "`nWould you like to attempt to stop these processes? (y/n)"
            if ($response -ne 'y' -and $response -ne 'Y') {
                Write-Host "Reservation cancelled." -ForegroundColor Red
                return
            }

            Write-Host "Attempting to stop processes..." -ForegroundColor Yellow
            $usedPortsInfo.GetEnumerator() | ForEach-Object {
                $processId = $_.Name
                $processName = $_.Value.ProcessName
                try {
                    Stop-Process -Id $processId -Force -ErrorAction Stop
                    Write-Host ("Successfully stopped process {0} (PID: {1})" -f $processName, $processId) -ForegroundColor Green
                } catch {
                    Write-Host ("Failed to stop process {0} (PID: {1}). Error: {2}" -f $processName, $processId, $_.Exception.Message) -ForegroundColor Red
                }
            }
            
            # Re-check after stopping to mitigate race condition
            $usedPortsInfo = Check-RangeUsage -startPort $startPort -numberOfPorts $numberOfPorts
            if ($usedPortsInfo) {
                Write-Host "Reservation cancelled due to remaining active port usage." -ForegroundColor Red
                return
            }
        }

        # [NEW CRITICAL STEP] Check for existing excluded ranges
        if (Test-ReservedRangeConflict -startPort $startPort -endPort $endPort) {
            Write-Host "Reservation stopped due to conflict with an existing reserved range." -ForegroundColor Red
            Write-Host "If this is a system-created range (e.g., Hyper-V/Docker), try Option 6 to restart WinNAT." -ForegroundColor Yellow
            return
        }
        
        # Try to reserve the range
        $result = netsh int ipv4 add excludedportrange protocol=tcp startport=$startPort numberofports=$numberOfPorts 2>&1
        
        if ($result -like "*Ok.*") {
            Write-Host ("Successfully reserved ports {0} to {1}." -f $startPort, $endPort) -ForegroundColor Green
            Write-Host ""
            Write-Host "Updated list of reserved port ranges:" -ForegroundColor Cyan
            netsh int ipv4 show excludedportrange protocol=tcp
        } else {
            Write-Host ("Failed to reserve ports: {0}" -f $result) -ForegroundColor Red

            # Check if the failure is likely due to Dynamic Port Range overlap
            $dynamicRange = netsh int ipv4 show dynamicportrange tcp
            $lines = $dynamicRange -split "`n"
            $dynamicStart = 0; $dynamicEnd = 0; $dynamicNum = 0
            $lines | Where-Object { $_ -match "Start Port" } | ForEach-Object { if ($_ -match ":\s*(\d+)") { $dynamicStart = [int]$matches[1] } }
            $lines | Where-Object { $_ -match "Number of Ports" } | ForEach-Object { if ($_ -match ":\s*(\d+)") { $dynamicNum = [int]$matches[1]; $dynamicEnd = $dynamicStart + $dynamicNum - 1 } }

            $dynamicConflict = $dynamicStart -and $dynamicEnd -and ($endPort -ge $dynamicStart -and $startPort -le $dynamicEnd)
            
            if ($dynamicConflict) {
                Write-Host "This failure is likely due to a conflict with the dynamic port range." -ForegroundColor Yellow
                Write-Host "Current dynamic port range:" -ForegroundColor Yellow
                $dynamicRange
                
                $response = Read-Host "Would you like to change the dynamic port range to avoid conflicts? (y/n)"
                if ($response -eq 'y' -or $response -eq 'Y') {
                    Change-DynamicPortRange
                }
            } else {
                # [NEW RECOMMENDATION] If it failed and it wasn't dynamic or an existing exclusion, suggest WinNAT restart
                Write-Host "The failure is not due to active processes, dynamic range, or a pre-existing reservation." -ForegroundColor Yellow
                Write-Host "This could be due to a kernel-level conflict, such as a transient WinNAT issue." -ForegroundColor Yellow
                $response = Read-Host "Would you like to try Option 6 (Restart Windows NAT Service) and attempt the reservation again? (y/n)"
                if ($response -eq 'y' -or $response -eq 'Y') {
                    Restart-WinNAT
                    Write-Host "Please attempt the reservation again using the menu." -ForegroundColor Yellow
                }
            }
        }
    } catch {
        Write-Host "An unexpected error occurred in the Reserve-PortRange function." -ForegroundColor Red
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "At: $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line)" -ForegroundColor Red
    }
}

function Delete-ReservedRange {
    $reservedRanges = Get-ReservedRanges

    if ($reservedRanges.Count -eq 0) {
        Write-Host "There are no reserved port ranges to delete." -ForegroundColor Yellow
        return
    }

    Write-Host "The following port ranges are currently reserved:" -ForegroundColor Yellow
    for ($i = 0; $i -lt $reservedRanges.Count; $i++) {
        $range = $reservedRanges[$i]
        Write-Host ("[{0}] Start: {1,-5} | End: {2,-5} | Count: {3}" -f ($i + 1), $range.StartPort, $range.EndPort, $range.NumberOfPorts)
    }
    Write-Host "[c] Cancel"

    $choice = Read-Host "`nEnter the number of the range to delete, or 'c' to cancel"

    if ($choice -eq 'c') {
        Write-Host "Delete operation cancelled." -ForegroundColor Yellow
        return
    }

    if (($choice -match '^\d+$') -and ([int]$choice -ge 1) -and ([int]$choice -le $reservedRanges.Count)) {
        $selectedRange = $reservedRanges[[int]$choice - 1]
        $startPort = $selectedRange.StartPort
        $numberOfPorts = $selectedRange.NumberOfPorts

        Write-Host ("`nAttempting to delete reserved ports {0} to {1}..." -f $startPort, $selectedRange.EndPort) -ForegroundColor Yellow
        $result = netsh int ipv4 delete excludedportrange protocol=tcp startport=$startPort numberofports=$numberOfPorts 2>&1
        
        if ($result -like "*Ok.*") {
            Write-Host ("Successfully deleted reserved ports {0} to {1}." -f $startPort, $selectedRange.EndPort) -ForegroundColor Green
        } else {
            Write-Host ("Failed to delete reserved ports: {0}" -f $result) -ForegroundColor Red
        }
    } else {
        Write-Host "Invalid selection. Please enter a number from the list." -ForegroundColor Red
    }
}

function Change-DynamicPortRange {
    $startPort = Read-Host "Enter the new starting port for dynamic range (default: 49152)"
    $numPorts = Read-Host "Enter the number of ports for dynamic range (default: 16384)"
    
    # Use defaults if empty
    if ([string]::IsNullOrWhiteSpace($startPort)) { $startPort = 49152 }
    if ([string]::IsNullOrWhiteSpace($numPorts)) { $numPorts = 16384 }
    
    # Validate input
    if (-not ($startPort -match '^\d+$') -or -not ($numPorts -match '^\d+$')) {
        Write-Host "Invalid input. Please enter valid numbers." -ForegroundColor Red
        return
    }
    
    $startPort = [int]$startPort
    $numPorts = [int]$numPorts
    $endPort = $startPort + $numPorts - 1
    
    Write-Host ("Setting dynamic port range to {0}-{1}..." -f $startPort, $endPort) -ForegroundColor Yellow
    
    $result = netsh int ipv4 set dynamicport tcp start=$startPort num=$numPorts 2>&1
    
    if ($result -like "*Ok.*") {
        Write-Host ("Successfully set dynamic port range to {0}-{1}." -f $startPort, $endPort) -ForegroundColor Green
        Write-Host "Note: A system restart may be required for this change to take full effect." -ForegroundColor Yellow
    } else {
        Write-Host ("Failed to set dynamic port range: {0}" -f $result) -ForegroundColor Red
    }
}

function Restart-WinNAT {
    Write-Host "Restarting Windows NAT service..." -ForegroundColor Yellow
    
    try {
        net stop winnat
        Write-Host "Windows NAT service stopped." -ForegroundColor Green
        
        net start winnat
        Write-Host "Windows NAT service started." -ForegroundColor Green
        
        Write-Host "Windows NAT service restarted successfully." -ForegroundColor Green
        Write-Host "NOTE: If you use WSL2, Docker, or Hyper-V, you may need to restart them or reboot for full network connectivity to resume." -ForegroundColor Yellow
    } catch {
        Write-Host ("Failed to restart Windows NAT service: {0}" -f $_) -ForegroundColor Red
    }
}

function Full-Diagnostics {
    $startPort = Read-Host "Enter the starting port number for diagnostics"
    $numberOfPorts = Read-Host "Enter the number of ports to check"
    
    # Validate input
    if (-not ($startPort -match '^\d+$') -or -not ($numberOfPorts -match '^\d+$')) {
        Write-Host "Invalid input. Please enter valid numbers." -ForegroundColor Red
        return
    }
    
    $startPort = [int]$startPort
    $numberOfPorts = [int]$numberOfPorts
    $endPort = $startPort + $numberOfPorts - 1
    
    Write-Host ("Running full diagnostics for ports {0} to {1}..." -f $startPort, $endPort) -ForegroundColor Yellow
    
    # Show current status
    Write-Host ""
    Write-Host "Current system status:" -ForegroundColor Green
    Show-PortStatus
    
    # Check for port usage
    Write-Host "" # Uses the now-optimized function
    $usedPorts = Check-RangeUsage -startPort $startPort -numberOfPorts $numberOfPorts
    
    # Check for conflicts with dynamic port range
    Write-Host ""
    Write-Host "Checking for conflicts with dynamic port range..." -ForegroundColor Yellow
    $dynamicRange = netsh int ipv4 show dynamicportrange tcp
    
    # Parse dynamic range
    $lines = $dynamicRange -split "`n"
    $startLine = $lines | Where-Object { $_ -match "Start Port" }
    $numLine = $lines | Where-Object { $_ -match "Number of Ports" }
    
    if ($startLine -match ":\s*(\d+)") {
        $dynamicStart = [int]$matches[1]
    }
    
    if ($numLine -match ":\s*(\d+)") {
        $dynamicNum = [int]$matches[1]
        $dynamicEnd = $dynamicStart + $dynamicNum - 1
    }
    
    if ($dynamicStart -and $dynamicEnd) {
        Write-Host ("Dynamic port range: {0}-{1}" -f $dynamicStart, $dynamicEnd) -ForegroundColor Yellow
        
        if ($endPort -ge $dynamicStart -and $startPort -le $dynamicEnd) {
            Write-Host "CONFLICT: Your desired port range overlaps with the dynamic port range!" -ForegroundColor Red
            Write-Host "Consider changing the dynamic port range using option 5." -ForegroundColor Yellow
        } else {
            Write-Host "No conflict with dynamic port range." -ForegroundColor Green
        }
    }
    
    # Check for existing reserved ranges that overlap
    Write-Host ""
    Write-Host "Checking for conflicts with existing reserved port ranges..." -ForegroundColor Yellow
    $reservedRanges = netsh int ipv4 show excludedportrange protocol=tcp
    
    $lines = $reservedRanges -split '[\r\n]+'
    $conflictsFound = $false

    # Detect netsh output format by checking the header
    $headerLine = $lines | Where-Object { $_ -match "Start Port" }
    $isEndPortFormat = $headerLine -like "*End Port*"
    
    foreach ($line in $lines) {
        if ($line -match "^\s*(\d+)\s+(\d+)") {
            $rangeStart = [int]$matches[1]
            $secondValue = [int]$matches[2]

            # Calculate End Port based on detected format
            $rangeEnd = if ($isEndPortFormat) { $secondValue } else { $rangeStart + $secondValue - 1 }

            if ($endPort -ge $rangeStart -and $startPort -le $rangeEnd) {
                Write-Host ("CONFLICT: Your desired port range overlaps with existing reserved range {0}-{1}!" -f $rangeStart, $rangeEnd) -ForegroundColor Red
                $conflictsFound = $true
            }
        }
    }
    
    if (-not $conflictsFound) {
        Write-Host "No conflicts with existing reserved port ranges." -ForegroundColor Green
    }
    
    # Summary
    Write-Host ""
    Write-Host "DIAGNOSTICS SUMMARY:" -ForegroundColor Cyan
    if ($usedPorts) {
        Write-Host ("- Ports in use: {0}" -f $usedPorts.Count) -ForegroundColor Red
    } else { 
        Write-Host "- No ports in use" -ForegroundColor Green
    }
    
    if ($conflictsFound) {
        Write-Host "- Conflicts with reserved ranges found" -ForegroundColor Red
    } else {
        Write-Host "- No conflicts with reserved ranges" -ForegroundColor Green
    }
    
    if ($dynamicStart -and $dynamicEnd -and $endPort -ge $dynamicStart -and $startPort -le $dynamicEnd) {
        Write-Host "- Conflict with dynamic port range" -ForegroundColor Red
    } else {
        Write-Host "- No conflict with dynamic port range" -ForegroundColor Green
    }
    
    if (-not $usedPorts -and -not $conflictsFound -and ($dynamicStart -and $dynamicEnd -and $endPort -lt $dynamicStart -or $startPort -gt $dynamicEnd)) {
        Write-Host ""
        Write-Host "Your port range appears to be available for reservation!" -ForegroundColor Green
        $response = Read-Host "Would you like to reserve it now? (y/n)"
        if ($response -eq 'y' -or $response -eq 'Y') {
            $result = netsh int ipv4 add excludedportrange protocol=tcp startport=$startPort numberofports=$numberOfPorts 2>&1
            
            if ($result -like "*Ok.*") {
                Write-Host ("Successfully reserved ports {0} to {1}." -f $startPort, $endPort) -ForegroundColor Green
            } else {
                Write-Host ("Failed to reserve ports: {0}" -f $result) -ForegroundColor Red
            }
        }
    }
}

# Main script loop
do {
    Show-Header
    $choice = Show-Menu
    
    switch ($choice) {
        "1" { 
            Show-PortStatus
            Read-Host "`nPress Enter to continue"
        }
        "2" { 
            $port = Read-Host "Enter the port number to check"
            if ($port -match '^\d+$') {
                Check-PortUsage -port ([int]$port)
            } else {
                Write-Host "Invalid port number." -ForegroundColor Red
            }
            Read-Host "`nPress Enter to continue"
        }
        "3" { Reserve-PortRange; Read-Host "`nPress Enter to return to the menu" }
        "4" { Delete-ReservedRange; Read-Host "`nPress Enter to return to the menu" }
        "5" { Change-DynamicPortRange; Read-Host "`nPress Enter to return to the menu" }
        "6" { Restart-WinNAT; Read-Host "`nPress Enter to return to the menu" }
        "7" { Full-Diagnostics; Read-Host "`nPress Enter to return to the menu" }
        "8" { 
            Write-Host "Exiting..." -ForegroundColor Yellow
            break
        }
        default { 
            Write-Host "Invalid choice. Please try again." -ForegroundColor Red
            Start-Sleep -Seconds 2
        }
    }
} while ($choice -ne "8")