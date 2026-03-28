<#
.SYNOPSIS
Close process that occupies specified port

.DESCRIPTION
This script is used to find and close processes that occupy specified port, supports PowerShell 5 environment

.PARAMETER Port
Specify the port number to release

.EXAMPLE
Close-PortProcess -Port 8080
Close process that occupies port 8080

.NOTES
Author: Trae AI Assistant
Version: 1.0
Date: 2026-03-28
#>

param(
    [Parameter(Mandatory=$true, HelpMessage="Please specify the port number to release")]
    [ValidateRange(1, 65535)]
    [int]$Port
)

# Function: Get process PID that occupies the port
function Get-PortProcess {
    param(
        [int]$Port
    )
    
    try {
        # Method 1: Use PowerShell native cmdlet
        Write-Host "Querying port occupation using Get-NetTCPConnection..."
        $connections = Get-NetTCPConnection -LocalPort $Port -ErrorAction Stop
        if ($connections) {
            return $connections | Select-Object -ExpandProperty OwningProcess
        }
    }
    catch {
        Write-Host "Get-NetTCPConnection command failed, trying netstat command..."
    }
    
    try {
        # Method 2: Use netstat command
        Write-Host "Querying port occupation using netstat command..."
        $netstatOutput = netstat -ano | findstr ":$Port"
        if ($netstatOutput) {
            $pids = @()
            foreach ($line in $netstatOutput) {
                $parts = $line -split '\s+'
                if ($parts[-1] -match '^\d+$') {
                    $pids += [int]$parts[-1]
                }
            }
            return $pids | Select-Object -Unique
        }
    }
    catch {
        Write-Host "netstat command failed: $($_.Exception.Message)"
    }
    
    return $null
}

# Function: Get process information
function Get-ProcessInfo {
    param(
        [int]$ProcessId
    )
    
    try {
        $process = Get-Process -Id $ProcessId -ErrorAction Stop
        return $process
    }
    catch {
        Write-Host "Failed to get process information: $($_.Exception.Message)"
        return $null
    }
}

# Function: Close process
function Stop-TargetProcess {
    param(
        [int]$ProcessId
    )
    
    try {
        Write-Host "Closing process with PID: $ProcessId..."
        Stop-Process -Id $ProcessId -Force -ErrorAction Stop
        Write-Host "Process closed successfully!"
        return $true
    }
    catch {
        Write-Host "Failed to close process: $($_.Exception.Message)"
        Write-Host "Administrator privileges may be required."
        Write-Host "Please try running PowerShell as administrator and re-execute this script."
        return $false
    }
}

# Function: Verify port release
function Test-PortRelease {
    param(
        [int]$Port
    )
    
    # Wait 1 second to ensure process is completely closed
    Start-Sleep -Seconds 1
    
    $pids = Get-PortProcess -Port $Port
    if ($pids) {
        Write-Host "Port $Port is still occupied, occupying process PIDs: $($pids -join ', ')"
        return $false
    } else {
        Write-Host "Port $Port has been successfully released!"
        return $true
    }
}

# Main script logic
Write-Host "========================================"
Write-Host "Port Occupation Process Close Tool"
Write-Host "========================================"
Write-Host "Processing port: $Port"
Write-Host ""

# 1. Query processes occupying the port
$pids = Get-PortProcess -Port $Port

if (-not $pids) {
    Write-Host "Port $Port is not occupied, no process needs to be closed."
    exit 0
}

Write-Host "Found processes occupying port $Port with PIDs: $($pids -join ', ')"
Write-Host ""

# 2. Get and display process information
foreach ($processId in $pids) {
    $process = Get-ProcessInfo -ProcessId $processId
    if ($process) {
        Write-Host "Process information:"
        Write-Host "- Process name: $($process.ProcessName)"
        Write-Host "- PID: $($process.Id)"
        Write-Host "- Start time: $($process.StartTime)"
        if ($process.Path) {
            Write-Host "- Executable path: $($process.Path)"
        }
        Write-Host ""
    }
}

# 3. Confirm whether to close processes
$confirm = Read-Host "Are you sure you want to close these processes? (Y/N)"
if ($confirm -ne 'Y' -and $confirm -ne 'y') {
    Write-Host "Operation cancelled."
    exit 0
}

# 4. Close processes
$allSuccess = $true
foreach ($processId in $pids) {
    $success = Stop-TargetProcess -ProcessId $processId
    if (-not $success) {
        $allSuccess = $false
    }
}

# 5. Verify port release
if ($allSuccess) {
    Test-PortRelease -Port $Port
}

Write-Host ""
Write-Host "Operation completed!"
Write-Host "========================================"