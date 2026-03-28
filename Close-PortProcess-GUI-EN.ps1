<#
.SYNOPSIS
Close process that occupies specified port (GUI version)

.DESCRIPTION
This script provides a graphical interface for entering port number and closing processes that occupy the port

.NOTES
Author: Trae AI Assistant
Version: 1.0
Date: 2026-03-28
#>

# Load Windows Forms
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Port Occupation Process Close Tool"
$form.Size = New-Object System.Drawing.Size(400, 200)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.MinimizeBox = $false

# Create label
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(20, 30)
$label.Size = New-Object System.Drawing.Size(100, 20)
$label.Text = "Enter port number:"
$form.Controls.Add($label)

# Create text box
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(120, 30)
$textBox.Size = New-Object System.Drawing.Size(150, 20)
$textBox.Text = "8080"  # Default port
$form.Controls.Add($textBox)

# Create error label
$errorLabel = New-Object System.Windows.Forms.Label
$errorLabel.Location = New-Object System.Drawing.Point(20, 60)
$errorLabel.Size = New-Object System.Drawing.Size(350, 20)
$errorLabel.ForeColor = [System.Drawing.Color]::Red
$errorLabel.Text = ""
$form.Controls.Add($errorLabel)

# Create OK button
$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(100, 100)
$okButton.Size = New-Object System.Drawing.Size(80, 30)
$okButton.Text = "OK"
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

# Create Cancel button
$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(200, 100)
$cancelButton.Size = New-Object System.Drawing.Size(80, 30)
$cancelButton.Text = "Cancel"
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

# Show form and get result
$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
    $port = $textBox.Text
    
    # Validate port number
    if (-not ($port -match '^\d+$' -and [int]$port -ge 1 -and [int]$port -le 65535)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a valid port number (1-65535)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        exit 1
    }
    
    $port = [int]$port
    
    # Function: Get process PID that occupies the port
    function Get-PortProcess {
        param(
            [int]$Port
        )
        
        try {
            # Method 1: Use PowerShell native cmdlet
            $connections = Get-NetTCPConnection -LocalPort $Port -ErrorAction Stop
            if ($connections) {
                return $connections | Select-Object -ExpandProperty OwningProcess
            }
        }
        catch {
            # Method 2: Use netstat command
            try {
                $netstatOutput = netstat -ano | findstr ":${Port}"
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
                return $null
            }
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
            return $null
        }
    }
    
    # Function: Close process
    function Stop-TargetProcess {
        param(
            [int]$ProcessId
        )
        
        try {
            Stop-Process -Id $ProcessId -Force -ErrorAction Stop
            return $true
        }
        catch {
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
            return $false
        } else {
            return $true
        }
    }
    
    # Main logic
    $pids = Get-PortProcess -Port $port
    
    if (-not $pids) {
        [System.Windows.Forms.MessageBox]::Show("Port $port is not occupied, no process needs to be closed.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        exit 0
    }
    
    # Display process information
    $processInfo = @()
    foreach ($processId in $pids) {
        $process = Get-ProcessInfo -Id $processId
        if ($process) {
            $processInfo += "Process name: $($process.ProcessName)`nPID: $($process.Id)`nStart time: $($process.StartTime)`nExecutable path: $($process.Path)`n"
        }
    }
    
    $processInfoText = $processInfo -join "`n"
    $result = [System.Windows.Forms.MessageBox]::Show("Found processes occupying port ${port}:`n`n$processInfoText`n`nAre you sure you want to close these processes?", "Confirmation", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        $allSuccess = $true
        foreach ($processId in $pids) {
            $success = Stop-TargetProcess -ProcessId $processId
            if (-not $success) {
                $allSuccess = $false
            }
        }
        
        if ($allSuccess) {
            $released = Test-PortRelease -Port $port
            if ($released) {
                [System.Windows.Forms.MessageBox]::Show("Port $port has been successfully released!", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } else {
                [System.Windows.Forms.MessageBox]::Show("Port $port is still occupied, please check if administrator privileges are required.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("Failed to close process, administrator privileges may be required. Please try running PowerShell as administrator and re-execute this script.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
}
