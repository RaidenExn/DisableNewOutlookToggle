<#
.SYNOPSIS
    Disables "New Outlook" toggle and enforces classic Outlook machine-wide.
.DESCRIPTION
    Applies registry changes to HKLM affecting all users, and forces close all Office apps to ensure policies take effect.
.NOTES
    Requires **administrative privileges** to modify HKLM and close apps.
#>

# Function to check if script is running as Administrator and re-launch it if not
function Ensure-Admin {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
    
    if (-not $isAdmin) {
        Write-Host "Script is not running as Administrator. Restarting with Administrator privileges..."
        Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
        exit
    }
}

# Function to clear the console screen
function Clear-Screen {
    Clear-Host
}

# Function to display a confirmation prompt
function Confirm-Prompt {
    param (
        [string]$Message
    )
    $Host.UI.RawUI.BackgroundColor = "Black"
    $Host.UI.RawUI.ForegroundColor = "White"
    Clear-Screen
    Write-Host "--------------------------------------------------" -ForegroundColor Cyan
    Write-Host $Message -ForegroundColor Green
    Write-Host "--------------------------------------------------" -ForegroundColor Cyan
    $confirm = Read-Host "Do you want to continue? (Y/N)"
    if ($confirm -match '^[Yy]') {
        return $true
    } else {
        Write-Host "Operation cancelled by user."
        exit
    }
}

# Function to close running Office apps gracefully, with confirmation before force kill
function Close-OfficeApps {
    $officeApps = @("OUTLOOK", "WINWORD", "EXCEL", "POWERPNT", "MSACCESS", "ONENOTE", "LYNC")
    foreach ($app in $officeApps) {
        $procs = Get-Process -Name $app -ErrorAction SilentlyContinue
        foreach ($proc in $procs) {
            Write-Host "Found running $($proc.ProcessName) (PID $($proc.Id))"
            $proc.CloseMainWindow() | Out-Null
            Start-Sleep -Seconds 3
            if (!$proc.HasExited) {
                $confirmKill = Read-Host "Process $($proc.ProcessName) (PID $($proc.Id)) is still running. Force close? (Y/N)"
                if ($confirmKill -match '^[Yy]') {
                    Write-Host "Force killing $($proc.ProcessName)..."
                    try {
                        $proc | Stop-Process -Force
                    } catch {
                        Write-Warning "Failed to force kill process $($proc.ProcessName): $_"
                    }
                } else {
                    Write-Host "Skipping force kill of $($proc.ProcessName)."
                }
            }
        }
    }
}

# Registry paths for machine-wide policy
$baseGeneral = "HKLM:\Software\Policies\Microsoft\Office\16.0\Outlook\Options\General"
$basePreferences = "HKLM:\Software\Policies\Microsoft\Office\16.0\Outlook\Preferences"

# Desired values for registry keys
$desiredHideToggle = 1
$desiredUseNewOutlook = 0

# Ensure the script is run as Administrator
Ensure-Admin

# Check if registry keys exist and values are set as desired
$hideToggleValue = $null
$useNewOutlookValue = $null

if (Test-Path $baseGeneral) {
    try {
        $hideToggleValue = (Get-ItemProperty -Path $baseGeneral -Name "HideNewOutlookToggle" -ErrorAction Stop).HideNewOutlookToggle
    } catch {}
}
if (Test-Path $basePreferences) {
    try {
        $useNewOutlookValue = (Get-ItemProperty -Path $basePreferences -Name "UseNewOutlook" -ErrorAction Stop).UseNewOutlook
    } catch {}
}

if ($hideToggleValue -eq $desiredHideToggle -and $useNewOutlookValue -eq $desiredUseNewOutlook) {
    Write-Host "Success: The machine-wide Outlook settings are already applied. No changes needed."
    exit 0
}

# Clear screen before running
Clear-Screen

# Confirmation prompt before proceeding
$confirmAction = Confirm-Prompt "This script will close running Office apps and apply machine-wide Outlook settings. Continue?"

if ($confirmAction) {
    # Close all running Office apps before applying settings
    Close-OfficeApps

    # Create and set registry keys
    New-Item -Path $baseGeneral -Force | Out-Null
    New-ItemProperty -Path $baseGeneral -Name "HideNewOutlookToggle" -Value $desiredHideToggle -PropertyType DWord -Force | Out-Null

    New-Item -Path $basePreferences -Force | Out-Null
    New-ItemProperty -Path $basePreferences -Name "UseNewOutlook" -Value $desiredUseNewOutlook -PropertyType DWord -Force | Out-Null

    Write-Host "Success: Machine-wide classic Outlook enforced and New Outlook toggle hidden."
    Write-Host "Please restart your computer or sign out and sign back in for the changes to fully take effect."
}

# Pause to ensure user can see the results before the script closes
Write-Host "Script execution completed. Press any key to exit..."
[System.Console]::ReadKey($true) | Out-Null
