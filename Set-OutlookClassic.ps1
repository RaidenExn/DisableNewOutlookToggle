<#
.SYNOPSIS
    Disables "New Outlook" toggle and enforces classic Outlook machine-wide.

.DESCRIPTION
    Applies registry changes to HKLM affecting all users, and forces close all Office apps to ensure policies take effect.

.NOTES
    Requires **administrative privileges** to modify HKLM and close apps.
#>

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

# Registry paths
$baseGeneral = "HKLM:\Software\Policies\Microsoft\Office\16.0\Outlook\Options\General"
$basePreferences = "HKLM:\Software\Policies\Microsoft\Office\16.0\Outlook\Preferences"

# Desired values
$desiredHideToggle = 1
$desiredUseNewOutlook = 0

# Check for Admin
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {
    Write-Warning "This script must be run as Administrator."
    exit 1
}

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
    Write-Host "✅ The machine-wide Outlook settings are already applied. No changes needed."
    exit 0
}

# Confirm before proceeding
$proceed = Read-Host "This script will close running Office apps and apply machine-wide Outlook settings. Continue? (Y/N)"
if ($proceed -notmatch '^[Yy]') {
    Write-Host "Operation cancelled by user."
    exit 0
}

# Close all running Office apps before applying settings
Close-OfficeApps

# Create and set registry keys
New-Item -Path $baseGeneral -Force | Out-Null
New-ItemProperty -Path $baseGeneral -Name "HideNewOutlookToggle" -Value $desiredHideToggle -PropertyType DWord -Force | Out-Null

New-Item -Path $basePreferences -Force | Out-Null
New-ItemProperty -Path $basePreferences -Name "UseNewOutlook" -Value $desiredUseNewOutlook -PropertyType DWord -Force | Out-Null

Write-Host "✅ Machine-wide classic Outlook enforced and New Outlook toggle hidden."
Write-Host "Please restart your computer or sign out and sign back in for the changes to fully take effect."
