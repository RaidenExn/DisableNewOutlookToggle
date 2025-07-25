# DisableNewOutlookToggle

This PowerShell script disables the new Outlook toggle feature in Microsoft Outlook by setting the required registry keys machine-wide. It also provides options to force close all Microsoft Office apps if needed.

---

## Features

- Applies the setting for all users on the machine (HKLM registry)
- Checks if the settings are already applied and notifies
- Force closes all Office apps
---

## Usage

### Running the script directly from the web

Run the following command in an elevated PowerShell window (Run as Administrator):

```powershell
irm https://github.com/RaidenExn/DisableNewOutlookToggle/raw/main/Set-OutlookClassic.ps1 | powershell -
