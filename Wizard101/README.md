# powershell-scripts

A collection of useful PowerShell scripts.

## Scripts

### Create-Wizard101Shortcut.ps1

Creates a desktop shortcut for Wizard101 from the running WizardGraphicalClient.exe process.

**Usage:**
```powershell
irm https://raw.githubusercontent.com/SquashyHydra/powershell-scripts/main/Wizard101/Create-Wizard101Shortcut.ps1 | iex
```

Or with a custom game path:
```powershell
& ([ScriptBlock]::Create((irm https://raw.githubusercontent.com/SquashyHydra/powershell-scripts/main/Wizard101/Create-Wizard101Shortcut.ps1))) -GamePath "C:\Path\To\Wizard101.exe"
```