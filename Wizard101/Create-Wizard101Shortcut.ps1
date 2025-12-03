<#
.SYNOPSIS
    Creates a desktop shortcut for Wizard101 from the running WizardGraphicalClient.exe process.

.DESCRIPTION
    This script finds the running WizardGraphicalClient.exe process, extracts its command line
    arguments, and creates a desktop shortcut with those same arguments. If the process is not
    running, it can optionally start Wizard101.exe and wait for the graphical client to appear.

.PARAMETER GamePath
    Optional path to Wizard101.exe to start the game if WizardGraphicalClient.exe is not running.

.EXAMPLE
    irm https://raw.githubusercontent.com/SquashyHydra/powershell-scripts/main/Create-Wizard101Shortcut.ps1 | iex

.EXAMPLE
    Create-Wizard101Shortcut.ps1 -GamePath "C:\ProgramData\KingsIsle Entertainment\Wizard101\Wizard101.exe"
#>

param(
    [string]$GamePath
)

# Try to find the running graphical client process
$proc = Get-CimInstance Win32_Process -Filter "Name='WizardGraphicalClient.exe'" | Select-Object -First 1

if (-not $proc) {
    # If GamePath was provided, use it; otherwise try to locate it
    $startPath = $GamePath
    
    if (-not $startPath) {
        # Common Wizard101 installation paths
        $commonPaths = @(
            "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\KingsIsle Entertainment\Wizard101\Play Wizard101.lnk",
            "C:\ProgramData\KingsIsle Entertainment\Wizard101\Wizard101.exe",
            "C:\Program Files (x86)\KingsIsle Entertainment\Wizard101\Wizard101.exe",
            "C:\Program Files\KingsIsle Entertainment\Wizard101\Wizard101.exe"
        )
        
        foreach ($path in $commonPaths) {
            if (Test-Path $path) {
                $startPath = $path
                break
            }
        }
    }
    
    if ($startPath) {
        # Check if Wizard101.exe is already running before launching
        $launcherRunning = Get-Process -Name "Wizard101" -ErrorAction SilentlyContinue
        
        if ($launcherRunning) {
            Write-Host "Wizard101.exe launcher is already running. Skipping launch." -ForegroundColor Yellow
        } else {
            Write-Host "Found starter executable: $startPath"
            try {
                Start-Process -FilePath $startPath
                Write-Host "Launched $startPath"
            } catch {
                Write-Host "Failed to start $startPath : $_" -ForegroundColor Red
            }
        }
    } else {
        Write-Host "Please start the application now and press Enter when it's running." -ForegroundColor Yellow
        Read-Host "Press Enter once you have started the application to continue."
    }
    
    Write-Host "Waiting for WizardGraphicalClient.exe to appear..."
    while (-not (Get-CimInstance Win32_Process -Filter "Name='WizardGraphicalClient.exe'")) {
        Start-Sleep -Seconds 2
    }
    
    # Re-fetch the process once it exists
    $proc = Get-CimInstance Win32_Process -Filter "Name='WizardGraphicalClient.exe'" | Select-Object -First 1
}

# If we have the process, extract its command line
if (-not $proc) {
    Write-Host "Failed to locate WizardGraphicalClient.exe process." -ForegroundColor Red
    exit 1
}

$cmd = $proc.CommandLine

# Split command line into executable and arguments (handles quoted path)
if ($cmd -match '^\s*"([^"]+)"\s*(.*)$') {
    $exe = $matches[1]
    $args = $matches[2]
} elseif ($cmd -match '^\s*([^ ]+)\s*(.*)$') {
    $exe = $matches[1]
    $args = $matches[2]
} else {
    $exe = $cmd
    $args = ''
}

# Create the shortcut
$shell = New-Object -ComObject WScript.Shell
$desktop = [Environment]::GetFolderPath("Desktop")
$linkPath = Join-Path $desktop "Wizard101.lnk"
$link = $shell.CreateShortcut($linkPath)
$link.TargetPath = $exe
$link.Arguments = $args
$link.WorkingDirectory = (Split-Path $exe)
$link.IconLocation = "$exe,0"
$link.Save()

Write-Host "Shortcut created at $linkPath" -ForegroundColor Green
