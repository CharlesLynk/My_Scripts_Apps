<#
.SYNOPSIS
    Install/upgrade modern PowerShell and enable inline PSReadLine Predictive IntelliSense.

.DESCRIPTION
    - Installs/upgrades stable PowerShell 7 using winget.
    - Optionally installs/upgrades Windows Terminal.
    - Updates PSReadLine inside PowerShell 7.
    - Creates/updates the PowerShell 7 profile.
    - Enables inline predictive text by default.
    - Keeps Windows PowerShell 5.1 intact for compatibility.

.NOTES
    Run from elevated Windows PowerShell or elevated Windows Terminal.
    Compatible with Windows PowerShell 5.1 as the launcher.
#>

param(
    [switch]$SkipWindowsTerminal
)

$ErrorActionPreference = "Stop"

function Assert-Administrator {
    $CurrentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $Principal = New-Object Security.Principal.WindowsPrincipal($CurrentIdentity)

    if (-not $Principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw "This script must be run as Administrator. Right-click PowerShell/Terminal and choose 'Run as administrator'."
    }
}

function Test-CommandExists {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Command
    )

    return $null -ne (Get-Command $Command -ErrorAction SilentlyContinue)
}

function Invoke-Winget {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Arguments
    )

    Write-Host "winget $($Arguments -join ' ')" -ForegroundColor DarkCyan

    $Process = Start-Process `
        -FilePath "winget.exe" `
        -ArgumentList $Arguments `
        -Wait `
        -PassThru `
        -NoNewWindow

    if ($Process.ExitCode -ne 0) {
        throw "winget failed with exit code $($Process.ExitCode): winget $($Arguments -join ' ')"
    }
}

function Install-OrUpgrade-WingetPackage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Id,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    Write-Host ""
    Write-Host "Checking $Name..." -ForegroundColor Cyan

    $ListOutput = & winget list --id $Id --exact --accept-source-agreements 2>$null
    $IsInstalled = ($LASTEXITCODE -eq 0) -and ($ListOutput -match [regex]::Escape($Id))

    if ($IsInstalled) {
        Write-Host "$Name appears to be installed. Attempting upgrade..." -ForegroundColor Yellow

        try {
            Invoke-Winget -Arguments @(
                "upgrade",
                "--id", $Id,
                "--exact",
                "--source", "winget",
                "--silent",
                "--accept-package-agreements",
                "--accept-source-agreements"
            )
        }
        catch {
            Write-Warning "Upgrade failed or no upgrade was available for $Name. Continuing. Details: $($_.Exception.Message)"
        }
    }
    else {
        Write-Host "$Name does not appear to be installed. Installing..." -ForegroundColor Yellow

        Invoke-Winget -Arguments @(
            "install",
            "--id", $Id,
            "--exact",
            "--source", "winget",
            "--silent",
            "--accept-package-agreements",
            "--accept-source-agreements"
        )
    }
}

function Find-Pwsh {
    $Candidates = @(
        "$env:ProgramFiles\PowerShell\7\pwsh.exe",
        "$env:LOCALAPPDATA\Microsoft\WindowsApps\pwsh.exe"
    )

    foreach ($Candidate in $Candidates) {
        if (Test-Path $Candidate) {
            return $Candidate
        }
    }

    $Command = Get-Command pwsh.exe -ErrorAction SilentlyContinue

    if ($Command) {
        return $Command.Source
    }

    return $null
}

Assert-Administrator

Write-Host "=== Modern PowerShell + Inline Predictive IntelliSense Setup ===" -ForegroundColor Green

if (-not (Test-CommandExists -Command "winget.exe")) {
    throw @"
winget.exe was not found.

On Windows 11, winget normally comes from 'App Installer'.
Install/update 'App Installer' from Microsoft Store, then rerun this script.

Alternative manual PowerShell 7 installer:
https://github.com/PowerShell/PowerShell/releases
"@
}

Write-Host ""
Write-Host "Installing/upgrading stable PowerShell 7..." -ForegroundColor Cyan
Install-OrUpgrade-WingetPackage -Id "Microsoft.PowerShell" -Name "PowerShell 7"

if (-not $SkipWindowsTerminal) {
    Write-Host ""
    Write-Host "Installing/upgrading Windows Terminal..." -ForegroundColor Cyan

    try {
        Install-OrUpgrade-WingetPackage -Id "Microsoft.WindowsTerminal" -Name "Windows Terminal"
    }
    catch {
        Write-Warning "Windows Terminal install/upgrade failed. Continuing. Details: $($_.Exception.Message)"
    }
}

$PwshPath = Find-Pwsh

if (-not $PwshPath) {
    Start-Sleep -Seconds 3
    $PwshPath = Find-Pwsh
}

if (-not $PwshPath) {
    throw "PowerShell 7 was installed/upgraded, but pwsh.exe could not be found. Reboot or open a new terminal and check PATH."
}

Write-Host ""
Write-Host "Found PowerShell 7 at: $PwshPath" -ForegroundColor Green

$InnerScript = @'
$ErrorActionPreference = "Stop"

Write-Host "Running inside PowerShell version: $($PSVersionTable.PSVersion)" -ForegroundColor Green

# Make PSGallery usable without interactive prompts where possible.
try {
    if (Get-Command Set-PSRepository -ErrorAction SilentlyContinue) {
        $Repo = Get-PSRepository -Name "PSGallery" -ErrorAction SilentlyContinue

        if ($Repo -and $Repo.InstallationPolicy -ne "Trusted") {
            Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
        }
    }
}
catch {
    Write-Warning "Could not mark PSGallery as trusted. Continuing. Details: $($_.Exception.Message)"
}

# Install/update PSReadLine.
# Prefer PSResourceGet when available; fall back to PowerShellGet.
try {
    if (Get-Command Install-PSResource -ErrorAction SilentlyContinue) {
        Write-Host "Installing/updating PSReadLine using Install-PSResource..." -ForegroundColor Cyan

        try {
            Install-PSResource -Name PSReadLine -Scope AllUsers -TrustRepository -Reinstall -Quiet
        }
        catch {
            Write-Warning "Install-PSResource with AllUsers failed. Retrying for current user. Details: $($_.Exception.Message)"
            Install-PSResource -Name PSReadLine -Scope CurrentUser -TrustRepository -Reinstall -Quiet
        }
    }
    elseif (Get-Command Install-Module -ErrorAction SilentlyContinue) {
        Write-Host "Installing/updating PSReadLine using Install-Module..." -ForegroundColor Cyan

        try {
            Install-Module -Name PSReadLine -Scope AllUsers -Force -AllowClobber -SkipPublisherCheck
        }
        catch {
            Write-Warning "Install-Module with AllUsers failed. Retrying for current user. Details: $($_.Exception.Message)"
            Install-Module -Name PSReadLine -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck
        }
    }
    else {
        Write-Warning "No supported module installer found. Using bundled PSReadLine if available."
    }
}
catch {
    Write-Warning "PSReadLine update failed. PowerShell 7 usually includes PSReadLine already. Details: $($_.Exception.Message)"
}

# Create/update PowerShell 7 current-user profile.
$ProfilePath = $PROFILE.CurrentUserAllHosts
$ProfileDir = Split-Path -Parent $ProfilePath

if (-not (Test-Path $ProfileDir)) {
    New-Item -ItemType Directory -Path $ProfileDir -Force | Out-Null
}

$ManagedBlockStart = "# >>> Chase PSReadLine Predictive IntelliSense >>>"
$ManagedBlockEnd   = "# <<< Chase PSReadLine Predictive IntelliSense <<<"

# This is intentionally a double-quoted here-string because the outer script is a single-quoted here-string.
$ManagedBlock = @"
# >>> Chase PSReadLine Predictive IntelliSense >>>
# Modern PowerShell command-line editing and inline predictive text.
# RightArrow accepts inline suggestions.
# F2 toggles prediction view style in supported PSReadLine versions.

Import-Module PSReadLine -ErrorAction SilentlyContinue

if (Get-Command Set-PSReadLineOption -ErrorAction SilentlyContinue) {
    try {
        Set-PSReadLineOption -PredictionSource HistoryAndPlugin
    }
    catch {
        Set-PSReadLineOption -PredictionSource History
    }

    try {
        Set-PSReadLineOption -PredictionViewStyle InlineView
    }
    catch {
        # Older PSReadLine versions may not support this option.
    }

    try {
        Set-PSReadLineOption -EditMode Windows
        Set-PSReadLineOption -BellStyle None
    }
    catch {
        # Ignore unsupported options.
    }
}

if (Get-Command Set-PSReadLineKeyHandler -ErrorAction SilentlyContinue) {
    try {
        Set-PSReadLineKeyHandler -Key Tab -Function MenuComplete
        Set-PSReadLineKeyHandler -Key Ctrl+Spacebar -Function MenuComplete
        Set-PSReadLineKeyHandler -Key Ctrl+r -Function ReverseSearchHistory
    }
    catch {
        # Ignore unsupported bindings.
    }

    try {
        Set-PSReadLineKeyHandler -Key Ctrl+f -Function ForwardWord
    }
    catch {
        # Ignore if unsupported.
    }
}
# <<< Chase PSReadLine Predictive IntelliSense <<<
"@

if (Test-Path $ProfilePath) {
    $Existing = Get-Content -Path $ProfilePath -Raw

    $Pattern = [regex]::Escape($ManagedBlockStart) + "(?s).*?" + [regex]::Escape($ManagedBlockEnd)

    if ($Existing -match $Pattern) {
        $Updated = [regex]::Replace($Existing, $Pattern, $ManagedBlock)
    }
    else {
        $Updated = $Existing.TrimEnd() + [Environment]::NewLine + [Environment]::NewLine + $ManagedBlock + [Environment]::NewLine
    }

    Set-Content -Path $ProfilePath -Value $Updated -Encoding UTF8
}
else {
    Set-Content -Path $ProfilePath -Value ($ManagedBlock + [Environment]::NewLine) -Encoding UTF8
}

Write-Host "Profile updated: $ProfilePath" -ForegroundColor Green

# Validate settings in this session.
Import-Module PSReadLine -ErrorAction SilentlyContinue

try {
    Set-PSReadLineOption -PredictionSource HistoryAndPlugin
}
catch {
    try {
        Set-PSReadLineOption -PredictionSource History
    }
    catch {}
}

try {
    Set-PSReadLineOption -PredictionViewStyle InlineView
}
catch {}

try {
    Set-PSReadLineOption -EditMode Windows
    Set-PSReadLineOption -BellStyle None
}
catch {}

Write-Host ""
Write-Host "PSReadLine version:" -ForegroundColor Cyan
Get-Module PSReadLine -ListAvailable |
    Sort-Object Version -Descending |
    Select-Object -First 1 Name, Version, Path |
    Format-List

Write-Host "Current prediction settings:" -ForegroundColor Cyan
Get-PSReadLineOption |
    Select-Object PredictionSource, PredictionViewStyle, EditMode |
    Format-List
'@

$TempInnerScript = Join-Path $env:TEMP "Configure-PowerShell7-PSReadLine.ps1"
Set-Content -Path $TempInnerScript -Value $InnerScript -Encoding UTF8

Write-Host ""
Write-Host "Configuring PSReadLine and PowerShell 7 profile..." -ForegroundColor Cyan

& $PwshPath -NoLogo -NoProfile -ExecutionPolicy Bypass -File $TempInnerScript

Remove-Item -Path $TempInnerScript -Force -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "=== Done ===" -ForegroundColor Green
Write-Host "Open a new Windows Terminal tab using PowerShell 7, or run: pwsh" -ForegroundColor Green
Write-Host "Inline predictive suggestions should appear after you build some command history." -ForegroundColor Green
Write-Host "RightArrow accepts inline suggestions. F2 toggles inline/list prediction view if supported." -ForegroundColor Green