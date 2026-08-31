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

$script:WingetPath = $null

function Write-Info {
    param([Parameter(Mandatory = $true)][string]$Message)
    Write-Host $Message -ForegroundColor Cyan
}

function Write-Notice {
    param([Parameter(Mandatory = $true)][string]$Message)
    Write-Host $Message -ForegroundColor Yellow
}

function Write-Success {
    param([Parameter(Mandatory = $true)][string]$Message)
    Write-Host $Message -ForegroundColor Green
}

function Write-Dim {
    param([Parameter(Mandatory = $true)][string]$Message)
    Write-Host $Message -ForegroundColor DarkGray
}

function Write-CommandLine {
    param([Parameter(Mandatory = $true)][string]$Message)
    Write-Host $Message -ForegroundColor DarkCyan
}

function Write-Warn {
    param([Parameter(Mandatory = $true)][string]$Message)
    Write-Warning $Message
}

function Assert-Administrator {
    $CurrentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $Principal = New-Object Security.Principal.WindowsPrincipal($CurrentIdentity)

    if (-not $Principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw "This script must be run as Administrator. Right-click PowerShell/Terminal and choose 'Run as administrator'."
    }
}

function Resolve-WingetPath {
    $Command = Get-Command winget.exe -ErrorAction SilentlyContinue

    if ($Command) {
        return $Command.Source
    }

    $Candidates = @(
        "$env:LOCALAPPDATA\Microsoft\WindowsApps\winget.exe",
        "$env:ProgramFiles\WindowsApps\Microsoft.DesktopAppInstaller_*\winget.exe"
    )

    foreach ($Candidate in $Candidates) {
        if ($Candidate.Contains("*")) {
            $Resolved = Get-Item -Path $Candidate -ErrorAction SilentlyContinue |
                Sort-Object LastWriteTime -Descending |
                Select-Object -First 1

            if ($Resolved) {
                return $Resolved.FullName
            }
        }
        elseif (Test-Path $Candidate) {
            return $Candidate
        }
    }

    return $null
}

function Assert-OsSupportsWinget {
    $Build = [System.Environment]::OSVersion.Version.Build

    if ($Build -lt 17763) {
        throw "WinGet requires Windows build 17763 or newer. Current build: $Build"
    }
}

function Assert-AppxDeploymentAvailable {
    Get-Command Add-AppxPackage -ErrorAction Stop | Out-Null
}

function Enable-Tls12 {
    [Net.ServicePointManager]::SecurityProtocol =
        [Net.ServicePointManager]::SecurityProtocol -bor
        [Net.SecurityProtocolType]::Tls12
}

function Test-TransientNetworkError {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    return $Message -match "sending the request|timed out|NameResolutionFailure|temporar|connection|TLS|SSL|503|502|504"
}

function Invoke-WithRetry {
    param(
        [Parameter(Mandatory = $true)]
        [ScriptBlock]$Action,

        [Parameter(Mandatory = $true)]
        [string]$Description,

        [int]$MaxAttempts = 3,

        [int]$DelaySeconds = 2,

        [switch]$RetryTransientOnly
    )

    $LastError = $null

    for ($Attempt = 1; $Attempt -le $MaxAttempts; $Attempt++) {
        try {
            return & $Action
        }
        catch {
            $LastError = $_
            $Message = $_.Exception.Message
            $IsTransient = Test-TransientNetworkError -Message $Message

            if ($RetryTransientOnly -and -not $IsTransient) {
                throw
            }

            if ($Attempt -ge $MaxAttempts) {
                break
            }

            Write-Warn "$Description failed (attempt $Attempt of $MaxAttempts). Retrying in $DelaySeconds second(s). Details: $Message"
            Start-Sleep -Seconds $DelaySeconds
        }
    }

    throw "$Description failed after $MaxAttempts attempt(s). Last error: $($LastError.Exception.Message)"
}

function Add-WindowsAppsToPath {
    $WindowsApps = Join-Path $env:LOCALAPPDATA "Microsoft\WindowsApps"

    if ($env:Path -notlike "*$WindowsApps*") {
        $env:Path = "$env:Path;$WindowsApps"
    }
}

function Initialize-PSGalleryAndNuGet {
    Write-Info "Ensuring NuGet provider and PSGallery trust..."

    Invoke-WithRetry -Description "Install-PackageProvider NuGet" -RetryTransientOnly -Action {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force | Out-Null
    }

    if (-not (Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue)) {
        Register-PSRepository -Default -ErrorAction SilentlyContinue
    }

    try {
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
    }
    catch {
        Write-Warn "Could not set PSGallery as trusted. Continuing. Details: $($_.Exception.Message)"
    }
}

function Register-AppInstaller {
    Write-Info "Registering App Installer package..."

    try {
        Add-AppxPackage -RegisterByFamilyName -MainPackage "Microsoft.DesktopAppInstaller_8wekyb3d8bbwe"
        return
    }
    catch {
        Write-Notice "RegisterByFamilyName failed, trying AppxManifest registration. Details: $($_.Exception.Message)"
    }

    $AppInstaller = Get-AppxPackage -AllUsers -Name Microsoft.DesktopAppInstaller |
        Sort-Object Version -Descending |
        Select-Object -First 1

    if (-not $AppInstaller -or -not $AppInstaller.InstallLocation) {
        throw "App Installer package is not present for manifest registration."
    }

    Add-AppxPackage -DisableDevelopmentMode -Register (Join-Path $AppInstaller.InstallLocation "AppxManifest.xml")
}

function Initialize-WingetWithMicrosoftClient {
    Assert-OsSupportsWinget
    Assert-AppxDeploymentAvailable
    Enable-Tls12
    Initialize-PSGalleryAndNuGet

    Write-Info "Installing Microsoft.WinGet.Client module..."
    Invoke-WithRetry -Description "Install-Module Microsoft.WinGet.Client" -RetryTransientOnly -Action {
        Install-Module -Name Microsoft.WinGet.Client -Repository PSGallery -Force -AllowClobber
    }
    Import-Module Microsoft.WinGet.Client -Force

    if (-not (Get-Command Repair-WinGetPackageManager -ErrorAction SilentlyContinue)) {
        throw "Repair-WinGetPackageManager is not available after importing Microsoft.WinGet.Client."
    }

    Write-Info "Repairing/installing App Installer and WinGet using Microsoft cmdlet..."
    try {
        Invoke-WithRetry -Description "Repair-WinGetPackageManager -Latest" -RetryTransientOnly -Action {
            Repair-WinGetPackageManager -AllUsers -Force -Latest
        }
    }
    catch {
        Write-Notice "-Latest repair path failed; trying standard repair path..."
        Invoke-WithRetry -Description "Repair-WinGetPackageManager" -RetryTransientOnly -Action {
            Repair-WinGetPackageManager -AllUsers -Force
        }
    }

    Register-AppInstaller
}

function Initialize-WingetAvailability {
    $ResolvedPath = Resolve-WingetPath

    if ($ResolvedPath) {
        $script:WingetPath = $ResolvedPath
        return
    }

    Write-Notice "winget.exe is not currently available. Bootstrapping with Microsoft.WinGet.Client..."
    Initialize-WingetWithMicrosoftClient

    Add-WindowsAppsToPath

    $ResolvedPath = Resolve-WingetPath

    if (-not $ResolvedPath) {
        throw @"
Microsoft bootstrap completed, but winget.exe is still unavailable.

Try signing out/in (or reboot), then rerun this script.
"@
    }

    $script:WingetPath = $ResolvedPath
}

function Invoke-Winget {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Arguments,

        [switch]$AllowNoApplicableUpgrade
    )

    Write-CommandLine "winget $($Arguments -join ' ')"

    $OutputLines = & $script:WingetPath @Arguments 2>&1
    $ExitCode = $LASTEXITCODE

    if ($OutputLines) {
        foreach ($Line in $OutputLines) {
            Write-Host $Line
        }
    }

    if ($ExitCode -ne 0) {
        $OutputText = ($OutputLines | Out-String)

        if (
            $AllowNoApplicableUpgrade -and
            $OutputText -match "No available upgrade found|No newer package versions are available"
        ) {
            Write-Dim "No upgrade available. Continuing."
            return
        }

        throw "winget failed with exit code $($ExitCode): winget $($Arguments -join ' ')"
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
    Write-Info "Checking $Name..."

    $ListOutput = & $script:WingetPath list --id $Id --exact --accept-source-agreements 2>$null
    $IsInstalled = ($LASTEXITCODE -eq 0) -and ($ListOutput -match [regex]::Escape($Id))

    if ($IsInstalled) {
        Write-Notice "$Name appears to be installed. Attempting upgrade..."

        Invoke-Winget -AllowNoApplicableUpgrade -Arguments @(
            "upgrade",
            "--id", $Id,
            "--exact",
            "--source", "winget",
            "--silent",
            "--accept-package-agreements",
            "--accept-source-agreements"
        )
    }
    else {
        Write-Notice "$Name does not appear to be installed. Installing..."

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

Write-Success "=== Modern PowerShell + Inline Predictive IntelliSense Setup ==="

Initialize-WingetAvailability

Write-Success "Using winget at: $script:WingetPath"

Write-Host ""
Write-Info "Installing/upgrading stable PowerShell 7..."
Install-OrUpgrade-WingetPackage -Id "Microsoft.PowerShell" -Name "PowerShell 7"

if (-not $SkipWindowsTerminal) {
    Write-Host ""
    Write-Info "Installing/upgrading Windows Terminal..."

    try {
        Install-OrUpgrade-WingetPackage -Id "Microsoft.WindowsTerminal" -Name "Windows Terminal"
    }
    catch {
        Write-Warn "Windows Terminal install/upgrade failed. Continuing. Details: $($_.Exception.Message)"
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
Write-Success "Found PowerShell 7 at: $PwshPath"

$InnerScript = @'
$ErrorActionPreference = "Stop"

function Write-Info {
    param([Parameter(Mandatory = $true)][string]$Message)
    Write-Host $Message -ForegroundColor Cyan
}

function Write-Notice {
    param([Parameter(Mandatory = $true)][string]$Message)
    Write-Host $Message -ForegroundColor Yellow
}

function Write-Success {
    param([Parameter(Mandatory = $true)][string]$Message)
    Write-Host $Message -ForegroundColor Green
}

function Write-Warn {
    param([Parameter(Mandatory = $true)][string]$Message)
    Write-Warning $Message
}

function Test-HostResolvable {
    param(
        [Parameter(Mandatory = $true)]
        [string]$HostName
    )

    try {
        [void][System.Net.Dns]::GetHostAddresses($HostName)
        return $true
    }
    catch {
        return $false
    }
}

Write-Success "Running inside PowerShell version: $($PSVersionTable.PSVersion)"

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
    Write-Warn "Could not mark PSGallery as trusted. Continuing. Details: $($_.Exception.Message)"
}

# Install/update PSReadLine.
# Prefer PSResourceGet when available; fall back to PowerShellGet.
try {
    $GalleryHosts = @(
        "www.powershellgallery.com",
        "cdn.powershellgallery.com"
    )

    $CanResolveGallery = $true

    foreach ($HostName in $GalleryHosts) {
        if (-not (Test-HostResolvable -HostName $HostName)) {
            $CanResolveGallery = $false
            break
        }
    }

    if (-not $CanResolveGallery) {
        Write-Notice "Skipping online PSReadLine update because PowerShell Gallery hosts are not reachable. Using installed module."
    }
    else {
        if (Get-Command Install-PSResource -ErrorAction SilentlyContinue) {
            Write-Info "Installing/updating PSReadLine using Install-PSResource..."

            try {
                Install-PSResource -Name PSReadLine -Scope AllUsers -TrustRepository -Reinstall -Quiet
            }
            catch {
                Write-Notice "Install-PSResource with AllUsers failed. Retrying for current user. Details: $($_.Exception.Message)"
                Install-PSResource -Name PSReadLine -Scope CurrentUser -TrustRepository -Reinstall -Quiet
            }
        }
        elseif (Get-Command Install-Module -ErrorAction SilentlyContinue) {
            Write-Info "Installing/updating PSReadLine using Install-Module..."

            try {
                Install-Module -Name PSReadLine -Scope AllUsers -Force -AllowClobber -SkipPublisherCheck
            }
            catch {
                Write-Notice "Install-Module with AllUsers failed. Retrying for current user. Details: $($_.Exception.Message)"
                Install-Module -Name PSReadLine -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck
            }
        }
        else {
            Write-Warn "No supported module installer found. Using bundled PSReadLine if available."
        }
    }
}
catch {
    Write-Notice "PSReadLine online update was skipped/failed. Using installed PSReadLine module. Details: $($_.Exception.Message)"
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

Write-Success "Profile updated: $ProfilePath"

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
Write-Info "PSReadLine version:"
Get-Module PSReadLine -ListAvailable |
    Sort-Object Version -Descending |
    Select-Object -First 1 Name, Version, Path |
    Format-List

Write-Info "Current prediction settings:"
Get-PSReadLineOption |
    Select-Object PredictionSource, PredictionViewStyle, EditMode |
    Format-List
'@

$TempInnerScript = Join-Path $env:TEMP "Configure-PowerShell7-PSReadLine.ps1"
Set-Content -Path $TempInnerScript -Value $InnerScript -Encoding UTF8

Write-Host ""
Write-Info "Configuring PSReadLine and PowerShell 7 profile..."

& $PwshPath -NoLogo -NoProfile -ExecutionPolicy Bypass -File $TempInnerScript

Remove-Item -Path $TempInnerScript -Force -ErrorAction SilentlyContinue

Write-Host ""
Write-Success "=== Done ==="
Write-Success "Open a new Windows Terminal tab using PowerShell 7, or run: pwsh"
Write-Success "Inline predictive suggestions should appear after you build some command history."
Write-Success "RightArrow accepts inline suggestions. F2 toggles inline/list prediction view if supported."