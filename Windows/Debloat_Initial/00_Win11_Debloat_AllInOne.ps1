[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [switch]$RemoveOneDrive,
    [switch]$DisableWidgets,
    [switch]$DisableCopilot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$script:PerUserHiveSummary = [ordered]@{
    Updated            = 0
    Skipped            = 0
    Failed             = 0
    ProfileLoaded      = 0
    ProfileLoadFailed  = 0
    DefaultHiveLockedSkipped = 0
    ProfileUnloaded    = 0
    ProfileUnloadFailed = 0
}
$script:PerUserHiveRoots = $null
$script:MountedHiveNames = @()
$script:PerUserHivesInitialized = $false

function Write-Step {
    param([string]$Message)
    Write-Output "`n=== $Message ==="
}

function Test-IsAdmin {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Invoke-RegExe {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string[]]$Arguments
    )

    $hasNativeErrorPref = Test-Path -Path "variable:PSNativeCommandUseErrorActionPreference"
    if ($hasNativeErrorPref) {
        $previousNativeErrorPref = $PSNativeCommandUseErrorActionPreference
        $PSNativeCommandUseErrorActionPreference = $false
    }

    try {
        $output = & reg.exe @Arguments 2>&1
        return [PSCustomObject]@{
            ExitCode = $LASTEXITCODE
            Output   = ($output | Out-String).Trim()
        }
    }
    finally {
        if ($hasNativeErrorPref) {
            $PSNativeCommandUseErrorActionPreference = $previousNativeErrorPref
        }
    }
}

function Set-DwordValue {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][int]$Value,
        [ref]$Outcome
    )

    $result = "Skipped"
    if ($PSCmdlet.ShouldProcess("$Path\\$Name", "Set DWORD value to $Value")) {
        if (-not (Test-Path -Path $Path)) {
            New-Item -Path $Path -Force | Out-Null
        }
        New-ItemProperty -Path $Path -Name $Name -PropertyType DWord -Value $Value -Force | Out-Null
        $result = "Updated"
    }

    if ($PSBoundParameters.ContainsKey('Outcome')) {
        $Outcome.Value = $result
    }
}

function Set-DwordValuePerUserHive {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)][string]$RelativePath,
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][int]$Value
    )

    if (-not $script:PerUserHivesInitialized) {
        $script:PerUserHiveRoots = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

        # Current session user hive.
        $null = $script:PerUserHiveRoots.Add("HKCU:")

        # All currently loaded user hives under HKU, excluding _Classes overlays.
        Get-ChildItem -Path "Registry::HKEY_USERS" -ErrorAction SilentlyContinue |
            Where-Object {
                $_.PSChildName -match "^S-1-(5-21|12-1)-" -and $_.PSChildName -notmatch "_Classes$"
            } |
            ForEach-Object {
                $null = $script:PerUserHiveRoots.Add("Registry::HKEY_USERS\$($_.PSChildName)")
            }

        # Load hives for existing local profiles not currently loaded.
        $profileListPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
        $profileKeys = Get-ChildItem -Path $profileListPath -ErrorAction SilentlyContinue
        foreach ($profileKey in $profileKeys) {
            $sid = $profileKey.PSChildName
            if ($sid -notmatch "^S-1-(5-21|12-1)-") {
                continue
            }

            $alreadyLoaded = $script:PerUserHiveRoots.Contains("Registry::HKEY_USERS\$sid")
            if ($alreadyLoaded) {
                continue
            }

            $profileInfo = Get-ItemProperty -Path $profileKey.PSPath -ErrorAction SilentlyContinue
            if (-not $profileInfo.ProfileImagePath) {
                continue
            }

            $expandedProfilePath = [Environment]::ExpandEnvironmentVariables($profileInfo.ProfileImagePath)
            if ($expandedProfilePath -match "\\defaultuser0$") {
                # OOBE/transient profile; commonly locked and not useful for persisted defaults.
                continue
            }

            $ntUserPath = Join-Path -Path $expandedProfilePath -ChildPath "NTUSER.DAT"
            if (-not (Test-Path -Path $ntUserPath)) {
                continue
            }

            $mountName = "TempUserHive_$([Guid]::NewGuid().ToString('N'))"
            if ($PSCmdlet.ShouldProcess("HKU\\$mountName", "Load hive from $ntUserPath")) {
                $regLoadResult = Invoke-RegExe -Arguments @("load", "HKU\$mountName", "$ntUserPath")
                if ($regLoadResult.ExitCode -eq 0) {
                    $script:MountedHiveNames += $mountName
                    $null = $script:PerUserHiveRoots.Add("Registry::HKEY_USERS\$mountName")
                    $script:PerUserHiveSummary.ProfileLoaded++
                }
                else {
                    $script:PerUserHiveSummary.ProfileLoadFailed++
                    Write-Warning "  Failed to load profile hive for SID $sid from $ntUserPath (reg.exe exit code: $($regLoadResult.ExitCode)). $($regLoadResult.Output)"
                }
            }
        }

        # Load Default User hive so future users inherit the values.
        $defaultNtUserPath = Join-Path -Path "$env:SystemDrive\Users\Default" -ChildPath "NTUSER.DAT"
        if (Test-Path -Path $defaultNtUserPath) {
            $defaultMountName = "TempDefaultHive_$([Guid]::NewGuid().ToString('N'))"
            if ($PSCmdlet.ShouldProcess("HKU\\$defaultMountName", "Load Default User hive from $defaultNtUserPath")) {
                $regDefaultLoadResult = Invoke-RegExe -Arguments @("load", "HKU\$defaultMountName", "$defaultNtUserPath")
                if ($regDefaultLoadResult.ExitCode -eq 0) {
                    $script:MountedHiveNames += $defaultMountName
                    $null = $script:PerUserHiveRoots.Add("Registry::HKEY_USERS\$defaultMountName")
                    $script:PerUserHiveSummary.ProfileLoaded++
                }
                else {
                    $regMessage = $regDefaultLoadResult.Output
                    if ($regMessage -match "being used by another process") {
                        $script:PerUserHiveSummary.DefaultHiveLockedSkipped++
                        Write-Output "  Default User hive is locked by another process; skipping future-user HKCU-style tweaks for this run."
                    }
                    else {
                        $script:PerUserHiveSummary.ProfileLoadFailed++
                        Write-Warning "  Failed to load Default User hive from $defaultNtUserPath (reg.exe exit code: $($regDefaultLoadResult.ExitCode)). $regMessage"
                    }
                }
            }
        }
        else {
            Write-Warning "  Default User hive not found at $defaultNtUserPath; future users may not inherit HKCU-style tweaks."
        }

        $script:PerUserHivesInitialized = $true
    }

    $updatedCount = 0
    $skippedCount = 0
    $failedCount = 0

    foreach ($root in $script:PerUserHiveRoots) {
        $targetPath = Join-Path -Path $root -ChildPath $RelativePath
        try {
            $writeOutcome = $null
            Set-DwordValue -Path $targetPath -Name $Name -Value $Value -Outcome ([ref]$writeOutcome)
            if ($writeOutcome -eq "Updated") {
                $updatedCount++
            }
            else {
                $skippedCount++
            }
        }
        catch {
            $failedCount++
            Write-Warning "  Failed to set $Name at $targetPath - $($_.Exception.GetType().Name): $($_.Exception.Message)"
        }
    }

    $script:PerUserHiveSummary.Updated += $updatedCount
    $script:PerUserHiveSummary.Skipped += $skippedCount
    $script:PerUserHiveSummary.Failed += $failedCount
}

function Complete-PerUserHiveWork {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()

    foreach ($mount in $script:MountedHiveNames) {
        if ($PSCmdlet.ShouldProcess("HKU\\$mount", "Unload temporary user hive")) {
            $regUnloadResult = Invoke-RegExe -Arguments @("unload", "HKU\$mount")
            if ($regUnloadResult.ExitCode -ne 0) {
                $script:PerUserHiveSummary.ProfileUnloadFailed++
                Write-Warning "  Failed to unload temporary hive HKU\\$mount (reg.exe exit code: $($regUnloadResult.ExitCode)). $($regUnloadResult.Output)"
            }
            else {
                $script:PerUserHiveSummary.ProfileUnloaded++
            }
        }
    }
}

function Remove-AppxAndProvisioned {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)][string[]]$Patterns,
        [string[]]$SkipPackageNames = @()
    )

    $seenInstalled = [System.Collections.Generic.HashSet[string]]::new()
    $seenProvisioned = [System.Collections.Generic.HashSet[string]]::new()

    Write-Step "Removing installed AppX packages for existing users"
    foreach ($pattern in $Patterns) {
        Write-Output "Pattern: $pattern"
        $installed = Get-AppxPackage -AllUsers -Name $pattern -ErrorAction SilentlyContinue
        foreach ($pkg in $installed) {
            if ($SkipPackageNames -contains $pkg.Name) {
                Write-Output "  Skipping protected package: $($pkg.Name)"
                continue
            }

            if (-not $seenInstalled.Add($pkg.PackageFullName)) {
                continue
            }

            try {
                if ($PSCmdlet.ShouldProcess($pkg.PackageFullName, "Remove-AppxPackage")) {
                    Remove-AppxPackage -Package $pkg.PackageFullName -AllUsers -ErrorAction Stop
                    Write-Output "  Removed installed: $($pkg.PackageFullName)"
                }
                else {
                    Write-Output "  Skipped installed removal (WhatIf/Confirm): $($pkg.PackageFullName)"
                }
            }
            catch {
                Write-Warning "  Failed installed removal: $($pkg.PackageFullName) - $($_.Exception.GetType().Name): $($_.Exception.Message)"
            }
        }
    }

    Write-Step "Removing provisioned AppX packages for future users"
    $allProvisioned = Get-AppxProvisionedPackage -Online
    foreach ($pattern in $Patterns) {
        Write-Output "Pattern: $pattern"
        $provisioned = $allProvisioned | Where-Object {
            $_.DisplayName -like $pattern -or $_.PackageName -like $pattern
        }

        foreach ($prov in $provisioned) {
            if ($SkipPackageNames -contains $prov.DisplayName) {
                Write-Output "  Skipping protected provisioned package: $($prov.DisplayName)"
                continue
            }

            if (-not $seenProvisioned.Add($prov.PackageName)) {
                continue
            }

            try {
                if ($PSCmdlet.ShouldProcess($prov.PackageName, "Remove-AppxProvisionedPackage")) {
                    Remove-AppxProvisionedPackage -Online -PackageName $prov.PackageName -ErrorAction Stop | Out-Null
                    Write-Output "  Deprovisioned: $($prov.PackageName)"
                }
                else {
                    Write-Output "  Skipped deprovision (WhatIf/Confirm): $($prov.PackageName)"
                }
            }
            catch {
                Write-Warning "  Failed deprovision: $($prov.PackageName) - $($_.Exception.GetType().Name): $($_.Exception.Message)"
            }
        }
    }
}

function Disable-ConsumerExperience {
    Write-Step "Disabling consumer content and app suggestions"

    Set-DwordValue -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent" -Name "DisableWindowsConsumerFeatures" -Value 1
    Set-DwordValue -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent" -Name "DisableCloudOptimizedContent" -Value 1
    Set-DwordValue -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent" -Name "DisableSoftLanding" -Value 1
    Set-DwordValue -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent" -Name "DisableTailoredExperiencesWithDiagnosticData" -Value 1
    Set-DwordValue -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AdvertisingInfo" -Name "DisabledByGroupPolicy" -Value 1

    $cdmPath = "Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"
    Set-DwordValuePerUserHive -RelativePath $cdmPath -Name "ContentDeliveryAllowed" -Value 0
    Set-DwordValuePerUserHive -RelativePath $cdmPath -Name "OemPreInstalledAppsEnabled" -Value 0
    Set-DwordValuePerUserHive -RelativePath $cdmPath -Name "PreInstalledAppsEnabled" -Value 0
    Set-DwordValuePerUserHive -RelativePath $cdmPath -Name "PreInstalledAppsEverEnabled" -Value 0
    Set-DwordValuePerUserHive -RelativePath $cdmPath -Name "SilentInstalledAppsEnabled" -Value 0
    Set-DwordValuePerUserHive -RelativePath $cdmPath -Name "RotatingLockScreenEnabled" -Value 0
    Set-DwordValuePerUserHive -RelativePath $cdmPath -Name "RotatingLockScreenOverlayEnabled" -Value 0
    Set-DwordValuePerUserHive -RelativePath $cdmPath -Name "SubscribedContent-310093Enabled" -Value 0
    Set-DwordValuePerUserHive -RelativePath $cdmPath -Name "SubscribedContent-338388Enabled" -Value 0
    Set-DwordValuePerUserHive -RelativePath $cdmPath -Name "SubscribedContent-338389Enabled" -Value 0
    Set-DwordValuePerUserHive -RelativePath $cdmPath -Name "SubscribedContent-353694Enabled" -Value 0
    Set-DwordValuePerUserHive -RelativePath $cdmPath -Name "SubscribedContent-353696Enabled" -Value 0
    Set-DwordValuePerUserHive -RelativePath $cdmPath -Name "SystemPaneSuggestionsEnabled" -Value 0
}

function Set-RecallPolicy {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()

    Write-Step "Disabling Recall availability and snapshot saving"
    if ($PSCmdlet.ShouldProcess("HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsAI", "Apply Recall hardening policy values")) {
        Set-DwordValue -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsAI" -Name "AllowRecallEnablement" -Value 0
        Set-DwordValue -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsAI" -Name "DisableAIDataAnalysis" -Value 1
    }
}

function Disable-GameBarAndDvr {
    Write-Step "Disabling Game Bar and Game DVR"
    Set-DwordValuePerUserHive -RelativePath "Software\Microsoft\Windows\CurrentVersion\GameDVR" -Name "AppCaptureEnabled" -Value 0
    Set-DwordValuePerUserHive -RelativePath "System\GameConfigStore" -Name "GameDVR_Enabled" -Value 0
    Set-DwordValue -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\GameDVR" -Name "AllowGameDVR" -Value 0
}

function Disable-WidgetsAndChat {
    Write-Step "Disabling Widgets and taskbar chat"
    Set-DwordValue -Path "HKLM:\SOFTWARE\Policies\Microsoft\Dsh" -Name "AllowNewsAndInterests" -Value 0
    Set-DwordValue -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Feeds" -Name "EnableFeeds" -Value 0
    Set-DwordValuePerUserHive -RelativePath "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "TaskbarDa" -Value 0
    Set-DwordValuePerUserHive -RelativePath "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "TaskbarMn" -Value 0
}

function Disable-Copilot {
    Write-Step "Disabling Windows Copilot"
    Set-DwordValuePerUserHive -RelativePath "Software\Policies\Microsoft\Windows\WindowsCopilot" -Name "TurnOffWindowsCopilot" -Value 1
    Set-DwordValue -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsCopilot" -Name "TurnOffWindowsCopilot" -Value 1
}

function Disable-GamingScheduledTask {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()

    Write-Step "Disabling Xbox/gaming scheduled tasks (if present)"

    $taskNames = @(
        "XblGameSaveTask",
        "XblGameSaveTaskLogon"
    )

    foreach ($taskName in $taskNames) {
        try {
            $task = Get-ScheduledTask -TaskName $taskName -ErrorAction Stop
            if ($PSCmdlet.ShouldProcess($task.TaskName, "Disable-ScheduledTask")) {
                Disable-ScheduledTask -InputObject $task | Out-Null
                Write-Output "  Disabled scheduled task: $($task.TaskPath)$($task.TaskName)"
            }
            else {
                Write-Output "  Skipped scheduled task disable (WhatIf/Confirm): $($task.TaskPath)$($task.TaskName)"
            }
        }
        catch {
            if ($_.Exception.Message -match "cannot find|not found|No MSFT_ScheduledTask objects") {
                Write-Output "  Task not found (skipped): $taskName"
            }
            else {
                Write-Warning "  Failed scheduled task operation for $taskName - $($_.Exception.GetType().Name): $($_.Exception.Message)"
            }
        }
    }
}

function Remove-OneDrive {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()

    Write-Step "Removing OneDrive"

    $candidates = @(
        "$env:SystemRoot\SysWOW64\OneDriveSetup.exe",
        "$env:SystemRoot\System32\OneDriveSetup.exe"
    )

    $setup = $candidates | Where-Object { Test-Path -Path $_ } | Select-Object -First 1
    if (-not $setup) {
        Write-Output "  OneDrive setup executable not found."
        return
    }

    if ($PSCmdlet.ShouldProcess("OneDrive", "Uninstall")) {
        Start-Process -FilePath $setup -ArgumentList "/uninstall" -Wait -NoNewWindow
        Write-Output "  OneDrive uninstall command executed."
    }
    else {
        Write-Output "  Skipped OneDrive uninstall (WhatIf/Confirm)."
    }
}

if (-not (Test-IsAdmin)) {
    throw "Run this script in an elevated PowerShell session (Run as Administrator)."
}

$appxPatterns = @(
    "*MicrosoftTeams*",
    "*MSTeams*",
    "*Microsoft.Office.Desktop*",
    "*Microsoft.MicrosoftOfficeHub*",
    "*Microsoft.Office.OneNote*",
    "*Microsoft.OutlookForWindows*",
    "*Microsoft.Todos*",
    "*Microsoft.Whiteboard*",
    "*Microsoft.PowerAutomateDesktop*",
    "*Microsoft.SkypeApp*",
    "*Microsoft.YourPhone*",
    "*Microsoft.WindowsCommunicationsApps*",
    "*Microsoft.GetHelp*",
    "*Microsoft.Getstarted*",
    "*Microsoft.People*",
    "*Microsoft.WindowsFeedbackHub*",
    "*Microsoft.BingNews*",
    "*Microsoft.BingWeather*",
    "*Microsoft.MicrosoftSolitaireCollection*",
    "*Microsoft.Clipchamp*",
    "*Clipchamp.Clipchamp*",
    "*Microsoft.ZuneMusic*",
    "*Microsoft.ZuneVideo*",
    "*Microsoft.GamingApp*",
    "*Microsoft.XboxApp*",
    "*Microsoft.Xbox.TCUI*",
    "*Microsoft.XboxGameOverlay*",
    "*Microsoft.XboxGamingOverlay*",
    "*Microsoft.XboxIdentityProvider*",
    "*Microsoft.XboxSpeechToTextOverlay*",
    "*Microsoft.549981C3F5F10*",
    "*Disney*",
    "*LinkedIn*",
    "*Netflix*",
    "*Spotify*",
    "*TikTok*",
    "*Facebook*",
    "*Instagram*",
    "*Twitter*"
)

$skipProtectedPackages = @(
    "Microsoft.XboxGameCallableUI",
    "Microsoft.DesktopAppInstaller",
    "Microsoft.StorePurchaseApp",
    "Microsoft.WindowsStore"
)

Write-Output "Starting Windows 11 all-in-one debloat routine..."

Disable-ConsumerExperience
Set-RecallPolicy
Disable-GameBarAndDvr
Disable-GamingScheduledTask

if ($DisableWidgets) {
    Disable-WidgetsAndChat
}

if ($DisableCopilot) {
    Disable-Copilot
}

Remove-AppxAndProvisioned -Patterns $appxPatterns -SkipPackageNames $skipProtectedPackages

if ($RemoveOneDrive) {
    Remove-OneDrive
}

Complete-PerUserHiveWork

Write-Step "Post-cleanup verification"
Write-Output "Installed app packages still matching major patterns:"
$installedMatches = Get-AppxPackage -AllUsers |
    Where-Object {
        $_.Name -match "Office|Microsoft365|OneNote|Outlook|Teams|MSTeams|Todo|Whiteboard|PowerAutomate|Skype|Xbox|GamingApp|Solitaire|Clipchamp|BingNews|BingWeather"
    }

$installedResidual = $installedMatches | Where-Object { $skipProtectedPackages -notcontains $_.Name }

if ($installedResidual) {
    $installedResidual |
        Select-Object Name, PackageFullName |
        Sort-Object Name
}
else {
    Write-Output "  None"
}

Write-Output "`nProvisioned app packages still matching major patterns:"
$provisionedMatches = Get-AppxProvisionedPackage -Online |
    Where-Object {
        $_.DisplayName -match "Office|Microsoft365|OneNote|Outlook|Teams|MSTeams|Todo|Whiteboard|PowerAutomate|Skype|Xbox|GamingApp|Solitaire|Clipchamp|BingNews|BingWeather"
    }

$provisionedResidual = $provisionedMatches | Where-Object { $skipProtectedPackages -notcontains $_.DisplayName }

if ($provisionedResidual) {
    $provisionedResidual |
        Select-Object DisplayName, PackageName |
        Sort-Object DisplayName
}
else {
    Write-Output "  None"
}

Write-Step "Per-user hive coverage summary"
Write-Output ("Updated: {0} | Skipped: {1} | Failed: {2}" -f $script:PerUserHiveSummary.Updated, $script:PerUserHiveSummary.Skipped, $script:PerUserHiveSummary.Failed)
Write-Output ("Hives loaded: {0} (load failed: {1}) | Hives unloaded: {2} (unload failed: {3})" -f $script:PerUserHiveSummary.ProfileLoaded, $script:PerUserHiveSummary.ProfileLoadFailed, $script:PerUserHiveSummary.ProfileUnloaded, $script:PerUserHiveSummary.ProfileUnloadFailed)
Write-Output ("Default hive locked skips: {0}" -f $script:PerUserHiveSummary.DefaultHiveLockedSkipped)

Write-Output "`nDebloat routine completed."
