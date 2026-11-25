[CmdletBinding()]
param(
    # UNC path, not P:, because mapped drives are not visible to elevated (admin) PowerShell windows.
    [string]$KeyFile = '\\FS01\Public\IT\Internal Software\AutoDIS\MbrohDrawingIndexCreator.Package_TemporaryKey.pfx',
    [switch]$SkipInstall,
    [switch]$Diagnostics,
    [switch]$NoDiagnostics,
    # Sign as yourself with a personal certificate, even if the Mbroh key is available.
    [switch]$SelfSign
)

$ErrorActionPreference = 'Stop'
$projectDir = Split-Path -Parent $PSScriptRoot
$project = Join-Path $projectDir 'MbrohDrawingIndexCreator.csproj'
$outDir = Join-Path $env:TEMP 'autodis_msix_signed'
$packageName = 'd2bbdbd4-3016-443a-9f1f-e2344ac6a034'
$markerRelative = 'AutoDIS\diagnostics.on'
$personalCerPath = Join-Path $env:LOCALAPPDATA 'AutoDIS\personal-signing.cer'

function Set-DiagnosticsMarker {
    param([string[]]$Paths, [bool]$On)
    foreach ($path in $Paths) {
        if ($On) {
            New-Item -ItemType Directory -Path (Split-Path -Parent $path) -Force | Out-Null
            New-Item -ItemType File -Path $path -Force | Out-Null
            Write-Host "Diagnostics ON: $path" -ForegroundColor Yellow
        }
        elseif (Test-Path -LiteralPath $path) {
            Remove-Item -LiteralPath $path -Force
            Write-Host "Diagnostics OFF: $path" -ForegroundColor Yellow
        }
    }
}

if ($NoDiagnostics -or $Diagnostics) {
    $on = [bool]$Diagnostics
    Set-DiagnosticsMarker -Paths @((Join-Path $env:TEMP $markerRelative)) -On $on
}

function Get-SigningMode {
    if ($SelfSign) { return 'Personal' }
    if (-not (Test-Path -LiteralPath $KeyFile)) {
        $reason = "not found at $KeyFile (check that \\FS01\Public is reachable)"
    }
    else {
        $key = Get-PfxCertificate -FilePath $KeyFile
        if ($key.NotAfter -ge (Get-Date)) { return 'Mbroh' }
        $reason = "expired on $($key.NotAfter.ToString('d'))"
    }
    $answer = Read-Host "The Mbroh signing key is $reason. Sign as yourself with a personal certificate instead? (y/N)"
    if ($answer -match '^[Yy]') { return 'Personal' }
    throw "Stopped: the Mbroh signing key is $reason."
}

function Get-PersonalCertificate {
    $subject = "CN=$env:USERNAME"
    $friendlyName = 'AutoDIS personal signing'
    $cert = Get-ChildItem -Path Cert:\CurrentUser\My | Where-Object {
        $_.Subject -eq $subject -and $_.FriendlyName -eq $friendlyName -and $_.NotAfter -gt (Get-Date)
    } | Select-Object -First 1
    if (-not $cert) {
        Write-Host "Creating personal signing certificate $subject (valid 3 years)..." -ForegroundColor Cyan
        $cert = New-SelfSignedCertificate -Type Custom -Subject $subject -KeyUsage DigitalSignature `
            -KeyAlgorithm RSA -KeyLength 2048 -HashAlgorithm SHA256 -CertStoreLocation Cert:\CurrentUser\My `
            -NotAfter (Get-Date).AddYears(3) -FriendlyName $friendlyName `
            -TextExtension @('2.5.29.37={text}1.3.6.1.5.5.7.3.3')
    }
    # Public certificate only. Windows needs it trusted machine-wide, which takes one admin import.
    New-Item -ItemType Directory -Path (Split-Path -Parent $personalCerPath) -Force | Out-Null
    Export-Certificate -Cert $cert -FilePath $personalCerPath | Out-Null
    return $cert
}

function Invoke-PackageBuild {
    param([string[]]$SigningArgs)
    Push-Location $projectDir
    try {
        dotnet build $project -c Release -p:Platform=x64 -p:EnableMsixTooling=true -p:GenerateAppxPackageOnBuild=true -p:AppxBundle=Never -p:UapAppxPackageBuildMode=SideloadOnly "-p:AppxPackageDir=$outDir\" @SigningArgs
        if ($LASTEXITCODE -ne 0) {
            throw "Build failed (exit code $LASTEXITCODE)."
        }
    }
    finally {
        Pop-Location
    }
}

$signingMode = Get-SigningMode

if (Test-Path -LiteralPath $outDir) {
    Remove-Item -LiteralPath $outDir -Recurse -Force
}

Write-Host "Building and signing package ($signingMode certificate)..." -ForegroundColor Cyan
if ($signingMode -eq 'Personal') {
    $signingCert = Get-PersonalCertificate
    $manifestPath = Join-Path $projectDir 'Package.appxmanifest'
    $manifestBackup = Join-Path $env:TEMP 'autodis_manifest_backup.appxmanifest'
    Copy-Item -LiteralPath $manifestPath -Destination $manifestBackup -Force
    try {
        # Publisher must match the certificate subject. The original manifest is restored in finally.
        $manifestText = [IO.File]::ReadAllText($manifestPath)
        $manifestText = [regex]::Replace($manifestText, 'Publisher="[^"]*"', "Publisher=`"$($signingCert.Subject)`"", 1)
        [IO.File]::WriteAllText($manifestPath, $manifestText, (New-Object System.Text.UTF8Encoding($false)))
        Invoke-PackageBuild -SigningArgs @('-p:AppxPackageSigningEnabled=false')
    }
    finally {
        Copy-Item -LiteralPath $manifestBackup -Destination $manifestPath -Force
        Remove-Item -LiteralPath $manifestBackup -Force
    }
}
else {
    $signingCert = Get-PfxCertificate -FilePath $KeyFile
    Invoke-PackageBuild -SigningArgs @('-p:AppxPackageSigningEnabled=true', "-p:PackageCertificateKeyFile=$KeyFile")
}

$msix = Get-ChildItem -Path $outDir -Recurse -Filter 'MbrohDrawingIndexCreator_*_x64.msix' | Select-Object -First 1
if (-not $msix) {
    throw 'Build finished but no .msix package was found.'
}

# The OCR fallback is silently useless if the models or native Tesseract libraries are not in the package.
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [IO.Compression.ZipFile]::OpenRead($msix.FullName)
try {
    $entries = $zip.Entries | ForEach-Object { $_.FullName -replace '\\', '/' }
    $required = 'tessdata/eng.traineddata', 'tessdata/eng_best.traineddata', 'x64/tesseract41.dll', 'x64/leptonica-1.80.0.dll', 'runtimes/win-x64/native/pdfium.dll'
    $missing = $required | Where-Object { $entries -notcontains $_ }
    if ($missing) {
        throw "The package is missing OCR files, so OCR would not work: $($missing -join ', ')"
    }
}
finally {
    $zip.Dispose()
}
Write-Host 'OCR models and native libraries are in the package.' -ForegroundColor Green

if ($signingMode -eq 'Personal') {
    $nugetRoot = dotnet msbuild $project -getProperty:NuGetPackageRoot
    if ($LASTEXITCODE -ne 0 -or -not $nugetRoot) {
        throw 'Could not determine the restored NuGet packages directory.'
    }
    $buildTools = Join-Path $nugetRoot 'microsoft.windows.sdk.buildtools'
    $signtool = Get-ChildItem -LiteralPath $buildTools -Recurse -Filter signtool.exe -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -like '*\x64\*' } | Sort-Object FullName -Descending | Select-Object -First 1
    if (-not $signtool) {
        throw "signtool.exe not found in '$buildTools'."
    }
    & $signtool.FullName sign /fd SHA256 /sha1 $signingCert.Thumbprint $msix.FullName
    if ($LASTEXITCODE -ne 0) {
        throw "signtool failed (exit code $LASTEXITCODE)."
    }
}
Write-Host "Built: $($msix.FullName)" -ForegroundColor Green

if ($SkipInstall) {
    return
}

# Checked before removing the old app, so a trust problem can't leave nothing installed.
$trustStores = 'LocalMachine\TrustedPeople', 'CurrentUser\TrustedPeople', 'LocalMachine\Root', 'CurrentUser\Root'
$trusted = foreach ($store in $trustStores) {
    Get-ChildItem -Path "Cert:\$store" | Where-Object { $_.Thumbprint -eq $signingCert.Thumbprint }
}
if (-not $trusted) {
    $adminCerPath = if ($signingMode -eq 'Personal') {
        $personalCerPath
    }
    else {
        Join-Path $projectDir 'reference\MbrohDrawingIndexCreator.Package_1.3.4.0_Test\MbrohDrawingIndexCreator.Package_1.3.4.0_x64.cer'
    }
    throw @"
The signing certificate ($($signingCert.Subject), thumbprint $($signingCert.Thumbprint)) is not trusted on this machine, so the install would fail.
Run this once in an ADMIN PowerShell window, then run the script again:
Import-Certificate -FilePath "$adminCerPath" -CertStoreLocation Cert:\LocalMachine\TrustedPeople
"@
}

Write-Host 'Removing previous install (close the app first if it is running)...' -ForegroundColor Cyan
Get-AppxPackage -Name $packageName -ErrorAction SilentlyContinue | Remove-AppxPackage

Write-Host 'Installing new package...' -ForegroundColor Cyan
Add-AppxPackage -Path $msix.FullName

Write-Host 'Installed. Launch "Mbroh Drawing Index Creator" from the Start menu.' -ForegroundColor Green

if ($Diagnostics -or $NoDiagnostics) {
    $installed = Get-AppxPackage -Name $packageName | Select-Object -First 1
    $packagedTemp = Join-Path $env:LOCALAPPDATA "Packages\$($installed.PackageFamilyName)\TempState"
    Set-DiagnosticsMarker -Paths @((Join-Path $packagedTemp $markerRelative)) -On ([bool]$Diagnostics)
}
