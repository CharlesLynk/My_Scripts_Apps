[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path $PSScriptRoot 'OMI_ViewAppLauncher.local.exe'),
    [string]$CompilerPath
)

$ErrorActionPreference = 'Stop'

$source = Join-Path $PSScriptRoot 'OMI_ViewAppLauncher.cs'
$output = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)

if ([IO.Path]::GetExtension($output) -ne '.exe') {
    throw 'OutputPath must name an .exe file.'
}
if (Test-Path -LiteralPath $output) {
    throw "Output already exists: $output. Choose a new -OutputPath."
}
if (-not (Test-Path -LiteralPath (Split-Path -Parent $output) -PathType Container)) {
    throw 'The OutputPath parent directory must already exist.'
}

if (-not $CompilerPath) {
    foreach ($framework in 'Framework64', 'Framework') {
        $candidate = Join-Path $env:WINDIR "Microsoft.NET\$framework\v4.0.30319\csc.exe"
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            $CompilerPath = $candidate
            break
        }
    }
}

if (-not (Test-Path -LiteralPath $source)) {
    throw "Source file not found: $source"
}

if (-not $CompilerPath -or -not (Test-Path -LiteralPath $CompilerPath -PathType Leaf)) {
    throw 'A .NET Framework C# compiler was not found. Supply -CompilerPath to csc.exe.'
}

$csc = (Get-Item -LiteralPath $CompilerPath).FullName
& $csc `
    /nologo `
    /target:winexe `
    "/out:$output" `
    /reference:System.Windows.Forms.dll `
    $source

if ($LASTEXITCODE -ne 0) {
    throw "Build failed with exit code $LASTEXITCODE"
}

Write-Host "Built executable:" $output