param(
    [Parameter(Mandatory)][uri]$Url,
    [Parameter(Mandatory)][ValidatePattern('^[0-9a-fA-F]{64}$')][string]$Sha256,
    [Parameter(Mandatory)][string]$Destination
)

$ErrorActionPreference = 'Stop'
if ($Url.Scheme -ne 'https') { throw 'OCR assets must be downloaded over HTTPS.' }

function Test-Asset {
    param([string]$Path)
    (Test-Path -LiteralPath $Path -PathType Leaf) -and
        ((Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash -eq $Sha256)
}

if (Test-Asset $Destination) { return }

$scratch = Join-Path ([System.IO.Path]::GetTempPath()) 'My_Scripts_Apps\AutoDIS'
$nonce = [guid]::NewGuid().ToString()
$download = Join-Path $scratch "$nonce.partial"
$staging = "$Destination.$nonce.partial"
$backup = "$Destination.$nonce.replaced"
try {
    [void][System.IO.Directory]::CreateDirectory($scratch)
    Write-Host "Restoring AutoDIS OCR asset from $Url"
    Invoke-WebRequest -UseBasicParsing -Uri $Url -OutFile $download -TimeoutSec 90
    if (-not (Test-Asset $download)) { throw "Download did not match SHA-256 $Sha256." }

    [void][System.IO.Directory]::CreateDirectory((Split-Path -Parent $Destination))
    [System.IO.File]::Copy($download, $staging)
    try {
        if ([System.IO.File]::Exists($Destination)) {
            [System.IO.File]::Replace($staging, $Destination, $backup)
        } else {
            [System.IO.File]::Move($staging, $Destination)
        }
    } catch [System.IO.IOException] {
        if (-not (Test-Asset $Destination)) { throw }
    }
} catch {
    throw "Cannot restore required OCR asset '$Destination' from '$Url'. Check network connectivity, the cache path, and the expected SHA-256. $($_.Exception.Message)"
} finally {
    [System.IO.File]::Delete($download)
    [System.IO.File]::Delete($staging)
    [System.IO.File]::Delete($backup)
}
