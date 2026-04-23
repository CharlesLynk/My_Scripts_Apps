<#
.SYNOPSIS
  Attempts to locate or download sourcemap files (*.map) for JS/CSS assets in a folder tree.

.DESCRIPTION
  For each .js/.mjs/.cjs (and optionally .css) file under -Root:
   1) Look for //# sourceMappingURL=... or /*# sourceMappingURL=... */
   2) If the map is local and exists, record it.
   3) If missing, attempt to download:
       - If sourceMappingURL is absolute http(s), download directly.
       - If relative and BaseUrl is provided (arg or prompt), resolve relative to the generated file URL.
   4) If no sourceMappingURL, try heuristic: "<filename>.map" in the same directory,
      and optionally download it if BaseUrl is provided.

.PARAMETER Root
  Root directory containing downloaded JS/CSS files.

.PARAMETER BaseUrl
  Optional site base URL (e.g., https://example.com). If omitted, the script can
  prompt for it (interactive mode) and then attempt automatic inference.

.PARAMETER IncludeCss
  Also process .css files.

.PARAMETER WriteReport
  Write sourcemap_report.csv in the Root directory.

.PARAMETER NonInteractive
  Do not prompt for input or confirmation. Useful for automation.

.PARAMETER Help
  Show short usage text.

.PARAMETER DryRun
  Print what would happen without downloading.

.PARAMETER TimeoutSec
  Web request timeout.

.PARAMETER MaxRetries
  Retry count for downloads.

.OUTPUTS
  Optionally writes a CSV report "sourcemap_report.csv" in Root when -WriteReport is used.
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)]
  [string]$Root,

  [string]$BaseUrl,

  [switch]$IncludeCss,

  [switch]$WriteReport,

  [switch]$NonInteractive,

  [switch]$DryRun,

  [int]$TimeoutSec = 25,

  [int]$MaxRetries = 2,

  [Alias('h')]
  [switch]$Help
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Show-Usage {
  Write-Host "Usage:"
  Write-Host "  pwsh -File .\MapFinder.ps1 -Root <path> [options]"
  Write-Host ""
  Write-Host "Options:"
  Write-Host "  -BaseUrl <url>      Optional base URL used to resolve relative sourceMappingURL values."
  Write-Host "  -IncludeCss         Also process .css files."
  Write-Host "  -WriteReport        Write sourcemap_report.csv in Root."
  Write-Host "  -DryRun             Simulate downloads without writing files."
  Write-Host "  -NonInteractive     Do not prompt for BaseUrl or confirmation."
  Write-Host "  -TimeoutSec <int>   Request timeout (default: 25)."
  Write-Host "  -MaxRetries <int>   Download retry count (default: 2)."
  Write-Host "  -Help               Show this help."
}

function Confirm-ExecutionPlan {
  param(
    [Parameter(Mandatory = $true)]
    [string]$RootValue,
    [string]$BaseUrlValue,
    [Parameter(Mandatory = $true)]
    [bool]$WillIncludeCss,
    [Parameter(Mandatory = $true)]
    [bool]$WillDryRun,
    [Parameter(Mandatory = $true)]
    [bool]$WillWriteReport
  )

  Write-Host "Run configuration:"
  Write-Host "  Root: $RootValue"
  Write-Host "  BaseUrl: $(if ($BaseUrlValue) { $BaseUrlValue } else { '<none>' })"
  Write-Host "  Include CSS: $WillIncludeCss"
  Write-Host "  Dry run: $WillDryRun"
  Write-Host "  Write report CSV: $WillWriteReport"

  while ($true) {
    $confirmation = Read-Host "Proceed? [y/N]"
    if ([string]::IsNullOrWhiteSpace($confirmation) -or $confirmation -match '^(?i)n(?:o)?$') {
      return $false
    }

    if ($confirmation -match '^(?i)y(?:es)?$') {
      return $true
    }

    Write-Host "Please enter 'y' or 'n'."
  }
}

function Get-RelativePath([string]$basePath, [string]$fullPath) {
  $base = (Resolve-Path $basePath).Path.TrimEnd('\')
  $full = (Resolve-Path $fullPath).Path
  if ($full.StartsWith($base, [System.StringComparison]::OrdinalIgnoreCase)) {
    return $full.Substring($base.Length).TrimStart('\')
  }
  return $null
}

function Get-PathSegments([string]$path) {
  return ($path -split '[\\/]') | Where-Object { $_ -ne '' }
}

function Test-IsLikelyHostSegment([string]$segment) {
  if (-not $segment) { return $false }
  return ($segment -match '^[A-Za-z0-9.-]+\.[A-Za-z]{2,}$')
}

function Get-HostFromRootPath([string]$path) {
  $parts = Get-PathSegments -path $path
  if ($parts.Count -eq 0) { return $null }

  for ($i = 0; $i -lt $parts.Count; $i++) {
    if ($parts[$i] -ieq 'top' -and ($i + 1) -lt $parts.Count -and (Test-IsLikelyHostSegment $parts[$i + 1])) {
      return $parts[$i + 1]
    }
  }

  for ($i = 0; $i -lt $parts.Count; $i++) {
    if (Test-IsLikelyHostSegment $parts[$i]) { return $parts[$i] }
  }

  return $null
}

function Get-CommonRootUrlPrefix([string]$rootPath) {
  if (-not $rootPath) { return $null }
  $parts = Get-PathSegments -path $rootPath
  if ($parts.Count -eq 0) { return $null }

  $strongAnchors = @('packs','assets','static','dist','build','public','_next','bundles','webpack','scripts','styles')
  for ($i = $parts.Count - 1; $i -ge 0; $i--) {
    if ($strongAnchors -contains $parts[$i].ToLowerInvariant()) {
      return ($parts[$i..($parts.Count - 1)] -join '/')
    }
  }

  $last = $parts[$parts.Count - 1].ToLowerInvariant()
  if ($last -in @('js','css')) {
    return $parts[$parts.Count - 1]
  }

  return $null
}

function Get-PathWithoutPrefixTail([string]$rootPath, [string]$prefix) {
  if (-not $rootPath -or -not $prefix) { return $null }

  $parts = Get-PathSegments -path $rootPath
  $tail = Get-PathSegments -path $prefix
  if ($parts.Count -le $tail.Count) { return $null }

  for ($i = 0; $i -lt $tail.Count; $i++) {
    $partIndex = $parts.Count - $tail.Count + $i
    if ($parts[$partIndex] -ine $tail[$i]) { return $null }
  }

  $keep = $parts[0..($parts.Count - $tail.Count - 1)]
  if ($keep.Count -eq 0) { return $null }
  return ($keep -join '\\')
}

function Get-RootUrlPrefix([string]$rootPath, [string]$host) {
  if (-not $rootPath) { return $null }

  if (-not $host) {
    return Get-CommonRootUrlPrefix -rootPath $rootPath
  }

  $parts = Get-PathSegments -path $rootPath
  $hostIndex = -1

  for ($i = 0; $i -lt $parts.Count; $i++) {
    if ($parts[$i] -ieq $host) {
      $hostIndex = $i
      break
    }
  }

  if ($hostIndex -lt 0) {
    return Get-CommonRootUrlPrefix -rootPath $rootPath
  }

  if ($hostIndex -ge ($parts.Count - 1)) {
    return Get-CommonRootUrlPrefix -rootPath $rootPath
  }

  $suffix = $parts[($hostIndex + 1)..($parts.Count - 1)]
  if ($suffix.Count -eq 0) { return $null }
  return ($suffix -join '/')
}

function Get-LocalUrlRoot([string]$rootPath, [string]$host, [string]$rootUrlPrefix) {
  if (-not $rootPath) { return $null }

  if (-not $host) {
    return Get-PathWithoutPrefixTail -rootPath $rootPath -prefix $rootUrlPrefix
  }

  $parts = Get-PathSegments -path $rootPath
  $hostIndex = -1

  for ($i = 0; $i -lt $parts.Count; $i++) {
    if ($parts[$i] -ieq $host) {
      $hostIndex = $i
      break
    }
  }

  if ($hostIndex -lt 0) {
    return Get-PathWithoutPrefixTail -rootPath $rootPath -prefix $rootUrlPrefix
  }

  return ($parts[0..$hostIndex] -join '\\')
}

function Resolve-MapUrl([string]$baseUrlNorm, [string]$rootResolved, [string]$rootUrlPrefix, [string]$mapRef, [string]$filePath) {
  if (-not $baseUrlNorm) { return $null }
  $rel = Get-RelativePath -basePath $rootResolved -fullPath $filePath
  if (-not $rel) { return $null }
  $relUrl = ($rel -replace '\\','/')
  $genUrl = if ($rootUrlPrefix) { "$baseUrlNorm/$rootUrlPrefix/$relUrl" } else { "$baseUrlNorm/$relUrl" }
  try {
    $baseUri = [System.Uri]::new($genUrl)
    $resolved = [System.Uri]::new($baseUri, $mapRef)
    return $resolved.AbsoluteUri
  } catch {
    return $null
  }
}

function ConvertTo-NormalizedMapReference([string]$mapRef) {
  if (-not $mapRef) { return $mapRef }
  $v = $mapRef.Trim()
  # Some build tools wrap sourceMappingURL values in single/double quotes.
  if (($v.StartsWith('"') -and $v.EndsWith('"')) -or ($v.StartsWith("'") -and $v.EndsWith("'"))) {
    if ($v.Length -ge 2) { $v = $v.Substring(1, $v.Length - 2) }
  }
  return $v
}

function Invoke-Download([string]$url, [string]$destFile) {
  $destDir = Split-Path -Parent $destFile
  if (-not (Test-Path $destDir)) { New-Item -ItemType Directory -Path $destDir | Out-Null }

  for ($i=0; $i -le $MaxRetries; $i++) {
    try {
      if ($DryRun) {
        Write-Host "[DryRun] Would download: $url -> $destFile"
        return $true
      }
      Write-Host "Downloading: $url"
      $iwrParams = @{
        Uri = $url
        OutFile = $destFile
        TimeoutSec = $TimeoutSec
      }
      # -UseBasicParsing exists in Windows PowerShell but not in PowerShell 7+.
      if ($PSVersionTable.PSVersion.Major -lt 6) { $iwrParams.UseBasicParsing = $true }
      Invoke-WebRequest @iwrParams
      return $true
    } catch {
      if ($i -ge $MaxRetries) {
        Write-Warning "Failed to download after retries: $url  Error: $($_.Exception.Message)"
        return $false
      }
      Start-Sleep -Seconds ([Math]::Min(3 * ($i+1), 10))
    }
  }
  return $false
}

if ($Help) {
  Show-Usage
  exit 0
}

if (-not (Test-Path -LiteralPath $Root)) {
  Write-Warning "Root not found: $Root"
  Write-Host ""
  Show-Usage
  exit 1
}

# Regex for either //# sourceMappingURL=... or /*# sourceMappingURL=... */
$rx = [regex]'(?im)^\s*(?://#|/\*#)\s*sourceMappingURL\s*=\s*(?<url>.+?)\s*(?:\*/)?\s*$'

$exts = @("*.js","*.mjs","*.cjs")
if ($IncludeCss) { $exts += "*.css" }

$rootResolved = (Resolve-Path $Root).Path.TrimEnd('\')
$baseUrlNorm = $null
$baseUrlInput = $BaseUrl

if (-not $baseUrlInput -and -not $NonInteractive) {
  try {
    $baseUrlInput = Read-Host "BaseUrl (optional, e.g. https://example.com). Press Enter to auto-infer"
  } catch {
    $baseUrlInput = $null
  }
}

if ($baseUrlInput) {
  $baseUrlInput = $baseUrlInput.Trim()
  if ($baseUrlInput) {
    $baseUrlNorm = $baseUrlInput.TrimEnd('/')
  }
}

$rootUrlPrefix = $null
$localUrlRoot = $rootResolved

if (-not $baseUrlNorm) {
  $inferredHost = Get-HostFromRootPath -path $rootResolved
  if ($inferredHost) {
    $baseUrlNorm = "https://$inferredHost"
    Write-Host "Inferred BaseUrl from path: $baseUrlNorm"
  }
}

if (-not $NonInteractive) {
  $confirmed = Confirm-ExecutionPlan -RootValue $rootResolved -BaseUrlValue $baseUrlNorm -WillIncludeCss $IncludeCss.IsPresent -WillDryRun $DryRun.IsPresent -WillWriteReport $WriteReport.IsPresent
  if (-not $confirmed) {
    Write-Host "Cancelled by user."
    exit 0
  }
}

if ($baseUrlNorm) {
  try {
    $baseUri = [System.Uri]::new($baseUrlNorm)
    $host = $baseUri.Host
    $rootUrlPrefix = Get-RootUrlPrefix -rootPath $rootResolved -host $host
    $maybeLocalUrlRoot = Get-LocalUrlRoot -rootPath $rootResolved -host $host -rootUrlPrefix $rootUrlPrefix
    if ($maybeLocalUrlRoot) { $localUrlRoot = $maybeLocalUrlRoot }
  } catch {
    # Keep defaults if BaseUrl cannot be parsed as URI.
  }
}

$results = New-Object System.Collections.Generic.List[object]

Get-ChildItem -Path $rootResolved -Recurse -File -Include $exts | ForEach-Object {
  $file = $_.FullName
  $dir  = $_.DirectoryName
  $name = $_.Name

  # Read only tail portion (maps are typically referenced near end); fall back to full if needed
  $text = $null
  try {
    $bytes = [System.IO.File]::ReadAllBytes($file)
    $take  = [Math]::Min($bytes.Length, 65536) # last 64KB
    $slice = New-Object byte[] $take
    [Array]::Copy($bytes, $bytes.Length - $take, $slice, 0, $take)
    $text = [System.Text.Encoding]::UTF8.GetString($slice)
  } catch {
    $text = Get-Content -LiteralPath $file -Raw -ErrorAction SilentlyContinue
  }

  $mapRef = $null
  $m = $rx.Matches($text)
  if ($m.Count -gt 0) {
    $mapRef = $m[$m.Count-1].Groups["url"].Value.Trim()
  }

  $status = "Unknown"
  $mapLocalPath = $null
  $mapUrl = $null
  $note = $null

  if ($mapRef) {
    # Ignore inline data: maps
    if ($mapRef -match '^\s*data:') {
      $status = "InlineMap"
      $note = "sourceMappingURL is a data: URI; no external .map to download."
    } else {
      # Determine local destination path for the map (relative paths create subfolders)
      $mapRef = ConvertTo-NormalizedMapReference -mapRef $mapRef
      $mapRefNoQuery = ($mapRef -split '\?')[0]
      $mapRefNoHash  = ($mapRefNoQuery -split '#')[0]

      # If absolute http(s)
      if ($mapRefNoHash -match '^(https?:)?//') {
        if ($mapRefNoHash -match '^//') { $mapUrl = "https:$mapRefNoHash" } else { $mapUrl = $mapRefNoHash }
        $leaf = [System.IO.Path]::GetFileName(([System.Uri]$mapUrl).AbsolutePath)
        if (-not $leaf) { $leaf = "$name.map" }
        $mapLocalPath = Join-Path $dir $leaf
      } elseif ($mapRefNoHash.StartsWith('/')) {
        # Root-relative map path: store in a local "maps" folder beside the source file.
        $leaf = [System.IO.Path]::GetFileName($mapRefNoHash)
        if (-not $leaf) { $leaf = "$name.map" }
        $mapLocalPath = Join-Path (Join-Path $dir 'maps') $leaf
        $mapUrl = if ($baseUrlNorm) { "$baseUrlNorm/$($mapRefNoHash.TrimStart('/'))" } else { $null }
      } else {
        # Relative path -> local sibling path
        $mapLocalPath = Join-Path $dir ($mapRefNoHash -replace '/','\')
        $mapUrl = Resolve-MapUrl -baseUrlNorm $baseUrlNorm -rootResolved $rootResolved -rootUrlPrefix $rootUrlPrefix -mapRef $mapRefNoHash -filePath $file
      }

      if (Test-Path -LiteralPath $mapLocalPath) {
        $status = "MapExistsLocal"
      } else {
        if ($mapUrl) {
          $ok = Invoke-Download -url $mapUrl -destFile $mapLocalPath
          $status = if ($ok) { "Downloaded" } else { "DownloadFailed" }
        } else {
          $status = "MissingNoBaseUrl"
          $note = "Relative sourceMappingURL but no BaseUrl was supplied or inferred to resolve it."
        }
      }
    }
  } else {
    # Heuristic: try "<filename>.map" in same folder
    $guess = Join-Path $dir ($name + ".map")
    if (Test-Path -LiteralPath $guess) {
      $status = "GuessedMapExistsLocal"
      $mapLocalPath = $guess
    } else {
      $mapLocalPath = $guess
      if ($baseUrlNorm) {
        $mapUrl = Resolve-MapUrl -baseUrlNorm $baseUrlNorm -rootResolved $rootResolved -rootUrlPrefix $rootUrlPrefix -mapRef ($name + ".map") -filePath $file
        if ($mapUrl) {
          $ok = Invoke-Download -url $mapUrl -destFile $mapLocalPath
          $status = if ($ok) { "DownloadedGuess" } else { "DownloadFailedGuess" }
        } else {
          $status = "GuessMissingUnresolvable"
        }
      } else {
        $status = "NoSourceMappingURL"
      }
    }
  }

  $results.Add([pscustomobject]@{
    File = $file
    SourceMappingURL = $mapRef
    MapLocalPath = $mapLocalPath
    MapUrl = $mapUrl
    Status = $status
    Note = $note
  }) | Out-Null
}

$processedCount = $results.Count
$downloadedCount = ($results | Where-Object { $_.Status -in @('Downloaded','DownloadedGuess') }).Count
$failedCount = ($results | Where-Object { $_.Status -in @('DownloadFailed','DownloadFailedGuess') }).Count

Write-Host "Processed files: $processedCount"
Write-Host "Downloaded maps: $downloadedCount"
Write-Host "Download failures: $failedCount"

if ($WriteReport) {
  $report = Join-Path $rootResolved "sourcemap_report.csv"
  if ($DryRun) {
    Write-Host "Dry run: report generation skipped"
    Write-Host "Report path (if generated): $report"
  } else {
    $results | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $report
    Write-Host "Report: $report"
  }
} else {
  Write-Host "Report: skipped (pass -WriteReport to generate CSV)"
}
