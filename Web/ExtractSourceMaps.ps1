<#
.SYNOPSIS
  Extracts original sources from sourcemap files into a destination folder.

.DESCRIPTION
  Scans .map files under -MapRoot and writes recovered source files into -OutRoot.
  Supports dry-run mode, conflict-safe naming, optional cleanup, and optional CSV reporting.

.PARAMETER MapRoot
  Root directory containing .map files to process.

.PARAMETER OutRoot
  Destination directory for extracted source files.
  If omitted, defaults to "<MapRoot parent>\_unminified_sources".

.PARAMETER Overwrite
  Overwrite existing extracted files. If omitted, conflict-safe file suffixes are used.

.PARAMETER CleanOutRoot
  Remove existing output files in OutRoot before extraction.

.PARAMETER WriteReport
  Write _source-map-extraction-report.csv in MapRoot.

.PARAMETER AllowOutRootInsideBundleTree
  Allow OutRoot inside the bundles tree under MapRoot's parent.

.PARAMETER DryRun
  Simulate extraction without writing files.

.PARAMETER NonInteractive
  Do not prompt for values or execution confirmation.

.PARAMETER Help
  Show short usage text.
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$MapRoot,
  [string]$OutRoot,
  [switch]$Overwrite,
  [switch]$CleanOutRoot,
  [switch]$WriteReport,
  [switch]$AllowOutRootInsideBundleTree,
  [switch]$DryRun,
  [switch]$NonInteractive,
  [Alias('h')]
  [switch]$Help
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Show-Usage {
  Write-Host "Usage:"
  Write-Host "  pwsh -File .\ExtractSourceMaps.ps1 -MapRoot <path> [-OutRoot <path>] [options]"
  Write-Host ""
  Write-Host "Options:"
  Write-Host "  -Overwrite                        Overwrite existing files instead of conflict suffix handling."
  Write-Host "  -CleanOutRoot                     Remove all existing content in OutRoot before extraction."
  Write-Host "  -WriteReport                      Write _source-map-extraction-report.csv in MapRoot."
  Write-Host "  -AllowOutRootInsideBundleTree     Allow OutRoot under packs/js."
  Write-Host "  -DryRun                           Simulate writes without changing files."
  Write-Host "  -NonInteractive                   Do not prompt for input or confirmation."
  Write-Host "  -Help                             Show this help."
  Write-Host ""
  Write-Host "Examples:"
  Write-Host "  pwsh -File .\ExtractSourceMaps.ps1 -MapRoot 'D:\site\packs\js\maps' -OutRoot 'D:\site\_unminified_sources' -CleanOutRoot"
  Write-Host "  pwsh -File .\ExtractSourceMaps.ps1 -MapRoot 'D:\site\packs\js\maps' -DryRun"
}

function Test-ReadAccess {
  param(
    [Parameter(Mandatory = $true)]
    [string]$PathValue
  )

  try {
    Get-ChildItem -LiteralPath $PathValue -File -ErrorAction Stop | Select-Object -First 1 | Out-Null
    return $true
  }
  catch {
    return $false
  }
}

function Test-WriteAccess {
  param(
    [Parameter(Mandatory = $true)]
    [string]$PathValue
  )

  $probe = Join-Path $PathValue (".write_probe_{0}.tmp" -f ([Guid]::NewGuid().ToString("N")))

  try {
    Set-Content -LiteralPath $probe -Value "probe" -Encoding ascii -ErrorAction Stop
    Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
    return $true
  }
  catch {
    if (Test-Path -LiteralPath $probe) {
      Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
    }

    return $false
  }
}

function Confirm-ExecutionPlan {
  param(
    [Parameter(Mandatory = $true)]
    [string]$MapRootValue,
    [Parameter(Mandatory = $true)]
    [string]$OutRootValue,
    [Parameter(Mandatory = $true)]
    [bool]$WillClean,
    [Parameter(Mandatory = $true)]
    [bool]$WillDryRun,
    [Parameter(Mandatory = $true)]
    [bool]$WillOverwrite,
    [Parameter(Mandatory = $true)]
    [bool]$WillWriteReport
  )

  Write-Host "Run configuration:"
  Write-Host "  MapRoot: $MapRootValue"
  Write-Host "  OutRoot: $OutRootValue"
  Write-Host "  Dry run: $WillDryRun"
  Write-Host "  Clean OutRoot first: $WillClean"
  Write-Host "  Overwrite mode: $WillOverwrite"
  Write-Host "  Write report CSV: $WillWriteReport"

  if ($WillClean -and -not $WillDryRun) {
    Write-Host "This will delete all content in directory '$OutRootValue'."
  }
  elseif ($WillClean -and $WillDryRun) {
    Write-Host "Dry run: OutRoot cleanup is enabled but no files will be deleted."
  }

  if (-not $WillOverwrite) {
    Write-Host "Conflict handling: keep all variants and suffix conflict files."
  }

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

function Normalize-SourcePath {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Source,
    [string]$SourceRoot,
    [Parameter(Mandatory = $true)]
    [string]$MapName,
    [Parameter(Mandatory = $true)]
    [int]$Index
  )

  $resolved = $Source
  if (-not [string]::IsNullOrWhiteSpace($SourceRoot) -and $Source -notmatch '^[a-zA-Z][a-zA-Z0-9+.-]*://') {
    $resolved = "$SourceRoot$Source"
  }

  $clean = $resolved
  $clean = $clean -replace '\\', '/'
  $clean = $clean -replace '^webpack:///?', ''
  $clean = $clean -replace '^file:///?', ''
  $clean = $clean -replace '^https?://', 'remote/'
  $clean = $clean -replace '[\?#].*$', ''
  $clean = $clean -replace '^\./', ''
  $clean = $clean -replace '^/', ''
  $clean = $clean -replace '/+$', ''
  $clean = $clean -replace '^[A-Za-z]:', ''

  # Keep relative source hierarchy while avoiding path traversal and invalid Windows chars.
  $clean = $clean -replace '\.\.(?:/|\\)', '_parent_/'
  $clean = $clean -replace '[<>:"|?*]', '_'

  $segments = @()
  foreach ($segment in ($clean -split '/')) {
    if ([string]::IsNullOrWhiteSpace($segment)) {
      continue
    }

    $safeSegment = $segment.Trim()
    $safeSegment = $safeSegment -replace '[\x00-\x1F]', '_'
    $safeSegment = $safeSegment -replace '[\. ]+$', ''

    if ([string]::IsNullOrWhiteSpace($safeSegment)) {
      $safeSegment = '_'
    }

    if ($safeSegment -match '^(?i)(con|prn|aux|nul|com[1-9]|lpt[1-9])$') {
      $safeSegment = $safeSegment + '_'
    }

    $segments += $safeSegment
  }

  if ($segments.Count -gt 0) {
    $clean = ($segments -join '/')
  }

  if ([string]::IsNullOrWhiteSpace($clean)) {
    return "unknown/$MapName/source_$Index.txt"
  }

  return $clean
}

function Resolve-NormalizedPath {
  param(
    [Parameter(Mandatory = $true)]
    [string]$PathValue
  )

  return [IO.Path]::GetFullPath($PathValue)
}

function Is-PathWithin {
  param(
    [Parameter(Mandatory = $true)]
    [string]$PathValue,
    [Parameter(Mandatory = $true)]
    [string]$BasePath
  )

  $normalizedPath = Resolve-NormalizedPath -PathValue $PathValue
  $normalizedBase = Resolve-NormalizedPath -PathValue $BasePath
  $baseWithSeparator = $normalizedBase.TrimEnd([IO.Path]::DirectorySeparatorChar) + [IO.Path]::DirectorySeparatorChar

  return $normalizedPath.StartsWith($baseWithSeparator, [System.StringComparison]::OrdinalIgnoreCase)
}

function New-FallbackTargetPath {
  param(
    [Parameter(Mandatory = $true)]
    [string]$OutRoot,
    [Parameter(Mandatory = $true)]
    [string]$MapBaseName,
    [Parameter(Mandatory = $true)]
    [int]$Index
  )

  $fallbackDir = Join-Path -Path $OutRoot -ChildPath "_invalid_paths/$MapBaseName"
  if (-not (Test-Path -LiteralPath $fallbackDir)) {
    New-Item -ItemType Directory -Force -Path $fallbackDir | Out-Null
  }

  return Join-Path -Path $fallbackDir -ChildPath ("source_{0}.txt" -f $Index)
}

function New-ConflictTargetPath {
  param(
    [Parameter(Mandatory = $true)]
    [string]$TargetPath,
    [Parameter(Mandatory = $true)]
    [string]$Suffix
  )

  $ext = [IO.Path]::GetExtension($TargetPath)
  $base = [IO.Path]::GetFileNameWithoutExtension($TargetPath)
  $dir = Split-Path -Path $TargetPath -Parent

  if ([string]::IsNullOrEmpty($ext)) {
    return Join-Path $dir ($base + $Suffix)
  }

  return Join-Path $dir ($base + $Suffix + $ext)
}

function Resolve-UniquePath {
  param(
    [Parameter(Mandatory = $true)]
    [string]$PathValue
  )

  if (-not (Test-Path -LiteralPath $PathValue)) {
    return $PathValue
  }

  $ext = [IO.Path]::GetExtension($PathValue)
  $base = [IO.Path]::GetFileNameWithoutExtension($PathValue)
  $dir = Split-Path -Path $PathValue -Parent
  $attempt = 1

  while ($true) {
    $candidate = if ([string]::IsNullOrEmpty($ext)) {
      Join-Path $dir ("{0}__dup_{1}" -f $base, $attempt)
    }
    else {
      Join-Path $dir ("{0}__dup_{1}{2}" -f $base, $attempt, $ext)
    }

    if (-not (Test-Path -LiteralPath $candidate)) {
      return $candidate
    }

    $attempt += 1
  }
}

if ($Help) {
  Show-Usage
  exit 0
}

$validationErrors = New-Object System.Collections.Generic.List[string]

if ([string]::IsNullOrWhiteSpace($MapRoot)) {
  $validationErrors.Add("MapRoot cannot be empty.") | Out-Null
}

if ($validationErrors.Count -gt 0) {
  Write-Warning "Invalid parameters:"
  foreach ($err in $validationErrors) {
    Write-Warning "  - $err"
  }

  Write-Host ""
  Show-Usage
  exit 1
}

if (-not (Test-Path -LiteralPath $MapRoot)) {
  Write-Warning "Invalid parameters:"
  Write-Warning "  - MapRoot not found: $MapRoot"
  Write-Host ""
  Show-Usage
  exit 1
}

$mapRootFull = Resolve-NormalizedPath -PathValue $MapRoot
$jsRoot = Split-Path -Path $mapRootFull -Parent
$defaultOutRoot = Join-Path (Split-Path -Path $mapRootFull -Parent) "_unminified_sources"

if ([string]::IsNullOrWhiteSpace($OutRoot)) {
  if ($NonInteractive) {
    $OutRoot = $defaultOutRoot
  }
  else {
    $outRootInput = Read-Host "OutRoot (optional). Press Enter to use default: $defaultOutRoot"
    if ([string]::IsNullOrWhiteSpace($outRootInput)) {
      $OutRoot = $defaultOutRoot
    }
    else {
      $OutRoot = $outRootInput
    }
  }
}

$outRootFull = Resolve-NormalizedPath -PathValue $OutRoot

if (-not $AllowOutRootInsideBundleTree) {
  if ($outRootFull -eq $jsRoot -or (Is-PathWithin -PathValue $outRootFull -BasePath $jsRoot)) {
    throw "OutRoot must be outside the original bundles folder ($jsRoot). Choose an output path outside packs/js, or pass -AllowOutRootInsideBundleTree if you intentionally want that."
  }
}

New-Item -ItemType Directory -Force -Path $outRootFull | Out-Null

if (-not (Test-ReadAccess -PathValue $mapRootFull)) {
  throw "MapRoot exists but could not be read: $mapRootFull"
}

if (-not (Test-WriteAccess -PathValue $outRootFull)) {
  throw "OutRoot exists but could not be written to: $outRootFull"
}

if (-not $NonInteractive) {
  $confirmed = Confirm-ExecutionPlan -MapRootValue $mapRootFull -OutRootValue $outRootFull -WillClean $CleanOutRoot.IsPresent -WillDryRun $DryRun.IsPresent -WillOverwrite $Overwrite.IsPresent -WillWriteReport $WriteReport.IsPresent
  if (-not $confirmed) {
    Write-Host "Cancelled by user."
    exit 0
  }
}

if ($CleanOutRoot) {
  if ($DryRun) {
    Write-Host "Dry run: would clear output directory: $outRootFull"
  }
  else {
    Get-ChildItem -LiteralPath $outRootFull -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force
    Write-Host "Cleared output directory: $outRootFull"
  }
}

$mappingRows = $null
if ($WriteReport) {
  $mappingRows = New-Object System.Collections.Generic.List[object]
}

$mapFiles = Get-ChildItem -LiteralPath $mapRootFull -Recurse -File -Filter '*.map'
if ($mapFiles.Count -eq 0) {
  Write-Host "No .map files found under: $mapRootFull"
  exit 0
}

$writtenCount = 0
$skippedNoSourceContent = 0
$skippedExisting = 0
$conflictCount = 0
$mapCount = 0
$firstWriteByTarget = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)

foreach ($mapFile in $mapFiles) {
  $mapCount += 1

  try {
    $raw = Get-Content -LiteralPath $mapFile.FullName -Raw
    $map = $raw | ConvertFrom-Json -Depth 100
  }
  catch {
    Write-Warning "Failed to parse map file: $($mapFile.FullName)"
    continue
  }

  if (-not $map.sources -or -not $map.sourcesContent) {
    continue
  }

  $sourceRootValue = $null
  if ($map.PSObject.Properties.Name -contains 'sourceRoot') {
    $sourceRootValue = [string]$map.sourceRoot
  }

  $limit = [Math]::Min($map.sources.Count, $map.sourcesContent.Count)
  for ($i = 0; $i -lt $limit; $i++) {
    $sourcePath = [string]$map.sources[$i]
    $sourceContent = $map.sourcesContent[$i]

    if ($null -eq $sourceContent -or $sourceContent -eq "") {
      $skippedNoSourceContent += 1
      continue
    }

    $relative = Normalize-SourcePath -Source $sourcePath -SourceRoot $sourceRootValue -MapName $mapFile.BaseName -Index $i
    $relativeFs = $relative -replace '/', [IO.Path]::DirectorySeparatorChar
    $target = Join-Path -Path $outRootFull -ChildPath $relativeFs
    $targetDir = Split-Path -Path $target -Parent

    if (-not (Is-PathWithin -PathValue $target -BasePath $outRootFull) -and (Resolve-NormalizedPath -PathValue $target) -ne $outRootFull) {
      throw "Refusing to write outside OutRoot. Calculated target: $target"
    }

    if (-not (Test-Path -LiteralPath $targetDir)) {
      New-Item -ItemType Directory -Force -Path $targetDir | Out-Null
    }

    $finalTarget = $target
    $targetKey = Resolve-NormalizedPath -PathValue $target
    if ((Test-Path -LiteralPath $target) -and -not $Overwrite) {
      $existing = Get-Content -LiteralPath $target -Raw
      if ($existing -eq $sourceContent) {
        $skippedExisting += 1
      }
      else {
        $conflictCount += 1
        if (Test-Path -LiteralPath $target) {
          $firstMeta = $null
          if ($firstWriteByTarget.ContainsKey($targetKey)) {
            $firstMeta = $firstWriteByTarget[$targetKey]
          }

          if ($null -ne $firstMeta -and -not $firstMeta.Moved) {
            $existingSuffix = "__from_$($firstMeta.MapBaseName)_$($firstMeta.Index)"
            $existingConflictTarget = Resolve-UniquePath -PathValue (New-ConflictTargetPath -TargetPath $target -Suffix $existingSuffix)

            if (-not $DryRun) {
              Move-Item -LiteralPath $target -Destination $existingConflictTarget -Force
            }

            if ($WriteReport -and $firstMeta.ReportRowIndex -ge 0) {
              $mappingRows[$firstMeta.ReportRowIndex].Output = $existingConflictTarget
            }

            $firstMeta.Moved = $true
            $firstMeta.MovedTo = $existingConflictTarget
            $firstWriteByTarget[$targetKey] = $firstMeta
          }
          elseif ($null -eq $firstMeta) {
            $existingConflictTarget = Resolve-UniquePath -PathValue (New-ConflictTargetPath -TargetPath $target -Suffix "__from_existing")
            if (-not $DryRun) {
              Move-Item -LiteralPath $target -Destination $existingConflictTarget -Force
            }
          }
        }

        $suffix = "__from_$($mapFile.BaseName)_$i"
        $finalTarget = Resolve-UniquePath -PathValue (New-ConflictTargetPath -TargetPath $target -Suffix $suffix)
      }
    }

    if ($DryRun) {
      $writtenCount += 1
    }
    else {
      try {
        Set-Content -LiteralPath $finalTarget -Value $sourceContent -Encoding utf8
      }
      catch {
        $fallbackTarget = New-FallbackTargetPath -OutRoot $outRootFull -MapBaseName $mapFile.BaseName -Index $i
        Set-Content -LiteralPath $fallbackTarget -Value $sourceContent -Encoding utf8
        $finalTarget = $fallbackTarget
      }

      $writtenCount += 1
    }

    $reportRowIndex = -1
    if ($WriteReport) {
      $mappingRows.Add([PSCustomObject]@{
        MapFile = $mapFile.FullName
        Source = $sourcePath
        Output = $finalTarget
      }) | Out-Null
      $reportRowIndex = $mappingRows.Count - 1
    }

    if ($finalTarget -eq $target -and -not $firstWriteByTarget.ContainsKey($targetKey)) {
      $firstWriteByTarget[$targetKey] = [PSCustomObject]@{
        MapBaseName = $mapFile.BaseName
        Index = $i
        Moved = $false
        MovedTo = $null
        ReportRowIndex = $reportRowIndex
      }
    }
  }
}

$reportPath = Join-Path $mapRootFull "_source-map-extraction-report.csv"
if ($WriteReport -and -not $DryRun) {
  $mappingRows | Export-Csv -LiteralPath $reportPath -NoTypeInformation -Encoding utf8
}

Write-Host "Processed maps: $mapCount"
Write-Host "Files written: $writtenCount"
Write-Host "Skipped (empty sourcesContent): $skippedNoSourceContent"
Write-Host "Skipped (existing same content): $skippedExisting"
Write-Host "Conflicts (renamed outputs): $conflictCount"
if ($DryRun) {
  Write-Host "Dry run: no files were written"
}
elseif ($WriteReport) {
  Write-Host "Report: $reportPath"
}
else {
  Write-Host "Report: skipped (pass -WriteReport to generate CSV)"
}

if ($DryRun -and $WriteReport) {
  Write-Host "Report path (if generated): $reportPath"
}
