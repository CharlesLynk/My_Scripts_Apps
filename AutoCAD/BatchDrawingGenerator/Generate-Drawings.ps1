[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$AccoreConsolePath,

    [string]$TemplatePath,

    [string]$CsvPath,

    [string]$OutputDirectory,

    [string]$BorderXrefPath,

    [int]$MaxDrawings = 0,

    [switch]$KeepLogs
)

$ErrorActionPreference = "Stop"

try {
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

function Select-InputFile {
    param(
        [string]$Title,
        [string]$Filter
    )

    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = $Title
    $dialog.Filter = $Filter
    $dialog.CheckFileExists = $true
    $dialog.Multiselect = $false
    try {
        if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
            throw "No file selected for: $Title"
        }
        return $dialog.FileName
    }
    finally {
        $dialog.Dispose()
    }
}

function Resolve-InputFile {
    param(
        [string]$Path,
        [string]$Title,
        [string]$Filter,
        [string]$Extension
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        $Path = Select-InputFile -Title $Title -Filter $Filter
    }
    $Path = [System.IO.Path]::GetFullPath($Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "$Title file does not exist: $Path"
    }
    if ([System.IO.Path]::GetExtension($Path) -ine $Extension) {
        throw "$Title must be a $Extension file: $Path"
    }
    return $Path
}

function Resolve-OutputDirectory {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = "Select the destination folder for generated drawings"
        $dialog.ShowNewFolderButton = $true
        try {
            if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
                throw "No destination folder selected."
            }
            $Path = $dialog.SelectedPath
        }
        finally {
            $dialog.Dispose()
        }
    }

    $Path = [System.IO.Path]::GetFullPath($Path)
    if (Test-Path -LiteralPath $Path -PathType Leaf) {
        throw "Destination path is a file, not a folder: $Path"
    }
    return $Path
}

function Test-ReadableFile {
    param([string]$Path)

    $stream = [System.IO.File]::OpenRead($Path)
    $stream.Dispose()
}

function Find-AccoreConsole {
    $installRoot = Join-Path $env:ProgramFiles "Autodesk"
    if (Test-Path -LiteralPath $installRoot -PathType Container) {
        $installations = Get-ChildItem -LiteralPath $installRoot -Directory -Filter "AutoCAD*" |
            Sort-Object Name -Descending
        foreach ($installation in $installations) {
            $candidate = Join-Path $installation.FullName "accoreconsole.exe"
            if (Test-Path -LiteralPath $candidate -PathType Leaf) {
                return $candidate
            }
        }
    }
    return Select-InputFile -Title "Locate AutoCAD Core Console (accoreconsole.exe)" -Filter "AutoCAD Core Console (accoreconsole.exe)|accoreconsole.exe|Executable files (*.exe)|*.exe"
}

function ConvertTo-LispString {
    param([AllowEmptyString()][string]$Value)

    if ($Value -match '[\x00-\x1F]') {
        throw "CSV values cannot contain control characters or line breaks."
    }

    return '"' + $Value.Replace('\', '\\').Replace('"', '\"') + '"'
}

function Write-Utf8File {
    param(
        [string]$Path,
        [string]$Content
    )

    [System.IO.File]::WriteAllText(
        $Path,
        $Content,
        (New-Object System.Text.UTF8Encoding($false))
    )
}

function Get-AutoCadErrorText {
    param([string]$LogText)

    $logText = $LogText.Replace([string][char]0, "")
    $match = [regex]::Match($logText, '(?im)error:\s*([^\r\n]+)')
    if ($match.Success) {
        $detail = $match.Groups[1].Value.Trim()
        $detail = $detail -replace '(?i)\s*;?\s*error:\s*quit\s*/\s*exit\s+abort\s*$', ''
        $mappingMatch = [regex]::Match(
            $detail,
            '(?i)^CSV attribute mapping failed\.\s*Missing mappings:\s*(.*)$'
        )
        if ($mappingMatch.Success) {
            $missingMappings = @(
                $mappingMatch.Groups[1].Value -split ';\s*(?=[^;]*\(missing on insert handle)' |
                    ForEach-Object { $_ -replace '\s+\(missing on insert handle\(s\):[^)]*\)\s*$', '' }
            )
            return "CSV attribute mapping failed.`nMissing mappings:`n  - " + ($missingMappings -join "`n  - ")
        }
        return $detail
    }
    return $null
}

function Invoke-CoreConsole {
    param(
        [string]$DrawingPath,
        [string]$ScriptPath
    )

    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = $AccoreConsolePath
    $startInfo.Arguments = '/p "accoreconsole" /i "' + $DrawingPath + '" /s "' + $ScriptPath + '"'
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.StandardOutputEncoding = [System.Text.Encoding]::Unicode
    $startInfo.StandardErrorEncoding = [System.Text.Encoding]::Unicode

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $startInfo
    try {
        if (-not $process.Start()) {
            throw "Could not start AutoCAD Core Console: $AccoreConsolePath"
        }

        $stdoutTask = $process.StandardOutput.ReadToEndAsync()
        $stderrTask = $process.StandardError.ReadToEndAsync()
        $process.WaitForExit()
        return [pscustomobject]@{
            ExitCode = $process.ExitCode
            StdOut   = $stdoutTask.Result
            StdErr   = $stderrTask.Result
        }
    }
    finally {
        $process.Dispose()
    }
}

function Save-CoreConsoleLogs {
    param(
        [string]$StdOut,
        [string]$StdErr,
        [string]$DestinationLog,
        [string]$DestinationErrorLog
    )

    $destinationDirectory = Split-Path -Parent $DestinationLog
    New-Item -ItemType Directory -Path $destinationDirectory -Force | Out-Null
    $encoding = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($DestinationLog, $StdOut, $encoding)
    [System.IO.File]::WriteAllText($DestinationErrorLog, $StdErr, $encoding)
    return "$DestinationLog and $DestinationErrorLog"
}

function ConvertTo-LispMappings {
    param(
        [object[]]$Mappings,
        [object[]]$Values
    )

    $groups = @(
        foreach ($group in ($Mappings | Group-Object -Property BlockName)) {
            $pairs = @(
                foreach ($mapping in $group.Group) {
                    $value = ($Values | Where-Object { $_.Header -ceq $mapping.Header } | Select-Object -First 1).Value
                    "(" + (ConvertTo-LispString $mapping.AttributeName) + " . " + (ConvertTo-LispString ([string]$value)) + ")"
                }
            )
            "(" + (ConvertTo-LispString $group.Name) + " " + ($pairs -join " ") + ")"
        }
    )
    return "(" + ($groups -join " ") + ")"
}

if ([string]::IsNullOrWhiteSpace($AccoreConsolePath)) {
    $AccoreConsolePath = Find-AccoreConsole
}
if (-not (Test-Path -LiteralPath $AccoreConsolePath -PathType Leaf) -or
    [System.IO.Path]::GetExtension($AccoreConsolePath) -ine ".exe" -or
    [System.IO.Path]::GetFileName($AccoreConsolePath) -ine "accoreconsole.exe") {
    throw "Select a valid accoreconsole.exe file: $AccoreConsolePath"
}

$AccoreConsolePath = [System.IO.Path]::GetFullPath($AccoreConsolePath)
$TemplatePath = Resolve-InputFile -Path $TemplatePath -Title "AutoCAD template" -Filter "AutoCAD template (*.dwt)|*.dwt" -Extension ".dwt"
$CsvPath = Resolve-InputFile -Path $CsvPath -Title "Drawing data" -Filter "CSV files (*.csv)|*.csv" -Extension ".csv"
$BorderXrefPath = Resolve-InputFile -Path $BorderXrefPath -Title "Border XREF" -Filter "AutoCAD drawing (*.dwg)|*.dwg" -Extension ".dwg"

foreach ($filePath in @($AccoreConsolePath, $TemplatePath, $BorderXrefPath)) {
    Test-ReadableFile -Path $filePath
}

$BorderXrefName = [System.IO.Path]::GetFileNameWithoutExtension($BorderXrefPath)
if ([string]::IsNullOrWhiteSpace($BorderXrefName) -or
    $BorderXrefName -match '[<>/\\":;?*|,=\x00-\x1F]') {
    throw "The selected border DWG filename cannot be used as an XREF block name: $BorderXrefName"
}

if ($MaxDrawings -lt 0) {
    throw "MaxDrawings cannot be negative."
}

$csvParser = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($CsvPath)
try {
    $csvParser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
    $csvParser.SetDelimiters(",")
    $csvParser.HasFieldsEnclosedInQuotes = $true
    $headers = $csvParser.ReadFields()
}
finally {
    $csvParser.Dispose()
}

if ($null -eq $headers -or $headers.Count -eq 0) {
    throw "CSV has no header row."
}

$seenHeaders = @{}
foreach ($header in $headers) {
    if ([string]::IsNullOrWhiteSpace($header)) {
        throw "CSV headers cannot be empty."
    }
    if ($seenHeaders.ContainsKey($header)) {
        throw "CSV contains duplicate column headers: $header"
    }
    $seenHeaders[$header] = $true
}

if ($headers -notcontains "DrawingFileName") {
    throw "CSV must contain a DrawingFileName column."
}

$mappings = @()
$seenMappings = @{}
foreach ($header in $headers) {
    if ($header -ceq "DrawingFileName") {
        continue
    }
    $parts = $header.Split('.')
    if ($parts.Count -ne 2 -or
        [string]::IsNullOrWhiteSpace($parts[0]) -or
        [string]::IsNullOrWhiteSpace($parts[1])) {
        throw "Attribute column '$header' must use the exact BLOCKNAME.ATTRIBUTENAME format."
    }
    if ($parts[0] -match '[<>/\\":;?*|,=\x00-\x1F]' -or
        $parts[1] -match '[<>/\\":;?*|,=\x00-\x1F.]') {
        throw "Invalid block name or attribute name in CSV column '$header'."
    }
    $mappingKey = ($parts[0] + "." + $parts[1]).ToUpperInvariant()
    if ($seenMappings.ContainsKey($mappingKey)) {
        throw "CSV contains duplicate block/attribute mappings: $header"
    }
    $seenMappings[$mappingKey] = $true
    $mappings += [pscustomobject]@{
        Header         = $header
        BlockName      = $parts[0]
        AttributeName  = $parts[1]
    }
}

if ($mappings.Count -eq 0) {
    throw "CSV must contain at least one BLOCKNAME.ATTRIBUTENAME column."
}

$rows = @(Import-Csv -LiteralPath $CsvPath -ErrorAction Stop)
if ($rows.Count -eq 0) {
    throw "CSV contains no drawing rows."
}

$resolvedRows = @()
$seenNames = @{}

foreach ($row in $rows) {
    $drawingName = [string]$row.DrawingFileName
    if ([string]::IsNullOrWhiteSpace($drawingName)) {
        throw "Every CSV row must have a DrawingFileName value."
    }

    if ([System.IO.Path]::GetFileName($drawingName) -cne $drawingName) {
        throw "DrawingFileName must be a filename only, not a path: $drawingName"
    }

    $extension = [System.IO.Path]::GetExtension($drawingName)
    if (-not $extension) {
        $drawingName += ".dwg"
    }
    elseif ($extension -ine ".dwg") {
        throw "DrawingFileName must end in .dwg or have no extension: $drawingName"
    }

    if ($drawingName.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars()) -ge 0) {
        throw "DrawingFileName contains invalid filename characters: $drawingName"
    }

    if ($seenNames.ContainsKey($drawingName)) {
        throw "Duplicate DrawingFileName in CSV: $drawingName"
    }
    $seenNames[$drawingName] = $true

    $values = @()
    foreach ($mapping in $mappings) {
        $value = $row.PSObject.Properties[$mapping.Header].Value
        if ($null -eq $value) {
            $value = ""
        }
        if ([string]$value -match '[\x00-\x1F]') {
            throw "CSV value for column '$($mapping.Header)' in drawing '$drawingName' contains control characters."
        }
        $values += [pscustomobject]@{
            Header = $mapping.Header
            Value  = [string]$value
        }
    }

    $resolvedRows += [pscustomobject]@{
        DrawingName = $drawingName
        Values      = $values
    }
}

$OutputDirectory = Resolve-OutputDirectory -Path $OutputDirectory

if ($MaxDrawings -gt 0) {
    $resolvedRows = @($resolvedRows | Select-Object -First $MaxDrawings)
}

$outputPath = [System.IO.Path]::GetFullPath($OutputDirectory)
$generatorDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$lispHelperPath = Join-Path $generatorDirectory "CreateFromTemplate.lsp"
if (-not (Test-Path -LiteralPath $lispHelperPath -PathType Leaf)) {
    throw "Required AutoLISP helper is missing: $lispHelperPath"
}

if ($WhatIfPreference) {
    foreach ($drawing in $resolvedRows) {
        Write-Host "What if: would create $(Join-Path $outputPath $drawing.DrawingName)"
    }
    return
}

New-Item -ItemType Directory -Path $outputPath -Force | Out-Null
$logDirectory = Join-Path $outputPath "logs"

$writeTestPath = Join-Path $outputPath (".batch-write-test-" + [guid]::NewGuid().ToString("N") + ".tmp")
try {
    [System.IO.File]::WriteAllText($writeTestPath, "")
}
finally {
    if (Test-Path -LiteralPath $writeTestPath) {
        Remove-Item -LiteralPath $writeTestPath -Force
    }
}

foreach ($drawing in $resolvedRows) {
    $targetPath = Join-Path $outputPath $drawing.DrawingName
    if (Test-Path -LiteralPath $targetPath) {
        throw "Output already exists; refusing to overwrite it: $targetPath"
    }
}

$approvedRows = @()
foreach ($drawing in $resolvedRows) {
    $targetPath = Join-Path $outputPath $drawing.DrawingName
    if ($PSCmdlet.ShouldProcess($targetPath, "Create drawing from template")) {
        $approvedRows += $drawing
    }
}
if ($approvedRows.Count -eq 0) {
    Write-Host "No drawings were selected for creation."
    return
}

$workRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("AutoCADBatch-" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $workRoot | Out-Null
$preflightGuid = [guid]::NewGuid().ToString("N")
$preflightDrawingPath = Join-Path $outputPath (".batch-preflight-$preflightGuid.dwg")
$stagedDrawings = @()

try {
    $preflightDirectory = Join-Path $workRoot "preflight"
    New-Item -ItemType Directory -Path $preflightDirectory | Out-Null
    $validatedTemplatePath = $preflightDrawingPath
    $preflightScriptPath = Join-Path $preflightDirectory "preflight.scr"
    $preflightResultPath = Join-Path $preflightDirectory "preflight-result.txt"
    $preflightLog = Join-Path $logDirectory "preflight.accoreconsole.log"
    $preflightErrorLog = Join-Path $logDirectory "preflight.accoreconsole.err"
    $preflightMappingList = ConvertTo-LispMappings -Mappings $mappings -Values $approvedRows[0].Values

    Copy-Item -LiteralPath $TemplatePath -Destination $validatedTemplatePath
    $preflightLines = @(
        "(load $(ConvertTo-LispString $lispHelperPath))"
        "(setq *batch-xref-path* $(ConvertTo-LispString ([System.IO.Path]::GetFullPath($BorderXrefPath))))"
        "(setq *batch-xref-name* $(ConvertTo-LispString $BorderXrefName))"
        "(setq *batch-mappings* '$preflightMappingList)"
        "(setq *batch-result-file* $(ConvertTo-LispString $preflightResultPath))"
        "(c:BATCHPREFLIGHT)"
    )
    Write-Utf8File -Path $preflightScriptPath -Content (($preflightLines -join "`r`n") + "`r`n")

    Write-Host "Running template and XREF preflight..."
    $preflightProcess = Invoke-CoreConsole -DrawingPath $validatedTemplatePath -ScriptPath $preflightScriptPath
    $preflightLogText = $preflightProcess.StdOut + [Environment]::NewLine + $preflightProcess.StdErr
    $preflightFailure = $null
    if ($preflightLogText -match '(?i)File load canceled|SECURELOAD') {
        $preflightFailure = "AutoCAD blocked loading the AutoLISP helper '$lispHelperPath'. Add '$generatorDirectory' to AutoCAD's Trusted Locations/TRUSTEDPATHS, or ask IT to approve that folder, then retry."
    }
    elseif ($preflightProcess.ExitCode -ne 0) {
        $preflightDetail = Get-AutoCadErrorText -LogText $preflightLogText
        $preflightFailure = "Core Console template/XREF preflight failed with exit code $($preflightProcess.ExitCode). $preflightDetail"
    }
    elseif (-not (Test-Path -LiteralPath $preflightResultPath -PathType Leaf) -or
        (Get-Content -LiteralPath $preflightResultPath -Raw).Trim() -ne "SUCCESS") {
        $preflightDetail = Get-AutoCadErrorText -LogText $preflightLogText
        $preflightFailure = "Template/XREF preflight did not report success. $preflightDetail"
    }
    elseif (-not (Test-Path -LiteralPath $validatedTemplatePath -PathType Leaf)) {
        $preflightFailure = "Core Console reported preflight success but the prepared template is missing: $validatedTemplatePath"
    }
    if ($preflightFailure) {
        Save-CoreConsoleLogs `
            -StdOut $preflightProcess.StdOut `
            -StdErr $preflightProcess.StdErr `
            -DestinationLog $preflightLog `
            -DestinationErrorLog $preflightErrorLog | Out-Null
        throw "$preflightFailure`nLogs saved to:`n  $preflightLog`n  $preflightErrorLog`nNo drawings were generated."
    }
    if ($KeepLogs) {
        Save-CoreConsoleLogs `
            -StdOut $preflightProcess.StdOut `
            -StdErr $preflightProcess.StdErr `
            -DestinationLog $preflightLog `
            -DestinationErrorLog $preflightErrorLog | Out-Null
    }

    Write-Host "Preflight passed. Generating $($approvedRows.Count) drawing(s) in one Core Console session..."
    $createdCount = 0
    $batchJobs = @()
    $batchScriptLines = @(
        "(load $(ConvertTo-LispString $lispHelperPath))"
        "(setq *batch-jobs* nil)"
    )
    foreach ($drawing in $approvedRows) {
        $targetPath = Join-Path $outputPath $drawing.DrawingName
        $stageId = [guid]::NewGuid().ToString("N")
        $stagedDrawing = Join-Path $outputPath (".batch-$stageId-" + $drawing.DrawingName)
        $stagedDrawings += $stagedDrawing
        $resultPath = Join-Path $preflightDirectory ("result-" + $stageId + ".txt")
        $mappingList = ConvertTo-LispMappings -Mappings $mappings -Values $drawing.Values

        $batchJobs += [pscustomobject]@{
            DrawingName = $drawing.DrawingName
            TargetPath  = $targetPath
            StagedPath  = $stagedDrawing
            ResultPath  = $resultPath
        }
        $batchScriptLines += "(setq *batch-jobs* (cons (list $(ConvertTo-LispString $stagedDrawing) $(ConvertTo-LispString $resultPath) '$mappingList) *batch-jobs*))"
    }
    $batchScriptLines += "(setq *batch-jobs* (reverse *batch-jobs*))"
    $batchScriptLines += "(c:CREATEBATCH)"
    $batchScriptPath = Join-Path $preflightDirectory "generate.scr"
    Write-Utf8File -Path $batchScriptPath -Content (($batchScriptLines -join "`r`n") + "`r`n")

    $batchLog = Join-Path $logDirectory "batch.accoreconsole.log"
    $batchErrorLog = Join-Path $logDirectory "batch.accoreconsole.err"
    $batchProcess = Invoke-CoreConsole -DrawingPath $validatedTemplatePath -ScriptPath $batchScriptPath
    $batchLogText = $batchProcess.StdOut + [Environment]::NewLine + $batchProcess.StdErr
    $batchError = Get-AutoCadErrorText -LogText $batchLogText
    $batchFailure = $null

    foreach ($job in $batchJobs) {
        $jobSucceeded = Test-Path -LiteralPath $job.ResultPath -PathType Leaf
        if ($jobSucceeded) {
            $jobSucceeded = (Get-Content -LiteralPath $job.ResultPath -Raw).Trim() -eq "SUCCESS"
        }
        if (-not $jobSucceeded) {
            $batchFailure = "Core Console did not report successful completion for '$($job.DrawingName)'."
            if ($batchError) {
                $batchFailure += " $batchError"
            }
            break
        }
        if (-not (Test-Path -LiteralPath $job.StagedPath -PathType Leaf)) {
            $batchFailure = "Core Console reported success for '$($job.DrawingName)', but its staged drawing is missing."
            break
        }
        if (Test-Path -LiteralPath $job.TargetPath) {
            $batchFailure = "Output appeared during processing; refusing to overwrite it: $($job.TargetPath)"
            break
        }

        Move-Item -LiteralPath $job.StagedPath -Destination $job.TargetPath
        Write-Host "Created $($job.TargetPath)"
        $createdCount++
    }

    if (-not $batchFailure -and $batchProcess.ExitCode -ne 0) {
        $batchFailure = "Core Console exited with code $($batchProcess.ExitCode) after completing the drawing sequence."
        if ($batchError) {
            $batchFailure += " $batchError"
        }
    }

    if ($batchFailure) {
        Save-CoreConsoleLogs `
            -StdOut $batchProcess.StdOut `
            -StdErr $batchProcess.StdErr `
            -DestinationLog $batchLog `
            -DestinationErrorLog $batchErrorLog | Out-Null
        throw "$batchFailure`nCreated $createdCount of $($approvedRows.Count) drawing(s) before stopping.`nLogs saved to:`n  $batchLog`n  $batchErrorLog"
    }

    if ($KeepLogs) {
        Save-CoreConsoleLogs `
            -StdOut $batchProcess.StdOut `
            -StdErr $batchProcess.StdErr `
            -DestinationLog $batchLog `
            -DestinationErrorLog $batchErrorLog | Out-Null
    }

    Write-Host "Batch complete. Created $createdCount drawing(s)."
}
finally {
    foreach ($stagedDrawing in $stagedDrawings) {
        if (Test-Path -LiteralPath $stagedDrawing) {
            Remove-Item -LiteralPath $stagedDrawing -Force
        }
    }
    if (Test-Path -LiteralPath $preflightDrawingPath) {
        Remove-Item -LiteralPath $preflightDrawingPath -Force
    }
    if (Test-Path -LiteralPath $workRoot) {
        Remove-Item -LiteralPath $workRoot -Recurse -Force
    }
}
}
catch {
    $failureMessage = $_.Exception.Message.Trim()
    $failureMessage = $failureMessage -replace '(?im)^\s*error:\s*quit\s*/\s*exit\s+abort\s*$', ''
    $failureMessage = $failureMessage.Trim()

    Write-Host ""
    Write-Host "Batch drawing generation failed:" -ForegroundColor Red
    foreach ($line in ($failureMessage -split "`r?`n")) {
        if (-not [string]::IsNullOrWhiteSpace($line)) {
            Write-Host "  $line"
        }
    }

    $global:LASTEXITCODE = 1
    Write-Error -Message "Batch drawing generation failed." -ErrorAction SilentlyContinue
}
