<#
Compress-XlfFormula.ps1

Purpose:
- Default behavior is entire formula minification.
- Optional behavior: keep header readable and minify only after a token line.

Quick usage:
1) Entirely minify (default mode):
   powershell -ExecutionPolicy Bypass -File "d:\Utilities\My_Scripts_Apps\Excel\Compress-XlfFormula.ps1" -InputPath "d:\Utilities\My_Scripts_Apps\Excel\Tag_Databases\Modbus\OldDescription.xlf"

2) Minify after a token line and copy to clipboard:
   powershell -ExecutionPolicy Bypass -File "d:\Utilities\My_Scripts_Apps\Excel\Compress-XlfFormula.ps1" -InputPath "d:\Utilities\My_Scripts_Apps\Excel\Tag_Databases\Modbus\OldDescription.xlf" -TokenString "cleanText," -CopyToClipboard

3) Explicitly request entire minify:
   powershell -ExecutionPolicy Bypass -File "d:\Utilities\My_Scripts_Apps\Excel\Compress-XlfFormula.ps1" -InputPath "d:\Utilities\My_Scripts_Apps\Excel\Tag_Databases\Modbus\OldDescription.xlf" -MinifyAll

Copilot reminder:
- If the user says "compress ... using token ...", pass -TokenString.
- If the user does not provide a token, run default entire minify or ask if they wanted token mode.
#>

param(
    [string]$InputPath,

    [string]$OutputPath,

    [string]$TokenString,

    [switch]$CopyToClipboard,

    [switch]$MinifyAll,

    [switch]$Help
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Show-Usage {
    Write-Output "Usage:"
    Write-Output "  Default (entirely minified):"
    Write-Output "    .\Compress-XlfFormula.ps1 -InputPath <file.xlf> [-OutputPath <file>] [-CopyToClipboard]"
    Write-Output ""
    Write-Output "  Token mode (header kept, minify after token line):"
    Write-Output "    .\Compress-XlfFormula.ps1 -InputPath <file.xlf> -TokenString <token line> [-OutputPath <file>] [-CopyToClipboard]"
    Write-Output ""
    Write-Output "  Notes:"
    Write-Output "  - Use either -TokenString or -MinifyAll, not both."
    Write-Output "  - If neither is provided, default is entire minify."
    Write-Output "  - Token must match a full line (leading/trailing whitespace ignored)."
    Write-Output "  - If -CopyToClipboard is used without -OutputPath, no output file is written."
}

function Compress-ExcelFormulaText {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text
    )

    $sb = New-Object System.Text.StringBuilder
    $inString = $false
    $bracketDepth = 0

    for ($i = 0; $i -lt $Text.Length; $i++) {
        $ch = $Text[$i]

        if ($inString) {
            [void]$sb.Append($ch)

            if ($ch -eq '"') {
                # Excel escaped quote inside string literal: ""
                if ($i + 1 -lt $Text.Length -and $Text[$i + 1] -eq '"') {
                    [void]$sb.Append($Text[$i + 1])
                    $i++
                }
                else {
                    $inString = $false
                }
            }
        }
        else {
            if ($ch -eq '"') {
                $inString = $true
                [void]$sb.Append($ch)
            }
            elseif ($ch -eq '[') {
                $bracketDepth++
                [void]$sb.Append($ch)
            }
            elseif ($ch -eq ']' -and $bracketDepth -gt 0) {
                $bracketDepth--
                [void]$sb.Append($ch)
            }
            elseif ($bracketDepth -gt 0) {
                # Preserve structured reference content exactly, including spaces.
                [void]$sb.Append($ch)
            }
            elseif ([char]::IsWhiteSpace($ch)) {
                continue
            }
            else {
                [void]$sb.Append($ch)
            }
        }
    }

    return $sb.ToString()
}

if ($Help) {
    Show-Usage
    return
}

if ([string]::IsNullOrWhiteSpace($InputPath)) {
    Show-Usage
    throw "Missing required -InputPath."
}

if (-not (Test-Path -LiteralPath $InputPath)) {
    throw "Input file not found: $InputPath"
}

$explicitOutputPath = $PSBoundParameters.ContainsKey('OutputPath')
$writeOutputFile = $true

if ($CopyToClipboard -and -not $explicitOutputPath) {
    $writeOutputFile = $false
}

if ($writeOutputFile -and -not $OutputPath) {
    $inputItem = Get-Item -LiteralPath $InputPath
    $OutputPath = Join-Path -Path $inputItem.DirectoryName -ChildPath ($inputItem.BaseName + ".paste" + $inputItem.Extension)
}

$useTokenMode = -not [string]::IsNullOrWhiteSpace($TokenString)

if ($useTokenMode -and $MinifyAll) {
    Show-Usage
    throw "Use either -TokenString or -MinifyAll, not both."
}

$raw = Get-Content -LiteralPath $InputPath -Raw

if ([string]::IsNullOrWhiteSpace($raw)) {
    throw "Input file is empty: $InputPath"
}

$minified = $null
$modeLabel = ""

if ($useTokenMode) {
    $tokenRegex = [regex]::new("(?m)^\s*" + [regex]::Escape($TokenString) + "\s*$")
    $match = $tokenRegex.Match($raw)

    if (-not $match.Success) {
        throw "Could not find token line '$TokenString' in $InputPath"
    }

    $header = $raw.Substring(0, $match.Index).TrimEnd("`r", "`n")
    $body = $raw.Substring($match.Index)
    $bodyMin = Compress-ExcelFormulaText -Text $body

    if ([string]::IsNullOrWhiteSpace($header)) {
        $minified = $bodyMin
    }
    else {
        # Keep header readable, minify only after token line.
        $minified = $header + "`r`n" + $bodyMin
    }

    $modeLabel = "minified after '$TokenString'"
}
else {
    $minified = Compress-ExcelFormulaText -Text $raw
    $modeLabel = "entirely minified"
}

if ($writeOutputFile) {
    Set-Content -LiteralPath $OutputPath -Value $minified -NoNewline
}

if ($CopyToClipboard) {
    Set-Clipboard -Value $minified
}

$sourceLen = $raw.Length
$outLen = $minified.Length
$delta = $sourceLen - $outLen
$pct = if ($sourceLen -gt 0) { [math]::Round(($delta / $sourceLen) * 100, 2) } else { 0 }

Write-Output "Mode: $modeLabel"
Write-Output "Input:  $InputPath"
if ($writeOutputFile) {
    Write-Output "Output: $OutputPath"
}
else {
    Write-Output "Output: (skipped; clipboard-only mode)"
}
Write-Output "Source chars: $sourceLen"
Write-Output "Output chars: $outLen"
Write-Output "Reduced by: $delta chars ($pct%)"

if ($CopyToClipboard) {
    Write-Output "Clipboard: updated"
}

if ($outLen -gt 8192) {
    Write-Warning "Output exceeds Excel 8192-character limit by $($outLen - 8192) chars."
}
else {
    Write-Output "Within Excel 8192-character limit (remaining: $([int](8192 - $outLen)))."
}
