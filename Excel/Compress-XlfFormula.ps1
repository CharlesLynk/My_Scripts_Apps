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

    [switch]$MinifyLetVariables,

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
    Write-Output "  - LET() variable names are NOT minified by default."
    Write-Output "  - Use -MinifyLetVariables to shrink LET names (aaa, aab, ...)."
    Write-Output "  - Token must match a full line (leading/trailing whitespace ignored)."
    Write-Output "  - If -CopyToClipboard is used without -OutputPath, no output file is written."
}

function New-CaseInsensitiveStringMap {
    return New-Object 'System.Collections.Generic.Dictionary[string,string]' ([System.StringComparer]::OrdinalIgnoreCase)
}

function New-ReservedLetNameSet {
    $set = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

    # Avoid names that can be parsed as built-in Excel functions or special tokens.
    foreach ($name in @(
            'ABS', 'ACOS', 'ACOSH', 'ACOT', 'ACOTH', 'AGGREGATE', 'AND', 'ASIN', 'ASINH', 'ATAN', 'ATAN2', 'ATANH',
            'AVERAGE', 'BYCOL', 'BYROW', 'CEILING', 'CHOOSE', 'CHOOSECOLS', 'CHOOSEROWS', 'CLEAN', 'COLUMN', 'COLUMNS',
            'CONCAT', 'COUNT', 'COUNTA', 'COUNTIF', 'COUNTIFS', 'DAY', 'DROP', 'EXP', 'FALSE', 'FILTER', 'FIND',
            'HSTACK', 'IF', 'IFERROR', 'IFNA', 'IFS', 'INDEX', 'INDIRECT', 'INT', 'ISBLANK', 'ISERROR', 'ISNUMBER',
            'LAMBDA', 'LEFT', 'LEN', 'LET', 'LOOKUP', 'LOWER', 'MAP', 'MATCH', 'MAX', 'MID', 'MIN', 'MMULT', 'MOD',
            'MONTH', 'NA', 'NOT', 'NOW', 'OFFSET', 'OR', 'RAND', 'RANDBETWEEN', 'REDUCE', 'REPLACE', 'RIGHT', 'ROUND',
            'ROW', 'ROWS', 'SCAN', 'SEQUENCE', 'SORT', 'SUBSTITUTE', 'SUM', 'SUMIF', 'SUMIFS', 'SWITCH', 'TAKE', 'T',
            'TEXT', 'TEXTAFTER', 'TEXTBEFORE', 'TEXTJOIN', 'TEXTSPLIT', 'TOCOL', 'TOROW', 'TRIM', 'TRUE', 'UNIQUE',
            'UPPER', 'VALUE', 'VLOOKUP', 'VSTACK', 'WRAPCOLS', 'WRAPROWS', 'XLOOKUP', 'XMATCH', 'YEAR', 'N'
        )) {
        [void]$set.Add($name)
    }

    return ,$set
}

function Test-IdentifierStart {
    param(
        [char]$Char
    )

    return ([char]::IsLetter($Char) -or $Char -eq '_' -or $Char -eq '\\')
}

function Test-IdentifierPart {
    param(
        [char]$Char
    )

    return ([char]::IsLetterOrDigit($Char) -or $Char -eq '_' -or $Char -eq '.' -or $Char -eq '\\')
}

function Skip-Whitespace {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text,

        [Parameter(Mandatory = $true)]
        [int]$Index,

        [Parameter(Mandatory = $true)]
        [int]$EndExclusive
    )

    $cursor = $Index
    while ($cursor -lt $EndExclusive -and [char]::IsWhiteSpace($Text[$cursor])) {
        $cursor++
    }

    return $cursor
}

function Test-CellReferenceLike {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Token
    )

    # Avoid rewriting invalid LET names and obvious A1/R1C1-like references.
    return ($Token -match '^(?:\$?[A-Za-z]{1,3}\$?[0-9]{1,7}|R(?:\[?-?[0-9]+\]?|)C(?:\[?-?[0-9]+\]?|))$')
}

function Get-LetVariableCandidate {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RawText
    )

    $candidate = $RawText.Trim()
    if ([string]::IsNullOrWhiteSpace($candidate)) {
        return $null
    }

    if ($candidate -notmatch '^[A-Za-z_\\][A-Za-z0-9_\\.]*$') {
        return $null
    }

    if ($candidate -match '^(TRUE|FALSE)$') {
        return $null
    }

    if (Test-CellReferenceLike -Token $candidate) {
        return $null
    }

    return $candidate
}

function New-MinifiedLetName {
    param(
        [Parameter(Mandatory = $true)]
        [int]$Index
    )

    if ($Index -lt 0) {
        throw "Name index cannot be negative: $Index"
    }

    $value = $Index
    $chars = New-Object System.Collections.Generic.List[char]

    do {
        [void]$chars.Add([char]([int][char]'a' + ($value % 26)))
        $value = [int]($value / 26)
    } while ($value -gt 0)

    $arr = $chars.ToArray()
    [array]::Reverse($arr)

    $name = -join $arr
    if ($name.Length -lt 3) {
        $name = ("a" * (3 - $name.Length)) + $name
    }

    return $name
}

function Get-NextSafeMinifiedLetName {
    param(
        [Parameter(Mandatory = $true)]
        [ref]$NameCounter,

        [System.Collections.Generic.HashSet[string]]$ReservedNames,

        [object[]]$Scopes
    )

    while ($true) {
        $candidate = New-MinifiedLetName -Index $NameCounter.Value
        $NameCounter.Value = $NameCounter.Value + 1

        if ($null -ne $ReservedNames -and $ReservedNames.Contains($candidate)) {
            continue
        }

        if ($null -ne $script:AllocatedLetNames -and $script:AllocatedLetNames.Contains($candidate)) {
            continue
        }

        if (Test-CellReferenceLike -Token $candidate) {
            continue
        }

        if ($null -ne $Scopes) {
            $scopeList = @()
            $scopeList += $Scopes
            $isInUse = $false

            for ($scopeIndex = $scopeList.Count - 1; $scopeIndex -ge 0; $scopeIndex--) {
                $scope = $scopeList[$scopeIndex]
                if ($null -ne $scope) {
                    if ($scope.ContainsKey($candidate)) {
                        $isInUse = $true
                        break
                    }

                    if ($scope.Values -contains $candidate) {
                        $isInUse = $true
                        break
                    }
                }
            }

            if ($isInUse) {
                continue
            }
        }

        if ($null -ne $script:AllocatedLetNames) {
            [void]$script:AllocatedLetNames.Add($candidate)
        }

        return $candidate
    }
}

function Resolve-ScopedName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Token,

        [object[]]$Scopes
    )

    if ($null -eq $Scopes) {
        return $Token
    }

    $scopeList = @()
    $scopeList += $Scopes

    for ($scopeIndex = $scopeList.Count - 1; $scopeIndex -ge 0; $scopeIndex--) {
        $scope = $scopeList[$scopeIndex]
        if ($null -ne $scope -and $scope.ContainsKey($Token)) {
            return $scope[$Token]
        }
    }

    return $Token
}

function Get-CallArguments {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text,

        [Parameter(Mandatory = $true)]
        [int]$OpenParenIndex,

        [Parameter(Mandatory = $true)]
        [int]$EndExclusive
    )

    $args = New-Object System.Collections.Generic.List[object]
    $argStart = $OpenParenIndex + 1
    $parenDepth = 1
    $inString = $false
    $bracketDepth = 0
    $braceDepth = 0
    $closeIndex = -1

    for ($cursor = $OpenParenIndex + 1; $cursor -lt $EndExclusive; $cursor++) {
        $ch = $Text[$cursor]

        if ($inString) {
            if ($ch -eq '"') {
                # Excel escaped quote inside string literal: ""
                if ($cursor + 1 -lt $EndExclusive -and $Text[$cursor + 1] -eq '"') {
                    $cursor++
                }
                else {
                    $inString = $false
                }
            }

            continue
        }

        if ($ch -eq '"') {
            $inString = $true
            continue
        }

        if ($ch -eq '[') {
            $bracketDepth++
            continue
        }

        if ($ch -eq ']' -and $bracketDepth -gt 0) {
            $bracketDepth--
            continue
        }

        if ($bracketDepth -gt 0) {
            continue
        }

        if ($ch -eq '{') {
            $braceDepth++
            continue
        }

        if ($ch -eq '}' -and $braceDepth -gt 0) {
            $braceDepth--
            continue
        }

        if ($ch -eq '(') {
            $parenDepth++
            continue
        }

        if ($ch -eq ')') {
            $parenDepth--

            if ($parenDepth -eq 0) {
                [void]$args.Add([pscustomobject]@{ Start = $argStart; End = $cursor })
                $closeIndex = $cursor
                break
            }

            continue
        }

        if ($ch -eq ',' -and $parenDepth -eq 1 -and $braceDepth -eq 0) {
            [void]$args.Add([pscustomobject]@{ Start = $argStart; End = $cursor })
            $argStart = $cursor + 1
        }
    }

    if ($closeIndex -lt 0) {
        return $null
    }

    return [pscustomobject]@{
        Arguments  = $args
        CloseIndex = $closeIndex
    }
}

function Try-RewriteLetCall {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text,

        [Parameter(Mandatory = $true)]
        [int]$TokenStart,

        [Parameter(Mandatory = $true)]
        [int]$TokenEnd,

        [Parameter(Mandatory = $true)]
        [int]$OpenParenIndex,

        [Parameter(Mandatory = $true)]
        [int]$EndExclusive,

        [object[]]$Scopes,

        [Parameter(Mandatory = $true)]
        [ref]$NameCounter
    )

    $callInfo = Get-CallArguments -Text $Text -OpenParenIndex $OpenParenIndex -EndExclusive $EndExclusive
    if ($null -eq $callInfo) {
        return $null
    }

    $args = $callInfo.Arguments
    $rewrittenArgs = New-Object System.Collections.Generic.List[string]
    $baseScopes = @()
    if ($null -ne $Scopes) {
        $baseScopes += $Scopes
    }

    if ($args.Count -ge 3 -and ($args.Count % 2 -eq 1)) {
        $declaredScope = New-CaseInsensitiveStringMap
        $pairCount = [int](($args.Count - 1) / 2)

        for ($pairIndex = 0; $pairIndex -lt $pairCount; $pairIndex++) {
            $nameArg = $args[$pairIndex * 2]
            $valueArg = $args[$pairIndex * 2 + 1]
            $currentScopes = @()
            $currentScopes += $baseScopes
            $currentScopes += $declaredScope

            $nameRaw = $Text.Substring($nameArg.Start, $nameArg.End - $nameArg.Start)
            $nameCandidate = Get-LetVariableCandidate -RawText $nameRaw
            $newName = $null

            if ($null -ne $nameCandidate) {
                $newName = Get-NextSafeMinifiedLetName -NameCounter $NameCounter -ReservedNames $script:ReservedLetNames -Scopes $currentScopes
                [void]$rewrittenArgs.Add($newName)
            }
            else {
                $rewrittenNameArg = Rewrite-FormulaSegment -Text $Text -StartIndex $nameArg.Start -EndExclusive $nameArg.End -Scopes $currentScopes -NameCounter $NameCounter
                [void]$rewrittenArgs.Add($rewrittenNameArg)
            }

            $rewrittenValueArg = Rewrite-FormulaSegment -Text $Text -StartIndex $valueArg.Start -EndExclusive $valueArg.End -Scopes $currentScopes -NameCounter $NameCounter
            [void]$rewrittenArgs.Add($rewrittenValueArg)

            if ($null -ne $nameCandidate -and $null -ne $newName) {
                $declaredScope[$nameCandidate] = $newName
            }
        }

        $finalArg = $args[$args.Count - 1]
        $finalScopes = @()
        $finalScopes += $baseScopes
        $finalScopes += $declaredScope
        $rewrittenFinalArg = Rewrite-FormulaSegment -Text $Text -StartIndex $finalArg.Start -EndExclusive $finalArg.End -Scopes $finalScopes -NameCounter $NameCounter
        [void]$rewrittenArgs.Add($rewrittenFinalArg)
    }
    else {
        foreach ($arg in $args) {
            $rewrittenArg = Rewrite-FormulaSegment -Text $Text -StartIndex $arg.Start -EndExclusive $arg.End -Scopes $baseScopes -NameCounter $NameCounter
            [void]$rewrittenArgs.Add($rewrittenArg)
        }
    }

    $tokenText = $Text.Substring($TokenStart, $TokenEnd - $TokenStart)
    $rewrittenCall = $tokenText + "(" + ([string]::Join(",", $rewrittenArgs)) + ")"

    return [pscustomobject]@{
        Text      = $rewrittenCall
        NextIndex = $callInfo.CloseIndex + 1
    }
}

function Rewrite-FormulaSegment {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text,

        [Parameter(Mandatory = $true)]
        [int]$StartIndex,

        [Parameter(Mandatory = $true)]
        [int]$EndExclusive,

        [object[]]$Scopes,

        [Parameter(Mandatory = $true)]
        [ref]$NameCounter
    )

    $sb = New-Object System.Text.StringBuilder
    $cursor = $StartIndex
    $inString = $false
    $bracketDepth = 0

    while ($cursor -lt $EndExclusive) {
        $ch = $Text[$cursor]

        if ($inString) {
            [void]$sb.Append($ch)

            if ($ch -eq '"') {
                # Excel escaped quote inside string literal: ""
                if ($cursor + 1 -lt $EndExclusive -and $Text[$cursor + 1] -eq '"') {
                    [void]$sb.Append($Text[$cursor + 1])
                    $cursor += 2
                    continue
                }

                $inString = $false
            }

            $cursor++
            continue
        }

        if ($ch -eq '"') {
            $inString = $true
            [void]$sb.Append($ch)
            $cursor++
            continue
        }

        if ($ch -eq '[') {
            $bracketDepth++
            [void]$sb.Append($ch)
            $cursor++
            continue
        }

        if ($ch -eq ']' -and $bracketDepth -gt 0) {
            $bracketDepth--
            [void]$sb.Append($ch)
            $cursor++
            continue
        }

        if ($bracketDepth -gt 0) {
            [void]$sb.Append($ch)
            $cursor++
            continue
        }

        if (Test-IdentifierStart -Char $ch) {
            $tokenStart = $cursor
            $cursor++

            while ($cursor -lt $EndExclusive -and (Test-IdentifierPart -Char $Text[$cursor])) {
                $cursor++
            }

            $token = $Text.Substring($tokenStart, $cursor - $tokenStart)
            $nextNonWs = Skip-Whitespace -Text $Text -Index $cursor -EndExclusive $EndExclusive

            if ($token -ieq 'LET' -and $nextNonWs -lt $EndExclusive -and $Text[$nextNonWs] -eq '(') {
                $rewrittenLet = Try-RewriteLetCall -Text $Text -TokenStart $tokenStart -TokenEnd $cursor -OpenParenIndex $nextNonWs -EndExclusive $EndExclusive -Scopes $Scopes -NameCounter $NameCounter

                if ($null -ne $rewrittenLet) {
                    [void]$sb.Append($rewrittenLet.Text)
                    $cursor = $rewrittenLet.NextIndex
                    continue
                }
            }

            $renamedToken = $token
            if (-not ($nextNonWs -lt $EndExclusive -and $Text[$nextNonWs] -eq '!')) {
                $renamedToken = Resolve-ScopedName -Token $token -Scopes $Scopes
            }

            [void]$sb.Append($renamedToken)
            continue
        }

        [void]$sb.Append($ch)
        $cursor++
    }

    return $sb.ToString()
}

function Minify-LetVariableNames {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text
    )

    if ([string]::IsNullOrEmpty($Text)) {
        return $Text
    }

    $script:AllocatedLetNames = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $nameCounter = 0
    return Rewrite-FormulaSegment -Text $Text -StartIndex 0 -EndExclusive $Text.Length -Scopes @() -NameCounter ([ref]$nameCounter)
}

$script:ReservedLetNames = New-ReservedLetNameSet

function Compress-ExcelFormulaText {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text,

        [switch]$MinifyLetVariables
    )

    $workText = $Text
    if ($MinifyLetVariables) {
        $workText = Minify-LetVariableNames -Text $workText
    }

    $sb = New-Object System.Text.StringBuilder
    $inString = $false
    $bracketDepth = 0

    for ($i = 0; $i -lt $workText.Length; $i++) {
        $ch = $workText[$i]

        if ($inString) {
            [void]$sb.Append($ch)

            if ($ch -eq '"') {
                # Excel escaped quote inside string literal: ""
                if ($i + 1 -lt $workText.Length -and $workText[$i + 1] -eq '"') {
                    [void]$sb.Append($workText[$i + 1])
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
$enableLetVariableMinify = $MinifyLetVariables

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
    $bodyMin = Compress-ExcelFormulaText -Text $body -MinifyLetVariables:$enableLetVariableMinify

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
    $minified = Compress-ExcelFormulaText -Text $raw -MinifyLetVariables:$enableLetVariableMinify
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
Write-Output "LET variable minify: $(if ($enableLetVariableMinify) { 'enabled' } else { 'disabled' })"
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
