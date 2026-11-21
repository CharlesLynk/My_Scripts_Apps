# Lists, sets, fits, and restores label FontSize values in VSD Craft profile manifests.
# Usage examples:
#   .\Set-StreamDockFontSize.ps1 -Mode List
#   .\Set-StreamDockFontSize.ps1 -Mode Fixed -ProfileName AutoCAD -Page 1 -Size 8 -WhatIf
#   .\Set-StreamDockFontSize.ps1 -Mode Fixed -ProfileName AutoCAD -Size 8 -Force
#   .\Set-StreamDockFontSize.ps1 -Mode Preset -PresetName Fit-Compact -ProfileName AutoCAD -Force
#   .\Set-StreamDockFontSize.ps1 -Mode Restore -BackupStamp 20250101-120000 -Force

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateSet('List', 'Fixed', 'Fit', 'Preset', 'Restore', 'Render', 'Apply')]
    [string]$Mode,

    [string]$ProfileName,

    # 0 means every page of the profile.
    [int]$Page = 0,

    [int]$Size,

    [int]$Min = 6,
    [int]$Max = 12,
    [int]$Margin = 4,

    [string]$PresetName,

    [string]$ProfilesRoot = 'C:\Users\clynk\AppData\Roaming\HotSpot\StreamDock\profiles',
    [string]$BackupRoot,
    [string]$BackupStamp,
    [string]$GeneratedRoot,
    [string]$RenderFont = 'Segoe UI',

    [switch]$Force
)

$ErrorActionPreference = 'Stop'

function Read-VsdJson {
    param([string]$Path)
    return (Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json)
}

# ConvertTo-Json on Windows PowerShell 5.1 escapes some characters, adds spacing after colons,
# leaves blank lines, and indents inconsistently. This rewrites its output into VSD Craft's style
# (4-space indent, "key": value) so an unchanged manifest saves back byte-identical.
function ConvertTo-VsdJson {
    param([string]$Json)
    $text = $Json -replace '\\u0027', "'" -replace '\\u003c', '<' -replace '\\u003e', '>' -replace '\\u0026', '&'
    $lines = $text -split "`r?`n" | Where-Object { $_.Trim() -ne '' }
    $sb = New-Object System.Text.StringBuilder
    $depth = 0
    foreach ($line in $lines) {
        $body = $line.Trim() -replace '^("(?:[^"\\]|\\.)*"):\s+', '$1: '
        if ($body -match '^[}\]]') { $depth = [Math]::Max(0, $depth - 1) }
        [void]$sb.Append(' ' * (4 * $depth)).Append($body).Append("`n")
        if ($body -match '[{\[]$') { $depth++ }
    }
    return $sb.ToString()
}

function Save-VsdJson {
    param([string]$Path, $Object)
    $json = ConvertTo-Json -InputObject $Object -Depth 100
    $text = ConvertTo-VsdJson -Json $json
    [System.IO.File]::WriteAllText($Path, $text, (New-Object System.Text.UTF8Encoding($false)))
}

function Get-VsdRoots {
    param([string]$Root)
    Get-ChildItem -LiteralPath $Root -Directory -Filter '*.sdProfile' -ErrorAction SilentlyContinue | ForEach-Object {
        $mf = Join-Path $_.FullName 'manifest.json'
        if (Test-Path -LiteralPath $mf) {
            $m = Read-VsdJson -Path $mf
            [pscustomobject]@{ Name = [string]$m.Name; Dir = $_.FullName; Manifest = $mf }
        }
    }
}

# Page 1 is always the root manifest. Other pages follow Pages.Pages, skipping the root's own entry.
function Get-VsdPages {
    param($Root)
    $m = Read-VsdJson -Path $Root.Manifest
    $rootLeaf = Split-Path -Leaf $Root.Dir
    $pages = @([pscustomobject]@{ Page = 1; Path = $Root.Manifest; Entry = $rootLeaf })
    $n = 1
    if ($m.Pages -and $m.Pages.Pages) {
        foreach ($e in @($m.Pages.Pages)) {
            if ($e -eq $rootLeaf) { continue }
            $n++
            $pages += [pscustomobject]@{
                Page  = $n
                Path  = (Join-Path $Root.Dir ('profiles\' + $e + '\manifest.json'))
                Entry = $e
            }
        }
    }
    return $pages
}

function Assert-VsdCraftClosed {
    $running = Get-Process -Name 'VSD Craft' -ErrorAction SilentlyContinue
    if ($running -and -not $Force) {
        throw 'VSD Craft is running and may overwrite edits. Close it, or pass -Force to write anyway.'
    }
    if ($running) { Write-Warning 'VSD Craft is running. Writing anyway because -Force was given.' }
}

function Backup-VsdFiles {
    param([string[]]$Paths, [string]$Root, [string]$Destination)
    $prefix = $Root.TrimEnd('\') + '\'
    foreach ($p in $Paths) {
        $rel = $p.Substring($prefix.Length)
        $target = Join-Path $Destination $rel
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $target) | Out-Null
        Copy-Item -LiteralPath $p -Destination $target -Force
    }
}

# Approximate glyph widths in em. Uncalibrated: replace with GDI measurement if fit results look wrong.
function Get-GlyphEm {
    param([char]$Ch)
    if ($Ch -eq ' ') { return 0.30 }
    if ([char]::IsUpper($Ch)) { return 0.70 }
    if ([char]::IsDigit($Ch)) { return 0.60 }
    if ([char]::IsLower($Ch)) { return 0.55 }
    return 0.45
}

# Returns the widest line in em and the number of lines.
function Get-LabelMetrics {
    param([string]$Title)
    $lines = @($Title -split "`r?`n")
    $widest = 0.0
    foreach ($line in $lines) {
        $w = 0.0
        foreach ($ch in $line.ToCharArray()) { $w += Get-GlyphEm -Ch $ch }
        if ($w -gt $widest) { $widest = $w }
    }
    return [pscustomobject]@{ WidthEm = $widest; Lines = $lines.Count }
}

# Fits when the widest line and the total line height both fit inside the key box.
function Test-LabelFits {
    param([string]$Title, [double]$Size, [double]$Box)
    $m = Get-LabelMetrics -Title $Title
    return (($m.WidthEm * $Size) -le $Box) -and (($m.Lines * $Size * 1.15) -le $Box)
}

# Largest whole size in [Min, Max] where every label in the list fits. Returns $null if none fit.
function Get-FitSize {
    param([string[]]$Titles, [int]$Min, [int]$Max, [double]$Box)
    $best = $null
    for ($s = $Max; $s -ge $Min; $s--) {
        $ok = $true
        foreach ($t in $Titles) {
            if (-not (Test-LabelFits -Title $t -Size $s -Box $Box)) { $ok = $false; break }
        }
        if ($ok) { $best = $s; break }
    }
    return $best
}

# Returns one record per titled state (label) across the given pages.
function Get-VsdLabels {
    param($Pages)
    foreach ($pg in $Pages) {
        if (-not (Test-Path -LiteralPath $pg.Path)) { continue }
        $m = Read-VsdJson -Path $pg.Path
        if (-not $m.Actions) { continue }
        foreach ($prop in $m.Actions.PSObject.Properties) {
            $states = @($prop.Value.States)
            for ($i = 0; $i -lt $states.Count; $i++) {
                $st = $states[$i]
                if ($st.Title -and ([string]$st.Title).Trim() -ne '') {
                    [pscustomobject]@{
                        Page     = $pg.Page
                        Path     = $pg.Path
                        Key      = $prop.Name
                        State    = $i
                        Title    = [string]$st.Title
                        FontSize = $st.FontSize
                    }
                }
            }
        }
    }
}

# Sets FontSize on every titled state of one manifest. $Size is a scriptblock that takes the title and returns a size.
# Keeps the original JSON type (string or number). Returns the number of states changed.
function Update-PageFontSize {
    param($Manifest, [scriptblock]$Size)
    $changed = 0
    foreach ($prop in $Manifest.Actions.PSObject.Properties) {
        foreach ($st in @($prop.Value.States)) {
            if ($st.Title -and ([string]$st.Title).Trim() -ne '') {
                $new = & $Size ([string]$st.Title)
                if ($null -eq $new) { continue }
                $hasSize = $st.PSObject.Properties['FontSize'] -ne $null
                $value = if ($hasSize -and $st.FontSize -is [string]) { [string]$new } else { [int]$new }
                if (-not $hasSize) {
                    Add-Member -InputObject $st -NotePropertyName FontSize -NotePropertyValue $value
                    $changed++
                }
                elseif ([string]$st.FontSize -ne [string]$value) {
                    $st.FontSize = $value
                    $changed++
                }
            }
        }
    }
    return $changed
}

function Get-RelativePath {
    param([string]$Root, [string]$Path)
    $prefix = $Root.TrimEnd('\') + '\'
    return $Path.Substring($prefix.Length)
}

# Finds the one profile whose display name matches, and its pages (optionally one page).
function Resolve-Scope {
    param([string]$Name, [int]$PageNo)
    $roots = @(Get-VsdRoots -Root $ProfilesRoot)
    $hits = @($roots | Where-Object { $_.Name -eq $Name })
    if ($hits.Count -eq 0) {
        throw "No profile named '$Name'. Available: $(($roots | ForEach-Object { $_.Name }) -join ', ')"
    }
    if ($hits.Count -gt 1) { throw "Profile name '$Name' is ambiguous ($($hits.Count) matches)." }
    $pages = @(Get-VsdPages -Root $hits[0])
    if ($PageNo -gt 0) { $pages = @($pages | Where-Object { $_.Page -eq $PageNo }) }
    if ($pages.Count -eq 0) { throw "Profile '$Name' has no page $PageNo." }
    return [pscustomobject]@{ Name = $Name; Root = $hits[0]; Pages = $pages }
}

# Reads each page, applies $Size, and reports the plan. Writes only when not a dry run.
# Backs up changed files first and restores them if a write fails.
function Invoke-SizeChange {
    param($Scope, [scriptblock]$Size, [string]$Describe)
    $plan = @()
    foreach ($pg in $Scope.Pages) {
        if (-not (Test-Path -LiteralPath $pg.Path)) {
            Write-Warning "Missing page file, skipped: $($pg.Path)"
            continue
        }
        $m = Read-VsdJson -Path $pg.Path
        $n = Update-PageFontSize -Manifest $m -Size $Size
        if ($n -gt 0) {
            $plan += [pscustomobject]@{ Page = $pg.Page; Path = $pg.Path; Manifest = $m; Changed = $n }
        }
    }
    $total = 0
    foreach ($p in $plan) { $total += $p.Changed }
    Write-Host "$Describe on '$($Scope.Name)': $total label(s) across $($plan.Count) file(s)."
    foreach ($p in $plan) { Write-Host "  page $($p.Page): $($p.Changed) label(s)  $(Get-RelativePath -Root $ProfilesRoot -Path $p.Path)" }
    if ($plan.Count -eq 0) { return }
    if ($DryRun) { Write-Host 'Dry run (-WhatIf): nothing written.'; return }

    Assert-VsdCraftClosed
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss-fff'
    $dest = Join-Path $BackupRoot $stamp
    Backup-VsdFiles -Paths @($plan | ForEach-Object { $_.Path }) -Root $ProfilesRoot -Destination $dest
    try {
        foreach ($p in $plan) { Save-VsdJson -Path $p.Path -Object $p.Manifest }
    }
    catch {
        Write-Warning "Write failed, restoring $($plan.Count) file(s) from $dest"
        foreach ($p in $plan) {
            $rel = Get-RelativePath -Root $ProfilesRoot -Path $p.Path
            Copy-Item -LiteralPath (Join-Path $dest $rel) -Destination $p.Path -Force
        }
        throw
    }
    Write-Host "Written. Backup: $dest"
    Write-Host "Undo with: -Mode Restore -BackupStamp $stamp -Force"
}

function Invoke-Fixed {
    param($Scope, [int]$Size)
    Invoke-SizeChange -Scope $Scope -Size ({ $Size }.GetNewClosure()) -Describe "Fixed $Size"
}

# Picks one size for every label in scope, so the whole profile matches. Uniform on purpose.
function Invoke-Fit {
    param($Scope, [int]$MinSize, [int]$MaxSize, [double]$Box)
    $titles = @(Get-VsdLabels -Pages $Scope.Pages | ForEach-Object { $_.Title })
    if ($titles.Count -eq 0) { Write-Warning 'No labels in scope.'; return }
    $fit = Get-FitSize -Titles $titles -Min $MinSize -Max $MaxSize -Box $Box
    if ($null -eq $fit) {
        Write-Warning "Some labels do not fit even at $MinSize. Using $MinSize."
        $fit = $MinSize
    }
    Invoke-SizeChange -Scope $Scope -Size ({ $fit }.GetNewClosure()) -Describe "Fit $fit (range $MinSize-$MaxSize)"
}

function Show-Labels {
    param($Scope)
    $rows = @(Get-VsdLabels -Pages $Scope.Pages | ForEach-Object {
        [pscustomobject]@{
            Profile  = $Scope.Name
            Page     = $_.Page
            Key      = $_.Key
            State    = $_.State
            FontSize = $_.FontSize
            Title    = ($_.Title -replace "`r?`n", ' / ')
        }
    })
    $rows | Format-Table -AutoSize | Out-String -Width 250 | Write-Host
}

function Invoke-Restore {
    param([string]$Stamp)
    if (-not $Stamp) {
        Write-Host 'Available backups:'
        Get-ChildItem -LiteralPath $BackupRoot -Directory -ErrorAction SilentlyContinue | Sort-Object Name | ForEach-Object { Write-Host "  $($_.Name)" }
        return
    }
    $src = Join-Path $BackupRoot $Stamp
    if (-not (Test-Path -LiteralPath $src)) { throw "No backup '$Stamp' under $BackupRoot" }
    $targets = @(Get-ChildItem -LiteralPath $src -Recurse -File | ForEach-Object {
        [pscustomobject]@{ Backup = $_.FullName; Live = (Join-Path $ProfilesRoot (Get-RelativePath -Root $src -Path $_.FullName)) }
    })
    Write-Host "Restore plan: $($targets.Count) file(s) from $Stamp"
    foreach ($t in $targets) { Write-Host "  $($t.Live)" }
    if ($DryRun) { Write-Host 'Dry run (-WhatIf): nothing restored.'; return }
    Assert-VsdCraftClosed
    foreach ($t in $targets) {
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $t.Live) | Out-Null
        Copy-Item -LiteralPath $t.Backup -Destination $t.Live -Force
    }
    Write-Host 'Restored.'
}

# ---- Image phase: bake labels into key PNGs (rendered on the PC, so the device font doesn't matter) ----

function Get-RenderFamily {
    param([string]$Preferred)
    Add-Type -AssemblyName System.Drawing
    if ([System.Drawing.FontFamily]::Families.Name -contains $Preferred) { return $Preferred }
    return 'Arial'
}

function Get-Measurer {
    if (-not $script:Measurer) {
        Add-Type -AssemblyName System.Drawing
        $bmp = New-Object System.Drawing.Bitmap 1, 1
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAlias
        $script:Measurer = $g
    }
    return $script:Measurer
}

# Largest pixel size (on a 72 px key) where every line fits inside the box. Null if none fit.
function Get-LabelPixelSize {
    param([string[]]$Lines, [string]$Family, [int]$Min, [int]$Max, [double]$Box)
    $g = Get-Measurer
    $sf = [System.Drawing.StringFormat]::new([System.Drawing.StringFormat]::GenericTypographic)
    $probe = [System.Drawing.SizeF]::new(10000, 10000)
    for ($px = $Max; $px -ge $Min; $px--) {
        $font = New-Object System.Drawing.Font($Family, [single]$px, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Pixel)
        try {
            $w = 0.0
            foreach ($ln in $Lines) {
                $sz = $g.MeasureString($ln, $font, $probe, $sf)
                if ($sz.Width -gt $w) { $w = $sz.Width }
            }
            $h = $Lines.Count * $font.GetHeight($g)
            if (($w -le $Box) -and ($h -le $Box)) { return $px }
        }
        finally { $font.Dispose() }
    }
    return $null
}

function Get-ImageFromFile {
    param([string]$Path)
    $bytes = [IO.File]::ReadAllBytes($Path)
    $ms = [IO.MemoryStream]::new($bytes)
    try { return [System.Drawing.Image]::FromStream($ms) } finally { $ms.Dispose() }
}

# Draws the base key image (if any), then the label, at 4x, and saves a 72x72 PNG.
function Save-LabelPng {
    param([string]$BasePath, [string[]]$Lines, [string]$Color, [string]$Family, [int]$Px, [string]$Align, [string]$OutPath)
    Add-Type -AssemblyName System.Drawing
    $scale = 4
    $dim = 72 * $scale
    $canvas = New-Object System.Drawing.Bitmap $dim, $dim
    $g = [System.Drawing.Graphics]::FromImage($canvas)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAlias
    $g.Clear([System.Drawing.Color]::Black)
    if ($BasePath -and (Test-Path -LiteralPath $BasePath)) {
        $src = Get-ImageFromFile -Path $BasePath
        try { $g.DrawImage($src, 0, 0, $dim, $dim) } finally { $src.Dispose() }
    }

    $font = New-Object System.Drawing.Font($Family, [single]($Px * $scale), [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Pixel)
    $brush = New-Object System.Drawing.SolidBrush ([System.Drawing.ColorTranslator]::FromHtml($Color))
    $sf = [System.Drawing.StringFormat]::new([System.Drawing.StringFormat]::GenericTypographic)
    $sf.Alignment = [System.Drawing.StringAlignment]::Center
    $sf.LineAlignment = [System.Drawing.StringAlignment]::Near
    $lineH = $font.GetHeight($g)
    $blockH = $lineH * $Lines.Count
    $y = if ($Align -eq 'bottom') { $dim - $blockH - 4 * $scale } else { ($dim - $blockH) / 2 }
    foreach ($ln in $Lines) {
        $rect = [System.Drawing.RectangleF]::new(0, $y, $dim, $lineH)
        $g.DrawString($ln, $font, $brush, $rect, $sf)
        $y += $lineH
    }

    $final = New-Object System.Drawing.Bitmap 72, 72
    $fg = [System.Drawing.Graphics]::FromImage($final)
    try {
        $fg.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $fg.DrawImage($canvas, 0, 0, 72, 72)
        $final.Save($OutPath, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    finally {
        $fg.Dispose(); $final.Dispose(); $brush.Dispose(); $font.Dispose(); $sf.Dispose(); $g.Dispose(); $canvas.Dispose()
    }
}

# One 2x preview of a page's first-state images, laid out on the key grid (6 columns x 3 rows).
function Save-ContactSheet {
    param([string]$OutPath, [hashtable]$Cells)
    Add-Type -AssemblyName System.Drawing
    $cell = 144; $gap = 8; $cols = 6; $rows = 3
    $w = $cols * $cell + ($cols + 1) * $gap
    $h = $rows * $cell + ($rows + 1) * $gap
    $sheet = New-Object System.Drawing.Bitmap $w, $h
    $g = [System.Drawing.Graphics]::FromImage($sheet)
    try {
        $g.Clear([System.Drawing.Color]::FromArgb(40, 40, 40))
        foreach ($k in $Cells.Keys) {
            $c, $r = $k.Split(',') | ForEach-Object { [int]$_ }
            $img = Get-ImageFromFile -Path $Cells[$k]
            try { $g.DrawImage($img, $gap + $c * ($cell + $gap), $gap + $r * ($cell + $gap), $cell, $cell) } finally { $img.Dispose() }
        }
        $sheet.Save($OutPath, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    finally { $g.Dispose(); $sheet.Dispose() }
}

function Get-ImageStem {
    param([string]$Key, [int]$StateIndex)
    return 'k' + $Key.Replace(',', '-') + "_s$StateIndex.png"
}

function Set-JsonProp {
    param($Obj, [string]$Name, $Value)
    if ($Obj.PSObject.Properties[$Name]) { $Obj.$Name = $Value }
    else { Add-Member -InputObject $Obj -NotePropertyName $Name -NotePropertyValue $Value }
}

function New-ImageName {
    param([string]$Dir)
    $chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789'
    do {
        $n = ((1..28 | ForEach-Object { $chars[(Get-Random -Maximum $chars.Length)] }) -join '') + '.png'
    } while (Test-Path -LiteralPath (Join-Path $Dir $n))
    return $n
}

# Renders every titled state on each page into GeneratedRoot. All labels in scope share one size (the smallest per-label fit).
function Invoke-Render {
    param($Scope, [string]$Family, [double]$Box)
    $outRoot = Join-Path $GeneratedRoot $Scope.Name
    $items = @()
    foreach ($pg in $Scope.Pages) {
        if (-not (Test-Path -LiteralPath $pg.Path)) { Write-Warning "Missing page file, skipped: $($pg.Path)"; continue }
        $m = Read-VsdJson -Path $pg.Path
        foreach ($prop in $m.Actions.PSObject.Properties) {
            $states = @($prop.Value.States)
            for ($i = 0; $i -lt $states.Count; $i++) {
                $st = $states[$i]
                if (-not ($st.Title -and ([string]$st.Title).Trim() -ne '')) { continue }
                $lines = @(([string]$st.Title) -split "`r?`n")
                $items += [pscustomobject]@{ Page = $pg; Manifest = $m; Key = $prop.Name; Index = $i; State = $st; Lines = $lines }
            }
        }
    }
    if ($items.Count -eq 0) { Write-Warning 'No labels in scope.'; return }

    $uniform = $null
    foreach ($it in $items) {
        $px = Get-LabelPixelSize -Lines $it.Lines -Family $Family -Min $Min -Max $Max -Box $Box
        if ($null -eq $px) { $px = $Min; Write-Warning "Label '$($it.Lines -join ' / ')' does not fit even at $Min px." }
        if ($null -eq $uniform -or $px -lt $uniform) { $uniform = $px }
    }
    Write-Host "Render '$($Scope.Name)': $($items.Count) label(s) at $uniform px ($Family)."

    $cellsByPage = @{}
    foreach ($it in $items) {
        $dir = Join-Path $outRoot ('p' + $it.Page.Page)
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
        $pageDir = Split-Path -Parent $it.Page.Path
        $base = if ($it.State.Image) { Join-Path (Join-Path $pageDir 'Images') ([string]$it.State.Image) } else { $null }
        $color = if ($it.State.TitleColor) { [string]$it.State.TitleColor } else { '#ffffff' }
        $out = Join-Path $dir (Get-ImageStem -Key $it.Key -StateIndex $it.Index)
        Save-LabelPng -BasePath $base -Lines $it.Lines -Color $color -Family $Family -Px $uniform -Align ([string]$it.State.TitleAlignment) -OutPath $out
        if ($it.Index -eq 0) {
            if (-not $cellsByPage.ContainsKey($it.Page.Page)) { $cellsByPage[$it.Page.Page] = @{} }
            $cellsByPage[$it.Page.Page][$it.Key] = $out
        }
    }
    foreach ($pno in $cellsByPage.Keys) {
        $dir = Join-Path $outRoot ('p' + $pno)
        Save-ContactSheet -OutPath (Join-Path $dir 'sheet.png') -Cells $cellsByPage[$pno]
        Write-Host "  page ${pno}: $(Join-Path $dir 'sheet.png')"
    }
}

# Copies the rendered PNGs into each page's Images folder, points the states at them, and hides the device title.
# Keeps Title and the old images so a restore from backup fully reverts.
function Invoke-Apply {
    param($Scope)
    $plan = @()
    foreach ($pg in $Scope.Pages) {
        if (-not (Test-Path -LiteralPath $pg.Path)) { Write-Warning "Missing page file, skipped: $($pg.Path)"; continue }
        $m = Read-VsdJson -Path $pg.Path
        $renderDir = Join-Path (Join-Path $GeneratedRoot $Scope.Name) ('p' + $pg.Page)
        $imgDir = Join-Path (Split-Path -Parent $pg.Path) 'Images'
        $jobs = @()
        foreach ($prop in $m.Actions.PSObject.Properties) {
            $states = @($prop.Value.States)
            for ($i = 0; $i -lt $states.Count; $i++) {
                $st = $states[$i]
                if (-not ($st.Title -and ([string]$st.Title).Trim() -ne '')) { continue }
                $src = Join-Path $renderDir (Get-ImageStem -Key $prop.Name -StateIndex $i)
                if (-not (Test-Path -LiteralPath $src)) { throw "No render for page $($pg.Page), key $($prop.Name), state $i. Run -Mode Render first." }
                $jobs += [pscustomobject]@{ State = $st; Src = $src; ImgDir = $imgDir }
            }
        }
        if ($jobs.Count -gt 0) { $plan += [pscustomobject]@{ Page = $pg.Page; Path = $pg.Path; Manifest = $m; Jobs = $jobs } }
    }
    $total = 0
    foreach ($p in $plan) { $total += $p.Jobs.Count }
    Write-Host "Apply '$($Scope.Name)': $total label image(s) across $($plan.Count) file(s)."
    if ($plan.Count -eq 0) { return }
    if ($DryRun) { Write-Host 'Dry run (-WhatIf): nothing written.'; return }

    Assert-VsdCraftClosed
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss-fff'
    $dest = Join-Path $BackupRoot $stamp
    Backup-VsdFiles -Paths @($plan | ForEach-Object { $_.Path }) -Root $ProfilesRoot -Destination $dest
    try {
        foreach ($p in $plan) {
            foreach ($j in $p.Jobs) {
                $name = New-ImageName -Dir $j.ImgDir
                New-Item -ItemType Directory -Force -Path $j.ImgDir | Out-Null
                Copy-Item -LiteralPath $j.Src -Destination (Join-Path $j.ImgDir $name)
                Set-JsonProp -Obj $j.State -Name 'Image' -Value $name
                Set-JsonProp -Obj $j.State -Name 'ShowTitle' -Value $false
            }
            Save-VsdJson -Path $p.Path -Object $p.Manifest
        }
    }
    catch {
        Write-Warning "Apply failed, restoring manifests from $dest"
        foreach ($p in $plan) {
            Copy-Item -LiteralPath (Join-Path $dest (Get-RelativePath -Root $ProfilesRoot -Path $p.Path)) -Destination $p.Path -Force
        }
        throw
    }
    Write-Host "Applied. Backup: $dest"
    Write-Host "Undo with: -Mode Restore -BackupStamp $stamp -Force"
}

if ($MyInvocation.InvocationName -ne '.') {
    if (-not $BackupRoot) { $BackupRoot = Join-Path (Split-Path -Parent $PSScriptRoot) 'Backups\Auto' }
    if (-not $GeneratedRoot) { $GeneratedRoot = Join-Path (Split-Path -Parent $PSScriptRoot) 'Generated' }
    # Images are drawn on a 72 px key, so the default cap is higher than the text-size default.
    if ($Mode -eq 'Render' -and -not $PSBoundParameters.ContainsKey('Max')) { $Max = 24 }
    $DryRun = [bool]$WhatIfPreference
    $box = 72 - 2 * $Margin

    switch ($Mode) {
        'List' {
            if ($ProfileName) {
                Show-Labels -Scope (Resolve-Scope -Name $ProfileName -PageNo $Page)
            }
            else {
                foreach ($r in @(Get-VsdRoots -Root $ProfilesRoot)) {
                    try { Show-Labels -Scope (Resolve-Scope -Name $r.Name -PageNo $Page) }
                    catch { Write-Warning "$($r.Name): $($_.Exception.Message)" }
                }
            }
        }
        'Fixed' {
            if (-not $ProfileName -or -not $Size) { throw 'Fixed mode needs -ProfileName and -Size.' }
            if ($Size -lt 4 -or $Size -gt 72) { throw 'Size must be between 4 and 72.' }
            Invoke-Fixed -Scope (Resolve-Scope -Name $ProfileName -PageNo $Page) -Size $Size
        }
        'Fit' {
            if (-not $ProfileName) { throw 'Fit mode needs -ProfileName.' }
            Invoke-Fit -Scope (Resolve-Scope -Name $ProfileName -PageNo $Page) -MinSize $Min -MaxSize $Max -Box $box
        }
        'Preset' {
            $presets = Read-VsdJson -Path (Join-Path $PSScriptRoot 'presets.json')
            if (-not $PresetName) {
                Write-Host 'Presets:'
                $presets.PSObject.Properties | ForEach-Object { Write-Host "  $($_.Name)" }
                break
            }
            $entry = $presets.PSObject.Properties[$PresetName]
            if (-not $entry) { throw "Unknown preset '$PresetName'." }
            if (-not $ProfileName) { throw 'Preset mode needs -ProfileName.' }
            $cfg = $entry.Value
            $scope = Resolve-Scope -Name $ProfileName -PageNo $Page
            switch ($cfg.Mode) {
                'Fixed' { Invoke-Fixed -Scope $scope -Size ([int]$cfg.Size) }
                'Fit'   { Invoke-Fit -Scope $scope -MinSize ([int]$cfg.MinSize) -MaxSize ([int]$cfg.MaxSize) -Box $box }
                default { throw "Preset '$PresetName' has unknown mode '$($cfg.Mode)'." }
            }
        }
        'Restore' {
            Invoke-Restore -Stamp $BackupStamp
        }
        'Render' {
            $scope = Resolve-Scope -Name $ProfileName -PageNo $Page
            Invoke-Render -Scope $scope -Family (Get-RenderFamily -Preferred $RenderFont) -Box $box
        }
        'Apply' {
            Invoke-Apply -Scope (Resolve-Scope -Name $ProfileName -PageNo $Page)
        }
    }
}
