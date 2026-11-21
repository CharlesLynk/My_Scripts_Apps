# Single GUI launcher for the AutoCAD batch drawing tools.
# Generate and block-attribute update modes run Generate-Drawings.ps1 in a new console window.
# The Plant 3D mode validates its CSV here; the update itself runs inside Plant 3D.

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$generatorPath = Join-Path $scriptDir 'Generate-Drawings.ps1'
$defaultAccore = 'C:\Program Files\Autodesk\AutoCAD 2026\accoreconsole.exe'

$modeGenerate = 'Generate new drawings from CSV'
$modeAttributes = 'Update block attributes in existing drawings'
$modePlant = 'Validate CSV for Plant 3D property update'

$form = New-Object System.Windows.Forms.Form
$form.Text = 'AutoCAD Batch Tools'
$form.Size = New-Object System.Drawing.Size(660, 560)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false

function Add-Label($text, $y) {
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $text
    $label.Location = New-Object System.Drawing.Point(15, ($y + 3))
    $label.Size = New-Object System.Drawing.Size(170, 20)
    $form.Controls.Add($label)
    return $label
}

function Add-PathRow($text, $y, $isFolder) {
    $label = Add-Label $text $y
    $box = New-Object System.Windows.Forms.TextBox
    $box.Location = New-Object System.Drawing.Point(190, $y)
    $box.Size = New-Object System.Drawing.Size(360, 22)
    $form.Controls.Add($box)
    $button = New-Object System.Windows.Forms.Button
    $button.Text = 'Browse...'
    $button.Location = New-Object System.Drawing.Point(558, ($y - 1))
    $button.Size = New-Object System.Drawing.Size(80, 24)
    $button.Tag = [pscustomobject]@{ Box = $box; IsFolder = $isFolder }
    $button.Add_Click({
        $info = $this.Tag
        if ($info.IsFolder) {
            $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
            if ($dialog.ShowDialog() -eq 'OK') { $info.Box.Text = $dialog.SelectedPath }
        } else {
            $dialog = New-Object System.Windows.Forms.OpenFileDialog
            $dialog.Filter = 'All supported|*.csv;*.dwt;*.dwg;*.exe|All files|*.*'
            if ($dialog.ShowDialog() -eq 'OK') { $info.Box.Text = $dialog.FileName }
        }
    })
    $form.Controls.Add($button)
    return [pscustomobject]@{ Label = $label; Box = $box; Button = $button }
}

$functionLabel = Add-Label 'Function:' 15
$modeBox = New-Object System.Windows.Forms.ComboBox
$modeBox.DropDownStyle = 'DropDownList'
$modeBox.Location = New-Object System.Drawing.Point(190, 12)
$modeBox.Size = New-Object System.Drawing.Size(448, 22)
$modeBox.Items.AddRange(@($modeGenerate, $modeAttributes, $modePlant))
$modeBox.SelectedIndex = 0
$form.Controls.Add($modeBox)

$accoreRow = Add-PathRow 'accoreconsole.exe:' 50 $false
$accoreRow.Box.Text = $defaultAccore
$csvRow = Add-PathRow 'CSV file:' 90 $false
$folderRow = Add-PathRow 'Output folder:' 130 $true
$templateRow = Add-PathRow 'Template DWT:' 170 $false
$borderRow = Add-PathRow 'Border XREF DWG (optional):' 210 $false

$maxLabel = Add-Label 'Max drawings (0 = all):' 250
$maxBox = New-Object System.Windows.Forms.NumericUpDown
$maxBox.Location = New-Object System.Drawing.Point(190, 248)
$maxBox.Size = New-Object System.Drawing.Size(80, 22)
$maxBox.Minimum = 0
$maxBox.Maximum = 100000
$form.Controls.Add($maxBox)

$dryRunBox = New-Object System.Windows.Forms.CheckBox
$dryRunBox.Text = 'Preview only (-WhatIf)'
$dryRunBox.Location = New-Object System.Drawing.Point(190, 285)
$dryRunBox.Size = New-Object System.Drawing.Size(200, 22)
$form.Controls.Add($dryRunBox)

$keepLogsBox = New-Object System.Windows.Forms.CheckBox
$keepLogsBox.Text = 'Keep logs for successful runs (-KeepLogs)'
$keepLogsBox.Location = New-Object System.Drawing.Point(190, 310)
$keepLogsBox.Size = New-Object System.Drawing.Size(320, 22)
$form.Controls.Add($keepLogsBox)

$runButton = New-Object System.Windows.Forms.Button
$runButton.Text = 'Run'
$runButton.Location = New-Object System.Drawing.Point(190, 345)
$runButton.Size = New-Object System.Drawing.Size(120, 30)
$form.Controls.Add($runButton)

$closeButton = New-Object System.Windows.Forms.Button
$closeButton.Text = 'Close'
$closeButton.Location = New-Object System.Drawing.Point(318, 345)
$closeButton.Size = New-Object System.Drawing.Size(120, 30)
$closeButton.Add_Click({ $form.Close() })
$form.Controls.Add($closeButton)

$outputBox = New-Object System.Windows.Forms.TextBox
$outputBox.Multiline = $true
$outputBox.ReadOnly = $true
$outputBox.ScrollBars = 'Vertical'
$outputBox.Location = New-Object System.Drawing.Point(15, 390)
$outputBox.Size = New-Object System.Drawing.Size(623, 120)
$form.Controls.Add($outputBox)

function Write-Output-Line($text) {
    $outputBox.AppendText($text + [Environment]::NewLine)
}

function Set-ModeLayout {
    $mode = $modeBox.SelectedItem
    $isPlant = ($mode -eq $modePlant)
    $isGenerate = ($mode -eq $modeGenerate)

    $accoreRow.Label.Enabled = -not $isPlant; $accoreRow.Box.Enabled = -not $isPlant; $accoreRow.Button.Enabled = -not $isPlant
    $templateRow.Label.Enabled = $isGenerate; $templateRow.Box.Enabled = $isGenerate; $templateRow.Button.Enabled = $isGenerate
    $borderRow.Label.Enabled = $isGenerate; $borderRow.Box.Enabled = $isGenerate; $borderRow.Button.Enabled = $isGenerate
    $folderRow.Box.Enabled = -not $isPlant; $folderRow.Button.Enabled = -not $isPlant
    $folderRow.Label.Text = $(if ($isGenerate) { 'Output folder:' } else { 'Drawing root folder:' })
    $folderRow.Label.Enabled = -not $isPlant
    $maxBox.Enabled = -not $isPlant
    $dryRunBox.Enabled = -not $isPlant
    $keepLogsBox.Enabled = -not $isPlant
    $runButton.Text = $(if ($isPlant) { 'Validate CSV' } else { 'Run' })
}

function Test-PlantCsv($path) {
    if (-not (Test-Path -LiteralPath $path)) { throw "CSV not found: $path" }
    $rows = @(Import-Csv -LiteralPath $path)
    if ($rows.Count -eq 0) { throw 'CSV has a header but no data rows.' }
    $columns = @($rows[0].PSObject.Properties.Name)
    if ($columns -notcontains 'DrawingRelativePath') {
        throw "CSV is missing the required DrawingRelativePath column. Found: $($columns -join ', ')"
    }
    $blank = 0
    foreach ($row in $rows) {
        if ([string]::IsNullOrWhiteSpace($row.DrawingRelativePath)) { $blank++ }
    }
    if ($blank -gt 0) { throw "$blank row(s) have an empty DrawingRelativePath." }
    $propertyColumns = @($columns | Where-Object { $_ -ne 'DrawingRelativePath' })
    return [pscustomobject]@{
        Rows = $rows.Count
        Properties = $propertyColumns
    }
}

function Invoke-PlantValidation {
    $csv = $csvRow.Box.Text.Trim('"').Trim()
    try {
        $result = Test-PlantCsv $csv
    } catch {
        Write-Output-Line "Validation failed: $($_.Exception.Message)"
        return
    }
    Write-Output-Line "CSV is valid: $($result.Rows) drawing row(s)."
    Write-Output-Line "Properties: $($result.Properties -join ', ')"
    Write-Output-Line ''
    Write-Output-Line 'Next step (inside Plant 3D, with the project open):'
    Write-Output-Line '  1. Type P3DUPDATEDRAWINGPROPERTIES at the command line.'
    Write-Output-Line "  2. When prompted for the CSV, enter: $csv"
    Write-Output-Line '  3. Review the preview and type UPDATE to apply.'
}

function Build-GeneratorArguments {
    $mode = $modeBox.SelectedItem
    $items = New-Object System.Collections.Generic.List[string]
    $items.Add('-NoProfile'); $items.Add('-ExecutionPolicy'); $items.Add('Bypass')
    $items.Add('-NoExit')
    $items.Add('-File'); $items.Add("`"$generatorPath`"")
    $items.Add('-AccoreConsolePath'); $items.Add("`"$($accoreRow.Box.Text.Trim())`"")
    $items.Add('-CsvPath'); $items.Add("`"$($csvRow.Box.Text.Trim())`"")
    if ($mode -eq $modeGenerate) {
        $items.Add('-OutputDirectory'); $items.Add("`"$($folderRow.Box.Text.Trim())`"")
        $items.Add('-TemplatePath'); $items.Add("`"$($templateRow.Box.Text.Trim())`"")
        $border = $borderRow.Box.Text.Trim()
        if ($border) { $items.Add('-BorderXrefPath'); $items.Add("`"$border`"") }
    } else {
        $items.Add('-UpdateExisting')
        $items.Add('-DrawingRoot'); $items.Add("`"$($folderRow.Box.Text.Trim())`"")
    }
    if ($maxBox.Value -gt 0) { $items.Add('-MaxDrawings'); $items.Add([string]$maxBox.Value) }
    if ($keepLogsBox.Checked) { $items.Add('-KeepLogs') }
    if ($dryRunBox.Checked) { $items.Add('-WhatIf') }
    return $items
}

function Test-GeneratorInputs {
    $mode = $modeBox.SelectedItem
    $checks = @(
        @{ Name = 'accoreconsole.exe'; Path = $accoreRow.Box.Text.Trim(); File = $true },
        @{ Name = 'CSV file'; Path = $csvRow.Box.Text.Trim(); File = $true },
        @{ Name = 'Generate-Drawings.ps1'; Path = $generatorPath; File = $true }
    )
    if ($mode -eq $modeGenerate) {
        $checks += @{ Name = 'Template DWT'; Path = $templateRow.Box.Text.Trim(); File = $true }
        $checks += @{ Name = 'Output folder'; Path = $folderRow.Box.Text.Trim(); File = $false }
    } else {
        $checks += @{ Name = 'Drawing root folder'; Path = $folderRow.Box.Text.Trim(); File = $false }
    }
    foreach ($check in $checks) {
        if ([string]::IsNullOrWhiteSpace($check.Path)) {
            return "Missing: $($check.Name)."
        }
        $exists = if ($check.File) { Test-Path -LiteralPath $check.Path -PathType Leaf } else { Test-Path -LiteralPath $check.Path -PathType Container }
        if (-not $exists -and $check.Name -ne 'Output folder') {
            return "Not found: $($check.Name) at $($check.Path)"
        }
    }
    return $null
}

function Invoke-GeneratorRun {
    $problem = Test-GeneratorInputs
    if ($problem) {
        Write-Output-Line $problem
        return
    }
    $arguments = Build-GeneratorArguments
    Write-Output-Line 'Starting in a new PowerShell window. Watch that window for progress.'
    Write-Output-Line ("powershell.exe " + ($arguments -join ' '))
    Start-Process -FilePath 'powershell.exe' -ArgumentList ($arguments -join ' ')
}

$modeBox.Add_SelectedIndexChanged({ Set-ModeLayout })

$runButton.Add_Click({
    $outputBox.Clear()
    if ($modeBox.SelectedItem -eq $modePlant) {
        Invoke-PlantValidation
    } else {
        Invoke-GeneratorRun
    }
})

Set-ModeLayout
[void]$form.ShowDialog()
