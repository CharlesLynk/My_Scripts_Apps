@echo off
setlocal
REM ============================================================
REM Overhauled_Publish.cmd — Phase 1 (scan + selection)
REM Goal:
REM   - Scan DWGs for named page setups (name + device + media + CTB)
REM   - Pick most popular by name; tie -> numbered prompt; cancel aborts
REM   - If zero setups or Client Standard: prompt CTB + page size (ANSI D/ARCH D/custom)
REM   - Prepare data for publish (Phase 2 will add DSD build + publish)
REM Notes:
REM   - Uses ACCoreConsole with an embedded LISP to scan via ObjectDBX
REM   - Non-destructive path preferred (Page Setup Override in DSD)
REM   - Config via Publish-With-CTB.config.json (optional) next to this file
REM   - Logging: PUBLISH_LOG=debug|info|quiet (dev flag consolidation)
REM ============================================================

REM --- Dev-friendly defaults ---
if "%PUBLISH_LOG%"=="" set "PUBLISH_LOG=debug"
if "%PUBLISH_DEBUG%"=="" set "PUBLISH_DEBUG=1"

set "PUBLISH_LOG=debug"
set "PUBLISH_DEBUG=1"

set "__PUBLISH_CMD_DIR=%~dp0"
set "__PUBLISH_CMD_PATH=%~f0"
set "PS_PAYLOAD=%TEMP%\Overhauled_Publish_inline.ps1"

REM Extract PowerShell payload after :__PS_PAYLOAD__ marker
for /f "tokens=1* delims=:" %%A in ('findstr /n /b ":__PS_PAYLOAD__" "%__PUBLISH_CMD_PATH%"') do set "LINE=%%A"
set /a LINE+=1
more +%LINE% "%__PUBLISH_CMD_PATH%" > "%PS_PAYLOAD%"

powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%PS_PAYLOAD%"
set "ERR=%ERRORLEVEL%"
del "%PS_PAYLOAD%" >nul 2>nul
echo [Overhauled_Publish] Exit code: %ERR%
if "%PUBLISH_DEBUG%"=="1" (
	echo.
	echo Press any key to close...
	pause >nul
)
endlocal
exit /b %ERR%

:__PS_PAYLOAD__
# ====================== PowerShell payload (Phase 1) ======================
param()
$ErrorActionPreference = 'Stop'

function Write-Log {
	param(
		[string]$Message,
		[ValidateSet('debug','info','warn','error')]
		[string]$Level = 'info'
	)
    $cur = if ($env:PUBLISH_LOG) { $env:PUBLISH_LOG } else { 'debug' }
    $cur = $cur.ToLowerInvariant()
	$order = @{ debug = 1; info = 2; warn = 3; error = 4 }
	if ($order[$Level] -ge $order[$cur]) {
		$prefix = switch($Level){ 'debug' {'[DBG]'} 'info' {'[INF]'} 'warn' {'[WRN]'} 'error' {'[ERR]'} }
		Write-Host "$prefix $Message"
	}
}

# Determine debug level once for conditional tracing
$IsDebug = $false
try {
	if ($env:PUBLISH_LOG) {
		$IsDebug = ($env:PUBLISH_LOG.ToLowerInvariant() -eq 'debug')
	} else { $IsDebug = $true }
} catch { $IsDebug = $true }

# Discover caller location
$ScriptHome = $env:__PUBLISH_CMD_DIR
if (-not $ScriptHome) { $ScriptHome = Split-Path -Parent $PSCommandPath }

# Config (optional JSON next to this .cmd)
$ConfigPath = Join-Path $ScriptHome 'Publish-With-CTB.config.json'
$Cfg = $null
if (Test-Path -LiteralPath $ConfigPath) {
	try { $Cfg = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json } catch { $Cfg = $null }
}

# Defaults
$AccorePath = $null
$Device     = 'MasterPwPdf.pc3'
$TiePref    = 'ANSI'
$OutPdfBase = [Environment]::ExpandEnvironmentVariables("$env:USERPROFILE\Documents\AutoCAD Plots\All_Layouts.pdf")
$FavDirs    = @('C:\Users\Public\Documents', 'C:\Users\%USERNAME%\Documents')
$ExtraDirs  = @()
$DefaultTemplate = $null
if ($Cfg) {
	$AccorePath = if ($Cfg.AccorePath) { "$($Cfg.AccorePath)" } else { $null }
	if ($Cfg.FavDirs) { $FavDirs = @($Cfg.FavDirs | ForEach-Object { "$_" }) }
	if ($Cfg.ExtraDirs) { $ExtraDirs = @($Cfg.ExtraDirs | ForEach-Object { "$_" }) }
	if ($Cfg.DefaultOutPdf) { $OutPdfBase = "$($Cfg.DefaultOutPdf)" }
    if ($Cfg.DefaultTemplate) { $DefaultTemplate = "$(($Cfg.DefaultTemplate))" }
}

function Resolve-AccoreConsole {
	param([string]$Hint)
	if ($Hint -and (Test-Path -LiteralPath $Hint)) { return (Resolve-Path -LiteralPath $Hint).Path }
	$candidates = @(
		'C:\Program Files\Autodesk\AutoCAD 2025\accoreconsole.exe',
		'C:\Program Files\Autodesk\AutoCAD 2024\accoreconsole.exe',
		'C:\Program Files\Autodesk\AutoCAD 2023\accoreconsole.exe'
	) + (Get-ChildItem 'C:\Program Files\Autodesk' -Recurse -ErrorAction SilentlyContinue -Filter accoreconsole.exe | ForEach-Object { $_.FullName })
	$hit = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
	return $hit
}

$AccorePath = Resolve-AccoreConsole -Hint $AccorePath
if (-not $AccorePath) { Write-Host "ACCoreConsole not found. Please set AccorePath in config." -ForegroundColor Red; exit 1 }
Write-Log "ACCoreConsole = $AccorePath" debug

# Source folder picker (console prompt, with optional browse)
Add-Type -AssemblyName System.Windows.Forms | Out-Null
function Pick-Folder {
	param([string]$Title = 'Select Source Folder')
	$dlg = New-Object System.Windows.Forms.FolderBrowserDialog
	$dlg.Description = $Title
	$dlg.ShowNewFolderButton = $false
	$r = $dlg.ShowDialog()
	if ($r -eq [System.Windows.Forms.DialogResult]::OK) { return $dlg.SelectedPath } else { return $null }
}

Write-Host ""
$SourceFolder = Read-Host "Paste DWG source folder path (or press Enter to browse)"
if (-not $SourceFolder) { $SourceFolder = Pick-Folder -Title 'Select DWG source folder' }
if (-not $SourceFolder -or -not (Test-Path -LiteralPath $SourceFolder)) { Write-Host "Invalid source folder." -ForegroundColor Red; exit 1 }
$SourceFolder = (Resolve-Path -LiteralPath $SourceFolder).Path
Write-Log "Source = $SourceFolder" info

# Standards Mode
Write-Host ""
Write-Host "Standards Mode:" -ForegroundColor Cyan
Write-Host "  1) Auto (use most popular page setup found in drawings)"
Write-Host "  2) Client Standard (choose CTB + page size)"
function Read-Mode12 {
	while ($true) {
		$v = Read-Host "Select 1 or 2"
		if ($v -in @('1','2')) { return [int]$v }
	}
}
$modeKey = Read-Mode12
$UseAuto = ($modeKey -eq 1)

# Prepare temp paths
$TempRoot = Join-Path $env:TEMP ('Overhauled_Publish_' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
New-Item -ItemType Directory -Force -Path $TempRoot | Out-Null
$ListPath = Join-Path $TempRoot 'dwgs.txt'
$CsvPath  = Join-Path $TempRoot 'page_setups.csv'
$LspPath  = Join-Path $TempRoot 'ScanPageSetups.lsp'
$ScrPath  = Join-Path $TempRoot 'scan.scr'
$Workers  = if ($Cfg -and $Cfg.Workers) { try { [int]$Cfg.Workers } catch { 4 } } else { [Math]::Min([Environment]::ProcessorCount, 6) }

# Enumerate DWGs (non-recursive) and write list
Get-ChildItem -LiteralPath $SourceFolder -File -Filter *.dwg | ForEach-Object { $_.FullName } | Set-Content -Encoding ASCII $ListPath
if (-not (Get-Item $ListPath).Length) { Write-Host "No DWG files found in folder." -ForegroundColor Yellow; exit 2 }

# COM-free: skip DBX scan of page setups; Auto mode uses existing saved layout
# page setups; Client Standard will synthesize and apply an override.
$selectionMode = if ($UseAuto) { 'auto' } else { 'client' }

# If client mode or zero setups: prompt CTB + page size
$ClientCtb = $null
$ClientMediaCanonical = $null
function Get-RoamPlotPaths {
	# Best effort; will be refined in Phase 2 if needed
	$paths = @()
	$appdata = $env:APPDATA
	if ($appdata) {
		$paths += Get-ChildItem (Join-Path $appdata 'Autodesk') -Recurse -Directory -ErrorAction SilentlyContinue |
			Where-Object { $_.Name -match 'Plot Styles|Plotters' } | Select-Object -ExpandProperty FullName
	}
	return $paths
}

function Pick-Ctb {
	param([string[]]$Roots)
	$roots = @($Roots | Where-Object { $_ -and (Test-Path $_) })
	$ctbs = @()
	foreach ($r in $roots) { $ctbs += Get-ChildItem -Path $r -Filter *.ctb -Recurse -ErrorAction SilentlyContinue }
	$ctbs = $ctbs | Sort-Object FullName -Unique
	$options = @()
	for ($i=0; $i -lt $ctbs.Count; $i++) { $options += "{0}) {1}" -f ($i+1), $ctbs[$i].FullName }
	$browseIdx = $options.Count + 1
	$options += "{0}) Browse..." -f $browseIdx
	Write-Host ""
	Write-Host "Select CTB:" -ForegroundColor Cyan
	$options | ForEach-Object { Write-Host "  $_" }
	$sel = Read-Host "Enter number"
	if ($sel -notmatch '^[0-9]+$') { return $null }
	$n = [int]$sel
	if ($n -eq $browseIdx) {
		$ofd = New-Object System.Windows.Forms.OpenFileDialog
		$ofd.Filter = 'Plot styles (*.ctb)|*.ctb|All files (*.*)|*.*'
		$ofd.Multiselect = $false
		if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $ofd.FileName } else { return $null }
	}
	if ($n -lt 1 -or $n -gt $ctbs.Count) { return $null }
	return $ctbs[$n-1].FullName
}

function Pick-PageSizeCanonical {
	# Return canonical media name string
	Write-Host ""
	Write-Host "Select Page Size:" -ForegroundColor Cyan
	Write-Host "  1) ANSI D (34 x 22)"
	Write-Host "  2) ARCH D (36 x 24)"
	Write-Host "  3) Custom (type canonical or WxH inches)"
	$sel = Read-Host "Enter 1/2/3"
	switch ($sel) {
		'1' { return 'ANSI_full_bleed_D_(34.00_x_22.00_Inches)' }
		'2' { return 'ARCH_full_bleed_D_(36.00_x_24.00_Inches)' }
		'3' {
			$txt = Read-Host 'Enter canonical name, or WxH (inches), e.g., 34x22'
			if ($txt -match '^(\d+(?:\.\d+)?)x(\d+(?:\.\d+)?)$') {
				$w = [double]$Matches[1]; $h = [double]$Matches[2]
				return ('UserDefined_{0:F2}_x_{1:F2}_Inches' -f $w,$h)
			} else {
				return $txt
			}
		}
		default { return $null }
	}
}

if ($selectionMode -eq 'client') {
	$roots = @($FavDirs + $ExtraDirs + (Get-RoamPlotPaths) + $SourceFolder)
	$ClientCtb = Pick-Ctb -Roots $roots
	if (-not $ClientCtb) { Write-Host "CTB selection canceled." -ForegroundColor Red; exit 6 }
	$ClientMediaCanonical = Pick-PageSizeCanonical
	if (-not $ClientMediaCanonical) { Write-Host "Page size selection canceled." -ForegroundColor Red; exit 6 }
}

# Decide override source and data to carry forward to Phase 2
$Plan = [ordered]@{}
if ($selectionMode -eq 'client') {
	$Plan.Mode            = 'Client'
	$Plan.Device          = $Device
	$Plan.CTB             = $ClientCtb
	$Plan.CanonicalMedia  = $ClientMediaCanonical
	$Plan.OverrideName    = 'AutoPublish_PS'  # fixed new setup name
	$Plan.OverrideSource  = $null            # to be created in Phase 2
}
else {
	$Plan.Mode            = 'Auto'
	$Plan.SelectedName    = $null
	$Plan.OverrideSource  = $null
	$Plan.OverrideName    = $null
	$Plan.Device          = $Device
	$Plan.CanonicalMedia  = $null
	$Plan.CTB             = $null
}

# Output PDF path with collision-avoidance (-001)
function Get-UniqueOutPath {
	param([string]$BasePath)
	$dir = Split-Path -Parent $BasePath
	$name = Split-Path -Leaf $BasePath
	$base = [System.IO.Path]::GetFileNameWithoutExtension($name)
	$ext = [System.IO.Path]::GetExtension($name)
	$p = Join-Path $dir $name
	if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
	if (-not (Test-Path -LiteralPath $p)) { return $p }
	for ($i=1; $i -lt 1000; $i++) {
		$p2 = Join-Path $dir ("{0}-{1:D3}{2}" -f $base,$i,$ext)
		if (-not (Test-Path -LiteralPath $p2)) { return $p2 }
	}
	return $p
}

$OutPdf = Get-UniqueOutPath -BasePath $OutPdfBase

Write-Host ""
Write-Host "Selection Summary" -ForegroundColor Green
if ($Plan.Mode -eq 'Client') {
	Write-Host "  Mode   : Client Standard"
	Write-Host "  Device : $($Plan.Device)"
	Write-Host "  CTB    : $($Plan.CTB)"
	Write-Host "  Media  : $($Plan.CanonicalMedia)"
	Write-Host "  PSName : $($Plan.OverrideName)"
}
else {
	Write-Host "  Mode   : Auto (use existing page setups)"
	Write-Host "  Device : $($Plan.Device)"
}
Write-Host "  Output : $OutPdf"

# ====================== Phase 2: Build DSD + Publish ======================

# Resolve PC3/CTB directories so CoreConsole can find resources
function Get-SearchRoots {
	param([string]$SourceFolder)
	$roots = @()
	if ($env:APPDATA)   { $roots += (Join-Path $env:APPDATA   'Autodesk') }
	if ($env:PROGRAMDATA) { $roots += (Join-Path $env:PROGRAMDATA 'Autodesk') }
	$roots += $FavDirs
	$roots += $ExtraDirs
	if ($SourceFolder) { $roots += $SourceFolder }
	$roots | Where-Object { Test-Path $_ } | Select-Object -Unique
}

function Resolve-FileInRoots {
	param([string]$FileName, [string[]]$Roots, [string]$Filter)
	$hits = @()
	foreach ($r in $Roots) {
		try {
			$hits += Get-ChildItem -Path $r -Recurse -File -Filter $Filter -ErrorAction SilentlyContinue |
							 Where-Object { $_.Name -ieq $FileName }
		} catch {}
	}
	$hits | Sort-Object LastWriteTime -Descending | Select-Object -First 1
}

$roots = Get-SearchRoots -SourceFolder $SourceFolder

# PC3
$pc3FileName = $Plan.Device
$pc3Hit = $null
if ($pc3FileName -and ($pc3FileName -match '\\.pc3$')) {
	$pc3Hit = Resolve-FileInRoots -FileName $pc3FileName -Roots $roots -Filter *.pc3
}
if ($pc3Hit) {
	$env:PUBLISH_PC3_DIR = $pc3Hit.DirectoryName
	Write-Log ("PC3 dir = {0}" -f $env:PUBLISH_PC3_DIR) debug
	Write-Log ("PC3 resolved path = {0}" -f $pc3Hit.FullName) debug
}

# CTB
$ctbPath = $Plan.CTB
if ($ctbPath) {
	if (Test-Path -LiteralPath $ctbPath) {
		$env:PUBLISH_CTB_DIR = (Split-Path -Parent $ctbPath)
		Write-Log ("CTB provided path = {0}" -f (Resolve-Path -LiteralPath $ctbPath).Path) debug
	} else {
		$ctbName = [System.IO.Path]::GetFileName($ctbPath)
		$ctbHit = Resolve-FileInRoots -FileName $ctbName -Roots $roots -Filter *.ctb
		if ($ctbHit) {
			$env:PUBLISH_CTB_DIR = $ctbHit.DirectoryName
			Write-Log ("CTB resolved path = {0}" -f $ctbHit.FullName) debug
		}
	}
}
if (-not $env:PUBLISH_CTB_DIR) { $env:PUBLISH_CTB_DIR = $SourceFolder }
Write-Log ("CTB dir = {0}" -f $env:PUBLISH_CTB_DIR) debug

# Helpers: find stock template and create seed DWG
function Get-StockTemplate {
	$cands = @()
	if ($env:APPDATA) {
		$cands += Get-ChildItem (Join-Path $env:APPDATA 'Autodesk') -Recurse -ErrorAction SilentlyContinue -Include acadiso.dwt, acad.dwt | Select-Object -ExpandProperty FullName
	}
	if ($env:PROGRAMDATA) {
		$cands += Get-ChildItem (Join-Path $env:PROGRAMDATA 'Autodesk') -Recurse -ErrorAction SilentlyContinue -Include acadiso.dwt, acad.dwt | Select-Object -ExpandProperty FullName
	}
	$cands | Select-Object -Unique | Select-Object -First 1
}

# If Client mode, use Modifiable_Template.dwg and modify an existing page setup
$TemplateWorkDwg = $null
$ChosenPsNameOut = Join-Path $TempRoot 'chosen_ps.txt'
if ($Plan.Mode -eq 'Client') {
	$TemplateSrc = Join-Path $ScriptHome 'Modifiable_Template.dwg'
	if (-not (Test-Path -LiteralPath $TemplateSrc)) {
		Write-Host "Required template 'Modifiable_Template.dwg' was not found next to the script: $ScriptHome" -ForegroundColor Red
		Write-Host "Please place Modifiable_Template.dwg alongside this .cmd and try again." -ForegroundColor Yellow
		exit 7
	}
	# Copy to temp working DWG (we will modify this one)
	$TemplateWorkDwg = Join-Path $TempRoot 'Modifiable_Template_work.dwg'
	Copy-Item -LiteralPath $TemplateSrc -Destination $TemplateWorkDwg -Force
	if (-not (Test-Path -LiteralPath $TemplateWorkDwg)) { Write-Host "Failed to prepare working template DWG." -ForegroundColor Red; exit 8 }

	# Also copy to discovered AutoCAD Template directories for future use
	function Get-AcadTemplateDirs {
		$dirs = @()
		if ($env:APPDATA) {
			$dirs += Get-ChildItem (Join-Path $env:APPDATA 'Autodesk') -Recurse -Directory -ErrorAction SilentlyContinue |
				Where-Object { $_.FullName -match 'Template(s)?' -and $_.FullName -match 'AutoCAD' } |
				Select-Object -ExpandProperty FullName
		}
		if ($env:PROGRAMDATA) {
			$dirs += Get-ChildItem (Join-Path $env:PROGRAMDATA 'Autodesk') -Recurse -Directory -ErrorAction SilentlyContinue |
				Where-Object { $_.FullName -match 'Template(s)?' -and $_.FullName -match 'AutoCAD' } |
				Select-Object -ExpandProperty FullName
		}
		$dirs | Sort-Object -Unique
	}
	$tplDirs = Get-AcadTemplateDirs
	foreach ($d in $tplDirs) {
		try {
			$dest = Join-Path $d 'Modifiable_Template.dwg'
			if (-not (Test-Path -LiteralPath $dest)) {
				Copy-Item -LiteralPath $TemplateSrc -Destination $dest -ErrorAction SilentlyContinue
				if (Test-Path -LiteralPath $dest) {
					Write-Log ("Copied Modifiable_Template.dwg to {0}" -f $d) info
				} else {
					Write-Log ("[WARN] Failed to copy Modifiable_Template.dwg to {0}" -f $d) warn
				}
			} else {
				Write-Log ("Template already present in {0}; skipping copy" -f $d) debug
			}
		} catch {
			Write-Log ("[WARN] Exception while ensuring template in {0}: {1}" -f $d, $_.Exception.Message) warn
		}
	}

	# Prepare to modify an existing named page setup inside the template
	$modifyScr = Join-Path $TempRoot 'modify_template.scr'
	$outNameQ = ($ChosenPsNameOut -replace '\\','/')
	$PsNamesPath = Join-Path $TempRoot 'psnames.txt'
	$outListQ = ($PsNamesPath -replace '\\','/')
	$devArg = $Plan.Device
	$devArgPath = if ($pc3Hit) { $pc3Hit.FullName } else { $devArg }
	$ctbArg = $Plan.CTB
	$mediaArg = $Plan.CanonicalMedia
	$prefName = $Plan.OverrideName  # try to use this if present in template
	$devArgQ = ($devArgPath -replace '\\','/')
	$ctbArgQ = ($ctbArg -replace '\\','/')
	$mediaArgEsc = $mediaArg.Replace('"','\"')
	$prefNameEsc = $prefName.Replace('"','\"')
	$ModifyTracePath = $null
	$ScrDbgMod = ''
	if ($IsDebug) {
		$ModifyTracePath = Join-Path $TempRoot 'prompt_trace_modify.log'
		$traceModQ = ($ModifyTracePath -replace '\\','/')
		$ScrDbgMod = '(setvar "LOGFILEMODE" 1)' + "`n" + '(setvar "LOGFILEPATH" "' + $traceModQ + '")' + "`n"
	}
	$ScrTxtM = @"
(setvar "FILEDIA" 0)
(setvar "CMDECHO" 0)
(setvar "SECURELOAD" 0)
(setvar "EXPERT" 5)
$ScrDbgMod
(setvar "BACKGROUNDPLOT" 0)
(if (/= (getvar "TILEMODE") 0) (setvar "TILEMODE" 0))
;; Ensure we are on a valid paper layout
(defun _layout_names_dxf (/ nod laydict pairs acc akey)
	(setq acc '())
	(if (and (setq nod (namedobjdict))
					 (setq laydict (cdr (assoc -1 (dictsearch nod "ACAD_LAYOUT")))))
		(progn
			(setq pairs (_dict_pairs laydict))
			(foreach p pairs
				(setq akey (car p))
				(if (and akey (/= (strcase akey) "MODEL"))
					(setq acc (cons akey acc))))) )
	(reverse acc))
(defun _layout_exists_p (nm)
	(vl-some '(lambda (s) (= (strcase s) (strcase nm))) (_layout_names_dxf)))
(defun _first_layout () (car (_layout_names_dxf)))
(if (_layout_exists_p "Layout1")
	(setvar "CTAB" "Layout1")
	(progn (if (_first_layout) (setvar "CTAB" (_first_layout)))))
; Echo params for diagnostics
(princ (strcat "\n[DBG] Modify params: Pref='" "$prefNameEsc" "' Device='" "$devArgQ" "' CTB='" "$ctbArgQ" "' Media='" "$mediaArgEsc" "'"))
;; ----- Begin embedded LISP payload -----
;; Utilities
(defun _onlyname (p)
	(if (and p (= (type p) 'STR)) (strcat (vl-filename-base p) (vl-filename-extension p)) ""))

;; Pair dictionary keys to their entries (name . ename)
(defun _dict_pairs (d / e out key)
	(setq e (entget d) out '() key nil)
	(foreach x e
		(cond
			((= (car x) 3) (setq key (cdr x)))
			((or (= (car x) 350) (= (car x) 360))
			 (if key (progn (setq out (cons (cons key (cdr x)) out)) (setq key nil))))
		)
	)
	(reverse out)
)

(defun _ps_names (/ nod dct pairs)
	(if (and (setq nod (namedobjdict)) (setq dct (dictsearch nod "ACAD_PLOTSETTINGS")))
		(progn (setq pairs (_dict_pairs (cdr (assoc -1 dct)))) (mapcar 'car pairs))
	)
)

(defun _ps_exists_p (nm)
	(vl-some '(lambda (s) (= (strcase s) (strcase nm))) (_ps_names)))

(defun _media_variants (m / v canon)
	(setq v '())
	(if (and m (> (strlen m) 0))
		(progn
			(setq canon m)
			(setq v (list
				canon
				(vl-string-translate "_" " " canon)
				(vl-string-subst " x " " x " (vl-string-translate "_" " " canon))
			))
			(if (wcmatch (strcase canon) "*ANSI*D*") (setq v (cons "ANSI D (34.00 x 22.00 Inches)" v)))
			(if (wcmatch (strcase canon) "*ARCH*D*") (setq v (cons "ARCH D (36.00 x 24.00 Inches)" v)))
		)
	)
	(reverse (vl-remove-if 'null v))
)

(defun _ensure_ctb (ctbFull / styles f)
	(if (and ctbFull (> (strlen ctbFull) 0))
		(progn
			(setq styles (getvar "STYLESMANAGER"))
			(if (and styles (findfile ctbFull))
				(progn
					(setq f (strcat (vl-string-right-trim "/\\" styles) "/" (_onlyname ctbFull)))
					(if (not (findfile f))
						(progn (vl-file-copy ctbFull f)
						       (princ (strcat "\n[DBG] CTB copied to: " f)))
						(princ (strcat "\n[DBG] CTB present: " f))
					)
				)
			)
		)
	)
	(_onlyname ctbFull)
)

(defun _ensure_pc3 (devPathOrName / devName plotters tgt probe)
	(setq devName (_onlyname devPathOrName))
	(if (= devName "") (setq devName devPathOrName))
	(if (wcmatch (strcase devName) "*.PC3")
		(progn
			(setq plotters (strcat (getvar "ROAMABLEROOTPREFIX") "Plotters"))
			(if (and plotters (findfile devPathOrName))
				(progn
					(setq tgt (strcat (vl-string-right-trim "/\\" plotters) "/" devName))
					(if (not (findfile tgt))
						(progn (vl-file-copy devPathOrName tgt)
						       (princ (strcat "\n[DBG] PC3 copied to: " tgt)))
						(princ (strcat "\n[DBG] PC3 present: " tgt))
					)
				)
			)
			;; Validate that the device file exists; otherwise fallback to DWG To PDF.pc3
			(setq probe (strcat (vl-string-right-trim "/\\" plotters) "/" devName))
			(if (not (findfile probe))
			  (progn
			    (princ (strcat "\n[WRN] Plot device not found: " devName "; falling back to DWG To PDF.pc3"))
			    (setq devName "DWG To PDF.pc3")
			  )
			)
			devName
		)
		devName
	)
)

(defun _ModifyTemplatePS (prefName dev ctbFull media outNamePath outListPath / oldEcho name names devName mediaList unitsTok fh f)
	(setvar "FILEDIA" 0)
	(setq oldEcho (getvar "CMDECHO"))
	(setvar "CMDECHO" 0)
	(setvar "EXPERT" 5)
	(setvar "BACKGROUNDPLOT" 0)
	;; Choose a target page setup
	(setq names (_ps_names))
	(setq name (cond
				((and prefName (> (strlen prefName) 0) (_ps_exists_p prefName)) prefName)
				((vl-some '(lambda (s) (if (= (strcase s) (strcase "NamedSetup")) s)) names))
				((vl-some '(lambda (s) (if (= (strcase s) (strcase "AutoPublish_PS")) s)) names))
				((vl-some '(lambda (s) (if (= (strcase s) (strcase "ClientStandard")) s)) names))
				((car names))
	))
	(princ (strcat "\n[DBG] Found " (itoa (length names)) " named page setup(s)."))
	(if name (princ (strcat "\n[DBG] Chosen page setup: " name)))
	;; Emit optional list of names for diagnostics
	(if outListPath (progn
		(setq fh (open outListPath "w"))
		(if fh (progn (foreach n names (write-line n fh)) (close fh)))
	))
	;; Bail if none present
	(if (not name)
		(progn
			(princ "\n[ERR] No named page setups found in template.")
			(if outNamePath (progn (setq f (open outNamePath "w")) (if f (progn (write-line "" f) (close f)))))
			(princ)
		)
	)
	;; Write the chosen name early so diagnostics reflect discovery even if modification fails
	(if (and name outNamePath)
		(progn (setq f (open outNamePath "w")) (if f (progn (write-line name f) (close f)))))
	;; Ensure PS mode is CTB after we've discovered names
	(if (= (getvar "PSTYLEMODE") 0) (command "._CONVERTPSTYLES" "_Y"))
	;; Ensure resources
	(setq devName (_ensure_pc3 dev))
	(if (or (not devName) (= (strlen devName) 0)) (setq devName "DWG To PDF.pc3"))
	(setq ctbName (_ensure_ctb ctbFull))
	(if (or (not ctbName) (= (strlen ctbName) 0)) (setq ctbName ""))
	(setq mediaList (_media_variants media))
	(princ (strcat "\n[DBG] Media candidates: " (vl-string-translate "\n" " " (apply 'strcat (mapcar '(lambda (m) (strcat "[" m "] ")) mediaList)))))
	(if (= (getvar "MEASUREMENT") 1)
		(setq unitsTok "Millimeters")
		(setq unitsTok "Inches")
	)
	;; Use first media candidate or blank
	(if (not mediaList) (setq mediaList (list "")))
	(princ (strcat "\n[DBG] Using device='" devName "' media='" (car mediaList) "' units=" unitsTok " CTB='" ctbName "'"))
	(princ "\n[DBG] About to call -PAGESETUP...")
	;; Prepare CTB token ('.' clears plot style)
	(setq ctbTok (if (> (strlen ctbName) 0) ctbName "."))
	;; Build attempt matrices: devices, media candidates, token variants
	(setq devCandidates (if (= (strcase devName) (strcase "DWG To PDF.pc3"))
									(list devName)
									(list devName "DWG To PDF.pc3")))
	(setq medCandidates mediaList)
	(setq success nil)
	(foreach d devCandidates
		(foreach m medCandidates
			(if (not success)
				(progn
					(princ (strcat "\n[DBG] Trying device='" d "' media='" m "' units=" unitsTok " (localized tokens)"))
					(setq err1 (vl-catch-all-apply 'vl-cmdf (list "._-PAGESETUP" d m unitsTok "Landscape" "No" "Layout" "1:1" "0.00,0.00" "Yes" ctbTok "Yes" "No" "No" "No")))
					(if (not (vl-catch-all-error-p err1)) (setq success (list d m "loc")))
				)
			)
			(if (and (not success))
				(progn
					(setq unitsTokI (if (= (strcase unitsTok) (strcase "Inches")) "_Inches" "_Millimeters"))
					(setq yesI "_Yes" noI "_No" landI "_Landscape" layI "_Layout")
					(princ (strcat "\n[DBG] Retrying device='" d "' media='" m "' (intl tokens)"))
					(setq err2 (vl-catch-all-apply 'vl-cmdf (list "._-PAGESETUP" d m unitsTokI landI noI layI "1:1" "0.00,0.00" yesI ctbTok yesI noI noI noI)))
					(if (not (vl-catch-all-error-p err2)) (setq success (list d m "intl")))
				)
			)
			;; Name-first variants as a last resort
			(if (and (not success) name)
				(progn
					(princ (strcat "\n[DBG] Name-first attempt for setup='" name "' device='" d "' media='" m "' (loc)"))
					(setq err3 (vl-catch-all-apply 'vl-cmdf (list "._-PAGESETUP" name d m unitsTok "Landscape" "No" "Layout" "1:1" "0.00,0.00" "Yes" ctbTok "Yes" "No" "No" "No")))
					(if (not (vl-catch-all-error-p err3)) (setq success (list d m "name-loc")))
				)
			)
			(if (and (not success) name)
				(progn
					(princ (strcat "\n[DBG] Name-first attempt for setup='" name "' device='" d "' media='" m "' (intl)"))
					(setq err4 (vl-catch-all-apply 'vl-cmdf (list "._-PAGESETUP" name d m unitsTokI landI noI layI "1:1" "0.00,0.00" yesI ctbTok yesI noI noI noI)))
					(if (not (vl-catch-all-error-p err4)) (setq success (list d m "name-intl")))
				)
			)
		)
	)
	(if (not success)
		(progn (princ "\n[ERR] All -PAGESETUP permutations failed.") (exit))
	)
	(princ (strcat "\n[DBG] -PAGESETUP succeeded with device='" (car success) "' media='" (cadr success) "' mode=" (caddr success)))
	;; Optionally sync to preferred named page setup (prefName), using the successful tokens
	(setq finalName name)
	(if (and prefName (> (strlen prefName) 0))
		(progn
			(setq dUsed (car success))
			(setq mUsed (cadr success))
			(setq modeTag (caddr success))
			(setq useIntl (wcmatch (strcase modeTag) "*INTL*"))
			(setq utok (if useIntl (if (= (strcase unitsTok) (strcase "Inches")) "_Inches" "_Millimeters") unitsTok))
			(setq landTok (if useIntl "_Landscape" "Landscape"))
			(setq noTok (if useIntl "_No" "No"))
			(setq yesTok (if useIntl "_Yes" "Yes"))
			(setq layTok (if useIntl "_Layout" "Layout"))
			(princ (strcat "\n[DBG] Syncing named page setup to '" prefName "' using device='" dUsed "' media='" mUsed "'"))
			(setq errSync (vl-catch-all-apply 'vl-cmdf (list "._-PAGESETUP" prefName dUsed mUsed utok landTok noTok layTok "1:1" "0.00,0.00" yesTok ctbTok yesTok noTok noTok noTok)))
			(if (not (vl-catch-all-error-p errSync))
				(progn (setq finalName prefName) (princ "\n[DBG] Named page setup synced."))
				(princ (strcat "\n[WRN] Failed to sync named page setup '" prefName "': " (vl-catch-all-error-message errSync)))
			)
        )
	)
	(princ "\n[OK] -PAGESETUP completed")
	(if outNamePath (progn (setq f (open outNamePath "w")) (if f (progn (write-line finalName f) (close f)))))
	(setvar "CMDECHO" oldEcho)
	(princ)
)

(_ModifyTemplatePS "$prefNameEsc" "$devArgQ" "$ctbArgQ" "$mediaArgEsc" "$outNameQ" "$outListQ")
;; ----- End embedded LISP payload -----
._QSAVE
._QUIT
"@
	$ScrTxtM | Set-Content -Encoding ASCII $modifyScr
	Write-Log "Modifying page setup in Modifiable_Template_work.dwg..." info
	$outM = & "$AccorePath" /product ACAD /l en-US /i "$TemplateWorkDwg" /s "$modifyScr" 2>&1
	# Persist full CoreConsole output for diagnostics
	try {
		$outM | Set-Content -LiteralPath (Join-Path $TempRoot 'accore_modify_console.log') -Encoding UTF8
	} catch {}
	$ChosenName = if (Test-Path -LiteralPath $ChosenPsNameOut) { (Get-Content -LiteralPath $ChosenPsNameOut -Raw).Trim() } else { '' }
	$okSentinel = ($outM | Where-Object { $_ -match "\[OK\] -PAGESETUP completed" })
	# If CoreConsole returned non-zero but we saw the success sentinel, continue with a warning
	if ($LASTEXITCODE -ne 0 -and -not $okSentinel) {
		Write-Host "Failed to modify page setup in template (accoreconsole error)." -ForegroundColor Red
		$outM | Out-Host
		exit 9
	}
	if ($LASTEXITCODE -ne 0 -and $okSentinel) {
		Write-Log "CoreConsole returned non-zero, but success sentinel detected; continuing." warn
	}
	if (-not $okSentinel) {
		Write-Host "-PAGESETUP did not reach completion; likely waiting for an unexpected prompt." -ForegroundColor Red
		if ($IsDebug -and $outM) {
			Write-Log "CoreConsole modify output (tail):" debug
			$outM | Select-Object -Last 120 | ForEach-Object { Write-Log $_ debug }
		}
		if ($IsDebug -and $ModifyTracePath -and (Test-Path -LiteralPath $ModifyTracePath)) {
			Write-Log ("Prompt trace (modify) saved to: {0}" -f $ModifyTracePath) debug
			Get-Content -LiteralPath $ModifyTracePath -Tail 200 | ForEach-Object { Write-Log $_ debug }
		}
		exit 9
	}
	if (-not $ChosenName) {
		Write-Host "Page setup modification did not complete in Modifiable_Template.dwg." -ForegroundColor Red
		Write-Host "A named page setup was discovered, but -PAGESETUP likely awaited input. See traces below." -ForegroundColor Yellow
		Write-Host ("Temp working directory: {0}" -f $TempRoot) -ForegroundColor DarkGray
		if (Test-Path -LiteralPath $PsNamesPath) {
			Write-Log ("Discovered page setup names (if any):" ) debug
			Get-Content -LiteralPath $PsNamesPath | ForEach-Object { Write-Log ("  - {0}" -f $_) debug }
		}
		# Dump CoreConsole output tail in debug
		if ($IsDebug -and $outM) {
			Write-Log "CoreConsole modify output (tail):" debug
			$outM | Select-Object -Last 80 | ForEach-Object { Write-Log $_ debug }
			Write-Log ("Full console log: {0}" -f (Join-Path $TempRoot 'accore_modify_console.log')) debug
		}
		if ($IsDebug -and $ModifyTracePath -and (Test-Path -LiteralPath $ModifyTracePath)) {
			Write-Log ("Prompt trace (modify) saved to: {0}" -f $ModifyTracePath) debug
			Get-Content -LiteralPath $ModifyTracePath -Tail 100 | ForEach-Object { Write-Log $_ debug }
		}
		exit 9
	}
	$Plan.OverrideSource = $TemplateWorkDwg
	$Plan.OverrideName   = $ChosenName
	Write-Log ("Using override: Setup='{0}' from '{1}'" -f $Plan.OverrideName,$Plan.OverrideSource) info
}

# COM-free enumeration: run one CoreConsole per DWG (throttled) to list paper layouts with content
$LayoutsLispPath = Join-Path $TempRoot 'ListLayoutsWithContent.lsp'
$LayoutsLisp = @'
(defun _dict_pairs (d / e out key)
	(setq e (entget d) out '() key nil)
	(foreach x e
		(cond
			((= (car x) 3)   (setq key (cdr x)))
			((or (= (car x) 350) (= (car x) 360))
			 (if key (progn (setq out (cons (cons key (cdr x)) out))
											(setq key nil))))
		)
	)
	(reverse out)
)
(defun _layout_names_dxf (/ nod laydict pairs acc akey)
	(setq acc '())
	(if (and (setq nod (namedobjdict))
					 (setq laydict (cdr (assoc -1 (dictsearch nod "ACAD_LAYOUT")))))
		(progn
			(setq pairs (_dict_pairs laydict))
			(foreach p pairs
				(setq akey (car p))
				(if (and akey (/= (strcase akey) "MODEL"))
					(setq acc (cons akey acc))
				)
			)
		)
	)
	(reverse acc)
)
(defun _fn_defined_p (sym)
	(and (= (type sym) 'SYM)
			 (vl-position sym (atoms-family 1)))
)
(if (not (_fn_defined_p 'layoutlist))
	(defun layoutlist (/) (_layout_names_dxf))
)
(defun _has_content_p (lyt / ss)
	;; Consider any entity in the layout as content (including viewports)
	(setq ss (ssget "_X" (list (cons 410 lyt))))
	(and ss (> (sslength ss) 0))
)
(defun _write_csv (pth rows / fh)
	(setq fh (open pth "w"))
	(if fh (progn (write-line "DWG,Layout" fh)
								(foreach r rows (write-line (strcat (car r) "," (cadr r)) fh))
								(close fh)))
)
(defun c:ListLayoutsWithContent_toCsv (outCsv / dwp dn full acc)
	(setvar "FILEDIA" 0)
	(setvar "CMDECHO" 0)
	(setq dwp (getvar "DWGPREFIX")
				dn  (getvar "DWGNAME")
				full (strcat dwp dn)
				acc '())
	(foreach l (layoutlist)
		(if (_has_content_p l)
			(setq acc (cons (list full l) acc))))
	(_write_csv outCsv (reverse acc))
	(princ)
)
'@
$LayoutsLisp | Set-Content -Encoding ASCII $LayoutsLispPath

$dwgs = Get-Content -LiteralPath $ListPath
$running = @()
$workerCsvs = @()
Write-Log ("Enumerating layouts with content using {0} workers..." -f $Workers) info
foreach ($dwg in $dwgs) {
	if (-not (Test-Path -LiteralPath $dwg)) { continue }
	$wdir = Join-Path $TempRoot ([IO.Path]::GetFileNameWithoutExtension($dwg))
	New-Item -ItemType Directory -Force -Path $wdir | Out-Null
	$csv  = Join-Path $wdir 'layouts.csv'
	$workerCsvs += $csv

	$lspQ = ($LayoutsLispPath -replace '\\','/')
	$csvQ = ($csv -replace '\\','/')

	$scr = @"
(setvar "FILEDIA" 0)
(setvar "CMDECHO" 0)
(setvar "SECURELOAD" 0)
(if (not (load "$lspQ" nil)) (progn (princ "\nERROR: failed to load lisp") (quit)))
(c:ListLayoutsWithContent_toCsv "$csvQ")
._QUIT _Y
"@
	$scrPath = Join-Path $wdir 'enumerate.scr'
	$scr | Set-Content -Encoding ASCII $scrPath

	$psi = New-Object System.Diagnostics.ProcessStartInfo
	$psi.FileName = $AccorePath
	$psi.Arguments = "/product ACAD /l en-US /i `"$dwg`" /s `"$scrPath`""
	$psi.WorkingDirectory = $wdir
	$psi.UseShellExecute = $true
	$p = [System.Diagnostics.Process]::Start($psi)
	$running += $p

	while ($running.Count -ge $Workers) {
		$running = $running | Where-Object { $_ -and -not $_.HasExited }
		Start-Sleep -Milliseconds 100
	}
}
foreach ($p in $running) { if ($p) { $null = $p.WaitForExit() } }

$LayoutsCsv = Join-Path $TempRoot 'layouts.csv'
"DWG,Layout" | Set-Content -Encoding ASCII $LayoutsCsv
foreach ($csv in $workerCsvs) {
	if (Test-Path -LiteralPath $csv) {
		(Get-Content -LiteralPath $csv | Select-Object -Skip 1) | Add-Content -Encoding ASCII $LayoutsCsv
	}
}
if (-not (Test-Path -LiteralPath $LayoutsCsv)) { Write-Host "No layouts CSV produced." -ForegroundColor Red; exit 10 }
$sheets = Import-Csv -LiteralPath $LayoutsCsv
if (-not $sheets -or $sheets.Count -eq 0) { Write-Host "No paper layouts found to publish." -ForegroundColor Yellow; exit 11 }

# Build DSD file
$DsdPath = Join-Path $TempRoot 'publish.dsd'
$LogPath = Join-Path $TempRoot 'publish.log'
$outDir  = Split-Path -Parent $OutPdf
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

function Write-DSD {
	param($Sheets, [string]$OutPdfPath, [string]$OverrideName, [string]$OverrideFile, [string]$LogFile)
    # Expand environment variables to avoid %USERPROFILE% in DSD
    $OutPdfPath = [Environment]::ExpandEnvironmentVariables($OutPdfPath)
    $OverrideFile = if ($OverrideFile) { [Environment]::ExpandEnvironmentVariables($OverrideFile) } else { $OverrideFile }
    $outFolder = Split-Path -Parent $OutPdfPath
	$lines = @()
	# Header matches "CorrectExample" structure
	$lines += '[DWF6Version]'
	$lines += 'Ver=1'
	$lines += '[DWF6MinorVersion]'
	$lines += 'MinorVer=1'

	foreach ($s in $Sheets) {
		$base = [IO.Path]::GetFileNameWithoutExtension($s.DWG)
		$title = "$base-$($s.Layout)"
		$lines += "[DWF6Sheet:$title]"
		$lines += "DWG=$($s.DWG)"
		$lines += "Layout=$($s.Layout)"
		if ($OverrideFile -and $OverrideName) {
			$lines += "Setup=$OverrideName|$OverrideFile"
		} else {
			$lines += 'Setup='
		}
		$lines += "OriginalSheetPath=$($s.DWG)"
		$lines += 'Has Plot Port=0'
		$lines += 'Has3DDWF=0'
	}

	# Target block
	if (-not $outFolder.EndsWith('\')) { $outFolder = $outFolder + '\' }
	$lines += '[Target]'
	$lines += 'Type=6'
	$lines += ("DWF={0}" -f $OutPdfPath)
	$lines += ("OUT={0}" -f $outFolder)
	$lines += 'PWD='

	# Optional MRU sections (benign placeholders)
	$lines += '[MRU Local]'
	$lines += 'MRU=0'
	$lines += '[MRU Sheet List]'
	$lines += 'MRU=1'
	$lines += ("File0={0}" -f $DsdPath)

	# Pdf options
	$lines += '[PdfOptions]'
	$lines += 'IncludeHyperlinks=FALSE'
	$lines += 'CreateBookmarks=TRUE'
	$lines += 'CaptureFontsInDrawing=TRUE'
	$lines += 'ConvertTextToGeometry=FALSE'
	$lines += 'VectorResolution=600'
	$lines += 'RasterResolution=400'

	# Block data section
	$lines += '[AutoCAD Block Data]'
	$lines += 'IncludeBlockInfo=0'
	$lines += 'BlockTmplFilePath='

	# Sheet set properties
	$lines += '[SheetSet Properties]'
	$lines += 'IsSheetSet=FALSE'
	$lines += 'IsHomogeneous=FALSE'
	$lines += 'SheetSet Name='
	$lines += 'NoOfCopies=1'
	$lines += 'PlotStampOn=FALSE'
	$lines += 'ViewFile=FALSE'
	$lines += 'JobID=0'
	$lines += 'SelectionSetName='
	$lines += 'AcadProfile=NewWorkflow'
	$lines += 'CategoryName='
	$lines += 'LogFilePath='
	$lines += 'IncludeLayer=TRUE'
	$lines += 'LineMerge=FALSE'
	$lines += 'CurrentPrecision='
	$lines += 'PromptForDwfName=TRUE'
	$lines += 'PwdProtectPublishedDWF=FALSE'
	$lines += 'PromptForPwd=FALSE'
	$lines += 'RepublishingMarkups=FALSE'
	$lines += 'PublishSheetSetMetadata=FALSE'
	$lines += 'PublishSheetMetadata=FALSE'
	$lines += '3DDWFOptions=0 0'

	Set-Content -LiteralPath $DsdPath -Value $lines -Encoding ASCII
}

Write-Log "Building DSD..." info
Write-DSD -Sheets $sheets -OutPdfPath $OutPdf -OverrideName $Plan.OverrideName -OverrideFile $Plan.OverrideSource -LogFile $LogPath

# Publish
$scrPublish = Join-Path $TempRoot 'publish.scr'
$dsdQ = ($DsdPath -replace '\\','/')
$PublishTracePath = $null
$ScrDbg2 = ''
if ($IsDebug) {
	$PublishTracePath = Join-Path $TempRoot 'prompt_trace_publish.log'
	$trace2Q = ($PublishTracePath -replace '\\','/')
	$ScrDbg2 = '(setvar "LOGFILEMODE" 1)' + "`n" + '(setvar "LOGFILEPATH" "' + $trace2Q + '")' + "`n"
}
$ScrTxt3 = @"
(setvar "FILEDIA" 0)
(setvar "CMDECHO" 0)
$ScrDbg2
_.-PUBLISH "$dsdQ"
._QUIT _Y
"@
$ScrTxt3 | Set-Content -Encoding ASCII $scrPublish

Write-Host "" 
Write-Host "Publishing (single multi-sheet PDF)..." -ForegroundColor Cyan
$out3 = & "$AccorePath" /product ACAD /l en-US /s "$scrPublish" 2>&1
if ($LASTEXITCODE -ne 0) {
	Write-Host "Publish failed." -ForegroundColor Red
	$out3 | Out-Host
	if (Test-Path -LiteralPath $LogPath) { Write-Host "See log: $LogPath" }
	if ($PublishTracePath -and (Test-Path -LiteralPath $PublishTracePath)) {
		Write-Log ("Prompt trace (publish) saved to: {0}" -f $PublishTracePath) debug
		Get-Content -LiteralPath $PublishTracePath -Tail 100 | ForEach-Object { Write-Log $_ debug }
	}
	exit 12
}

Write-Host "Publish complete: $OutPdf" -ForegroundColor Green
if (Test-Path -LiteralPath $LogPath) { Write-Log "Publish log at $LogPath" info }
if ($PublishTracePath -and (Test-Path -LiteralPath $PublishTracePath)) {
	Write-Log ("Prompt trace (publish) saved to: {0}" -f $PublishTracePath) debug
}

exit 0