# ManifestGui_Rebound.ps1
# WPF + XAML + data binding (MVVM-style) for tri-state checkbox propagation.
# - Parents show Indeterminate when children are mixed (computed from children).
# - Users cannot directly select Indeterminate (users toggle only true/false).
# - No external UI libraries required.

$ErrorActionPreference = 'Stop'

# ------------------------------------------------------------
# Load core WPF assemblies
# ------------------------------------------------------------
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xml

# System.Xaml is required for WPF controls (IQueryAmbient, etc.). Try by name, then by path fallback.
try { Add-Type -AssemblyName System.Xaml -ErrorAction Stop } catch {
    # In PowerShell 7+ do NOT load .NET Framework 4.x System.Xaml from WINDIR;
    # that causes mixed-assembly version errors at C# compile time.
    if ($PSVersionTable.PSEdition -eq 'Desktop') {
        $xamlPath = Join-Path $env:WINDIR "Microsoft.NET\Framework\v4.0.30319\WPF\System.Xaml.dll"
        if (-not (Test-Path $xamlPath)) {
            $xamlPath = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\WPF\System.Xaml.dll"
        }
        if (Test-Path $xamlPath) { Add-Type -Path $xamlPath }
    }
}

# ------------------------------------------------------------
# C# model (compatible with older compilers—no auto-prop initializers)
# Use a unique namespace per run to avoid stale types in the same session.
# ------------------------------------------------------------
$csNamespace = "ManifestGuiModel_{0}" -f ([System.Guid]::NewGuid().ToString("N"))

$cs = @"
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Xml;
using System.Windows.Controls;

namespace $csNamespace
{
    public enum NodeKind { Template, InstanceGroup, Instance, Folder }

    public class TreeNode : INotifyPropertyChanged
    {
        private bool? _isChecked;
        private bool _explicitlyUnchecked;
        private TreeNode _parent;
        private ObservableCollection<TreeNode> _children;
        private bool _isExpanded = true;

        public static int BulkUpdateDepth = 0;

        public TreeNode()
        {
            _children = new ObservableCollection<TreeNode>();
        }

        public string Name { get; set; }
        public NodeKind Kind { get; set; }

        public ObservableCollection<TreeNode> Children
        {
            get { return _children; }
            set { _children = value; }
        }

        public TreeNode Parent
        {
            get { return _parent; }
            private set { _parent = value; }
        }

        public XmlNode XmlRef { get; set; }
        public bool IsCommented { get; set; }
        public string SourceXml { get; set; }
        public string IndexPath { get; set; }
        public string ContainerCommentPath { get; set; }
        public string ContainerSourceXml { get; set; }

        public bool IsExpanded
        {
            get { return _isExpanded; }
            set
            {
                if (_isExpanded == value) return;
                _isExpanded = value;
                OnPropertyChanged("IsExpanded");
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public bool? IsChecked
        {
            get { return _isChecked; }
            set { SetIsChecked(value, true, true); }
        }

        public void Attach(TreeNode child)
        {
            child.Parent = this;
            _children.Add(child);
        }

        public void SetIsChecked(bool? value, bool updateChildren, bool updateParent)
        {
            if (_isChecked == value)
            {
                // Still honor explicit intent and force propagation when requested.
                if (value == false && updateChildren)
                    _explicitlyUnchecked = true;
                else if (value == true)
                    _explicitlyUnchecked = false;

                if (updateChildren && value.HasValue)
                {
                    BulkUpdateDepth++;
                    try
                    {
                        foreach (var c in _children)
                            c.SetIsChecked(value, true, false);
                    }
                    finally
                    {
                        BulkUpdateDepth--;
                    }
                }
                return;
            }

            _isChecked = value;
            // Track explicit uncheck intent for this node when a definite value is set
            if (value == false && updateChildren)
                _explicitlyUnchecked = true;
            else if (value == true)
                _explicitlyUnchecked = false;
            OnPropertyChanged("IsChecked");

            // Push down to children only when value is definite (true/false).
            if (updateChildren && value.HasValue)
            {
                BulkUpdateDepth++;
                try
                {
                    foreach (var c in _children)
                        c.SetIsChecked(value, true, false);
                }
                finally
                {
                    BulkUpdateDepth--;
                }
            }

            // Recompute parent
            if (updateParent && _parent != null)
                _parent.RecomputeFromChildren();
        }

        public void RecomputeFromChildren()
        {
            if (_children == null || _children.Count == 0)
                return;

            int count  = _children.Count;
            int trues  = _children.Count(c => c.IsChecked == true);
            int falses = _children.Count(c => c.IsChecked == false);

            bool allTrue  = trues  == count;
            bool allFalse = falses == count;

            bool? newVal;
            if (allTrue)
            {
                newVal = true;
                _explicitlyUnchecked = false;
            }
            else if (allFalse)
            {
                // For instance folders/groups: if all children unchecked, show false.
                if (this.Kind == NodeKind.Folder || this.Kind == NodeKind.InstanceGroup)
                {
                    newVal = false;
                    _explicitlyUnchecked = true;
                }
                else
                {
                    // For templates: only keep false when explicitly unchecked; otherwise indeterminate.
                    newVal = _explicitlyUnchecked ? (bool?)false : (bool?)null;
                }
            }
            else
            {
                // Mixed children -> indeterminate
                newVal = null;
                _explicitlyUnchecked = false;
            }

            if (_isChecked != newVal)
            {
                _isChecked = newVal;
                OnPropertyChanged("IsChecked");
            }

            if (_parent != null) _parent.RecomputeFromChildren();
        }

        protected void OnPropertyChanged(string name)
        {
            var handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(name));
        }
    }

    // Prevent users from landing on Indeterminate: toggle only between true/false.
    public class TwoStateUserCheckBox : CheckBox
    {
        protected override void OnToggle()
        {
            bool next = this.IsChecked != true; // true -> false, false/null -> true
            this.IsChecked = next;
        }
    }
}
"@

# Compile the C# types with explicit references from the currently loaded runtime
# to avoid version-mismatch issues (especially in PowerShell 7+).
$refAssemblies = @(
    [System.Object].Assembly.Location,
    [System.Linq.Enumerable].Assembly.Location,
    [System.ComponentModel.INotifyPropertyChanged].Assembly.Location,
    [System.Xml.XmlDocument].Assembly.Location,
    [System.Windows.DependencyObject].Assembly.Location,
    [System.Windows.Media.Visual].Assembly.Location,
    [System.Windows.Controls.Control].Assembly.Location
)

$systemXamlAsm = [AppDomain]::CurrentDomain.GetAssemblies() |
    Where-Object { $_.GetName().Name -eq 'System.Xaml' } |
    Select-Object -First 1
if ($systemXamlAsm -and $systemXamlAsm.Location) {
    $refAssemblies += $systemXamlAsm.Location
}

$refAssemblies = $refAssemblies | Where-Object { $_ } | Select-Object -Unique

Add-Type -TypeDefinition $cs -ReferencedAssemblies $refAssemblies

# Bind dynamic type references
$script:TreeNodeType = ("$csNamespace.TreeNode" -as [type])
$script:NodeKindType = ("$csNamespace.NodeKind" -as [type])
$script:NodeKind_Template = [System.Enum]::Parse($script:NodeKindType, "Template")
$script:NodeKind_InstanceGroup = [System.Enum]::Parse($script:NodeKindType, "InstanceGroup")
$script:NodeKind_Instance = [System.Enum]::Parse($script:NodeKindType, "Instance")
$script:NodeKind_Folder = [System.Enum]::Parse($script:NodeKindType, "Folder")

# Grab the dynamic assembly name so XAML can resolve local: types
$asmName = $script:TreeNodeType.Assembly.GetName().Name

# ------------------------------------------------------------
# XAML with HierarchicalDataTemplate (data-bound tree)
# IMPORTANT: include the dynamic assembly name in the xmlns
# ------------------------------------------------------------
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:local="clr-namespace:$csNamespace;assembly=$asmName"
        Title="Manifest Selector (Data-bound)" Height="700" Width="900" WindowStartupLocation="CenterScreen">
    <Window.Resources>
    <!-- Render any TreeNode with a tri-state checkbox + text -->
    <HierarchicalDataTemplate DataType="{x:Type local:TreeNode}" ItemsSource="{Binding Children}">
      <StackPanel Orientation="Horizontal">
        <local:TwoStateUserCheckBox
            IsThreeState="True"
            IsChecked="{Binding IsChecked, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"
            Focusable="False"
            VerticalAlignment="Center"/>
        <TextBlock Text="{Binding Name}" Margin="6,0,0,0" VerticalAlignment="Center"/>
      </StackPanel>
    </HierarchicalDataTemplate>
        <Style TargetType="TreeViewItem">
            <Setter Property="IsExpanded" Value="{Binding IsExpanded, Mode=TwoWay}" />
        </Style>
  </Window.Resources>

  <Grid Margin="10">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

        <Grid Grid.Row="0" Margin="0,0,0,8">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

            <StackPanel Grid.Row="0" Orientation="Horizontal">
                <Button x:Name="btnLoad" Content="Load Manifest.xml" Width="150" Margin="0,0,8,0"/>
                <Button x:Name="btnSave" Content="Save Manifest.xml" Width="170" Margin="0,0,8,0" IsEnabled="False"/>
                <CheckBox x:Name="chkOverwriteOriginal" Content="Overwrite original on save" IsChecked="False" VerticalAlignment="Center" Margin="0,0,8,0"/>
                <CheckBox x:Name="chkHardDeleteAapkg" Content="AAPKG hard delete unchecked (remove XML + files)" IsChecked="False" IsEnabled="False" VerticalAlignment="Center" Margin="0,0,8,0"/>
            </StackPanel>

            <TextBox Grid.Row="1" x:Name="txtStatus" IsReadOnly="True" BorderThickness="0" Background="Transparent" Padding="0" Margin="0,6,0,0" TextWrapping="Wrap" AcceptsReturn="True" VerticalScrollBarVisibility="Disabled" HorizontalScrollBarVisibility="Disabled"/>
            <TextBox Grid.Row="2" x:Name="txtOutputPath" IsReadOnly="True" BorderThickness="0" Background="Transparent" Padding="0" Margin="0,4,0,0" Foreground="#444" TextWrapping="Wrap" AcceptsReturn="True" VerticalScrollBarVisibility="Disabled" HorizontalScrollBarVisibility="Disabled"/>
        </Grid>

        <Border Grid.Row="1" BorderBrush="#CCC" BorderThickness="1" CornerRadius="4" Padding="6">
            <ScrollViewer x:Name="svManifest" VerticalScrollBarVisibility="Auto">
        <TreeView x:Name="treeManifest"/>
      </ScrollViewer>
    </Border>

    <TextBlock Grid.Row="2" Margin="0,8,0,0" Foreground="#666">
      Tip: Parents become indeterminate when children are mixed. Click toggles only between checked and unchecked.
    </TextBlock>
  </Grid>
</Window>
"@

# Helper to load XAML
$reader  = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
$window  = [Windows.Markup.XamlReader]::Load($reader)
$btnLoad = $window.FindName("btnLoad")
$btnSave = $window.FindName("btnSave")
$chkOverwriteOriginal = $window.FindName("chkOverwriteOriginal")
$chkHardDeleteAapkg = $window.FindName("chkHardDeleteAapkg")
$tree    = $window.FindName("treeManifest")
$sv      = $window.FindName("svManifest")
$txt     = $window.FindName("txtStatus")
$txtOutput = $window.FindName("txtOutputPath")

# ------------------------------------------------------------
# Script-level state andM binding hookup
# ------------------------------------------------------------
$script:ManifestXml = $null
$script:RootNodes   = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$script:SourceAapkgPath = $null
$script:SourceCabPath = $null
$script:SourceManifestPath = $null
$script:SourceXmlPath = $null
$script:SourceTempRoot = $null
$script:SevenZipPath = $null

# Synthetic, non-visual root that parents all top-level templates
$script:HiddenRoot = New-Object $script:TreeNodeType.FullName
$script:HiddenRoot.Name = "(root)"
$script:HiddenRoot.Kind = $script:NodeKind_Template
$script:HiddenRoot.SetIsChecked($true, $false, $false)

# Bind the TreeView to the ObservableCollection
$tree.ItemsSource = $script:RootNodes

# Ensure mouse wheel scrolls the ScrollViewer when the TreeView has focus
$tree.AddHandler(
    [System.Windows.UIElement]::PreviewMouseWheelEvent,
    [System.Windows.Input.MouseWheelEventHandler]{
        if ($sv) {
            $sv.ScrollToVerticalOffset($sv.VerticalOffset - $_.Delta)
            $_.Handled = $true
        }
    },
    $true
)

# Normalize after any Check/Uncheck in the tree (mouse or keyboard)
$tree.AddHandler(
    [System.Windows.Controls.Primitives.ToggleButton]::CheckedEvent,
    [System.Windows.RoutedEventHandler]{ if ($script:TreeNodeType::BulkUpdateDepth -eq 0) { Normalize-All $script:HiddenRoot } },
    $true
)
$tree.AddHandler(
    [System.Windows.Controls.Primitives.ToggleButton]::UncheckedEvent,
    [System.Windows.RoutedEventHandler]{ if ($script:TreeNodeType::BulkUpdateDepth -eq 0) { Normalize-All $script:HiddenRoot } },
    $true
)

# ------------------------------------------------------------
# Utilities
# ------------------------------------------------------------
function Get-NodeLabel([System.Xml.XmlNode]$node) {
    if (-not $node) { return "(missing)" }
    foreach ($n in "tag_name","tagName","Name","name","Id","id") {
        if ($node.Attributes[$n]) { return $node.Attributes[$n].Value }
    }
    return $node.Name
}

function Find-7Zip {
    $scriptDir = $null

    if ($PSScriptRoot) {
        $scriptDir = $PSScriptRoot
    }
    elseif ($PSCommandPath) {
        $scriptDir = Split-Path -Parent $PSCommandPath
    }
    elseif ($MyInvocation -and $MyInvocation.MyCommand -and $MyInvocation.MyCommand.Path) {
        $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    }

    $paths = New-Object System.Collections.Generic.List[string]

    if ($scriptDir) {
        $paths.Add((Join-Path $scriptDir "7-ZipPortable\App\7-Zip\7z.exe"))
        $paths.Add((Join-Path $scriptDir "7z.exe"))
    }

    if ($env:ProgramFiles) {
        $paths.Add((Join-Path $env:ProgramFiles "7-Zip\7z.exe"))
    }
    if (${env:ProgramFiles(x86)}) {
        $paths.Add((Join-Path ${env:ProgramFiles(x86)} "7-Zip\7z.exe"))
    }

    $paths.Add("C:\Program Files\7-Zip\7z.exe")
    $paths.Add("C:\Program Files (x86)\7-Zip\7z.exe")

    foreach ($p in $paths) { if ($p -and (Test-Path -LiteralPath $p)) { return $p } }
    $cmd = Get-Command 7z -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    return $null
}

function Invoke-7Zip {
    param(
        [string]$SevenZip,
        [string[]]$Args
    )
    if (-not $SevenZip) { throw "7-Zip not found." }

    # Capture output to help debug but do not spam UI
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $SevenZip
    $psi.Arguments = ($Args | ForEach-Object {
        if ($_ -match '[\s"]') { '"' + ($_ -replace '"', '\\"') + '"' } else { $_ }
    }) -join ' '
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $psi
    [void]$p.Start()
    $stdout = $p.StandardOutput.ReadToEnd()
    $stderr = $p.StandardError.ReadToEnd()
    $p.WaitForExit()

    # 7-Zip exit codes: 0=ok, 1=warnings (still usually OK), >=2 error
    if ($p.ExitCode -ge 2) {
        throw "7-Zip failed (exit $($p.ExitCode)). Args: $($psi.Arguments)`nSTDOUT:`n$stdout`nSTDERR:`n$stderr"
    }

    # Return logs if you want to display or log later
    return @{ ExitCode=$p.ExitCode; StdOut=$stdout; StdErr=$stderr }
}

function Invoke-7ZipCapture {
    param(
        [string]$SevenZip,
        [string[]]$Args
    )
    if (-not $SevenZip) { throw "7-Zip not found." }
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $SevenZip
    $quoted = $Args | ForEach-Object {
        if ($_ -match '[\s"]') {
            '"' + ($_ -replace '"', '\\"') + '"'
        } else { $_ }
    }
    $psi.Arguments = ($quoted -join ' ')
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $psi
    [void]$p.Start()
    $stdout = $p.StandardOutput.ReadToEnd()
    $stderr = $p.StandardError.ReadToEnd()
    $p.WaitForExit()
    return @{ ExitCode=$p.ExitCode; StdOut=$stdout; StdErr=$stderr }
}

function Expand-ZipWithDotNet {
    param(
        [Parameter(Mandatory)][string]$ZipPath,
        [Parameter(Mandatory)][string]$Destination
    )
    if (-not (Test-Path $Destination)) {
        $null = New-Item -ItemType Directory -Path $Destination -Force
    }
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($ZipPath)
    try {
        foreach ($entry in $zip.Entries) {
            $target = Join-Path $Destination $entry.FullName
            $dir = Split-Path $target -Parent
            if ($dir -and -not (Test-Path $dir)) { $null = New-Item -ItemType Directory -Path $dir -Force }
            if ($entry.FullName.EndsWith('/')) { continue } # directory
            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $target, $true)
        }
    } finally {
        $zip.Dispose()
    }
}

function Normalize-AapkgPath {
    param(
        [Parameter(Mandatory)][string]$Path
    )
    $clean = $Path.Trim('"')
    if (-not (Test-Path -LiteralPath $clean)) { throw "File not found: $Path" }
    return (Resolve-Path -LiteralPath $clean).Path
}

function Wait-ForCabFile {
    param(
        [Parameter(Mandatory)][string]$Folder,
        [int]$TimeoutSeconds = 10
    )
    $existing = Get-ChildItem -Path $Folder -Recurse -Force -File -Filter *.cab -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($existing) { return $existing }

    $watcher = New-Object System.IO.FileSystemWatcher
    $watcher.Path = $Folder
    $watcher.Filter = "*.cab"
    $watcher.IncludeSubdirectories = $true
    $watcher.EnableRaisingEvents = $true

    $signal = New-Object System.Threading.ManualResetEvent($false)
    $created = Register-ObjectEvent -InputObject $watcher -EventName Created -Action { $signal.Set() } | Select-Object -First 1
    $renamed = Register-ObjectEvent -InputObject $watcher -EventName Renamed -Action { $signal.Set() } | Select-Object -First 1

    try {
        $signal.WaitOne([TimeSpan]::FromSeconds($TimeoutSeconds)) | Out-Null
    }
    finally {
        if ($created) { Unregister-Event -SourceIdentifier $created.Name -ErrorAction SilentlyContinue }
        if ($renamed) { Unregister-Event -SourceIdentifier $renamed.Name -ErrorAction SilentlyContinue }
        $watcher.EnableRaisingEvents = $false
        $watcher.Dispose()
        $signal.Dispose()
    }

    return (Get-ChildItem -Path $Folder -Recurse -Force -File -Filter *.cab -ErrorAction SilentlyContinue | Select-Object -First 1)
}

function Get-CabListingInfo {
    param(
        [Parameter(Mandatory)][string]$SevenZip,
        [Parameter(Mandatory)][string]$CabPath
    )

    $listing = & $SevenZip l -slt $CabPath 2>&1
    if ($LASTEXITCODE -ne 0) { throw "7-Zip listing failed:`n$listing" }

    $archiveTypeLine = $listing | Where-Object { $_ -like 'Type = *' } | Select-Object -First 1
    $archiveType = if ($archiveTypeLine) { ($archiveTypeLine -replace '^\s*Type\s*=\s*','').Trim() } else { $null }

    $paths = foreach ($line in $listing) {
        if ($line -like 'Path = *') { $line.Substring(7) }
    }

    return @{ Type=$archiveType; Paths=@($paths); RawListing=$listing }
}

function Get-ManifestRelPath {
    param(
        [Parameter(Mandatory)][string[]]$Paths
    )

    $manifestRel =
        $Paths | Where-Object { $_ -match '(?i)(^|[\\/])manifest\.xml$' } | Select-Object -First 1
    if (-not $manifestRel) {
        $manifestRel =
            $Paths | Where-Object { $_ -match '(?i)(^|[\\/])AppManifest\.xml$' } | Select-Object -First 1
    }
    if (-not $manifestRel) {
        $manifestRel =
            $Paths | Where-Object { $_ -match '(?i)(^|[\\/])manifest.*\.xml$' } | Select-Object -First 1
    }

    return $manifestRel
}

function Extract-ManifestFromCab {
    param(
        [Parameter(Mandatory)][string]$SevenZip,
        [Parameter(Mandatory)][string]$CabPath,
        [Parameter(Mandatory)][string]$ManifestOut
    )
    if (-not (Test-Path $ManifestOut)) { New-Item -ItemType Directory -Path $ManifestOut | Out-Null }

    $listingInfo = Get-CabListingInfo -SevenZip $SevenZip -CabPath $CabPath
    $manifestRel = Get-ManifestRelPath -Paths $listingInfo.Paths
    if (-not $manifestRel) { throw "No 'manifest.xml' (or 'AppManifest.xml') found inside inner archive." }

    $leafName = Split-Path $manifestRel -Leaf
    $sevenZipArgs = @('e')
    if ($listingInfo.Type -match '(?i)^zip$') { $sevenZipArgs += '-tzip' }
    $sevenZipArgs += @(
        $CabPath,
        $leafName,
        "-o$ManifestOut",
        '-y'
    )

    $extractOut = & $SevenZip @sevenZipArgs 2>&1
    if ($LASTEXITCODE -ne 0) { throw "7-Zip extract failed:`n$extractOut" }

    $manifestPath = Join-Path $ManifestOut $leafName
    if (-not (Test-Path -LiteralPath $manifestPath)) {
        $fallback = Get-ChildItem -Path $ManifestOut -Recurse -Force -File -Filter $leafName | Select-Object -First 1
        if ($fallback) { $manifestPath = $fallback.FullName }
    }
    if (-not (Test-Path -LiteralPath $manifestPath)) { throw "After extraction, '$leafName' not found in '$ManifestOut'." }

    return $manifestPath
}

function Wait-ForFileUnlock {
    param(
        [Parameter(Mandatory)][string]$Path,
        [int]$TimeoutSeconds = 10
    )
    if (-not (Test-Path -LiteralPath $Path)) { return $true }
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        try {
            $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
            $fs.Dispose()
            return $true
        } catch {
            Start-Sleep -Milliseconds 200
        }
    }
    return $false
}

function Resolve-ManifestFromAapkg {
    param(
        [string]$AapkgPath
    )
    $sevenZip = Find-7Zip
    if (-not $sevenZip) { throw "7-Zip is required to open .aapkg files." }

    $resolvedAapkg = Normalize-AapkgPath -Path $AapkgPath

    $tempRoot = Join-Path $env:TEMP ("aapkg_extract_{0}" -f ([DateTime]::Now.ToString("yyyyMMdd_HHmmss")))
    if (-not (Test-Path $tempRoot)) { New-Item -ItemType Directory -Path $tempRoot | Out-Null }

    # Extract the entire AAPKG into tempRoot using .NET (same as simpleTest.ps1)
    Expand-ZipWithDotNet -ZipPath $resolvedAapkg -Destination $tempRoot

    $cabPath = Get-ChildItem -Path $tempRoot -Recurse -Force -File -Filter *.cab | Select-Object -First 1
    if (-not $cabPath) { $cabPath = Wait-ForCabFile -Folder $tempRoot -TimeoutSeconds 10 }
    if (-not $cabPath) { throw "No CAB found inside AAPKG." }

    $manifestOut = Join-Path $tempRoot "manifest"
    $manifestPath = Extract-ManifestFromCab -SevenZip $sevenZip -CabPath $cabPath.FullName -ManifestOut $manifestOut

    return @{ ManifestPath=$manifestPath; CabPath=$cabPath.FullName; TempRoot=$tempRoot; SevenZip=$sevenZip }
}

function Update-Cab-With7Zip {
    param(
        [string]$SevenZip,
        [string]$CabPath,
        [string]$ManifestPath
    )
    if (-not (Test-Path -LiteralPath $CabPath)) { throw "CAB not found: $CabPath" }
    if (-not (Test-Path -LiteralPath $ManifestPath)) { throw "Manifest not found: $ManifestPath" }

    if (-not (Wait-ForFileUnlock -Path $CabPath -TimeoutSeconds 10)) {
        throw "CAB is locked by another process: $CabPath"
    }

    $listingInfo = Get-CabListingInfo -SevenZip $SevenZip -CabPath $CabPath
    $manifestRel = Get-ManifestRelPath -Paths $listingInfo.Paths
    if (-not $manifestRel) { $manifestRel = "Manifest.xml" }

    $detectedType = $listingInfo.Type

    if ($detectedType -match '(?i)^zip$') {
        $manifestRelNorm = $manifestRel -replace '[\\/]+','\\' | ForEach-Object { $_.TrimStart('\','/','.',' ') }
        $stageRoot = Join-Path $env:TEMP ("cab_stage_{0}" -f ([guid]::NewGuid().ToString("N")))

        if ($manifestRelNorm.Contains('\\')) {
            $dir = Split-Path $manifestRelNorm -Parent
            $full = Join-Path $stageRoot $dir
            New-Item -ItemType Directory -Force -Path $full | Out-Null
            $stagePath = Join-Path $full (Split-Path $manifestRelNorm -Leaf)
        } else {
            New-Item -ItemType Directory -Force -Path $stageRoot | Out-Null
            $stagePath = Join-Path $stageRoot $manifestRelNorm
        }

        Copy-Item -LiteralPath $ManifestPath -Destination $stagePath -Force

        $updateArgs = @('u','-tzip',$CabPath,$stagePath,'-y')
        $updateOut = & $SevenZip @updateArgs 2>&1
        if ($LASTEXITCODE -ne 0) { throw "7-Zip update failed:`n$updateOut" }

        try { Remove-Item -LiteralPath $stageRoot -Recurse -Force -ErrorAction SilentlyContinue } catch {}
    }
    else {
        $fullWork = Join-Path $env:TEMP ("cab_full_{0}" -f ([guid]::NewGuid().ToString("N")))
        $null = New-Item -ItemType Directory -Path $fullWork -Force

        $xArgs = @('x', $CabPath, "-o$fullWork", '-y')
        $xOut = & $SevenZip @xArgs 2>&1
        if ($LASTEXITCODE -ne 0) { throw "7-Zip full extract failed:`n$xOut" }

        $manifestRelNorm = $manifestRel -replace '[\\/]+','\\' | ForEach-Object { $_.TrimStart('\','/','.',' ') }
        $targetPath = Join-Path $fullWork $manifestRelNorm
        if (-not (Test-Path -LiteralPath (Split-Path $targetPath -Parent))) {
            New-Item -ItemType Directory -Force -Path (Split-Path $targetPath -Parent) | Out-Null
        }
        Copy-Item -LiteralPath $ManifestPath -Destination $targetPath -Force

        $backupPath = "$CabPath.bak"
        if (-not (Test-Path -LiteralPath $backupPath)) {
            Copy-Item -LiteralPath $CabPath -Destination $backupPath -Force
        }

        $ddfPath = Join-Path $fullWork "_cab_build.ddf"
        $destDir = [IO.Path]::GetDirectoryName($CabPath)
        $cabName = [IO.Path]::GetFileName($CabPath)

        $ddfHeader = @(
            '.OPTION EXPLICIT',
            '.Set Cabinet=on',
            '.Set Compress=on',
            '.Set CompressionType=LZX',
            '.Set CompressionMemory=21',
            ".Set CabinetNameTemplate=$cabName",
            ".Set DiskDirectoryTemplate=$destDir",
            ".Set SourceDir=$fullWork",
            '.Set DestinationDir=.'
        ) -join "`r`n"

        $relLines = Get-ChildItem -LiteralPath $fullWork -Recurse -File | ForEach-Object {
            $rel = $_.FullName.Substring($fullWork.Length).TrimStart('\','/')
            "`"$rel`" `"$rel`""
        }

        $ddfContent = $ddfHeader + "`r`n" + ($relLines -join "`r`n") + "`r`n"
        Set-Content -LiteralPath $ddfPath -Value $ddfContent -Encoding ASCII

        $makecabExe = Join-Path $env:SystemRoot 'System32\makecab.exe'
        $makecabOut = & $makecabExe '/f' "$ddfPath" 2>&1
        if ($LASTEXITCODE -ne 0) { throw "makecab.exe failed:`n$makecabOut" }

        try { Remove-Item -LiteralPath $fullWork -Recurse -Force -ErrorAction SilentlyContinue } catch {}
    }
}

function Get-NextVersionedPath {
    param(
        [Parameter(Mandatory)][string]$SourcePath
    )

    $baseDir = Split-Path -Parent $SourcePath
    $baseNameNoExt = [System.IO.Path]::GetFileNameWithoutExtension($SourcePath)
    $ext = [System.IO.Path]::GetExtension($SourcePath)

    $m = [regex]::Match($BaseNameNoExt, '(?i)_v(\d+)$')
    if ($m.Success) {
        $root = $BaseNameNoExt.Substring(0, $m.Index)
        $n = [int]$m.Groups[1].Value + 1
    } else {
        $root = $BaseNameNoExt
        $n = 1
    }

    do {
        $candidate = Join-Path $baseDir ("{0}_v{1}{2}" -f $root, $n, $ext)
        $n++
    } while (Test-Path -LiteralPath $candidate)

    return $candidate
}

function Repack-AapkgFromFolder {
    param(
        [string]$SevenZip,
        [string]$ExtractRoot,
        [string]$AapkgPath,
        [string[]]$ExcludeFolders = @()
    )
    $zipPath = [System.IO.Path]::ChangeExtension($AapkgPath, ".zip")
    if (Test-Path $zipPath) { Remove-Item -LiteralPath $zipPath -Force }

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $zip = [IO.Compression.ZipFile]::Open($zipPath, [IO.Compression.ZipArchiveMode]::Create)
    try {
        $root = (Resolve-Path -LiteralPath $ExtractRoot).Path
        $files = Get-ChildItem -Path $root -Recurse -Force -File
        foreach ($file in $files) {
            $relative = $file.FullName.Substring($root.Length).TrimStart('\','/')
            $segments = $relative -split '[\\/]'
            if ($ExcludeFolders -and ($segments | Where-Object { $ExcludeFolders -contains $_ })) { continue }

            $entry = $zip.CreateEntry($relative, [IO.Compression.CompressionLevel]::Optimal)
            $entryStream = $entry.Open()
            try {
                $fs = [System.IO.File]::OpenRead($file.FullName)
                try {
                    $fs.CopyTo($entryStream)
                } finally {
                    $fs.Dispose()
                }
            } finally {
                $entryStream.Dispose()
            }
        }
    } finally {
        $zip.Dispose()
    }

    if (-not (Test-Path $zipPath)) {
        throw "Failed to create $zipPath from $ExtractRoot"
    }

    $bakPath = "$AapkgPath.bak"
    if (Test-Path $bakPath) { Remove-Item -LiteralPath $bakPath -Force }
    if (Test-Path $AapkgPath) {
        if (-not (Wait-ForFileUnlock -Path $AapkgPath -TimeoutSeconds 10)) {
            throw "AAPKG is locked by another process: $AapkgPath"
        }
        Rename-Item -LiteralPath $AapkgPath -NewName $bakPath
    }
    Move-Item -LiteralPath $zipPath -Destination $AapkgPath
}

function Get-InstanceFolderSegments([System.Xml.XmlNode]$node) {
    if (-not $node) { return @() }
    $toolset = $null
    foreach ($n in "toolset_name","toolset","Toolset","ToolsetName") {
        if ($node.Attributes[$n]) { $toolset = $node.Attributes[$n].Value; break }
    }
    if (-not $toolset) { return @() }

    # Strip galaxy prefix if it is the first token before '$' and has no spaces.
    $firstDollar = $toolset.IndexOf('$')
    if ($firstDollar -gt 0) {
        $prefix = $toolset.Substring(0, $firstDollar)
        if ($prefix -notmatch "\s") {
            $toolset = $toolset.Substring($firstDollar + 1)
        }
    }

    $parts = $toolset -split "\\|/|\$" | Where-Object { $_ -and $_.Trim().Length -gt 0 }
    if (-not $parts -or $parts.Count -eq 0) { return @() }

    return $parts
}

function Ensure-FolderChild {
    param(
        [object]$ParentNode,
        [string]$FolderName
    )
    foreach ($c in $ParentNode.Children) {
        if ($c.Kind -eq $script:NodeKind_Folder -and $c.Name -eq $FolderName) { return $c }
    }
    $fn = New-Object $script:TreeNodeType.FullName
    $fn.Name = $FolderName
    $fn.Kind = $script:NodeKind_Folder
    $fn.SetIsChecked($true, $false, $false)
    $ParentNode.Attach($fn)
    return $fn
}

function Add-InstanceNodeToGroup {
    param(
        [object]$GroupNode,
        [object]$InstanceNode,
        [string[]]$FolderSegments
    )
    $cursor = $GroupNode
    if ($FolderSegments -and $FolderSegments.Count -gt 0) {
        foreach ($seg in $FolderSegments) {
            $cursor = Ensure-FolderChild -ParentNode $cursor -FolderName $seg
        }
    }
    $cursor.Attach($InstanceNode)
}

function Sort-ChildrenAlphabetically {
    param([object]$Node)
    if (-not $Node -or -not $Node.Children) { return }
    $sorted = @($Node.Children | Sort-Object -Property Name)
    $Node.Children.Clear()
    foreach ($c in $sorted) { $Node.Children.Add($c) }
    foreach ($c in $Node.Children) { Sort-ChildrenAlphabetically $c }
}

function Get-NodeIndexPath([System.Xml.XmlNode]$node) {
    if (-not $node -or -not $node.ParentNode) { return $null }
    if (-not $script:ManifestXml) { return $null }
    if ($node.OwnerDocument -ne $script:ManifestXml) { return $null }
    $segments = New-Object System.Collections.Generic.List[string]
    $current = $node
    while ($current.ParentNode -and $current.ParentNode.NodeType -ne [System.Xml.XmlNodeType]::Document) {
        $parent = $current.ParentNode
        $idx = 0
        for ($i = 0; $i -lt $parent.ChildNodes.Count; $i++) {
            if ($parent.ChildNodes[$i] -eq $current) { $idx = $i; break }
        }
        $segments.Insert(0, [string]$idx)
        $current = $parent
    }
    return ($segments -join "/")
}

function Resolve-NodeByIndexPath {
    param(
        [xml]$Doc,
        [string]$Path
    )
    if (-not $Doc -or -not $Doc.DocumentElement -or -not $Path) { return $null }
    $parts = $Path -split '/'
    $current = $Doc.DocumentElement
    foreach ($p in $parts) {
        if ($p -eq "") { continue }
        $idx = [int]$p
        if ($idx -ge $current.ChildNodes.Count) { return $null }
        $current = $current.ChildNodes[$idx]
    }
    return $current
}

function Try-Parse-CommentElement {
    param(
        [System.Xml.XmlComment]$Comment,
        [ref]$ElementOut
    )
    if (-not $Comment) { return $false }
    $val = $Comment.Value
    if (-not $val) { return $false }
    $val = $val.Trim()
    if (-not $val.StartsWith("<")) { return $false }
    try {
        [xml]$tmp = "<root>$val</root>"
        $el = $tmp.DocumentElement.ChildNodes | Where-Object { $_.NodeType -eq [System.Xml.XmlNodeType]::Element } | Select-Object -First 1
        if ($el) {
            $ElementOut.Value = $el
            return $true
        }
    } catch {}
    return $false
}

function New-ElementFromXml {
    param(
        [xml]$Doc,
        [string]$XmlText
    )
    if (-not $Doc -or -not $XmlText) { return $null }
    [xml]$tmp = $XmlText
    return $Doc.ImportNode($tmp.DocumentElement, $true)
}

function Find-MatchingTemplateNode {
    param(
        [System.Xml.XmlNode]$ParentElement,
        [object]$TemplateNode
    )
    if (-not $ParentElement -or -not $TemplateNode) { return $null }

    $targetName = $TemplateNode.Name
    foreach ($child in $ParentElement.ChildNodes) {
        if ($child.NodeType -eq [System.Xml.XmlNodeType]::Element -and $child.Name -eq "template") {
            if ((Get-NodeLabel $child) -eq $targetName) { return $child }
        } elseif ($child.NodeType -eq [System.Xml.XmlNodeType]::Comment) {
            $el = $null
            if (Try-Parse-CommentElement -Comment $child -ElementOut ([ref]$el)) {
                if ($el.Name -eq "template" -and (Get-NodeLabel $el) -eq $targetName) {
                    return $child
                }
            }
        }
    }
    return $null
}

function Find-MatchingInstanceNode {
    param(
        [System.Xml.XmlNode]$ContainerElement,
        [object]$InstNode
    )
    if (-not $ContainerElement -or -not $InstNode) { return $null }

    $targetName = $InstNode.Name
    foreach ($child in $ContainerElement.ChildNodes) {
        if ($child.NodeType -eq [System.Xml.XmlNodeType]::Element -and $child.Name -eq "instance") {
            if ((Get-NodeLabel $child) -eq $targetName) { return $child }
        } elseif ($child.NodeType -eq [System.Xml.XmlNodeType]::Comment) {
            $el = $null
            if (Try-Parse-CommentElement -Comment $child -ElementOut ([ref]$el)) {
                if ($el.Name -eq "instance" -and (Get-NodeLabel $el) -eq $targetName) {
                    return $child
                }
            }
        }
    }
    return $null
}

function Find-MatchingDerivedInstancesNode {
    param(
        [System.Xml.XmlNode]$TemplateElement,
        [object]$GroupNode
    )
    if (-not $TemplateElement -or -not $GroupNode) { return $null }

    $elementCount = @($TemplateElement.SelectNodes('./derived_instances')).Count
    if ($elementCount -eq 1) {
        return $TemplateElement.SelectSingleNode('./derived_instances')
    }

    $sourceXml = $GroupNode.SourceXml
    if ($sourceXml) {
        foreach ($comment in $TemplateElement.SelectNodes('./comment()')) {
            $el = $null
            if (Try-Parse-CommentElement -Comment $comment -ElementOut ([ref]$el)) {
                if ($el.Name -eq 'derived_instances' -and $el.OuterXml -eq $sourceXml) {
                    return $comment
                }
            }
        }
    }

    if ($elementCount -gt 1) {
        foreach ($el in $TemplateElement.SelectNodes('./derived_instances')) {
            if ($sourceXml -and $el.OuterXml -eq $sourceXml) { return $el }
        }
    }

    return $null
}

function Replace-Node-With-Comment {
    param(
        [xml]$Doc,
        [System.Xml.XmlNode]$Node,
        [string]$XmlText
    )
    if (-not $Doc -or -not $Node -or -not $Node.ParentNode) { return $null }
    if (-not $XmlText) { return $null }
    if ($XmlText -like "*--*") { throw "Cannot create XML comment because content contains '--'." }
    $comment = $Doc.CreateComment($XmlText)
    [void]$Node.ParentNode.ReplaceChild($comment, $Node)
    return $comment
}

function Replace-Comment-With-Element {
    param(
        [xml]$Doc,
        [System.Xml.XmlNode]$CommentNode,
        [string]$XmlText
    )
    if (-not $Doc -or -not $CommentNode -or -not $CommentNode.ParentNode) { return $null }
    if (-not $XmlText) { return $null }
    $el = New-ElementFromXml -Doc $Doc -XmlText $XmlText
    if (-not $el) { return $null }
    [void]$CommentNode.ParentNode.ReplaceChild($el, $CommentNode)
    return $el
}

function Expand-ParseableXmlComments {
    param(
        [xml]$Doc
    )
    if (-not $Doc -or -not $Doc.DocumentElement) { return }

    # Repeat until no parseable XML comments remain.
    # Replacing comment -> element keeps sibling positions stable.
    do {
        $changed = $false
        $comments = @($Doc.SelectNodes('//comment()'))
        foreach ($comment in $comments) {
            if (-not $comment) { continue }
            $el = $null
            if (Try-Parse-CommentElement -Comment $comment -ElementOut ([ref]$el)) {
                $raw = $comment.Value
                if ($raw) {
                    $repl = Replace-Comment-With-Element -Doc $Doc -CommentNode $comment -XmlText $raw
                    if ($repl) { $changed = $true }
                }
            }
        }
    } while ($changed)
}

function Ensure-ContainerElement {
    param(
        [xml]$Doc,
        [System.Xml.XmlNode]$ParentElement,
        [string]$ContainerName,
        [string]$ContainerCommentPath,
        [string]$ContainerSourceXml
    )
    if (-not $ParentElement -or -not $ContainerName) { return $null }
    if ($ContainerCommentPath) {
        $commentNode = Resolve-NodeByIndexPath -Doc $Doc -Path $ContainerCommentPath
        if ($commentNode -and $commentNode.NodeType -eq [System.Xml.XmlNodeType]::Comment) {
            $repl = Replace-Comment-With-Element -Doc $Doc -CommentNode $commentNode -XmlText $ContainerSourceXml
            if ($repl) { return $repl }
        }
    }
    $existing = $ParentElement.SelectSingleNode("./$ContainerName")
    if ($existing) { return $existing }
    $new = $Doc.CreateElement($ContainerName)
    [void]$ParentElement.AppendChild($new)
    return $new
}

function Build-TemplateNode {
    param(
        [System.Xml.XmlNode]$templateNode,
        [bool]$isCommented = $false,
        [string]$sourceXml = $null,
        [string]$indexPath = $null,
        [string]$containerCommentPath = $null,
        [string]$containerSourceXml = $null
    )
    $tn         = New-Object $script:TreeNodeType.FullName
    $tn.Name    = Get-NodeLabel $templateNode
    $tn.Kind    = $script:NodeKind_Template
    $tn.XmlRef  = $templateNode
    $tn.IsCommented = $isCommented
    $tn.SourceXml = $sourceXml
    $tn.IndexPath = $indexPath
    $tn.ContainerCommentPath = $containerCommentPath
    $tn.ContainerSourceXml = $containerSourceXml
    # default selection
    if ($isCommented) {
        $tn.SetIsChecked($false, $false, $false)
    } else {
        $tn.SetIsChecked($true, $false, $false)
    }

    # Instances group (if any)
    $groupNodes = @()
    foreach ($di in $templateNode.SelectNodes("./derived_instances")) {
        $grp        = New-Object $script:TreeNodeType.FullName
        $grp.Name   = "Instances"
        $grp.Kind   = $script:NodeKind_InstanceGroup
        $grp.IsCommented = $false
        $grp.SourceXml = $di.OuterXml
        $grp.IndexPath = Get-NodeIndexPath $di
        $grp.SetIsChecked($true, $false, $false)

        foreach ($child in $di.ChildNodes) {
            if ($child.NodeType -eq [System.Xml.XmlNodeType]::Element -and $child.Name -eq "instance") {
                $in         = New-Object $script:TreeNodeType.FullName
                $in.Name    = Get-NodeLabel $child
                $in.Kind    = $script:NodeKind_Instance
                $in.XmlRef  = $child
                $in.IsCommented = $false
                $in.SourceXml = $child.OuterXml
                $in.IndexPath = Get-NodeIndexPath $child
                $in.SetIsChecked($true, $false, $false)
                $segments = Get-InstanceFolderSegments $child
                Add-InstanceNodeToGroup -GroupNode $grp -InstanceNode $in -FolderSegments $segments
            } elseif ($child.NodeType -eq [System.Xml.XmlNodeType]::Comment) {
                $el = $null
                if (Try-Parse-CommentElement -Comment $child -ElementOut ([ref]$el)) {
                    if ($el.Name -eq "instance") {
                        $in         = New-Object $script:TreeNodeType.FullName
                        $in.Name    = Get-NodeLabel $el
                        $in.Kind    = $script:NodeKind_Instance
                        $in.IsCommented = $true
                        $in.SourceXml = $el.OuterXml
                        $in.IndexPath = Get-NodeIndexPath $child
                        $in.SetIsChecked($false, $false, $false)
                        $segments = Get-InstanceFolderSegments $el
                        Add-InstanceNodeToGroup -GroupNode $grp -InstanceNode $in -FolderSegments $segments
                    }
                }
            }
        }
        if ($grp.Children.Count -gt 0) { $groupNodes += $grp }
    }

    foreach ($comment in $templateNode.SelectNodes("./comment()")) {
        $el = $null
        if (Try-Parse-CommentElement -Comment $comment -ElementOut ([ref]$el)) {
            if ($el.Name -eq "derived_instances") {
                $grp        = New-Object $script:TreeNodeType.FullName
                $grp.Name   = "Instances"
                $grp.Kind   = $script:NodeKind_InstanceGroup
                $grp.IsCommented = $true
                $grp.SourceXml = $el.OuterXml
                $grp.IndexPath = Get-NodeIndexPath $comment
                $grp.SetIsChecked($false, $false, $false)

                foreach ($inst in $el.SelectNodes("./instance")) {
                    $in         = New-Object $script:TreeNodeType.FullName
                    $in.Name    = Get-NodeLabel $inst
                    $in.Kind    = $script:NodeKind_Instance
                    $in.IsCommented = $true
                    $in.SourceXml = $inst.OuterXml
                    $in.SetIsChecked($false, $false, $false)
                    $segments = Get-InstanceFolderSegments $inst
                    Add-InstanceNodeToGroup -GroupNode $grp -InstanceNode $in -FolderSegments $segments
                }
                if ($grp.Children.Count -gt 0) { $groupNodes += $grp }
            }
        }
    }

    foreach ($g in $groupNodes) { 
        Sort-ChildrenAlphabetically $g
        $tn.Attach($g) 
    }

    # Derived templates
    $derivedList = New-Object System.Collections.Generic.List[object]
    foreach ($child in $templateNode.SelectNodes("./derived_templates/template")) {
        $derivedList.Add(@{ Element=$child; IsCommented=$false; IndexPath=(Get-NodeIndexPath $child); SourceXml=$child.OuterXml })
    }
    foreach ($comment in $templateNode.SelectNodes("./derived_templates/comment()")) {
        $el = $null
        if (Try-Parse-CommentElement -Comment $comment -ElementOut ([ref]$el)) {
            if ($el.Name -eq "template") {
                $derivedList.Add(@{ Element=$el; IsCommented=$true; IndexPath=(Get-NodeIndexPath $comment); SourceXml=$el.OuterXml })
            }
        }
    }
    foreach ($comment in $templateNode.SelectNodes("./comment()")) {
        $el = $null
        if (Try-Parse-CommentElement -Comment $comment -ElementOut ([ref]$el)) {
            if ($el.Name -eq "template") {
                $derivedList.Add(@{ Element=$el; IsCommented=$true; IndexPath=(Get-NodeIndexPath $comment); SourceXml=$el.OuterXml })
            } elseif ($el.Name -eq "derived_templates") {
                foreach ($t in $el.SelectNodes("./template")) {
                    $derivedList.Add(@{ Element=$t; IsCommented=$true; IndexPath=$null; SourceXml=$t.OuterXml; ContainerCommentPath=(Get-NodeIndexPath $comment); ContainerSourceXml=$el.OuterXml })
                }
            }
        }
    }
    $derivedList = @($derivedList | Sort-Object { Get-NodeLabel $_.Element })
    foreach ($d in $derivedList) {
        $tn.Attach( (Build-TemplateNode -templateNode $d.Element -isCommented:$d.IsCommented -sourceXml $d.SourceXml -indexPath $d.IndexPath -containerCommentPath $d.ContainerCommentPath -containerSourceXml $d.ContainerSourceXml) )
    }

    # If this template itself is commented out in the source manifest, force the
    # entire subtree to unchecked so load state matches file state.
    if ($isCommented) {
        $tn.SetIsChecked($false, $true, $false)
    }

    Sort-ChildrenAlphabetically $tn
    return $tn
}

function Build-TreeFromManifest {
    $script:RootNodes.Clear()
    if (-not $script:ManifestXml -or -not $script:ManifestXml.DocumentElement) { return }

    # Start fresh: detach all children of hidden root (if any)
    # (We can't Clear() Children from PowerShell directly, so rebuild HiddenRoot)
    $script:HiddenRoot = New-Object $script:TreeNodeType.FullName
    $script:HiddenRoot.Name = "(root)"
    $script:HiddenRoot.Kind = $script:NodeKind_Template
    $script:HiddenRoot.SetIsChecked($true, $false, $false)

    $topNodes = New-Object System.Collections.Generic.List[object]
    foreach ($node in $script:ManifestXml.DocumentElement.ChildNodes) {
        if ($node.NodeType -eq [System.Xml.XmlNodeType]::Element -and $node.Name -eq "template") {
            $tNode = Build-TemplateNode -templateNode $node -isCommented:$false -sourceXml $node.OuterXml -indexPath (Get-NodeIndexPath $node)
            $topNodes.Add($tNode)
        } elseif ($node.NodeType -eq [System.Xml.XmlNodeType]::Comment) {
            $el = $null
            if (Try-Parse-CommentElement -Comment $node -ElementOut ([ref]$el)) {
                if ($el.Name -eq "template") {
                    $tNode = Build-TemplateNode -templateNode $el -isCommented:$true -sourceXml $el.OuterXml -indexPath (Get-NodeIndexPath $node)
                    $topNodes.Add($tNode)
                }
            }
        }
    }
    $topNodes = @($topNodes | Sort-Object -Property Name)
    foreach ($tNode in $topNodes) {
        $script:HiddenRoot.Attach($tNode)
        $script:RootNodes.Add($tNode)
    }
    Sort-ChildrenAlphabetically $script:HiddenRoot
    Normalize-All $script:HiddenRoot
}

function Normalize-All {
    param([object]$node)
    foreach ($c in $node.Children) { Normalize-All $c }
    # Recompute this node from its direct children
    $node.RecomputeFromChildren()
}

function Update-TotalObjectCount {
    param(
        [xml]$Doc
    )
    if (-not $Doc -or -not $Doc.DocumentElement) { return 0 }

    $templateCount = @($Doc.SelectNodes("//*[local-name()='template']")).Count
    $instanceCount = @($Doc.SelectNodes("//*[local-name()='instance']")).Count
    $total = $templateCount + $instanceCount

    $root = $Doc.DocumentElement
    $existing = @($root.SelectNodes("./*[local-name()='TotalObjectCount']"))

    $node = $null
    if ($existing.Count -gt 0) {
        $node = $existing[0]
        if ($existing.Count -gt 1) {
            for ($i = 1; $i -lt $existing.Count; $i++) {
                if ($existing[$i] -and $existing[$i].ParentNode) {
                    [void]$existing[$i].ParentNode.RemoveChild($existing[$i])
                }
            }
        }
    } else {
        if ($root.NamespaceURI) {
            $node = $Doc.CreateElement("TotalObjectCount", $root.NamespaceURI)
        } else {
            $node = $Doc.CreateElement("TotalObjectCount")
        }
    }

    $node.SetAttribute("objectcount", [string]$total)

    # Keep it as a direct child at the end of document element for consistency.
    if ($node.ParentNode) {
        [void]$node.ParentNode.RemoveChild($node)
    }
    [void]$root.AppendChild($node)

    return $total
}

function Get-FileNamesFromObjectElement {
    param(
        [System.Xml.XmlNode]$Element
    )
    $list = New-Object System.Collections.Generic.List[string]
    if (-not $Element) { return $list }

    $nodes = @($Element.SelectNodes("descendant-or-self::*[(local-name()='template' or local-name()='instance') and @file_name]"))
    foreach ($n in $nodes) {
        if ($n.Attributes -and $n.Attributes['file_name']) {
            $fn = $n.Attributes['file_name'].Value
            if ($fn -and $fn.Trim().Length -gt 0) { [void]$list.Add($fn.Trim()) }
        }
    }

    return $list
}

function Remove-CommentedUncheckedNodesAndCollectFiles {
    param(
        [xml]$Doc
    )
    $files = New-Object System.Collections.Generic.List[string]
    if (-not $Doc -or -not $Doc.DocumentElement) { return $files }

    $comments = @($Doc.SelectNodes('//comment()'))
    foreach ($comment in $comments) {
        if (-not $comment -or -not $comment.ParentNode) { continue }

        $el = $null
        if (-not (Try-Parse-CommentElement -Comment $comment -ElementOut ([ref]$el))) { continue }
        if (-not $el) { continue }

        $name = $el.Name
        if ($name -notin @('template','instance','derived_instances','derived_templates')) { continue }

        foreach ($fn in (Get-FileNamesFromObjectElement -Element $el)) {
            [void]$files.Add($fn)
        }

        [void]$comment.ParentNode.RemoveChild($comment)
    }

    return $files
}

function Update-Cab-WithManifestAndDeletes {
    param(
        [string]$SevenZip,
        [string]$CabPath,
        [string]$ManifestPath,
        [string[]]$FileNames
    )
    if (-not (Test-Path -LiteralPath $CabPath)) { throw "CAB not found: $CabPath" }
    if (-not (Test-Path -LiteralPath $ManifestPath)) { throw "Manifest not found: $ManifestPath" }

    $listingInfo = Get-CabListingInfo -SevenZip $SevenZip -CabPath $CabPath
    if ($listingInfo.Type -notmatch '(?i)^zip$') {
        throw "Hard-delete mode currently supports ZIP-type CAB archives only."
    }

    $manifestRel = Get-ManifestRelPath -Paths $listingInfo.Paths
    if (-not $manifestRel) { $manifestRel = "Manifest.xml" }
    $manifestRelNorm = ($manifestRel -replace '/', '\\').TrimStart('\','/')

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    if (-not (Wait-ForFileUnlock -Path $CabPath -TimeoutSeconds 10)) {
        throw "CAB is locked by another process: $CabPath"
    }

    $requestedFull = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)
    $requestedLeaf = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($raw in ($FileNames | Select-Object -Unique)) {
        if (-not $raw) { continue }
        $name = $raw.Trim().Trim('"')
        if (-not $name) { continue }
        $norm = ($name -replace '/', '\\').TrimStart('\','/')
        if ($norm) { [void]$requestedFull.Add($norm) }
        $leaf = [System.IO.Path]::GetFileName($norm)
        if ($leaf) { [void]$requestedLeaf.Add($leaf) }
    }

    $manifestEntryPath = ($manifestRelNorm -replace '\\','/')
    $removedCount = 0

    $zip = [IO.Compression.ZipFile]::Open($CabPath, [IO.Compression.ZipArchiveMode]::Update)
    try {
        $entries = @($zip.Entries)
        foreach ($entry in $entries) {
            if (-not $entry -or [string]::IsNullOrWhiteSpace($entry.FullName)) { continue }

            $entryNorm = ($entry.FullName -replace '/', '\\').TrimStart('\','/')
            $entryLeaf = [System.IO.Path]::GetFileName($entryNorm)
            $isManifest = ($entryNorm -ieq $manifestRelNorm)

            if (-not $isManifest -and ($requestedFull.Contains($entryNorm) -or ($entryLeaf -and $requestedLeaf.Contains($entryLeaf)))) {
                $entry.Delete()
                $removedCount++
            }
        }

        $existingManifest = $zip.Entries | Where-Object {
            (($_.FullName -replace '/', '\\').TrimStart('\','/')) -ieq $manifestRelNorm
        } | Select-Object -First 1
        if ($existingManifest) { $existingManifest.Delete() }

        $newManifest = $zip.CreateEntry($manifestEntryPath, [IO.Compression.CompressionLevel]::Optimal)
        $entryStream = $newManifest.Open()
        try {
            $fs = [System.IO.File]::OpenRead($ManifestPath)
            try { $fs.CopyTo($entryStream) } finally { $fs.Dispose() }
        }
        finally {
            $entryStream.Dispose()
        }
    }
    finally {
        $zip.Dispose()
    }

    return $removedCount
}

function Apply-InstanceNode {
    param(
        [object]$InstNode,
        [System.Xml.XmlNode]$ContainerElement,
        [xml]$WorkingDoc
    )
    if (-not $InstNode -or -not $ContainerElement) { return }
    $target = $null
    if ($InstNode.IndexPath) { $target = Resolve-NodeByIndexPath -Doc $WorkingDoc -Path $InstNode.IndexPath }
    if (-not $target) { $target = Find-MatchingInstanceNode -ContainerElement $ContainerElement -InstNode $InstNode }

    if ($InstNode.IsChecked -eq $false) {
        if ($target -and $target.NodeType -eq [System.Xml.XmlNodeType]::Element) {
            Replace-Node-With-Comment -Doc $WorkingDoc -Node $target -XmlText ($target.OuterXml) | Out-Null
        } elseif (-not $target -and $InstNode.SourceXml) {
            if ($InstNode.SourceXml -like "*--*") { throw "Cannot create XML comment because content contains '--'." }
            [void]$ContainerElement.AppendChild($WorkingDoc.CreateComment($InstNode.SourceXml))
        }
        return
    }

    if ($target -and $target.NodeType -eq [System.Xml.XmlNodeType]::Comment) {
        $target = Replace-Comment-With-Element -Doc $WorkingDoc -CommentNode $target -XmlText $InstNode.SourceXml
    } elseif (-not $target -and $InstNode.SourceXml) {
        $el = New-ElementFromXml -Doc $WorkingDoc -XmlText $InstNode.SourceXml
        if ($el) { [void]$ContainerElement.AppendChild($el) }
    }
}

function Apply-InstanceGroup {
    param(
        [object]$GroupNode,
        [System.Xml.XmlNode]$TemplateElement,
        [xml]$WorkingDoc
    )
    if (-not $GroupNode -or -not $TemplateElement) { return }
    $container = $null
    if ($GroupNode.IndexPath) { $container = Resolve-NodeByIndexPath -Doc $WorkingDoc -Path $GroupNode.IndexPath }
    if (-not $container) { $container = Find-MatchingDerivedInstancesNode -TemplateElement $TemplateElement -GroupNode $GroupNode }

    if ($GroupNode.IsChecked -eq $false) {
        if ($container -and $container.NodeType -eq [System.Xml.XmlNodeType]::Element) {
            Replace-Node-With-Comment -Doc $WorkingDoc -Node $container -XmlText ($container.OuterXml) | Out-Null
        } elseif (-not $container -and $GroupNode.SourceXml) {
            if ($GroupNode.SourceXml -like "*--*") { throw "Cannot create XML comment because content contains '--'." }
            [void]$TemplateElement.AppendChild($WorkingDoc.CreateComment($GroupNode.SourceXml))
        }
        return
    }

    if ($container -and $container.NodeType -eq [System.Xml.XmlNodeType]::Comment) {
        $container = Replace-Comment-With-Element -Doc $WorkingDoc -CommentNode $container -XmlText $GroupNode.SourceXml
    } elseif (-not $container) {
        $container = Ensure-ContainerElement -Doc $WorkingDoc -ParentElement $TemplateElement -ContainerName "derived_instances"
    }

    $stack = New-Object System.Collections.Generic.Stack[object]
    foreach ($c in $GroupNode.Children) { $stack.Push($c) }
    while ($stack.Count -gt 0) {
        $n = $stack.Pop()
        if ($n.Kind -eq $script:NodeKind_Instance) {
            Apply-InstanceNode -InstNode $n -ContainerElement $container -WorkingDoc $WorkingDoc
        } else {
            foreach ($c in $n.Children) { $stack.Push($c) }
        }
    }
}

function Apply-TemplateNode {
    param(
        [object]$TemplateNode,
        [xml]$WorkingDoc,
        [System.Xml.XmlNode]$ParentElement
    )
    if (-not $TemplateNode -or -not $ParentElement) { return }

    $target = $null
    if ($TemplateNode.IndexPath) { $target = Resolve-NodeByIndexPath -Doc $WorkingDoc -Path $TemplateNode.IndexPath }
    if (-not $target) { $target = Find-MatchingTemplateNode -ParentElement $ParentElement -TemplateNode $TemplateNode }

    if ($TemplateNode.IsChecked -eq $false) {
        if ($target -and $target.NodeType -eq [System.Xml.XmlNodeType]::Element) {
            Replace-Node-With-Comment -Doc $WorkingDoc -Node $target -XmlText ($target.OuterXml) | Out-Null
        } elseif (-not $target -and $TemplateNode.SourceXml) {
            if ($TemplateNode.SourceXml -like "*--*") { throw "Cannot create XML comment because content contains '--'." }
            [void]$ParentElement.AppendChild($WorkingDoc.CreateComment($TemplateNode.SourceXml))
        }
        return
    }

    if ($target -and $target.NodeType -eq [System.Xml.XmlNodeType]::Comment) {
        $target = Replace-Comment-With-Element -Doc $WorkingDoc -CommentNode $target -XmlText $TemplateNode.SourceXml
    } elseif (-not $target) {
        $el = New-ElementFromXml -Doc $WorkingDoc -XmlText $TemplateNode.SourceXml
        if ($el) {
            [void]$ParentElement.AppendChild($el)
            $target = $el
        }
    }

    if (-not $target) { return }

    foreach ($grp in $TemplateNode.Children | Where-Object { $_.Kind -eq $script:NodeKind_InstanceGroup }) {
        Apply-InstanceGroup -GroupNode $grp -TemplateElement $target -WorkingDoc $WorkingDoc
    }

    $childTemplates = @($TemplateNode.Children | Where-Object { $_.Kind -eq $script:NodeKind_Template })
    if ($childTemplates.Count -gt 0) {
        $containerName = "derived_templates"
        foreach ($child in $childTemplates) {
            $container = Ensure-ContainerElement -Doc $WorkingDoc -ParentElement $target -ContainerName $containerName -ContainerCommentPath $child.ContainerCommentPath -ContainerSourceXml $child.ContainerSourceXml
            Apply-TemplateNode -TemplateNode $child -WorkingDoc $WorkingDoc -ParentElement $container
        }
    }
}



# ------------------------------------------------------------
# Load / Save handlers
# ------------------------------------------------------------
$chkHardDeleteAapkg.Add_Checked({
    if ($chkOverwriteOriginal) {
        $chkOverwriteOriginal.IsChecked = $false
        $chkOverwriteOriginal.IsEnabled = $false
    }
})

$chkHardDeleteAapkg.Add_Unchecked({
    if ($chkOverwriteOriginal) {
        $chkOverwriteOriginal.IsEnabled = $true
    }
})

$btnLoad.Add_Click({
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = "XML or AAPKG (*.xml;*.aapkg)|*.xml;*.aapkg|All files (*.*)|*.*"
    $dlg.Title  = "Select Manifest.xml or .aapkg"
    if ($dlg.ShowDialog()) {
        try {
            $script:SourceAapkgPath = $null
            $script:SourceCabPath = $null
            $script:SourceManifestPath = $null
            $script:SourceXmlPath = $null
            $script:SourceTempRoot = $null
            $script:SevenZipPath = $null
            $chkHardDeleteAapkg.IsChecked = $false
            $chkHardDeleteAapkg.IsEnabled = $false
            $chkOverwriteOriginal.IsEnabled = $true

            $path = $dlg.FileName
            if ($path.ToLower().EndsWith(".aapkg")) {
                $resolved = Resolve-ManifestFromAapkg -AapkgPath $path
                $script:SourceTempRoot = $resolved.TempRoot
                $script:SourceAapkgPath = $path
                $script:SourceCabPath = $resolved.CabPath
                $script:SourceManifestPath = $resolved.ManifestPath
                $script:SevenZipPath = $resolved.SevenZip
                [xml]$script:ManifestXml = Get-Content -Raw -LiteralPath $resolved.ManifestPath
                $txt.Text = "Loaded from AAPKG: $path"
                $chkHardDeleteAapkg.IsEnabled = $true
                if ($txtOutput) { $txtOutput.Text = "" }
            } else {
                $script:SourceXmlPath = $path
                [xml]$script:ManifestXml = Get-Content -Raw -LiteralPath $path
                $txt.Text = "Loaded: $path"
                if ($txtOutput) { $txtOutput.Text = "" }
            }
            Build-TreeFromManifest
            $btnSave.IsEnabled = $true
        }
        catch {
            [System.Windows.MessageBox]::Show($_.Exception.Message, "Load failed", "OK", "Error") | Out-Null
        }
    }
})

$btnSave.Add_Click({
    if (-not $script:ManifestXml) { return }
    try {
        # 0) Fresh working copy
        [xml]$working = $script:ManifestXml.OuterXml

        # 1) Start from a normalized baseline: convert existing parseable XML comments
        #    back into elements, then apply current checkbox state by commenting
        #    unchecked entries.
        Expand-ParseableXmlComments -Doc $working
        foreach ($root in $script:RootNodes) {
            Apply-TemplateNode -TemplateNode $root -WorkingDoc $working -ParentElement $working.DocumentElement
        }

        $hardDeleteMode = ($script:SourceAapkgPath -and $chkHardDeleteAapkg -and $chkHardDeleteAapkg.IsChecked -eq $true)
        $removedFileCount = 0
        $filesToDelete = @()
        if ($hardDeleteMode) {
            # Convert newly-commented unchecked entries into permanent deletion.
            $filesToDelete = @(Remove-CommentedUncheckedNodesAndCollectFiles -Doc $working)
        }

        # 2) Recompute total selected object count (templates + instances)
        [void](Update-TotalObjectCount -Doc $working)

        $overwrite = ($chkOverwriteOriginal -and $chkOverwriteOriginal.IsChecked -eq $true)

        if ($script:SourceAapkgPath) {
            $targetAapkgPath = if ($overwrite) { $script:SourceAapkgPath } else { Get-NextVersionedPath -SourcePath $script:SourceAapkgPath }

            if (-not $script:SourceManifestPath) {
                throw "Source manifest path is not available. Reload the AAPKG and save again."
            }

            # Save XML to the extracted manifest path, then push it into CAB and repack AAPKG.
            $working.Save($script:SourceManifestPath)
            if ($hardDeleteMode) {
                $removedFileCount = Update-Cab-WithManifestAndDeletes -SevenZip $script:SevenZipPath -CabPath $script:SourceCabPath -ManifestPath $script:SourceManifestPath -FileNames $filesToDelete
            } else {
                Update-Cab-With7Zip -SevenZip $script:SevenZipPath -CabPath $script:SourceCabPath -ManifestPath $script:SourceManifestPath
            }

            if ($script:SourceTempRoot) {
                Repack-AapkgFromFolder -SevenZip $script:SevenZipPath -ExtractRoot $script:SourceTempRoot -AapkgPath $targetAapkgPath -ExcludeFolders @("manifest")
            } else {
                throw "AAPKG source folder is not available for repack. Reload the AAPKG and save again."
            }

            $script:SourceAapkgPath = $targetAapkgPath
            if ($txtOutput) {
                if ($hardDeleteMode) {
                    $txtOutput.Text = "Saved AAPKG successfully (hard delete mode). Output AAPKG: $targetAapkgPath  |  Files removed from package: $removedFileCount"
                } else {
                    $txtOutput.Text = "Saved AAPKG successfully. Output AAPKG: $targetAapkgPath"
                }
            }
        }
        else {
            if (-not $script:SourceXmlPath) {
                throw "Source XML path is not available. Reload the XML and save again."
            }

            $targetXmlPath = if ($overwrite) { $script:SourceXmlPath } else { Get-NextVersionedPath -SourcePath $script:SourceXmlPath }
            $working.Save($targetXmlPath)
            $script:SourceXmlPath = $targetXmlPath
            if ($txtOutput) { $txtOutput.Text = "Saved XML successfully. Output XML: $targetXmlPath" }
        }
    }
    catch {
        [System.Windows.MessageBox]::Show($_.Exception.Message, "Save failed", "OK", "Error") | Out-Null
    }
})

# ------------------------------------------------------------
# Show window
# ------------------------------------------------------------
$window.ShowDialog() | Out-Null
