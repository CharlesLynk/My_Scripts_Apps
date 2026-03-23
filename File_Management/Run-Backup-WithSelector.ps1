<#
.SYNOPSIS
    Interactive, selector-based backup script.

.DESCRIPTION
    Shows a WPF GUI so the user can choose backup locations per run.

    High-level behavior:
        - Jobs (SSD): incremental sync, stage backup-only items, clean empty directories.
        - Jobs (OneDrive): sync to "1 - Current Jobs", move missing top-level jobs to "2 - Old Jobs",
            and stage deep removals.
        - Default custom mappings include AppData, Downloads, and Pictures.
        - AppData defaults run as full-replace style mirrors (robocopy /MIR with no retries/waits).
        - Other Documents subfolders and custom mappings: mirrored backup.

    Timestamped snapshot folders are created under:
        <BackupRoot>\_Backup_YYYYMMDD

    All labels, destination conventions, selector items, and robocopy behavior are
    centralized in the configuration section near the top of this script.

.PARAMETER BackupRoot
  Root folder containing backup runs and the "To Be Deleted" folder.

.PARAMETER NoProgress
  Disables Write-Progress UI to reduce overhead (useful for Task Scheduler).

.NOTES
  PowerShell 5.1 compatible.
  Supports -WhatIf for dry-run.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter()]
    [string]$BackupRoot = 'F:\BACKUP_AND_SETTINGS',

    [Parameter()]
    [string]$BackupTarget = '',

    [Parameter()]
    [switch]$NoProgress,

    [Parameter()]
    [switch]$DryRunDiffCheck,

    [Parameter()]
    [switch]$RunBackupBeforeDiffCheck,

    [Parameter()]
    [ValidateRange(1, 12)]
    [int]$DryRunMaxDepth = 3
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ============================================================
# SCRIPT CONFIGURATION & CONSTANTS
# ============================================================
# Progress behavior: update less frequently to reduce overhead on huge trees
$script:ProgressEnabled = -not $NoProgress
$script:ProgressUpdateEvery = 250  # update progress every N items

# Progress IDs for Write-Progress hierarchy
$script:ProgressIds = @{
    Main                 = 1      # Top-level backup run progress
    JobsSync             = 10     # Jobs: scanning source
    JobsDirCreate        = 11     # Jobs: creating directories
    JobsFileCopy         = 12     # Jobs: copying files
    JobsDirCheck         = 13     # Jobs: checking removed directories
    JobsFileCheck        = 14     # Jobs: checking removed files
    JobsEmptyDirs        = 15     # Jobs: removing empty directories
    CustomBase           = 200    # Custom mappings: auto-assigned starting at this ID
}

# Date and time format strings
$script:DateFormat = 'yyyyMMdd'
$script:DateTimeFormat = 'yyyyMMdd_HHmmss'

# Dry diff timestamp tolerance (seconds).
# Robocopy runs with /FFT which allows 2-second timestamp granularity differences,
# so dry diff should honor the same tolerance to avoid false positives (especially on OneDrive).
$script:DryDiffTimestampToleranceSeconds = 2

# OneDrive backup root (used when BackupTarget = OneDrive)
$script:OneDriveBackupRoot = 'C:\Users\lynkcn\OneDrive - CDM Smith\Documents\BACKUP_AND_SETTINGS'

# Centralized path and label constants
$script:PathConfig = @{
    DocumentsRoot              = 'Documents'
    JobsRelative               = 'Documents\Jobs'
    JobsCurrentRelative        = 'Documents\Jobs\1 - Current Jobs'
    JobsOldRelative            = 'Documents\Jobs\2 - Old Jobs'
    AppDataBackupRelativeRoot  = '_AppData'
    DownloadsRelative          = 'Downloads'
    PicturesRelative           = 'Pictures'
    DeletionRootName           = 'To Be Deleted'
    DeletionRunPrefix          = 'Run_'
    JobsRemovedFromLiveLabelSSD             = 'Jobs_RemovedFromLive_SSD'
    JobsRemovedFromLiveLabelOneDriveCurrent = 'Jobs_RemovedFromLive_OneDriveCurrent'
}

# Centralized selector path tokens used throughout selection and processing logic
$script:SelectorPaths = @{
    Jobs            = 'DocumentsJobs'
    CustomPrefix    = 'Custom_'
    DocumentsPrefix = 'Documents_'
}

# Centralized main-flow display/status text
$script:MainText = @{
    ProgressActivityMain = 'Backup run'
    ProgressStart        = 'Starting...'
    ProgressFinalize     = 'Finalizing...'
    ProgressJobsSync     = 'Jobs: syncing...'
    ProgressJobsRecon    = 'Jobs: reconciling Current vs live ({0} + {1})...'
    ProgressJobsMoveOnly = 'Jobs: moving backup-only items...'
    ProgressJobsClean    = 'Jobs: cleaning empty dirs...'
    ProgressGenericCopy  = '{0} : copying...'

    HeaderLine           = '========================================'
    HeaderTitle          = 'Backup with Selective Items'
    HeaderSelectTarget   = 'Select backup target:'
    HeaderSelectorLaunch = 'Launching GUI selector...'
    SelectedItemsTitle   = 'Selected items for backup:'
    MessageCancelled     = 'Backup cancelled by user.'
    MessageNoSelection   = 'No items selected for backup.'
    MessageCompleted     = 'Backup run completed.'
    MessageReviewDeletes = 'Review staged removals here: {0}'
}

# Centralized operation progress text
$script:ProgressText = @{
    CompleteActivity            = 'Complete'
    EnumeratingStatus           = 'Enumerating...'
    ProcessingDirsStatus        = 'Processing directories...'
    EnumeratingFilesStatus      = 'Enumerating files...'
    ScanningStatus              = 'Scanning...'
    CountStatusFormat           = '{0}/{1}'

    JobsScanActivity              = 'Jobs: scanning source'
    JobsCreateDirsActivity        = 'Jobs: creating directories'
    JobsCopyFilesActivity         = 'Jobs: copying files'
    JobsOneDriveDeltaActivity     = 'Jobs: OneDrive delta sync'
    JobsOneDriveCompareStatus     = 'Comparing metadata (full depth, no hydration copy pass yet)...'
    JobsOneDriveSkipStatus        = 'Already in sync. Skipping copy pass.'
    JobsOneDriveCopyStatus        = 'Copying detected deltas...'
    JobsCheckRemovedDirsActivity  = 'Jobs: checking for removed directories'
    JobsCheckRemovedFilesActivity = 'Jobs: checking for removed files'
    JobsRemoveEmptyActivity       = 'Jobs: removing empty directories'

    AppDataReplaceActivity      = 'AppData: replacing {0}'
    BackupCopyActivity          = 'Backup: copying {0}'

    SyncSkipStatus              = 'Already in sync (timestamp-only). Skipping.'
    Stage1Status                = 'Stage 1/3: moving backup-only directories...'
    Stage1bStatus               = 'Stage 1b/3: staging backup-only files...'
    Stage2Status                = 'Stage 2/3: syncing with robocopy...'
    Stage3DoneStatus            = 'Stage 3/3: done.'
}

# Centralized robocopy argument templates by operation mode.
# Custom mapping option RobocopyMode can select one of these modes:
#   - MirrorStandard : /MIR with default retry behavior (best for normal resilient copies).
#   - MirrorNoRetry  : /MIR /R:0 /W:0 (fail fast; useful for volatile/locked trees like AppData).
#
# Internal-only modes:
#   - CopyNoMirror : used for non-delete copy scenarios.
#   - CompareList  : dry listing mode used by diff/compare routines.
$script:RobocopyModes = @{
    MirrorStandard = @('/MIR','/COPY:DAT','/DCOPY:DAT','/XJ','/FFT','/NP','/NFL','/NDL','/NJH','/NJS')
    MirrorNoRetry  = @('/MIR','/R:0','/W:0','/COPY:DAT','/DCOPY:DAT','/XJ','/FFT','/NP','/NFL','/NDL','/NJH','/NJS')
    CopyNoMirror   = @('/E','/R:0','/W:0','/COPY:DAT','/DCOPY:DAT','/XJ','/FFT','/NP','/NFL','/NDL','/NJH','/NJS')
    CompareList    = @('/E','/L','/R:0','/W:0','/COPY:DAT','/DCOPY:DAT','/XJ','/FFT','/NP','/NFL','/NDL','/NJH','/NJS')
}

# Deletion staging behavior
$script:MirrorDeletionStructure = $false
$script:DeletionMirrorBasePath = ''

# Backup folder naming pattern
$script:BackupFolderPattern = '_Backup_\d{8}'
$script:BackupFolderPrefix = '_Backup_'

# Centralized target roots (editable in one place)
$script:BackupTargets = @{
    SSD      = 'SSD'
    OneDrive = 'OneDrive'
}

# Canonical destination-base tokens
$script:DestinationBases = @{
    Root = 'BackupRoot'
    Run  = 'BackupRunRoot'
}

# Canonical target configuration map (single source of truth)
$script:BackupTargetConfig = @{
    SSD = @{
        RootPath                     = $BackupRoot
        DefaultCustomDestinationBase = $script:DestinationBases.Root
    }
    OneDrive = @{
        RootPath                     = $script:OneDriveBackupRoot
        DefaultCustomDestinationBase = $script:DestinationBases.Root
    }
}

# Centralized selector mapping and exclusions
$script:SelectorConfig = @{
    StaticItems = @()
    DocumentsExclusions = @('My Music', 'My Pictures', 'My Videos')
}

# Centralized ignore lists
$script:IgnoredNames = @{
    StageDirectoriesNotOnLive = @($script:PathConfig.DeletionRootName)
}

# Optional simple source overrides for default custom mappings.
# Keep these explicit/hardcoded if preferred.
$script:CustomSourceOverrides = @{
    Downloads = 'C:\Users\lynkcn\Downloads'
    Pictures  = 'C:\Users\lynkcn\OneDrive - CDM Smith\Pictures'
}

# Centralized custom mappings
# Add/edit entries here to back up folders, including the default AppData/Downloads/Pictures mappings.
#
# Mapping field reference (new schema):
#   Name         : Display label shown in selector and logs.
#   SelectorPath : Stable unique key (recommend prefix: 'Custom_').
#   IsEnabled    : If $false, entry is hidden and ignored.
#   IsChecked    : Default checkbox state when selector opens.
#   Targets      : Per-target config table (SSD and/or OneDrive).
#
# Target config fields:
#   SourceFolder             (required) : Live source path.
#   DestinationRelative      (required) : Relative destination path under selected base.
#   DestinationBase          (optional) :
#       - $script:DestinationBases.Root (default from target config)
#       - $script:DestinationBases.Run  (timestamped snapshot folder _Backup_YYYYMMDD)
#   RobocopyMode             (optional) : 'MirrorStandard' (default) or 'MirrorNoRetry'.
#       - MirrorStandard keeps robocopy retry behavior.
#       - MirrorNoRetry sets /R:0 /W:0 to fail fast (AppData-friendly).
#   ProgressActivityTemplate (optional) : format string used for activity text; receives Name as {0}.
#       - Default: $script:ProgressText.BackupCopyActivity
#       - Example: $script:ProgressText.AppDataReplaceActivity
#   IncludeTimestampMismatch (optional, default $true) :
#       - $true  => stage backup-only files that only differ by timestamp before mirroring.
#       - $false => only stage files missing from source.
#
# Notes:
#   - If a target key (SSD/OneDrive) is omitted, mapping is hidden in that mode.
#   - Invalid/unknown RobocopyMode values fall back to 'MirrorStandard'.
#
# Template A: Same source/destination logic for both SSD and OneDrive
# @{
#     Name         = 'Team Resources\CAD Standards'
#     SelectorPath = 'Custom_CADStandards'
#     IsEnabled    = $true
#     IsChecked    = $false
#     Targets      = @{
#         SSD = @{
#             SourceFolder        = 'D:\Team\CAD Standards'
#             DestinationRelative = 'Team Resources\CAD Standards'
#             # DestinationBase omitted -> uses target default from $script:BackupTargetConfig
#             # RobocopyMode omitted -> defaults to 'MirrorStandard'
#             # IncludeTimestampMismatch omitted -> defaults to $true
#         }
#         OneDrive = @{
#             SourceFolder        = 'D:\Team\CAD Standards'
#             DestinationRelative = 'Team Resources\CAD Standards'
#             # DestinationBase omitted -> uses target default from $script:BackupTargetConfig
#             # RobocopyMode omitted -> defaults to 'MirrorStandard'
#             # IncludeTimestampMismatch omitted -> defaults to $true
#         }
#     }
# }
#
# Template B: Only map for one target mode (this example: SSD only)
# @{
#     Name         = 'Local Scratch\Exports'
#     SelectorPath = 'Custom_Exports'
#     IsEnabled    = $true
#     IsChecked    = $false
#     Targets      = @{
#         SSD = @{
#             SourceFolder        = 'E:\Exports'
#             DestinationRelative = 'Scratch\Exports'
#             DestinationBase     = $script:DestinationBases.Run
#             RobocopyMode        = 'MirrorStandard'
#             # ProgressActivityTemplate = $script:ProgressText.BackupCopyActivity
#             # IncludeTimestampMismatch = $true
#         }
#         # OneDrive omitted intentionally -> this entry is hidden in OneDrive mode
#     }
# }
#
# Template C: AppData-style behavior (fail fast + custom activity text)
# @{
#     Name         = 'AppData\Roaming\VendorX'
#     SelectorPath = 'Custom_AppDataRoamingVendorX'
#     IsEnabled    = $true
#     IsChecked    = $true
#     Targets      = @{
#         SSD = @{
#             SourceFolder             = (Join-Path $env:APPDATA 'VendorX')
#             DestinationRelative      = "$($script:PathConfig.AppDataBackupRelativeRoot)\Roaming\VendorX"
#             DestinationBase          = $script:DestinationBases.Run
#             RobocopyMode             = 'MirrorNoRetry'
#             ProgressActivityTemplate = $script:ProgressText.AppDataReplaceActivity
#             IncludeTimestampMismatch = $true
#         }
#         OneDrive = @{
#             SourceFolder             = (Join-Path $env:APPDATA 'VendorX')
#             DestinationRelative      = "$($script:PathConfig.AppDataBackupRelativeRoot)\Roaming\VendorX"
#             DestinationBase          = $script:DestinationBases.Run
#             RobocopyMode             = 'MirrorNoRetry'
#             ProgressActivityTemplate = $script:ProgressText.AppDataReplaceActivity
#             IncludeTimestampMismatch = $true
#         }
#     }
# }

$script:CustomMappings = @(
    @{
        Name         = 'AppData\Local\Autodesk'
        SelectorPath = 'Custom_AppDataLocalAutodesk'
        IsEnabled    = $true
        IsChecked    = $true
        Targets      = @{
            SSD = @{
                SourceFolder             = (Join-Path $env:LOCALAPPDATA 'Autodesk')
                DestinationRelative      = "$($script:PathConfig.AppDataBackupRelativeRoot)\Local\Autodesk"
                DestinationBase          = $script:DestinationBases.Run
                RobocopyMode             = 'MirrorNoRetry'
                ProgressActivityTemplate = $script:ProgressText.AppDataReplaceActivity
                IncludeTimestampMismatch = $true
            }
            OneDrive = @{
                SourceFolder             = (Join-Path $env:LOCALAPPDATA 'Autodesk')
                DestinationRelative      = "$($script:PathConfig.AppDataBackupRelativeRoot)\Local\Autodesk"
                DestinationBase          = $script:DestinationBases.Run
                RobocopyMode             = 'MirrorNoRetry'
                ProgressActivityTemplate = $script:ProgressText.AppDataReplaceActivity
                IncludeTimestampMismatch = $true
            }
        }
    }
    @{
        Name         = 'AppData\Roaming\Autodesk'
        SelectorPath = 'Custom_AppDataRoamingAutodesk'
        IsEnabled    = $true
        IsChecked    = $true
        Targets      = @{
            SSD = @{
                SourceFolder             = (Join-Path $env:APPDATA 'Autodesk')
                DestinationRelative      = "$($script:PathConfig.AppDataBackupRelativeRoot)\Roaming\Autodesk"
                DestinationBase          = $script:DestinationBases.Run
                RobocopyMode             = 'MirrorNoRetry'
                ProgressActivityTemplate = $script:ProgressText.AppDataReplaceActivity
                IncludeTimestampMismatch = $true
            }
            OneDrive = @{
                SourceFolder             = (Join-Path $env:APPDATA 'Autodesk')
                DestinationRelative      = "$($script:PathConfig.AppDataBackupRelativeRoot)\Roaming\Autodesk"
                DestinationBase          = $script:DestinationBases.Run
                RobocopyMode             = 'MirrorNoRetry'
                ProgressActivityTemplate = $script:ProgressText.AppDataReplaceActivity
                IncludeTimestampMismatch = $true
            }
        }
    }
    @{
        Name         = 'AppData\Roaming\Logishrd'
        SelectorPath = 'Custom_AppDataRoamingLogishrd'
        IsEnabled    = $true
        IsChecked    = $true
        Targets      = @{
            SSD = @{
                SourceFolder             = (Join-Path $env:APPDATA 'Logishrd')
                DestinationRelative      = "$($script:PathConfig.AppDataBackupRelativeRoot)\Roaming\Logishrd"
                DestinationBase          = $script:DestinationBases.Run
                RobocopyMode             = 'MirrorNoRetry'
                ProgressActivityTemplate = $script:ProgressText.AppDataReplaceActivity
                IncludeTimestampMismatch = $true
            }
            OneDrive = @{
                SourceFolder             = (Join-Path $env:APPDATA 'Logishrd')
                DestinationRelative      = "$($script:PathConfig.AppDataBackupRelativeRoot)\Roaming\Logishrd"
                DestinationBase          = $script:DestinationBases.Run
                RobocopyMode             = 'MirrorNoRetry'
                ProgressActivityTemplate = $script:ProgressText.AppDataReplaceActivity
                IncludeTimestampMismatch = $true
            }
        }
    }
    @{
        Name         = 'AppData\Roaming\Notepad++'
        SelectorPath = 'Custom_AppDataRoamingNotepadPP'
        IsEnabled    = $true
        IsChecked    = $true
        Targets      = @{
            SSD = @{
                SourceFolder             = (Join-Path $env:APPDATA 'Notepad++')
                DestinationRelative      = "$($script:PathConfig.AppDataBackupRelativeRoot)\Roaming\Notepad++"
                DestinationBase          = $script:DestinationBases.Run
                RobocopyMode             = 'MirrorNoRetry'
                ProgressActivityTemplate = $script:ProgressText.AppDataReplaceActivity
                IncludeTimestampMismatch = $true
            }
            OneDrive = @{
                SourceFolder             = (Join-Path $env:APPDATA 'Notepad++')
                DestinationRelative      = "$($script:PathConfig.AppDataBackupRelativeRoot)\Roaming\Notepad++"
                DestinationBase          = $script:DestinationBases.Run
                RobocopyMode             = 'MirrorNoRetry'
                ProgressActivityTemplate = $script:ProgressText.AppDataReplaceActivity
                IncludeTimestampMismatch = $true
            }
        }
    }
    @{
        Name         = $script:PathConfig.DownloadsRelative
        SelectorPath = 'Custom_Downloads'
        IsEnabled    = $true
        IsChecked    = $true
        Targets      = @{
            SSD = @{
                SourceFolder        = [string]$script:CustomSourceOverrides.Downloads
                DestinationRelative = $script:PathConfig.DownloadsRelative
                DestinationBase     = $script:DestinationBases.Run
            }
            OneDrive = @{
                SourceFolder        = [string]$script:CustomSourceOverrides.Downloads
                DestinationRelative = $script:PathConfig.DownloadsRelative
                DestinationBase     = $script:DestinationBases.Run
            }
        }
    }
    @{
        Name         = $script:PathConfig.PicturesRelative
        SelectorPath = 'Custom_Pictures'
        IsEnabled    = $true
        IsChecked    = $true
        Targets      = @{
            SSD = @{
                SourceFolder        = [string]$script:CustomSourceOverrides.Pictures
                DestinationRelative = $script:PathConfig.PicturesRelative
                DestinationBase     = $script:DestinationBases.Run
            }
            OneDrive = @{
                SourceFolder        = [string]$script:CustomSourceOverrides.Pictures
                DestinationRelative = $script:PathConfig.PicturesRelative
                DestinationBase     = $script:DestinationBases.Run
            }
        }
    }
    @{
        Name         = 'Utilities'
        SelectorPath = 'Custom_Utilities'
        IsEnabled    = $true
        IsChecked    = $false
        Targets      = @{
            SSD = @{
                SourceFolder        = 'D:\Utilities'
                DestinationRelative = 'Personal Resources\Utilities'
            }
            OneDrive = @{
                SourceFolder        = 'D:\Utilities'
                DestinationRelative = 'Personal Resources\Utilities'
            }
        }
    }
)

# ============================================================
# WPF-BASED SELECTOR (adapted from ManifestGui.ps1)
# ============================================================
# Load assemblies first - this handles version resolution automatically
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xml, System.ObjectModel

# C# model for tri-state checkboxes
$csNamespace = "BackupSelectorModel_{0}" -f ([System.Guid]::NewGuid().ToString("N"))

$csharpCode = @"
using System;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Linq;
using System.Windows.Controls;

namespace $csNamespace
{
    public enum ItemKind { Category, Location }

    public class BackupItem : INotifyPropertyChanged
    {
        private bool? _isChecked = true;
        private ObservableCollection<BackupItem> _children;
        private bool _isExpanded = true;
        private BackupItem _parent;
        private bool _suppressChildRollup;

        public BackupItem()
        {
            _children = new ObservableCollection<BackupItem>();
            _children.CollectionChanged += Children_CollectionChanged;
        }

        public string Name { get; set; }
        public string DisplayPath { get; set; }
        public ItemKind Kind { get; set; }
        public string Path { get; set; }
        public bool IsEnabled { get; set; } = true;

        public BackupItem Parent
        {
            get { return _parent; }
            set { _parent = value; }
        }

        public ObservableCollection<BackupItem> Children
        {
            get { return _children; }
            set
            {
                if (ReferenceEquals(_children, value)) return;

                if (_children != null)
                {
                    _children.CollectionChanged -= Children_CollectionChanged;
                    foreach (var child in _children) UnhookChild(child);
                }

                _children = value ?? new ObservableCollection<BackupItem>();
                _children.CollectionChanged += Children_CollectionChanged;

                foreach (var child in _children) HookChild(child);
                RecalculateFromChildren();
            }
        }

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

        public bool? IsChecked
        {
            get { return _isChecked; }
            set
            {
                SetIsChecked(value, true, true);
            }
        }

        private void SetIsChecked(bool? value, bool propagateToChildren, bool rollupToParent)
        {
            if (_isChecked == value) return;

            _isChecked = value;
            OnPropertyChanged("IsChecked");

            if (propagateToChildren && value.HasValue && _children != null && _children.Count > 0)
            {
                _suppressChildRollup = true;
                foreach (var child in _children)
                {
                    child.SetIsChecked(value.Value, true, false);
                }
                _suppressChildRollup = false;
            }

            if (rollupToParent && _parent != null)
            {
                _parent.RecalculateFromChildren();
            }
        }

        private void Children_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (e.OldItems != null)
            {
                foreach (BackupItem oldChild in e.OldItems) UnhookChild(oldChild);
            }

            if (e.NewItems != null)
            {
                foreach (BackupItem newChild in e.NewItems) HookChild(newChild);
            }

            RecalculateFromChildren();
        }

        private void HookChild(BackupItem child)
        {
            if (child == null) return;
            child.Parent = this;
            child.PropertyChanged += Child_PropertyChanged;
        }

        private void UnhookChild(BackupItem child)
        {
            if (child == null) return;
            child.PropertyChanged -= Child_PropertyChanged;
            if (ReferenceEquals(child.Parent, this)) child.Parent = null;
        }

        private void Child_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (_suppressChildRollup) return;
            if (e.PropertyName == "IsChecked")
            {
                RecalculateFromChildren();
            }
        }

        private void RecalculateFromChildren()
        {
            if (_children == null || _children.Count == 0) return;

            bool anyTrue = false;
            bool anyFalse = false;
            bool anyNull = false;

            foreach (var child in _children)
            {
                if (!child.IsChecked.HasValue) anyNull = true;
                else if (child.IsChecked.Value) anyTrue = true;
                else anyFalse = true;

                if ((anyTrue && anyFalse) || anyNull) break;
            }

            bool? rolledUp;
            if (anyNull || (anyTrue && anyFalse)) rolledUp = null;
            else if (anyTrue) rolledUp = true;
            else rolledUp = false;

            if (_isChecked != rolledUp)
            {
                _isChecked = rolledUp;
                OnPropertyChanged("IsChecked");
            }

            if (_parent != null)
            {
                _parent.RecalculateFromChildren();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }

    public class TwoStateUserCheckBox : CheckBox
    {
        protected override void OnToggle()
        {
            bool next = this.IsChecked != true;
            this.IsChecked = next;
        }
    }
}
"@

# Get the loaded assembly paths for proper compilation
$asm_System = [System.Object].Assembly
$asm_Core = [System.Linq.Enumerable].Assembly
$asm_WindowsBase = [System.Windows.Threading.Dispatcher].Assembly
$asm_PresentationCore = [System.Windows.Media.ImageSource].Assembly
$asm_PresentationFramework = [System.Windows.Controls.Button].Assembly
$asm_Collections = [System.Collections.ObjectModel.ObservableCollection[int]].Assembly
$asm_ComponentModel = [System.ComponentModel.INotifyPropertyChanged].Assembly

# Compile C# with loaded assembly references to match versions exactly
Add-Type -TypeDefinition $csharpCode -ReferencedAssemblies @(
    $asm_System.Location,
    $asm_Core.Location,
    $asm_Collections.Location,
    $asm_ComponentModel.Location,
    $asm_WindowsBase.Location,
    $asm_PresentationCore.Location,
    $asm_PresentationFramework.Location
)

$BackupItemType = ("$csNamespace.BackupItem" -as [type])
$ItemKindType = ("$csNamespace.ItemKind" -as [type])
$ItemKind_Category = [System.Enum]::Parse($ItemKindType, "Category")
$ItemKind_Location = [System.Enum]::Parse($ItemKindType, "Location")

$asmName = $BackupItemType.Assembly.GetName().Name

# XAML for the selector GUI
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:local="clr-namespace:$csNamespace;assembly=$asmName"
        Title="Backup Selector" Height="500" Width="600" WindowStartupLocation="CenterScreen">
    <Window.Resources>
        <HierarchicalDataTemplate DataType="{x:Type local:BackupItem}" ItemsSource="{Binding Children}">
            <StackPanel Orientation="Horizontal">
                <local:TwoStateUserCheckBox
                    IsThreeState="True"
                    IsChecked="{Binding IsChecked, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"
                    IsEnabled="{Binding IsEnabled}"
                    Focusable="False"
                    VerticalAlignment="Center"/>
                <TextBlock Text="{Binding Name}" Margin="6,0,0,0" VerticalAlignment="Center" FontWeight="Normal"/>
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

        <StackPanel Orientation="Horizontal" Grid.Row="0" Margin="0,0,0,8">
            <TextBlock VerticalAlignment="Center" Foreground="#666">Select items to back up:</TextBlock>
        </StackPanel>

        <Border Grid.Row="1" BorderBrush="#CCC" BorderThickness="1" CornerRadius="4" Padding="6">
            <ScrollViewer VerticalScrollBarVisibility="Auto">
                <TreeView x:Name="treeBackupItems"/>
            </ScrollViewer>
        </Border>

        <StackPanel Orientation="Horizontal" Grid.Row="2" Margin="0,8,0,0" HorizontalAlignment="Right">
            <Button x:Name="btnOK" Content="Start Backup" Width="120" Margin="0,0,8,0" IsDefault="True"/>
            <Button x:Name="btnCancel" Content="Cancel" Width="120" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
"@

function Show-BackupSelector {
    [xml]$xamlDoc = $xaml
    $reader = New-Object System.Xml.XmlNodeReader($xamlDoc)
    
    try {
        $window = [Windows.Markup.XamlReader]::Load($reader)
    } catch {
        Write-Error "Failed to load XAML: $($_.Exception.Message)"
        return $null
    }

    # Get controls
    $treeView = $window.FindName("treeBackupItems")
    $btnOK = $window.FindName("btnOK")
    $btnCancel = $window.FindName("btnCancel")
    
    if (-not $treeView -or -not $btnOK -or -not $btnCancel) {
        Write-Error "Failed to find XAML controls in window"
        return $null
    }

    # Build tree structure
    $root = New-Object $BackupItemType.FullName
    $root.Name = "Backup Locations"
    $root.DisplayPath = "Backup Locations"
    $root.Kind = $ItemKind_Category
    $root.IsExpanded = $true

    function Get-OrCreateCategoryNode {
        param(
            [Parameter(Mandatory)]$ParentNode,
            [Parameter(Mandatory)][string]$CategoryName,
            [Parameter(Mandatory)][string]$CategoryDisplayPath
        )

        foreach ($existing in $ParentNode.Children) {
            if ($existing.Kind -eq $ItemKind_Category -and [string]::Equals([string]$existing.Name, $CategoryName, [System.StringComparison]::OrdinalIgnoreCase)) {
                return $existing
            }
        }

        $node = New-Object $BackupItemType.FullName
        $node.Name = $CategoryName
        $node.DisplayPath = $CategoryDisplayPath
        $node.Kind = $ItemKind_Category
        $node.IsExpanded = $true
        $node.IsEnabled = $true
        $ParentNode.Children.Add($node)
        return $node
    }

    function Add-SelectorEntryToTree {
        param(
            [Parameter(Mandatory)]$RootNode,
            [Parameter(Mandatory)][string]$DisplayPath,
            [Parameter(Mandatory)][string]$SelectorPath,
            [Parameter(Mandatory)][bool]$IsChecked,
            [Parameter(Mandatory)][bool]$IsEnabled
        )

        if ([string]::IsNullOrWhiteSpace($DisplayPath) -or [string]::IsNullOrWhiteSpace($SelectorPath)) { return }

        $segments = @($DisplayPath -split '\\' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        if ($segments.Count -eq 0) { return }

        $cursor = $RootNode
        $pathSoFar = ''

        if ($segments.Count -gt 1) {
            for ($i = 0; $i -lt ($segments.Count - 1); $i++) {
                $seg = [string]$segments[$i]
                $pathSoFar = if ([string]::IsNullOrWhiteSpace($pathSoFar)) { $seg } else { ([string]$pathSoFar) + '\' + ([string]$seg) }
                $cursor = Get-OrCreateCategoryNode -ParentNode $cursor -CategoryName $seg -CategoryDisplayPath $pathSoFar
            }
        }

        $leafName = [string]$segments[$segments.Count - 1]
        $leaf = New-Object $BackupItemType.FullName
        $leaf.Name = $leafName
        $leaf.DisplayPath = $DisplayPath
        $leaf.Kind = $ItemKind_Location
        $leaf.Path = $SelectorPath
        $leaf.IsChecked = $IsChecked
        $leaf.IsEnabled = $IsEnabled
        $cursor.Children.Add($leaf)
    }

    function Get-TopLevelDisplaySegment {
        param(
            [Parameter(Mandatory)][string]$DisplayPath
        )

        $segments = @($DisplayPath -split '\\' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        if ($segments.Count -eq 0) { return [string]$DisplayPath }
        return [string]$segments[0]
    }

    function Add-TopLevelSourcePathLabel {
        param(
            [Parameter(Mandatory)]$RootNode,
            [Parameter(Mandatory)][hashtable]$TopLevelSourceMap
        )

        foreach ($child in $RootNode.Children) {
            $topLevelName = [string]$child.DisplayPath
            if ([string]::IsNullOrWhiteSpace($topLevelName)) {
                $topLevelName = [string]$child.Name
            }

            if (-not $TopLevelSourceMap.ContainsKey($topLevelName)) { continue }

            $sources = @($TopLevelSourceMap[$topLevelName])
            if ($sources.Count -eq 0) { continue }

            $sourceLabel = ($sources | Sort-Object) -join ' | '
            $child.Name = "{0}  [{1}]" -f ([string]$child.Name), $sourceLabel
        }
    }

    function Sort-BackupTree {
        param([Parameter(Mandatory)]$Node)

        $ordered = @(
            $Node.Children |
            Sort-Object -Property @{ Expression = {
                ([string]$_.Name).ToLowerInvariant()
            } }
        )

        $Node.Children.Clear()
        foreach ($child in $ordered) {
            $Node.Children.Add($child)
        }

        foreach ($child in $Node.Children) {
            if ($child.Kind -eq $ItemKind_Category) {
                Sort-BackupTree -Node $child
            }
        }
    }

    function Add-CheckedSelectionsRecursive {
        param(
            [Parameter(Mandatory)]$Node,
            [Parameter(Mandatory)][ref]$SelectionOut
        )

        foreach ($child in $Node.Children) {
            if ($child.Kind -eq $ItemKind_Location) {
                if ($child.IsChecked -eq $true) {
                    $SelectionOut.Value += @{
                        Name = if (-not [string]::IsNullOrWhiteSpace([string]$child.DisplayPath)) { [string]$child.DisplayPath } else { [string]$child.Name }
                        Path = [string]$child.Path
                    }
                }
            } else {
                Add-CheckedSelectionsRecursive -Node $child -SelectionOut $SelectionOut
            }
        }
    }

    $selectorEntries = @()
    $topLevelSourceSets = @{}

    function Register-TopLevelSourcePath {
        param(
            [Parameter(Mandatory)][string]$DisplayPath,
            [Parameter()][string]$SourcePath = ''
        )

        if ([string]::IsNullOrWhiteSpace($SourcePath)) { return }

        $topLevel = Get-TopLevelDisplaySegment -DisplayPath $DisplayPath
        if ([string]::IsNullOrWhiteSpace($topLevel)) { return }

        if (-not $topLevelSourceSets.ContainsKey($topLevel)) {
            $topLevelSourceSets[$topLevel] = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        }

        $null = $topLevelSourceSets[$topLevel].Add([string]$SourcePath)
    }

    # Static selector items from centralized mapping
    foreach ($cfg in $script:SelectorConfig.StaticItems) {
        $sourcePath = switch ([string]$cfg.Path) {
            $script:SelectorPaths.Jobs      { Join-Path $env:USERPROFILE $script:PathConfig.JobsRelative; break }
            default { '' }
        }

        $selectorEntries += [pscustomobject]@{
            DisplayPath = [string]$cfg.Name
            SelectorPath = [string]$cfg.Path
            IsChecked = [bool]$cfg.IsChecked
            IsEnabled = [bool]$cfg.IsEnabled
        }

        Register-TopLevelSourcePath -DisplayPath ([string]$cfg.Name) -SourcePath ([string]$sourcePath)
    }

    # Custom selector items from centralized mappings
    $selectorIsOneDrive = [bool]$script:IsOneDriveMode
    foreach ($cfg in $script:CustomMappings) {
        if (-not [bool]$cfg.IsEnabled) { continue }

        $modeCfg = Resolve-CustomMappingModeConfig -Mapping $cfg -IsOneDrive $selectorIsOneDrive
        if ($null -eq $modeCfg) { continue }

        $selectorEntries += [pscustomobject]@{
            DisplayPath = [string]$cfg.Name
            SelectorPath = [string]$cfg.SelectorPath
            IsChecked = [bool]$cfg.IsChecked
            IsEnabled = $true
        }

        Register-TopLevelSourcePath -DisplayPath ([string]$cfg.Name) -SourcePath ([string]$modeCfg.SourceFolder)
    }

    # Documents - enumerate subfolders (filtered only by configured exclusions)
    $docsPath = Join-Path $env:USERPROFILE $script:PathConfig.DocumentsRoot
    if (Test-Path -LiteralPath $docsPath) {
        $docsDirs = Get-ChildItem -LiteralPath $docsPath -Directory -Force -ErrorAction SilentlyContinue | Where-Object { $_.Name -notin $script:SelectorConfig.DocumentsExclusions }
        foreach ($dir in $docsDirs) {
            $displayPath = ([string]$script:PathConfig.DocumentsRoot) + '\' + ([string]$dir.Name)
            $selectorPath = ([string]$script:SelectorPaths.DocumentsPrefix) + ([string]$dir.Name)
            $selectorEntries += [pscustomobject]@{
                DisplayPath = $displayPath
                SelectorPath = $selectorPath
                IsChecked = $true
                IsEnabled = $true
            }

            Register-TopLevelSourcePath -DisplayPath $displayPath -SourcePath $docsPath
        }
    }

    foreach ($entry in ($selectorEntries | Sort-Object -Property @{ Expression = { [string]$_.DisplayPath } })) {
        Add-SelectorEntryToTree -RootNode $root -DisplayPath ([string]$entry.DisplayPath) -SelectorPath ([string]$entry.SelectorPath) -IsChecked ([bool]$entry.IsChecked) -IsEnabled ([bool]$entry.IsEnabled)
    }

    Sort-BackupTree -Node $root

    $topLevelSources = @{}
    foreach ($k in $topLevelSourceSets.Keys) {
        $topLevelSources[[string]$k] = @($topLevelSourceSets[$k])
    }
    Add-TopLevelSourcePathLabel -RootNode $root -TopLevelSourceMap $topLevelSources
    Sort-BackupTree -Node $root

    # Show top-level nodes directly (Explorer-like hierarchy beneath)
    $treeView.ItemsSource = $root.Children

    # Button handlers - direct variable references work in this function scope
    $btnOK.Add_Click({
        # Collect selected leaf locations recursively
        $script:SelectedItems = @()
        Add-CheckedSelectionsRecursive -Node $root -SelectionOut ([ref]$script:SelectedItems)
        
        $window.DialogResult = $true
        $window.Close()
    })

    $btnCancel.Add_Click({
        $window.DialogResult = $false
        $window.Close()
    })

    # Initialize the array before showing dialog
    $script:SelectedItems = @()

    try {
        $result = $window.ShowDialog()
    } catch {
        Write-Error "Failed to show dialog: $($_.Exception.Message)"
        return $null
    }

    if ($result -eq $true) {
        return $script:SelectedItems
    } else {
        return $null
    }
}

# ============================================================
# BACKUP LOGIC
# ============================================================

# Progress helpers
function Write-ProgressSafe {
    param(
        [int]$Id = 1,
        [int]$ParentId = -1,
        [string]$Activity,
        [string]$Status,
        [int]$PercentComplete = -1
    )
    if (-not $script:ProgressEnabled) { return }

    if ($ParentId -ge 0) {
        if ($PercentComplete -ge 0) {
            Write-Progress -Id $Id -ParentId $ParentId -Activity $Activity -Status $Status -PercentComplete $PercentComplete
        } else {
            Write-Progress -Id $Id -ParentId $ParentId -Activity $Activity -Status $Status
        }
    } else {
        if ($PercentComplete -ge 0) {
            Write-Progress -Id $Id -Activity $Activity -Status $Status -PercentComplete $PercentComplete
        } else {
            Write-Progress -Id $Id -Activity $Activity -Status $Status
        }
    }
}

function Complete-ProgressSafe {
    param([int]$Id = 1)
    if (-not $script:ProgressEnabled) { return }
    Write-Progress -Id $Id -Activity $script:ProgressText.CompleteActivity -Completed
}

# Core helpers
function Ensure-Directory {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Path-ContainsProtectedFolder {
    param([Parameter(Mandatory)][string]$Path)
    $current = $Path
    while ($true) {
        $parent = Split-Path -Parent $current
        if ([string]::IsNullOrWhiteSpace($parent) -or ($parent -eq $current)) { break }
        $leaf = Split-Path -Leaf $current
        if (-not [string]::IsNullOrWhiteSpace($leaf) -and ($leaf -match '(?i)\bbit\s*-?\s*lock')) {
            return $true
        }
        $current = $parent
    }
    return $false
}

function Get-RelativePathSafe {
    param(
        [Parameter(Mandatory)][string]$FullPath,
        [Parameter(Mandatory)][string]$BasePath
    )
    $base = ([System.IO.Path]::GetFullPath($BasePath)).TrimEnd('\\')
    $full = ([System.IO.Path]::GetFullPath($FullPath))
    if ($full.Length -le $base.Length) { return '' }
    if ($full.Substring(0, $base.Length).TrimEnd('\\') -ne $base) { return '' }
    return $full.Substring($base.Length).TrimStart('\\')
}

function Get-BackupSnapshotFolder {
    <#
    .SYNOPSIS
    Gets a backup snapshot folder path under BackupRoot.
    
    .DESCRIPTION
    Modes:
    - CurrentOrCreate: returns today's _Backup_YYYYMMDD folder, renaming the most
      recent older snapshot to today's date when appropriate.
    - LatestExisting: returns the most recent existing _Backup_YYYYMMDD snapshot
      without creating or renaming folders.
    
    .PARAMETER BackupRootPath
    The root backup directory (e.g., F:\BACKUP_AND_SETTINGS)

    .PARAMETER Mode
    Snapshot selection behavior. Valid values are 'CurrentOrCreate' or 'LatestExisting'.
    
    .OUTPUTS
    [string] Snapshot folder path, or empty string when Mode='LatestExisting' and none exist.
    #>
    param(
        [Parameter(Mandatory)][string]$BackupRootPath,
        [Parameter()][ValidateSet('CurrentOrCreate','LatestExisting')][string]$Mode = 'CurrentOrCreate'
    )

    if ($Mode -eq 'CurrentOrCreate') {
        Ensure-Directory -Path $BackupRootPath

        $dateTag = Get-Date -Format $script:DateFormat
        $todayFolder = Join-Path $BackupRootPath ("{0}{1}" -f $script:BackupFolderPrefix, $dateTag)
        if (Test-Path -LiteralPath $todayFolder) {
            return $todayFolder
        }

        $backupFolders = @(Get-ChildItem -LiteralPath $BackupRootPath -Directory -Force -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match "^$($script:BackupFolderPattern)$" } |
            Sort-Object Name -Descending)

        if ($backupFolders.Count -gt 0) {
            $mostRecent = $backupFolders[0]
            $newName = "{0}{1}" -f $script:BackupFolderPrefix, $dateTag
            $newPath = Join-Path $BackupRootPath $newName
            Rename-Item -LiteralPath $mostRecent.FullName -NewName $newName -Force
            return $newPath
        }

        return $todayFolder
    }

    if (-not (Test-Path -LiteralPath $BackupRootPath)) {
        return ''
    }

    $latestFolders = @(Get-ChildItem -LiteralPath $BackupRootPath -Directory -Force -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match "^$($script:BackupFolderPattern)$" } |
        Sort-Object Name -Descending)

    if ($latestFolders.Count -gt 0) {
        return [string]$latestFolders[0].FullName
    }

    return ''
}

function Move-ToDeletionArea {
    param(
        [Parameter(Mandatory)][string]$PathToMove,
        [Parameter(Mandatory)][string]$DeletionRoot,
        [Parameter(Mandatory)][string]$RelLabel
    )
    if (-not (Test-Path -LiteralPath $PathToMove)) { return }
    if (Path-ContainsProtectedFolder -Path $PathToMove) {
        Write-Warning "Protected path detected (BitLocker-like folder in path): '$PathToMove'. Skipping move to deletion area."
        return
    }

    # Construct destination path, ensuring parent directories exist.
    # In OneDrive mode, prefer mirroring original structure relative to backup root.
    $effectiveRelLabel = $RelLabel
    if ($script:MirrorDeletionStructure -and -not [string]::IsNullOrWhiteSpace($script:DeletionMirrorBasePath)) {
        $mirroredRel = Get-RelativePathSafe -FullPath $PathToMove -BasePath $script:DeletionMirrorBasePath
        if (-not [string]::IsNullOrWhiteSpace($mirroredRel)) {
            $effectiveRelLabel = $mirroredRel
        }
    }

    $dest = Join-Path $DeletionRoot $effectiveRelLabel
    $parentDir = Split-Path -Parent $dest
    Ensure-Directory -Path $parentDir

    # If destination already exists, it could mean:
    # 1. This exact item was already moved (don't move again)
    # 2. A parent directory was already moved (file should have moved with parent)
    # 3. Partial directory from a previous incomplete operation
    # 
    # For safety, check if the destination is a directory that's empty or nearly empty
    # If so, it might be a directory that was partially moved. We should not attempt to move again.
    if (Test-Path -LiteralPath $dest) {
        $destItem = Get-Item -LiteralPath $dest -Force
        if ($destItem -is [System.IO.DirectoryInfo]) {
            # Destination is a directory. If source is also a directory, don't move again
            $sourceItem = Get-Item -LiteralPath $PathToMove -Force
            if ($sourceItem -is [System.IO.DirectoryInfo]) {
                # Both are directories - destination already exists, skip to avoid duplicates
                Write-Warning "Destination directory already exists (previously moved?), skipping: '$dest'"
                return
            }
        }
        # For files, if destination exists, it's a duplicate file - skip it
        Write-Warning "Destination already exists, skipping move to avoid loss: '$dest'"
        return
    }

    if ($PSCmdlet.ShouldProcess($PathToMove, "Move to '$dest'")) {
        try {
            Move-Item -LiteralPath $PathToMove -Destination $dest -Force
        }
        catch {
            Write-Warning "Failed to move '$PathToMove' to deletion staging. Error: $($_.Exception.Message)"
        }
    }
}

function Copy-FileIfDifferent {
    param(
        [Parameter(Mandatory)][System.IO.FileInfo]$SourceFile,
        [Parameter(Mandatory)][string]$DestFilePath
    )

    $copyNeeded = $true
    if (Test-Path -LiteralPath $DestFilePath) {
        $destInfo = Get-Item -LiteralPath $DestFilePath -Force
        $srcTime = $SourceFile.LastWriteTimeUtc
        $dstTime = $destInfo.LastWriteTimeUtc

        # Timestamp-only comparison (ignore file size)
        if ($srcTime -eq $dstTime) {
            $copyNeeded = $false
        }
    }

    if ($copyNeeded) {
        Ensure-Directory -Path (Split-Path -Parent $DestFilePath)
        if ($PSCmdlet.ShouldProcess($DestFilePath, "Copy from '$($SourceFile.FullName)'")) {
            try {
                Copy-Item -LiteralPath $SourceFile.FullName -Destination $DestFilePath -Force -ErrorAction Stop
                (Get-Item -LiteralPath $DestFilePath -Force).LastWriteTimeUtc = $SourceFile.LastWriteTimeUtc
            }
            catch {
                Write-Warning "Copy failed (skipping): '$($SourceFile.FullName)' -> '$DestFilePath'. Error: $($_.Exception.Message)"
                return $false
            }
        }
    }

    return $copyNeeded
}

function Test-FolderInSyncByTimestamp {
    param(
        [Parameter(Mandatory)][string]$SourceFolder,
        [Parameter(Mandatory)][string]$DestFolder
    )

    if (-not (Test-Path -LiteralPath $SourceFolder) -or -not (Test-Path -LiteralPath $DestFolder)) {
        return $false
    }

    $srcFiles = @(Get-ChildItem -LiteralPath $SourceFolder -File -Recurse -Force -ErrorAction SilentlyContinue)
    $dstFiles = @(Get-ChildItem -LiteralPath $DestFolder -File -Recurse -Force -ErrorAction SilentlyContinue)

    if ($srcFiles.Count -ne $dstFiles.Count) { return $false }

    $srcMap = New-Object 'System.Collections.Generic.Dictionary[string,long]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($f in $srcFiles) {
        $rel = Get-RelativePathSafe -FullPath $f.FullName -BasePath $SourceFolder
        if ([string]::IsNullOrWhiteSpace($rel)) { return $false }
        $srcMap[$rel] = $f.LastWriteTimeUtc.Ticks
    }

    foreach ($f in $dstFiles) {
        $rel = Get-RelativePathSafe -FullPath $f.FullName -BasePath $DestFolder
        if ([string]::IsNullOrWhiteSpace($rel)) { return $false }
        if (-not $srcMap.ContainsKey($rel)) { return $false }
        if ($srcMap[$rel] -ne $f.LastWriteTimeUtc.Ticks) { return $false }
    }

    $srcDirs = @(Get-ChildItem -LiteralPath $SourceFolder -Directory -Recurse -Force -ErrorAction SilentlyContinue)
    $dstDirs = @(Get-ChildItem -LiteralPath $DestFolder -Directory -Recurse -Force -ErrorAction SilentlyContinue)

    if ($srcDirs.Count -ne $dstDirs.Count) { return $false }

    $srcDirSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($d in $srcDirs) {
        $rel = Get-RelativePathSafe -FullPath $d.FullName -BasePath $SourceFolder
        if ([string]::IsNullOrWhiteSpace($rel)) { return $false }
        $null = $srcDirSet.Add($rel)
    }

    foreach ($d in $dstDirs) {
        $rel = Get-RelativePathSafe -FullPath $d.FullName -BasePath $DestFolder
        if ([string]::IsNullOrWhiteSpace($rel)) { return $false }
        if (-not $srcDirSet.Contains($rel)) { return $false }
    }

    return $true
}

function Get-RelativeDepth {
    param([Parameter(Mandatory)][string]$RelativePath)
    if ([string]::IsNullOrWhiteSpace($RelativePath)) { return 0 }
    return @($RelativePath -split '\\').Count
}

function Test-RelativePathHasAnyPrefix {
    param(
        [Parameter(Mandatory)][string]$RelativePath,
        [Parameter()][string[]]$Prefixes = @()
    )

    if ([string]::IsNullOrWhiteSpace($RelativePath)) { return $false }
    if ($null -eq $Prefixes -or $Prefixes.Count -eq 0) { return $false }

    foreach ($p in $Prefixes) {
        if ([string]::IsNullOrWhiteSpace([string]$p)) { continue }
        $prefix = ([string]$p).Trim('\')
        if ([string]::IsNullOrWhiteSpace($prefix)) { continue }

        if ($RelativePath.Equals($prefix, [System.StringComparison]::OrdinalIgnoreCase) -or
            $RelativePath.StartsWith(($prefix + '\'), [System.StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }

    return $false
}

function Get-StructureSnapshotLimited {
    param(
        [Parameter(Mandatory)][string]$RootFolder,
        [Parameter()][bool]$ExcludeReparsePointDirectories = $true
    )

    $dirs = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $files = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $fileTimes = New-Object 'System.Collections.Generic.Dictionary[string,long]' ([System.StringComparer]::OrdinalIgnoreCase)
    $reparseDirs = New-Object 'System.Collections.Generic.List[string]'

    if (-not (Test-Path -LiteralPath $RootFolder)) {
        return @{ Dirs = $dirs; Files = $files; FileTimes = $fileTimes }
    }

    $rootPath = [System.IO.Path]::GetFullPath($RootFolder)

    $allDirs = @(Get-ChildItem -LiteralPath $rootPath -Directory -Recurse -Force -ErrorAction SilentlyContinue)
    foreach ($d in $allDirs) {
        if ($ExcludeReparsePointDirectories -and (($d.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0)) {
            $reparseDirs.Add($d.FullName.TrimEnd('\'))
            continue
        }

        $rel = Get-RelativePathSafe -FullPath $d.FullName -BasePath $rootPath
        if ([string]::IsNullOrWhiteSpace($rel)) { continue }
        $null = $dirs.Add($rel)
    }

    $allFiles = @(Get-ChildItem -LiteralPath $rootPath -File -Recurse -Force -ErrorAction SilentlyContinue)
    foreach ($f in $allFiles) {
        if ($ExcludeReparsePointDirectories -and $reparseDirs.Count -gt 0) {
            $isUnderReparseDir = $false
            foreach ($r in $reparseDirs) {
                if ($f.FullName.StartsWith($r + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
                    $isUnderReparseDir = $true
                    break
                }
            }
            if ($isUnderReparseDir) { continue }
        }

        $rel = Get-RelativePathSafe -FullPath $f.FullName -BasePath $rootPath
        if ([string]::IsNullOrWhiteSpace($rel)) { continue }
        $null = $files.Add($rel)
        $fileTimes[$rel] = [long]$f.LastWriteTimeUtc.Ticks
    }

    return @{ Dirs = $dirs; Files = $files; FileTimes = $fileTimes }
}

function Get-ReportBucketPath {
    param(
        [Parameter(Mandatory)][string]$ParentDirectory,
        [Parameter(Mandatory)][int]$MaxDepth
    )

    if ([string]::IsNullOrWhiteSpace($ParentDirectory) -or $ParentDirectory -eq '.') { return '.' }

    $depth = Get-RelativeDepth -RelativePath $ParentDirectory
    if ($depth -le $MaxDepth) {
        return $ParentDirectory
    }

    $parts = @($ParentDirectory -split '\\')
    $head = ($parts | Select-Object -First $MaxDepth) -join '\\'
    return "$head\\..."
}

function Compare-FolderStructureLimited {
    param(
        [Parameter(Mandatory)][string]$SourceFolder,
        [Parameter(Mandatory)][string]$DestFolder,
        [Parameter(Mandatory)][string]$DisplayName,
        [Parameter(Mandatory)][int]$MaxDepth,
        [Parameter()][ValidateRange(0, 60)][int]$TimestampToleranceSeconds = 2,
        [Parameter()][string[]]$SourceExcludePrefixes = @(),
        [Parameter()][string[]]$DestExcludePrefixes = @(),
        [Parameter()][bool]$ExcludeReparsePointDirectories = $true
    )

    $src = Get-StructureSnapshotLimited -RootFolder $SourceFolder -ExcludeReparsePointDirectories:$ExcludeReparsePointDirectories
    $dst = Get-StructureSnapshotLimited -RootFolder $DestFolder -ExcludeReparsePointDirectories:$ExcludeReparsePointDirectories

    # Apply optional subtree exclusions (relative path prefixes)
    if ($SourceExcludePrefixes -and $SourceExcludePrefixes.Count -gt 0) {
        $srcDirsToRemove = @($src.Dirs | Where-Object { Test-RelativePathHasAnyPrefix -RelativePath ([string]$_) -Prefixes $SourceExcludePrefixes })
        foreach ($rel in $srcDirsToRemove) { $null = $src.Dirs.Remove([string]$rel) }

        $srcFilesToRemove = @($src.Files | Where-Object { Test-RelativePathHasAnyPrefix -RelativePath ([string]$_) -Prefixes $SourceExcludePrefixes })
        foreach ($rel in $srcFilesToRemove) {
            $null = $src.Files.Remove([string]$rel)
            if ($src.FileTimes.ContainsKey([string]$rel)) { $null = $src.FileTimes.Remove([string]$rel) }
        }
    }

    if ($DestExcludePrefixes -and $DestExcludePrefixes.Count -gt 0) {
        $dstDirsToRemove = @($dst.Dirs | Where-Object { Test-RelativePathHasAnyPrefix -RelativePath ([string]$_) -Prefixes $DestExcludePrefixes })
        foreach ($rel in $dstDirsToRemove) { $null = $dst.Dirs.Remove([string]$rel) }

        $dstFilesToRemove = @($dst.Files | Where-Object { Test-RelativePathHasAnyPrefix -RelativePath ([string]$_) -Prefixes $DestExcludePrefixes })
        foreach ($rel in $dstFilesToRemove) {
            $null = $dst.Files.Remove([string]$rel)
            if ($dst.FileTimes.ContainsKey([string]$rel)) { $null = $dst.FileTimes.Remove([string]$rel) }
        }
    }

    $timestampToleranceTicks = [long]([TimeSpan]::FromSeconds([double]$TimestampToleranceSeconds).Ticks)

    $diffs = New-Object System.Collections.ArrayList

    foreach ($rel in $src.Dirs) {
        if (-not $dst.Dirs.Contains($rel)) {
            $parent = Split-Path -Parent $rel
            if ([string]::IsNullOrWhiteSpace($parent)) { $parent = '.' }
            $parentDepth = if ($parent -eq '.') { 0 } else { Get-RelativeDepth -RelativePath $parent }
            $bucket = Get-ReportBucketPath -ParentDirectory $parent -MaxDepth $MaxDepth
            [void]$diffs.Add([pscustomobject]@{ EntryType = 'Directory'; RelativePath = $rel; Difference = 'MissingInDestination'; ParentDirectory = $parent; ParentDepth = $parentDepth; ReportBucket = $bucket })
        }
    }
    foreach ($rel in $dst.Dirs) {
        if (-not $src.Dirs.Contains($rel)) {
            $parent = Split-Path -Parent $rel
            if ([string]::IsNullOrWhiteSpace($parent)) { $parent = '.' }
            $parentDepth = if ($parent -eq '.') { 0 } else { Get-RelativeDepth -RelativePath $parent }
            $bucket = Get-ReportBucketPath -ParentDirectory $parent -MaxDepth $MaxDepth
            [void]$diffs.Add([pscustomobject]@{ EntryType = 'Directory'; RelativePath = $rel; Difference = 'ExtraInDestination'; ParentDirectory = $parent; ParentDepth = $parentDepth; ReportBucket = $bucket })
        }
    }
    foreach ($rel in $src.Files) {
        if (-not $dst.Files.Contains($rel)) {
            $parent = Split-Path -Parent $rel
            if ([string]::IsNullOrWhiteSpace($parent)) { $parent = '.' }
            $parentDepth = if ($parent -eq '.') { 0 } else { Get-RelativeDepth -RelativePath $parent }
            $bucket = Get-ReportBucketPath -ParentDirectory $parent -MaxDepth $MaxDepth
            [void]$diffs.Add([pscustomobject]@{ EntryType = 'File'; RelativePath = $rel; Difference = 'MissingInDestination'; ParentDirectory = $parent; ParentDepth = $parentDepth; ReportBucket = $bucket })
        }
    }
    foreach ($rel in $dst.Files) {
        if (-not $src.Files.Contains($rel)) {
            $parent = Split-Path -Parent $rel
            if ([string]::IsNullOrWhiteSpace($parent)) { $parent = '.' }
            $parentDepth = if ($parent -eq '.') { 0 } else { Get-RelativeDepth -RelativePath $parent }
            $bucket = Get-ReportBucketPath -ParentDirectory $parent -MaxDepth $MaxDepth
            [void]$diffs.Add([pscustomobject]@{ EntryType = 'File'; RelativePath = $rel; Difference = 'ExtraInDestination'; ParentDirectory = $parent; ParentDepth = $parentDepth; ReportBucket = $bucket })
        }
    }

    foreach ($rel in $src.FileTimes.Keys) {
        if ($dst.FileTimes.ContainsKey($rel)) {
            $deltaTicks = [long]$src.FileTimes[$rel] - [long]$dst.FileTimes[$rel]
            if ([math]::Abs([double]$deltaTicks) -gt [double]$timestampToleranceTicks) {
                $parent = Split-Path -Parent $rel
                if ([string]::IsNullOrWhiteSpace($parent)) { $parent = '.' }
                $parentDepth = if ($parent -eq '.') { 0 } else { Get-RelativeDepth -RelativePath $parent }
                $bucket = Get-ReportBucketPath -ParentDirectory $parent -MaxDepth $MaxDepth
                [void]$diffs.Add([pscustomobject]@{ EntryType = 'FileTimestamp'; RelativePath = $rel; Difference = 'TimestampMismatch'; ParentDirectory = $parent; ParentDepth = $parentDepth; ReportBucket = $bucket; DeltaSeconds = [math]::Round(([math]::Abs([double]$deltaTicks) / 10000000.0), 3) })
            }
        }
    }

    $differenceTypeSummary = @(
        $diffs |
        Group-Object -Property EntryType, Difference |
        Sort-Object Count -Descending |
        Select-Object @{Name='DifferenceType'; Expression={$_.Name}}, @{Name='Count'; Expression={$_.Count}}
    )

    $directorySummary = @(
        $diffs |
        Group-Object -Property ReportBucket |
        Sort-Object Count -Descending |
        Select-Object @{Name='DirectoryPath'; Expression={$_.Name}}, @{Name='DifferenceCount'; Expression={$_.Count}}, @{Name='ExamplePaths'; Expression={ (($_.Group | Select-Object -First 3 -ExpandProperty RelativePath) -join '; ') }}
    )

    $deeperDiffCount = @($diffs | Where-Object { $_.ParentDepth -gt $MaxDepth }).Count

    return [pscustomobject]@{
        DisplayName       = $DisplayName
        SourceFolder      = $SourceFolder
        DestFolder        = $DestFolder
        MaxDepth          = $MaxDepth
        TotalDifferences  = $diffs.Count
        DeeperDiffCount   = $deeperDiffCount
        SourceDirectoryCount = $src.Dirs.Count
        SourceFileCount      = $src.Files.Count
        DestDirectoryCount   = $dst.Dirs.Count
        DestFileCount        = $dst.Files.Count
        TimestampToleranceSeconds = $TimestampToleranceSeconds
        ExcludeReparsePointDirectories = $ExcludeReparsePointDirectories
        DifferenceTypeSummary = $differenceTypeSummary
        DirectorySummary  = $directorySummary
        Differences       = @($diffs)
    }
}

function Write-StructureDiffReport {
    param([Parameter(Mandatory)]$Report)

    if ($Report.TotalDifferences -eq 0) {
        Write-Host ("[OK] {0}: 0 differences (structure + file timestamps, tolerance <= {1}s) (full-depth query, report depth <= {2})" -f $Report.DisplayName, $Report.TimestampToleranceSeconds, $Report.MaxDepth) -ForegroundColor Green
        return
    }

    Write-Host ("[DIFF] {0}: {1} differences (structure + file timestamps, tolerance <= {2}s) (full-depth query, report depth <= {3})" -f $Report.DisplayName, $Report.TotalDifferences, $Report.TimestampToleranceSeconds, $Report.MaxDepth) -ForegroundColor Yellow
    Write-Host ("  Source: {0}" -f $Report.SourceFolder) -ForegroundColor DarkGray
    Write-Host ("  Dest:   {0}" -f $Report.DestFolder) -ForegroundColor DarkGray
    Write-Host ("  Source counts: {0} dirs, {1} files" -f $Report.SourceDirectoryCount, $Report.SourceFileCount) -ForegroundColor DarkGray
    Write-Host ("  Dest counts:   {0} dirs, {1} files" -f $Report.DestDirectoryCount, $Report.DestFileCount) -ForegroundColor DarkGray
    Write-Host ("  Reparse-point dirs excluded: {0}" -f $Report.ExcludeReparsePointDirectories) -ForegroundColor DarkGray
    Write-Host ("  Differences deeper than report depth: {0}" -f $Report.DeeperDiffCount) -ForegroundColor DarkGray

    if ($Report.PSObject.Properties.Name -contains 'DifferenceTypeSummary' -and $Report.DifferenceTypeSummary) {
        Write-Host "  Difference type summary:" -ForegroundColor DarkGray
        foreach ($typeRow in $Report.DifferenceTypeSummary) {
            Write-Host ("    {0} -> {1}" -f $typeRow.DifferenceType, $typeRow.Count) -ForegroundColor DarkGray
        }
    }

    foreach ($row in $Report.DirectorySummary) {
        $label = if ($row.DirectoryPath -eq '.') { '<root>' } else { $row.DirectoryPath }
        Write-Host ("    {0}  ->  {1}" -f $label, $row.DifferenceCount) -ForegroundColor Gray
        if ($row.PSObject.Properties.Name -contains 'ExamplePaths' -and -not [string]::IsNullOrWhiteSpace([string]$row.ExamplePaths)) {
            Write-Host ("      e.g. {0}" -f $row.ExamplePaths) -ForegroundColor DarkGray
        }
    }
}

function Resolve-CustomMapping {
    param(
        [Parameter(Mandatory)]$Mapping,
        [Parameter(Mandatory)][bool]$IsOneDrive,
        [Parameter(Mandatory)][string]$BackupRoot,
        [Parameter(Mandatory)][string]$BackupSnapshotRoot
    )

    $modeConfig = Resolve-CustomMappingModeConfig -Mapping $Mapping -IsOneDrive $IsOneDrive
    if ($null -eq $modeConfig) { return $null }

    $destBase = if ($modeConfig.DestinationBase -eq $script:DestinationBases.Run) { $BackupSnapshotRoot } else { $BackupRoot }
    $dest = if (-not [string]::IsNullOrWhiteSpace([string]$destBase)) {
        Join-Path $destBase $modeConfig.DestinationRelative
    } else {
        ''
    }

    return [pscustomobject]@{
        Name                     = [string]$Mapping.Name
        SelectorPath             = [string]$Mapping.SelectorPath
        SourceFolder             = [string]$modeConfig.SourceFolder
        DestFolder               = $dest
        DeletionLabel            = [string]$modeConfig.DestinationRelative
        RobocopyMode             = [string]$modeConfig.RobocopyMode
        ProgressActivityTemplate = [string]$modeConfig.ProgressActivityTemplate
        IncludeTimestampMismatch = [bool]$modeConfig.IncludeTimestampMismatch
    }
}

function Resolve-CustomSelectionMapping {
    param(
        [Parameter(Mandatory)]$SelectionItem,
        [Parameter(Mandatory)][hashtable]$MappingsByPath,
        [Parameter(Mandatory)][hashtable]$MappingsByName
    )

    $selectionPath = [string]$SelectionItem.Path
    $selectionName = [string]$SelectionItem.Name

    if ($MappingsByPath.ContainsKey($selectionPath)) {
        return $MappingsByPath[$selectionPath]
    }

    if ($MappingsByName.ContainsKey($selectionName)) {
        return $MappingsByName[$selectionName]
    }

    return $null
}

function Get-DocumentsDiffExclusionsFromCustomMappings {
    param(
        [Parameter(Mandatory)][string]$DocSourceRoot,
        [Parameter(Mandatory)][string]$DocDestRoot,
        [Parameter(Mandatory)][hashtable]$CustomMappingsByPath
    )

    $sourceSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $destSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($customKey in $CustomMappingsByPath.Keys) {
        $customEntry = $CustomMappingsByPath[[string]$customKey]
        if ($null -eq $customEntry) { continue }

        $customSourcePath = [string]$customEntry.SourceFolder
        $customDestPath = [string]$customEntry.DestFolder

        if (-not [string]::IsNullOrWhiteSpace($customSourcePath) -and (Test-Path -LiteralPath $DocSourceRoot)) {
            $relCustomSource = Get-RelativePathSafe -FullPath $customSourcePath -BasePath $DocSourceRoot
            if (-not [string]::IsNullOrWhiteSpace($relCustomSource)) {
                $null = $sourceSet.Add($relCustomSource)
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($customDestPath) -and (Test-Path -LiteralPath $DocDestRoot)) {
            $relCustomDest = Get-RelativePathSafe -FullPath $customDestPath -BasePath $DocDestRoot
            if (-not [string]::IsNullOrWhiteSpace($relCustomDest)) {
                $null = $destSet.Add($relCustomDest)
            }
        }
    }

    return [pscustomobject]@{
        SourceExcludePrefixes = @($sourceSet)
        DestExcludePrefixes   = @($destSet)
    }
}

function Resolve-CustomMappingModeConfig {
    <#
    Resolves effective custom-mapping settings for the active target mode.

    Behavior summary:
      - Prefers new schema: Mapping.Targets.<SSD|OneDrive>.*
      - Applies defaults/safety fallbacks for optional specialized fields:
          RobocopyMode             -> 'MirrorStandard'
          ProgressActivityTemplate -> $script:ProgressText.BackupCopyActivity
          IncludeTimestampMismatch -> $true
      - Unknown/invalid destination base or robocopy mode values are normalized
        to safe defaults.
    #>
    param(
        [Parameter(Mandatory)]$Mapping,
        [Parameter(Mandatory)][bool]$IsOneDrive
    )

    $targetKey = if ($IsOneDrive) { [string]$script:BackupTargets.OneDrive } else { [string]$script:BackupTargets.SSD }
    $targetDefaults = $script:BackupTargetConfig[$targetKey]
    $defaultBase = [string]$targetDefaults.DefaultCustomDestinationBase

    $targets = $null
    if ($Mapping -is [System.Collections.IDictionary]) {
        foreach ($k in $Mapping.Keys) {
            if ([string]::Equals([string]$k, 'Targets', [System.StringComparison]::OrdinalIgnoreCase)) {
                $targets = $Mapping[$k]
                break
            }
        }
    } elseif ($Mapping.PSObject.Properties.Name -contains 'Targets') {
        $targets = $Mapping.Targets
    }
    if ($null -eq $targets) { return $null }

    $modeCfg = $null
    if ($targets -is [System.Collections.IDictionary]) {
        foreach ($k in $targets.Keys) {
            if ([string]::Equals([string]$k, $targetKey, [System.StringComparison]::OrdinalIgnoreCase)) {
                $modeCfg = $targets[$k]
                break
            }
        }
    } else {
        foreach ($p in $targets.PSObject.Properties) {
            if ([string]::Equals([string]$p.Name, $targetKey, [System.StringComparison]::OrdinalIgnoreCase)) {
                $modeCfg = $p.Value
                break
            }
        }
    }
    if ($null -eq $modeCfg) { return $null }

    $readProp = {
        param($Obj, [string]$PropName)
        if ($null -eq $Obj) { return $null }

        if ($Obj -is [System.Collections.IDictionary]) {
            foreach ($k in $Obj.Keys) {
                if ([string]::Equals([string]$k, $PropName, [System.StringComparison]::OrdinalIgnoreCase)) {
                    return $Obj[$k]
                }
            }
            return $null
        }

        foreach ($p in $Obj.PSObject.Properties) {
            if ([string]::Equals([string]$p.Name, $PropName, [System.StringComparison]::OrdinalIgnoreCase)) {
                return $p.Value
            }
        }

        return $null
    }

    $source = [string](& $readProp $modeCfg 'SourceFolder')
    $destRel = [string](& $readProp $modeCfg 'DestinationRelative')
    $basePref = [string](& $readProp $modeCfg 'DestinationBase')
    $robocopyMode = [string](& $readProp $modeCfg 'RobocopyMode')
    $progressActivityTemplate = [string](& $readProp $modeCfg 'ProgressActivityTemplate')

    $includeTimestampMismatch = $true
    $includeRaw = & $readProp $modeCfg 'IncludeTimestampMismatch'
    if ($null -ne $includeRaw) {
        if ($includeRaw -is [bool]) {
            $includeTimestampMismatch = [bool]$includeRaw
        } else {
            $parsedBool = $false
            if ([bool]::TryParse([string]$includeRaw, [ref]$parsedBool)) {
                $includeTimestampMismatch = $parsedBool
            }
        }
    }

    if ([string]::IsNullOrWhiteSpace($basePref)) { $basePref = $defaultBase }
    if ($basePref -notin @($script:DestinationBases.Root, $script:DestinationBases.Run)) { $basePref = $defaultBase }
    if ($robocopyMode -notin @('MirrorStandard', 'MirrorNoRetry')) { $robocopyMode = 'MirrorStandard' }
    if ([string]::IsNullOrWhiteSpace($progressActivityTemplate)) { $progressActivityTemplate = $script:ProgressText.BackupCopyActivity }

    if ([string]::IsNullOrWhiteSpace($source) -or [string]::IsNullOrWhiteSpace($destRel)) { return $null }

    return [pscustomobject]@{
        SourceFolder             = $source
        DestinationRelative      = $destRel
        DestinationBase          = $basePref
        RobocopyMode             = $robocopyMode
        ProgressActivityTemplate = $progressActivityTemplate
        IncludeTimestampMismatch = [bool]$includeTimestampMismatch
    }
}

function Stage-BackupOnlyItems {
    # Stages backup-only content into deletion area.
    # ItemKinds controls whether directories, files, or both are processed.
    # Files can optionally include timestamp mismatch staging.
    param(
        [Parameter(Mandatory=$true)][string]$SourceFolder,
        [Parameter(Mandatory=$true)][string]$DestFolder,
        [Parameter(Mandatory=$true)][string]$DeletionRoot,
        [Parameter(Mandatory=$true)][string]$DeletionLabel,
        [ValidateSet('Directories','Files','Both')][string]$ItemKinds = 'Both',
        [bool]$IncludeTimestampMismatch = $true
    )

    if (-not (Test-Path -LiteralPath $DestFolder)) { return }

    $processDirectories = $ItemKinds -in @('Directories','Both')
    $processFiles = $ItemKinds -in @('Files','Both')

    if ($processDirectories) {
        $destDirs = @(
            Get-ChildItem -LiteralPath $DestFolder -Directory -Recurse -Force -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -notin $script:IgnoredNames.StageDirectoriesNotOnLive }
        ) | Sort-Object { @($_.FullName -split '\\').Count }

        foreach ($dir in $destDirs) {
            if (-not (Test-Path -LiteralPath $dir.FullName)) { continue }

            $rel = Get-RelativePathSafe -FullPath $dir.FullName -BasePath $DestFolder
            $sourcePath = Join-Path $SourceFolder $rel
            if (-not (Test-Path -LiteralPath $sourcePath)) {
                $relLabel = Join-Path $DeletionLabel $rel
                Move-ToDeletionArea -PathToMove $dir.FullName -DeletionRoot $DeletionRoot -RelLabel $relLabel
            }
        }
    }

    if ($processFiles) {
        $destFiles = @(Get-ChildItem -LiteralPath $DestFolder -File -Recurse -Force -ErrorAction SilentlyContinue)

        foreach ($file in $destFiles) {
            $rel = Get-RelativePathSafe -FullPath $file.FullName -BasePath $DestFolder
            $sourcePath = Join-Path $SourceFolder $rel

            if (-not (Test-Path -LiteralPath $sourcePath)) {
                $relLabel = Join-Path $DeletionLabel $rel
                Move-ToDeletionArea -PathToMove $file.FullName -DeletionRoot $DeletionRoot -RelLabel $relLabel
                continue
            }

            if ($IncludeTimestampMismatch) {
                $srcInfo = Get-Item -LiteralPath $sourcePath -Force
                if ($srcInfo.LastWriteTimeUtc -ne $file.LastWriteTimeUtc) {
                    $relLabel = Join-Path $DeletionLabel $rel
                    Move-ToDeletionArea -PathToMove $file.FullName -DeletionRoot $DeletionRoot -RelLabel $relLabel
                }
            }
        }
    }
}

function Get-RobocopyArgs {
    param(
        [Parameter(Mandatory)][string]$SourceFolder,
        [Parameter(Mandatory)][string]$DestFolder,
        [Parameter(Mandatory)][ValidateSet('MirrorStandard','MirrorNoRetry','CopyNoMirror','CompareList')][string]$Mode
    )

    $modeArgs = $script:RobocopyModes[$Mode]
    if (-not $modeArgs) {
        throw "Unknown robocopy mode '$Mode'."
    }

    return @($SourceFolder, $DestFolder) + @($modeArgs)
}

function Test-RobocopyDeltaNeeded {
    <#
    .SYNOPSIS
    Uses robocopy list mode to detect whether a copy pass is needed.

    .DESCRIPTION
    Performs a full-depth metadata comparison between source and destination
    without mirroring deletions. Uses robocopy /L so no content is copied.

    Returns $true when a sync pass is needed, $false when already in sync.
    If robocopy reports warnings/failures (exit code >= 8), returns $true
    to allow the normal sync pass to proceed.
    #>
    param(
        [Parameter(Mandatory)][string]$SourceFolder,
        [Parameter(Mandatory)][string]$DestFolder,
        [Parameter(Mandatory)][string]$DisplayName
    )

    if (-not (Test-Path -LiteralPath $DestFolder)) {
        return $true
    }

    $probeArgs = Get-RobocopyArgs -SourceFolder $SourceFolder -DestFolder $DestFolder -Mode CompareList

    try {
        $null = & robocopy @probeArgs
        $rc = $LASTEXITCODE
        if ($rc -ge 8) {
            Write-Warning "Robocopy preflight compare reported warnings/failures for $DisplayName (exit code $rc). Proceeding with sync pass."
            return $true
        }

        $hasCopyDelta = (($rc -band 1) -ne 0)
        $hasMismatchDelta = (($rc -band 4) -ne 0)
        return ($hasCopyDelta -or $hasMismatchDelta)
    }
    catch {
        Write-Warning "Robocopy preflight compare failed for $DisplayName. Proceeding with sync pass. Error: $($_.Exception.Message)"
        return $true
    }
}

function Test-IsJobsSelectionPath {
    param([Parameter(Mandatory)][string]$SelectionPath)

    $documentsJobsPath = "{0}Jobs" -f $script:SelectorPaths.DocumentsPrefix
    return ($SelectionPath -eq $script:SelectorPaths.Jobs) -or ($SelectionPath -eq $documentsJobsPath)
}

# Jobs operations (incremental)
function Sync-JobsIncremental {
    <#
    .SYNOPSIS
    Performs incremental Jobs sync in either Standard or OneDriveDelta mode.

    .DESCRIPTION
    Modes:
      - Standard:
          Enumerates source folders/files, creates destination directories, and
          copies files that differ by timestamp.
      - OneDriveDelta:
          Uses robocopy (non-mirror) to copy only new/changed files to reduce
          OneDrive hydration/download churn.

    .PARAMETER SourceRoot
    The live system Documents\Jobs folder.

    .PARAMETER DestRoot
    Destination Jobs folder (snapshot Jobs path for Standard, Current Jobs path for OneDriveDelta).

    .PARAMETER Mode
    Sync behavior mode: Standard or OneDriveDelta.

    .PARAMETER ProgressParentId
    Parent progress ID for nested progress display.
    #>
    param(
        [Parameter(Mandatory)][string]$SourceRoot,
        [Parameter(Mandatory)][string]$DestRoot,
        [ValidateSet('Standard','OneDriveDelta')][string]$Mode = 'Standard',
        [int]$ProgressParentId = 1
    )

    if (-not (Test-Path -LiteralPath $SourceRoot)) {
        Write-Warning "Jobs source folder not found: $SourceRoot"
        return
    }

    Ensure-Directory -Path (Split-Path -Parent $DestRoot)
    Ensure-Directory -Path $DestRoot

    if ($Mode -eq 'OneDriveDelta') {
        Write-ProgressSafe -Id $script:ProgressIds.JobsSync -ParentId $ProgressParentId -Activity $script:ProgressText.JobsOneDriveDeltaActivity -Status $script:ProgressText.JobsOneDriveCompareStatus -PercentComplete 30

        $needsSync = Test-RobocopyDeltaNeeded -SourceFolder $SourceRoot -DestFolder $DestRoot -DisplayName 'Jobs (OneDrive Current)'
        if (-not $needsSync) {
            Write-ProgressSafe -Id $script:ProgressIds.JobsSync -ParentId $ProgressParentId -Activity $script:ProgressText.JobsOneDriveDeltaActivity -Status $script:ProgressText.JobsOneDriveSkipStatus -PercentComplete 100
            Complete-ProgressSafe -Id $script:ProgressIds.JobsSync
            return
        }

        Write-ProgressSafe -Id $script:ProgressIds.JobsSync -ParentId $ProgressParentId -Activity $script:ProgressText.JobsOneDriveDeltaActivity -Status $script:ProgressText.JobsOneDriveCopyStatus -PercentComplete 70

        # Important: no /MIR here. We only push new/changed items to reduce OneDrive churn.
        # Robocopy compares metadata and skips unchanged files.
        $robocopyArgs = Get-RobocopyArgs -SourceFolder $SourceRoot -DestFolder $DestRoot -Mode CopyNoMirror

        if ($PSCmdlet.ShouldProcess($DestRoot, "Robocopy from '$SourceRoot' (delta, no mirror)")) {
            try {
                $null = & robocopy @robocopyArgs
                $rc = $LASTEXITCODE
                if ($rc -ge 8) {
                    Write-Warning "Robocopy reported warnings/failures for Jobs OneDrive sync (exit code $rc)"
                }
            }
            catch {
                Write-Warning "Robocopy invocation failed for Jobs OneDrive sync. Error: $($_.Exception.Message)"
            }
        }

        Complete-ProgressSafe -Id $script:ProgressIds.JobsSync
        return
    }

    Write-ProgressSafe -Id $script:ProgressIds.JobsSync -ParentId $ProgressParentId -Activity $script:ProgressText.JobsScanActivity -Status $script:ProgressText.EnumeratingStatus
    $dirs  = @(Get-ChildItem -LiteralPath $SourceRoot -Directory -Recurse -Force)
    $files = @(Get-ChildItem -LiteralPath $SourceRoot -File -Recurse -Force)
    Complete-ProgressSafe -Id $script:ProgressIds.JobsSync

    # Create directory structure
    $dirCount = $dirs.Count
    for ($i = 0; $i -lt $dirCount; $i++) {
        $shouldUpdateDirProgress = $script:ProgressEnabled -and (
            ($dirCount -le 0) -or
            ($i -eq 0) -or
            (($i % $script:ProgressUpdateEvery) -eq 0) -or
            ($i -eq ($dirCount - 1))
        )
        if ($shouldUpdateDirProgress) {
            $dirPct = if ($dirCount -le 0) { 100 } else { [int][math]::Min(100, [math]::Max(0, [math]::Floor(100 * (($i + 1) / [double]$dirCount)))) }
            Write-ProgressSafe -Id $script:ProgressIds.JobsDirCreate -ParentId $ProgressParentId -Activity $script:ProgressText.JobsCreateDirsActivity -Status ($script:ProgressText.CountStatusFormat -f ($i + 1), $dirCount) -PercentComplete $dirPct
        }
        $rel = Get-RelativePathSafe -FullPath $dirs[$i].FullName -BasePath $SourceRoot
        Ensure-Directory -Path (Join-Path $DestRoot $rel)
    }
    Complete-ProgressSafe -Id $script:ProgressIds.JobsDirCreate

    # Copy files that differ
    $fileCount = $files.Count
    for ($i = 0; $i -lt $fileCount; $i++) {
        $shouldUpdateFileProgress = $script:ProgressEnabled -and (
            ($fileCount -le 0) -or
            ($i -eq 0) -or
            (($i % $script:ProgressUpdateEvery) -eq 0) -or
            ($i -eq ($fileCount - 1))
        )
        if ($shouldUpdateFileProgress) {
            $filePct = if ($fileCount -le 0) { 100 } else { [int][math]::Min(100, [math]::Max(0, [math]::Floor(100 * (($i + 1) / [double]$fileCount)))) }
            Write-ProgressSafe -Id $script:ProgressIds.JobsFileCopy -ParentId $ProgressParentId -Activity $script:ProgressText.JobsCopyFilesActivity -Status ($script:ProgressText.CountStatusFormat -f ($i + 1), $fileCount) -PercentComplete $filePct
        }
        $rel = Get-RelativePathSafe -FullPath $files[$i].FullName -BasePath $SourceRoot
        $destPath = Join-Path $DestRoot $rel
        Copy-FileIfDifferent -SourceFile $files[$i] -DestFilePath $destPath | Out-Null
    }
    Complete-ProgressSafe -Id $script:ProgressIds.JobsFileCopy
}

function Move-CurrentJobsFoldersToOld {
    <#
    .SYNOPSIS
    Reconciles OneDrive Current Jobs against live Jobs with split handling for top-level and deep removals.
    
    .DESCRIPTION
    Behavior:
    1) Top-level folder missing on live:
         - Merge that entire Current top-level job folder into "2 - Old Jobs".
         - If destination exists, conflicts are resolved by newest file timestamp (UTC).
         - Destination-only files in Old Jobs are preserved.
         - After successful merge, remove the source top-level folder from Current.

    2) Top-level folder still exists on live:
         - Reconcile deep removals inside Current by comparing to live.
         - Missing subfolders/files are moved to "To Be Deleted" (not Old Jobs).
         - Path in "To Be Deleted" mirrors original path relative to backup root.

    This keeps Old Jobs reserved for removed top-level jobs while still pruning
    stale deep content from Current Jobs.
    
    .PARAMETER SourceRoot
    The live system Documents\Jobs folder
    
    .PARAMETER CurrentRoot
    The OneDrive "1 - Current Jobs" folder
    
    .PARAMETER OldRoot
    The OneDrive "2 - Old Jobs" folder
    
    .PARAMETER DeletionRoot
    The root folder where deep removed items are staged.
    
    .PARAMETER RemovalLabel
    Relative label prefix used when staging deep removals from Current Jobs.
    #>
    param(
        [Parameter(Mandatory)][string]$SourceRoot,
        [Parameter(Mandatory)][string]$CurrentRoot,
        [Parameter(Mandatory)][string]$OldRoot,
        [Parameter(Mandatory)][string]$DeletionRoot,
        [Parameter(Mandatory)][string]$RemovalLabel
    )

    if (-not (Test-Path -LiteralPath $CurrentRoot)) { return }

    function Merge-DirectoryPreferNewest {
        param(
            [Parameter(Mandatory)][string]$SourceFolder,
            [Parameter(Mandatory)][string]$DestFolder
        )

        if (-not (Test-Path -LiteralPath $SourceFolder)) { return }

        Ensure-Directory -Path $DestFolder

        $hadCopyErrors = $false
        $sourceFiles = @(Get-ChildItem -LiteralPath $SourceFolder -File -Recurse -Force -ErrorAction SilentlyContinue)

        foreach ($srcFile in $sourceFiles) {
            $rel = Get-RelativePathSafe -FullPath $srcFile.FullName -BasePath $SourceFolder
            if ([string]::IsNullOrWhiteSpace($rel)) { continue }

            $destFile = Join-Path $DestFolder $rel
            $copySource = $false

            if (-not (Test-Path -LiteralPath $destFile)) {
                $copySource = $true
            } else {
                try {
                    $destInfo = Get-Item -LiteralPath $destFile -Force -ErrorAction Stop
                    if ($srcFile.LastWriteTimeUtc -gt $destInfo.LastWriteTimeUtc) {
                        $copySource = $true
                    }
                }
                catch {
                    # If destination cannot be read, attempt replacement from source
                    $copySource = $true
                }
            }

            if ($copySource) {
                Ensure-Directory -Path (Split-Path -Parent $destFile)
                if ($PSCmdlet.ShouldProcess($destFile, "Copy newest version from '$($srcFile.FullName)'")) {
                    try {
                        Copy-Item -LiteralPath $srcFile.FullName -Destination $destFile -Force -ErrorAction Stop
                        (Get-Item -LiteralPath $destFile -Force).LastWriteTimeUtc = $srcFile.LastWriteTimeUtc
                    }
                    catch {
                        $hadCopyErrors = $true
                        Write-Warning "Failed to copy newest-wins merge file '$($srcFile.FullName)' -> '$destFile'. Error: $($_.Exception.Message)"
                    }
                }
            }
        }

        if ($hadCopyErrors) {
            Write-Warning "Merge completed with copy errors. Source folder retained for safety: '$SourceFolder'"
            return
        }

        if (Test-Path -LiteralPath $SourceFolder) {
            if ($PSCmdlet.ShouldProcess($SourceFolder, "Remove source folder after newest-wins merge")) {
                try {
                    Remove-Item -LiteralPath $SourceFolder -Recurse -Force -ErrorAction Stop
                }
                catch {
                    Write-Warning "Failed to remove source folder after newest-wins merge: '$SourceFolder'. Error: $($_.Exception.Message)"
                }
            }
        }
    }

    Ensure-Directory -Path $OldRoot

    $liveTopDirs = @(Get-ChildItem -LiteralPath $SourceRoot -Directory -Force -ErrorAction SilentlyContinue)
    $liveTopSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($d in $liveTopDirs) { $null = $liveTopSet.Add($d.Name) }

    $currentTopDirs = @(Get-ChildItem -LiteralPath $CurrentRoot -Directory -Force -ErrorAction SilentlyContinue)

    # PHASE 1: top-level folders missing on live -> merge to Old Jobs (newest wins)
    foreach ($topDir in $currentTopDirs) {
        if (-not $liveTopSet.Contains($topDir.Name)) {
            $dest = Join-Path $OldRoot $topDir.Name

            if ($PSCmdlet.ShouldProcess($topDir.FullName, "Merge top-level to '$dest' (newest timestamp wins)")) {
                try {
                    Merge-DirectoryPreferNewest -SourceFolder $topDir.FullName -DestFolder $dest
                }
                catch {
                    Write-Warning "Failed to merge top-level '$($topDir.FullName)' to Old Jobs with newest-wins logic. Error: $($_.Exception.Message)"
                }
            }
        }
    }

    # PHASE 2: deep reconciliation for top-level folders that still exist on live
    $currentTopDirs = @(Get-ChildItem -LiteralPath $CurrentRoot -Directory -Force -ErrorAction SilentlyContinue)
    foreach ($topDir in $currentTopDirs) {
        if (-not $liveTopSet.Contains($topDir.Name)) { continue }

        $liveTopPath = Join-Path $SourceRoot $topDir.Name

        # Move deep missing directories to To Be Deleted (shallowest first)
        $subDirs = @(Get-ChildItem -LiteralPath $topDir.FullName -Directory -Recurse -Force -ErrorAction SilentlyContinue)
        $subDirs = $subDirs | Sort-Object { @($_.FullName -split '\\').Count }
        $movedSubDirs = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

        foreach ($dir in $subDirs) {
            if (-not (Test-Path -LiteralPath $dir.FullName)) { continue }

            $relFromTop = Get-RelativePathSafe -FullPath $dir.FullName -BasePath $topDir.FullName
            if ([string]::IsNullOrWhiteSpace($relFromTop)) { continue }

            $livePath = Join-Path $liveTopPath $relFromTop
            if (-not (Test-Path -LiteralPath $livePath)) {
                Move-ToDeletionArea -PathToMove $dir.FullName -DeletionRoot $DeletionRoot -RelLabel (Join-Path $RemovalLabel (Join-Path $topDir.Name $relFromTop))
                if (-not (Test-Path -LiteralPath $dir.FullName)) {
                    $null = $movedSubDirs.Add($dir.FullName)
                }
            }
        }

        # Move deep missing files to To Be Deleted
        $subFiles = @(Get-ChildItem -LiteralPath $topDir.FullName -File -Recurse -Force -ErrorAction SilentlyContinue)
        foreach ($file in $subFiles) {
            $skip = $false
            foreach ($movedDir in $movedSubDirs) {
                if ($file.FullName.StartsWith($movedDir + '\\')) {
                    $skip = $true
                    break
                }
            }
            if ($skip) { continue }

            $relFromTop = Get-RelativePathSafe -FullPath $file.FullName -BasePath $topDir.FullName
            if ([string]::IsNullOrWhiteSpace($relFromTop)) { continue }

            $livePath = Join-Path $liveTopPath $relFromTop
            if (-not (Test-Path -LiteralPath $livePath)) {
                Move-ToDeletionArea -PathToMove $file.FullName -DeletionRoot $DeletionRoot -RelLabel (Join-Path $RemovalLabel (Join-Path $topDir.Name $relFromTop))
            }
        }
    }

    # Cleanup remaining empty directories in Current that are not on live
    Remove-EmptyDirsNotOnLiveJobs -SourceRoot $SourceRoot -DestRoot $CurrentRoot
}

function Remove-EmptyDirsNotOnLiveJobs {
    <#
    .SYNOPSIS
    Removes empty directories from the Jobs backup that don't exist on the live system.
    
    .DESCRIPTION
    Iteratively removes directories that are:
    1. Empty (contain no files or subdirectories)
    2. Don't exist on the live system
    
    This is typically run after moving backup-only items to cleanup the remaining structure.
    
    .PARAMETER SourceRoot
    The live system Documents\Jobs folder
    
    .PARAMETER DestRoot
    The backup Documents\Jobs folder
    
    .PARAMETER ProgressParentId
    The parent progress ID for nested progress display
    #>
    param(
        [Parameter(Mandatory)][string]$SourceRoot,
        [Parameter(Mandatory)][string]$DestRoot,
        [int]$ProgressParentId = 1
    )

    if (-not (Test-Path -LiteralPath $SourceRoot) -or -not (Test-Path -LiteralPath $DestRoot)) { return }

    Write-ProgressSafe -Id $script:ProgressIds.JobsEmptyDirs -ParentId $ProgressParentId -Activity $script:ProgressText.JobsRemoveEmptyActivity -Status $script:ProgressText.ScanningStatus

    # Attempt to remove empty directories up to 5 times (in case of cascading empties)
    $attempt = 0
    while ($attempt -lt 5) {
        $attempt++
        $empty = @(Get-ChildItem -LiteralPath $DestRoot -Directory -Recurse -Force | Where-Object { @(Get-ChildItem -LiteralPath $_.FullName -Recurse -Force).Count -eq 0 })
        
        if ($empty.Count -eq 0) { break }

        foreach ($dir in $empty) {
            $rel = Get-RelativePathSafe -FullPath $dir.FullName -BasePath $DestRoot
            $sourcePath = Join-Path $SourceRoot $rel
            
            # Only remove if directory doesn't exist on live system
            if (-not (Test-Path -LiteralPath $sourcePath)) {
                if ($PSCmdlet.ShouldProcess($dir.FullName, "Remove empty directory")) {
                    try {
                        Remove-Item -LiteralPath $dir.FullName -Force
                    }
                    catch {
                        Write-Warning "Failed to remove empty dir: '$($dir.FullName)'. Error: $($_.Exception.Message)"
                    }
                }
            }
        }
    }
    Complete-ProgressSafe -Id $script:ProgressIds.JobsEmptyDirs
}

function Copy-FolderWithRobocopy {
    <#
    .SYNOPSIS
    Mirrors a folder from live system to backup using robocopy.
    
    .DESCRIPTION
    Three-stage process:
    1. Moves directories that don't exist on live to deletion area
    2. Moves files that don't exist or differ from live to deletion area
    3. Uses robocopy for final sync with configurable mode
    
    This approach ensures backup-only items are safely preserved before robocopy
    makes any deletions.
    
    .PARAMETER SourceFolder
    The live system folder to back up
    
    .PARAMETER DestFolder
    The backup destination folder
    
    .PARAMETER DisplayName
    Display name for progress messages
    
    .PARAMETER DeletionRoot
    The root folder where items to be deleted are staged
    
    .PARAMETER DeletionLabel
    The relative path label within DeletionRoot
    
    .PARAMETER ProgressParentId
    The parent progress ID for nested progress display
    
    .PARAMETER ProgressId
    The progress ID to use for this folder's operations

    .PARAMETER RobocopyMode
    Robocopy mode from $script:RobocopyModes. Use MirrorNoRetry for AppData-like behavior.

    .PARAMETER ProgressActivityTemplate
    Activity text format string that accepts DisplayName (e.g. 'Backup: copying {0}').

    .PARAMETER IncludeTimestampMismatch
    When true, backup-only files with timestamp mismatch are staged before mirroring.
    #>
    param(
        [Parameter(Mandatory)][string]$SourceFolder,
        [Parameter(Mandatory)][string]$DestFolder,
        [Parameter(Mandatory)][string]$DisplayName,
        [Parameter(Mandatory)][string]$DeletionRoot,
        [Parameter(Mandatory)][string]$DeletionLabel,
        [int]$ProgressParentId = 1,
        [int]$ProgressId = 50,
        [ValidateSet('MirrorStandard','MirrorNoRetry')][string]$RobocopyMode = 'MirrorStandard',
        [string]$ProgressActivityTemplate = $script:ProgressText.BackupCopyActivity,
        [bool]$IncludeTimestampMismatch = $true
    )

    $activityText = $ProgressActivityTemplate -f $DisplayName

    if (-not (Test-Path -LiteralPath $SourceFolder)) {
        Write-Warning "Source folder not found: $SourceFolder"
        return
    }

    if (Test-Path -LiteralPath $DestFolder) {
        if (Test-FolderInSyncByTimestamp -SourceFolder $SourceFolder -DestFolder $DestFolder) {
            Write-ProgressSafe -Id $ProgressId -ParentId $ProgressParentId -Activity $activityText -Status $script:ProgressText.SyncSkipStatus -PercentComplete 100
            Complete-ProgressSafe -Id $ProgressId
            return
        }
    }

    Write-ProgressSafe -Id $ProgressId -ParentId $ProgressParentId -Activity $activityText -Status $script:ProgressText.Stage1Status -PercentComplete 15

    # Move entire backup-only directories first (one operation per dir, not per file)
    # This prevents a 10000-file directory from being processed file-by-file
    Stage-BackupOnlyItems -SourceFolder $SourceFolder -DestFolder $DestFolder -DeletionRoot $DeletionRoot -DeletionLabel $DeletionLabel -ItemKinds Directories

    Write-ProgressSafe -Id $ProgressId -ParentId $ProgressParentId -Activity $activityText -Status $script:ProgressText.Stage1bStatus -PercentComplete 30

    # Move individual files BEFORE robocopy runs, so they aren't deleted by robocopy's /MIR flag
    Stage-BackupOnlyItems -SourceFolder $SourceFolder -DestFolder $DestFolder -DeletionRoot $DeletionRoot -DeletionLabel $DeletionLabel -ItemKinds Files -IncludeTimestampMismatch:$IncludeTimestampMismatch

    Ensure-Directory -Path (Split-Path -Parent $DestFolder)
    Ensure-Directory -Path $DestFolder

    Write-ProgressSafe -Id $ProgressId -ParentId $ProgressParentId -Activity $activityText -Status $script:ProgressText.Stage2Status -PercentComplete 50

    $robocopyArgs = Get-RobocopyArgs -SourceFolder $SourceFolder -DestFolder $DestFolder -Mode $RobocopyMode

    if ($PSCmdlet.ShouldProcess($DestFolder, "Robocopy from '$SourceFolder'")) {
        try {
            $null = & robocopy @robocopyArgs
            $rc = $LASTEXITCODE
            if ($rc -ge 8) {
                Write-Warning "Robocopy reported warnings/failures for $DisplayName (exit code $rc)"
            }
        }
        catch {
            Write-Warning "Robocopy invocation failed for $DisplayName. Error: $($_.Exception.Message)"
        }
    }

    Write-ProgressSafe -Id $ProgressId -ParentId $ProgressParentId -Activity $activityText -Status $script:ProgressText.Stage3DoneStatus -PercentComplete 100
    Complete-ProgressSafe -Id $ProgressId
}

# ============================================================
# MAIN EXECUTION
# ============================================================

Write-Host $script:MainText.HeaderLine -ForegroundColor Cyan
Write-Host $script:MainText.HeaderTitle -ForegroundColor Cyan
Write-Host $script:MainText.HeaderLine -ForegroundColor Cyan
Write-Host ""

$ssdTargetRootDisplay = [string]$script:BackupTargetConfig[$script:BackupTargets.SSD].RootPath
$oneDriveTargetRootDisplay = [string]$script:BackupTargetConfig[$script:BackupTargets.OneDrive].RootPath

# Select backup target (SSD vs OneDrive)
$selectedTarget = $BackupTarget
if ([string]::IsNullOrWhiteSpace($selectedTarget)) {
    Write-Host $script:MainText.HeaderSelectTarget -ForegroundColor Cyan
    Write-Host "  1) SSD ($ssdTargetRootDisplay)" -ForegroundColor Gray
    Write-Host "  2) OneDrive ($oneDriveTargetRootDisplay)" -ForegroundColor Gray
    $choice = Read-Host "Enter 1 or 2"
    switch ($choice) {
        '2' { $selectedTarget = $script:BackupTargets.OneDrive }
        default { $selectedTarget = $script:BackupTargets.SSD }
    }
} else {
    if ($selectedTarget -notin @($script:BackupTargets.SSD, $script:BackupTargets.OneDrive)) {
        throw "Invalid BackupTarget '$selectedTarget'. Valid values are '$($script:BackupTargets.SSD)' or '$($script:BackupTargets.OneDrive)'."
    }
}

$IsOneDrive = ($selectedTarget -eq $script:BackupTargets.OneDrive)
$script:IsOneDriveMode = $IsOneDrive
$ActiveBackupRoot = [string]$script:BackupTargetConfig[$selectedTarget].RootPath
if ($IsOneDrive) {
    $script:MirrorDeletionStructure = $true
    $script:DeletionMirrorBasePath = $ActiveBackupRoot

    if ($PSVersionTable.PSVersion.Major -ge 7) {
        Write-Warning "OneDrive Files On-Demand can behave differently under PowerShell 7+. If you still see heavy hydration, run this script in Windows PowerShell 5.1 and compare behavior."
    }
} else {
    $script:MirrorDeletionStructure = $false
    $script:DeletionMirrorBasePath = ''
}

# Show selector GUI
Write-Host $script:MainText.HeaderSelectorLaunch -ForegroundColor Cyan
$selectedItems = Show-BackupSelector

if (-not $selectedItems) {
    Write-Host $script:MainText.MessageCancelled -ForegroundColor Yellow
    exit 0
}

if ($selectedItems.Count -eq 0) {
    Write-Host $script:MainText.MessageNoSelection -ForegroundColor Yellow
    exit 0
}

Write-Host $script:MainText.SelectedItemsTitle -ForegroundColor Green
foreach ($item in $selectedItems) {
    Write-Host "  - $($item.Name)"
}
Write-Host ""

$runBackupOperations = (-not $DryRunDiffCheck) -or $RunBackupBeforeDiffCheck

# Setup directories
Ensure-Directory -Path $ActiveBackupRoot
if ($runBackupOperations) {
    $BackupRunRoot = Get-BackupSnapshotFolder -BackupRootPath $ActiveBackupRoot -Mode CurrentOrCreate
    Ensure-Directory -Path $BackupRunRoot

    $runStamp = Get-Date -Format $script:DateTimeFormat
    $DeletionSessionRoot = Join-Path (Join-Path $ActiveBackupRoot $script:PathConfig.DeletionRootName) ("{0}{1}" -f $script:PathConfig.DeletionRunPrefix, $runStamp)
    Ensure-Directory -Path $DeletionSessionRoot
} else {
    $BackupRunRoot = Get-BackupSnapshotFolder -BackupRootPath $ActiveBackupRoot -Mode LatestExisting
    $DeletionSessionRoot = ''
}

$hasBackupRunRoot = -not [string]::IsNullOrWhiteSpace([string]$BackupRunRoot)

# Define source paths (live system locations)
$JobsSource         = Join-Path $env:USERPROFILE $script:PathConfig.JobsRelative

# Define destination paths (in timestamped snapshot folder)
$JobsDest           = if ($hasBackupRunRoot) { Join-Path $BackupRunRoot $script:PathConfig.JobsRelative } else { '' }
$JobsCurrentDest    = Join-Path $ActiveBackupRoot $script:PathConfig.JobsCurrentRelative
$JobsOldDest        = Join-Path $ActiveBackupRoot $script:PathConfig.JobsOldRelative

# Build custom mapping lookups (selector path + display name fallback)
$customMappingLookup = @{}
$customMappingNameLookup = @{}
$customProgressOffset = 0
foreach ($cfg in $script:CustomMappings) {
    if (-not [bool]$cfg.IsEnabled) { continue }

    $resolved = Resolve-CustomMapping -Mapping $cfg -IsOneDrive $IsOneDrive -BackupRoot $ActiveBackupRoot -BackupSnapshotRoot $BackupRunRoot
    if ($null -eq $resolved) { continue }

    $resolved | Add-Member -NotePropertyName ProgressId -NotePropertyValue ($script:ProgressIds.CustomBase + $customProgressOffset)
    $customProgressOffset++

    $customMappingLookup[[string]$cfg.SelectorPath] = $resolved
    if (-not [string]::IsNullOrWhiteSpace([string]$resolved.Name)) {
        $customMappingNameLookup[[string]$resolved.Name] = $resolved
    }
}

Write-Host "Backup root:           $ActiveBackupRoot" -ForegroundColor Gray
Write-Host "Backup target:         $selectedTarget" -ForegroundColor Gray
if ($IsOneDrive) {
    Write-Host "OneDrive sync mode:    Jobs keep Current/Old behavior; other folders mirror with staged deletions" -ForegroundColor Gray
}
if ($hasBackupRunRoot) {
    Write-Host "Backup snapshot folder (_Backup_YYYYMMDD): $BackupRunRoot" -ForegroundColor Gray
} else {
    Write-Host "Backup snapshot folder (_Backup_YYYYMMDD): <none found>" -ForegroundColor Yellow
}
if ($runBackupOperations) {
    Write-Host "Deletion session:      $DeletionSessionRoot" -ForegroundColor Gray
}
Write-Host ""

if ($runBackupOperations) {
    Write-ProgressSafe -Id $script:ProgressIds.Main -Activity $script:MainText.ProgressActivityMain -Status $script:MainText.ProgressStart -PercentComplete 0

    $progressStep = 0
    $totalItems = ($selectedItems | Measure-Object).Count

    # ============================================================
    # PROCESS DOCUMENTS\JOBS (if selected)
    # ============================================================
    $jobsSelected = $selectedItems | Where-Object { Test-IsJobsSelectionPath -SelectionPath ([string]$_.Path) }
    if ($jobsSelected) {
        if ($IsOneDrive) {
            Write-Host "=== Documents\Jobs (OneDrive: Current/Old) ===" -ForegroundColor Cyan
        } else {
            Write-Host "=== Documents\Jobs (incremental) ===" -ForegroundColor Cyan
        }

        Write-ProgressSafe -Id $script:ProgressIds.Main -Activity $script:MainText.ProgressActivityMain -Status $script:MainText.ProgressJobsSync -PercentComplete (15 * $progressStep / $totalItems)
        $progressStep++

        if ($IsOneDrive) {
            Sync-JobsIncremental -SourceRoot $JobsSource -DestRoot $JobsCurrentDest -Mode OneDriveDelta -ProgressParentId $script:ProgressIds.Main

            $jobsReconStatus = $script:MainText.ProgressJobsRecon -f (Split-Path -Leaf $script:PathConfig.JobsOldRelative), $script:PathConfig.DeletionRootName
            Write-ProgressSafe -Id $script:ProgressIds.Main -Activity $script:MainText.ProgressActivityMain -Status $jobsReconStatus -PercentComplete (25 * $progressStep / $totalItems)
            Move-CurrentJobsFoldersToOld -SourceRoot $JobsSource -CurrentRoot $JobsCurrentDest -OldRoot $JobsOldDest -DeletionRoot $DeletionSessionRoot -RemovalLabel $script:PathConfig.JobsRemovedFromLiveLabelOneDriveCurrent
        } else {
            Sync-JobsIncremental -SourceRoot $JobsSource -DestRoot $JobsDest -Mode Standard -ProgressParentId $script:ProgressIds.Main

            Write-ProgressSafe -Id $script:ProgressIds.Main -Activity $script:MainText.ProgressActivityMain -Status $script:MainText.ProgressJobsMoveOnly -PercentComplete (35 * $progressStep / $totalItems)
            if ((Test-Path -LiteralPath $JobsSource) -and (Test-Path -LiteralPath $JobsDest)) {
                Write-ProgressSafe -Id $script:ProgressIds.JobsDirCheck -ParentId $script:ProgressIds.Main -Activity $script:ProgressText.JobsCheckRemovedDirsActivity -Status $script:ProgressText.ProcessingDirsStatus
                Stage-BackupOnlyItems -SourceFolder $JobsSource -DestFolder $JobsDest -DeletionRoot $DeletionSessionRoot -DeletionLabel $script:PathConfig.JobsRemovedFromLiveLabelSSD -ItemKinds Directories
                Complete-ProgressSafe -Id $script:ProgressIds.JobsDirCheck

                Write-ProgressSafe -Id $script:ProgressIds.JobsFileCheck -ParentId $script:ProgressIds.Main -Activity $script:ProgressText.JobsCheckRemovedFilesActivity -Status $script:ProgressText.EnumeratingFilesStatus
                Stage-BackupOnlyItems -SourceFolder $JobsSource -DestFolder $JobsDest -DeletionRoot $DeletionSessionRoot -DeletionLabel $script:PathConfig.JobsRemovedFromLiveLabelSSD -ItemKinds Files -IncludeTimestampMismatch:$false
                Complete-ProgressSafe -Id $script:ProgressIds.JobsFileCheck
            }

            Write-ProgressSafe -Id $script:ProgressIds.Main -Activity $script:MainText.ProgressActivityMain -Status $script:MainText.ProgressJobsClean -PercentComplete (50 * $progressStep / $totalItems)
            Remove-EmptyDirsNotOnLiveJobs -SourceRoot $JobsSource -DestRoot $JobsDest -ProgressParentId $script:ProgressIds.Main
        }

        Write-Host "Jobs completed." -ForegroundColor Green
        Write-Host ""
    } else {
        Write-Host "Documents\Jobs (skipped by user)" -ForegroundColor Yellow
        Write-Host ""
    }

    # ============================================================
    # PROCESS OPTIONAL BACKUP ITEMS (custom mappings + Documents\*)
    # ============================================================
    $optionalItems = @($selectedItems | Where-Object { -not (Test-IsJobsSelectionPath -SelectionPath ([string]$_.Path)) })
    $optionalTotal = $optionalItems.Count
    $itemIndex = 0
    $customSelectorPattern = "{0}*" -f $script:SelectorPaths.CustomPrefix
    $documentsSelectorPattern = "{0}*" -f $script:SelectorPaths.DocumentsPrefix
    foreach ($item in $optionalItems) {
        $itemIndex++
        $currentProgress = if ($optionalTotal -gt 0) {
            50 + (50 * ($itemIndex - 1) / $optionalTotal)
        } else {
            50
        }

        $customCfg = Resolve-CustomSelectionMapping -SelectionItem $item -MappingsByPath $customMappingLookup -MappingsByName $customMappingNameLookup

        if ($item.Path -like $customSelectorPattern -or $null -ne $customCfg) {
            if ($null -eq $customCfg) {
                Write-Warning "No custom mapping found for selector path '$($item.Path)' (name '$($item.Name)'). Skipping."
                continue
            }

            $cfg = $customCfg
            $customSource = $cfg.SourceFolder
            $customDest = $cfg.DestFolder

            if ([string]::IsNullOrWhiteSpace([string]$customDest)) {
                Write-Warning "No destination is available for '$($cfg.Name)' in this target mode. Skipping."
                continue
            }

            Write-Host "=== $($cfg.Name) (custom mapping) ===" -ForegroundColor Cyan
            Write-ProgressSafe -Id $script:ProgressIds.Main -Activity $script:MainText.ProgressActivityMain -Status ($script:MainText.ProgressGenericCopy -f $cfg.Name) -PercentComplete $currentProgress

            Copy-FolderWithRobocopy -SourceFolder $customSource -DestFolder $customDest -DeletionRoot $DeletionSessionRoot -DeletionLabel $cfg.DeletionLabel -DisplayName $cfg.Name -ProgressParentId $script:ProgressIds.Main -ProgressId $cfg.ProgressId -RobocopyMode ([string]$cfg.RobocopyMode) -ProgressActivityTemplate ([string]$cfg.ProgressActivityTemplate) -IncludeTimestampMismatch:([bool]$cfg.IncludeTimestampMismatch)

            Write-Host "$($cfg.Name) completed." -ForegroundColor Green
            Write-Host ""
            continue
        }

        # Handle Documents\* folders (except Jobs)
        if ($item.Path -like $documentsSelectorPattern) {
            $folderName = $item.Path -replace ("^{0}" -f [regex]::Escape($script:SelectorPaths.DocumentsPrefix)), ''
            $docSource = Join-Path $env:USERPROFILE ("{0}\{1}" -f $script:PathConfig.DocumentsRoot, $folderName)
            $docDest = Join-Path $ActiveBackupRoot $folderName

            Write-Host "=== Documents\$folderName (mirror backup) ===" -ForegroundColor Cyan
            Write-ProgressSafe -Id $script:ProgressIds.Main -Activity $script:MainText.ProgressActivityMain -Status ($script:MainText.ProgressGenericCopy -f ("Documents\{0}" -f $folderName)) -PercentComplete $currentProgress

            Copy-FolderWithRobocopy -SourceFolder $docSource -DestFolder $docDest -DeletionRoot $DeletionSessionRoot -DeletionLabel $folderName -DisplayName ("Documents\{0}" -f $folderName) -ProgressParentId $script:ProgressIds.Main -ProgressId (52 + $itemIndex)

            Write-Host "Documents\$folderName completed." -ForegroundColor Green
            Write-Host ""
        }
    }

    Write-ProgressSafe -Id $script:ProgressIds.Main -Activity $script:MainText.ProgressActivityMain -Status $script:MainText.ProgressFinalize -PercentComplete 100
    Complete-ProgressSafe -Id $script:ProgressIds.Main

    Write-Host $script:MainText.HeaderLine -ForegroundColor Cyan
    Write-Host $script:MainText.MessageCompleted -ForegroundColor Cyan
    Write-Host $script:MainText.HeaderLine -ForegroundColor Cyan
    Write-Host ($script:MainText.MessageReviewDeletes -f $DeletionSessionRoot) -ForegroundColor Green
    Write-Host ""
} else {
    Write-Host "Dry diff check only mode: skipping backup/copy/move operations." -ForegroundColor Yellow
    Write-Host "Tip: add -RunBackupBeforeDiffCheck to run backup first, then compare." -ForegroundColor Gray
    Write-Host ""
}

if ($DryRunDiffCheck) {
    $excludeReparseForDiff = -not $IsOneDrive

    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Post-run dry diff check" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ("Query depth: full | Report depth: <= {0} | Timestamp tolerance: <= {1}s | Exclude reparse dirs: {2}" -f $DryRunMaxDepth, $script:DryDiffTimestampToleranceSeconds, $excludeReparseForDiff) -ForegroundColor Gray
    if (-not $hasBackupRunRoot) {
        Write-Warning "No existing backup snapshot folder (_Backup_YYYYMMDD) was found. Items mapped to the snapshot destination base will be skipped in dry diff check."
    }

    $compareTargets = @()
    $customSelectorPattern = "{0}*" -f $script:SelectorPaths.CustomPrefix
    $documentsSelectorPattern = "{0}*" -f $script:SelectorPaths.DocumentsPrefix

    foreach ($item in $selectedItems) {
        if (Test-IsJobsSelectionPath -SelectionPath ([string]$item.Path)) {
            $dest = if ($IsOneDrive) {
                $JobsCurrentDest
            } elseif ($hasBackupRunRoot) {
                $JobsDest
            } else {
                ''
            }

            if ([string]::IsNullOrWhiteSpace([string]$dest)) {
                Write-Warning "Skipping dry diff check for '$($script:PathConfig.JobsRelative)' because no backup snapshot destination (_Backup_YYYYMMDD) is available."
                continue
            }

            $compareTargets += [pscustomobject]@{ Name = $script:PathConfig.JobsRelative; Source = $JobsSource; Dest = $dest }
            continue
        }

        $customCfg = Resolve-CustomSelectionMapping -SelectionItem $item -MappingsByPath $customMappingLookup -MappingsByName $customMappingNameLookup

        if ($item.Path -like $customSelectorPattern -or $null -ne $customCfg) {
            if ($null -eq $customCfg) {
                Write-Warning "Skipping dry diff check for '$($item.Name)' because no matching custom mapping was found for selector path '$($item.Path)'."
                continue
            }

            $cfg = $customCfg
            $customDest = $cfg.DestFolder

            if (-not (Test-Path -LiteralPath ([string]$cfg.SourceFolder))) {
                Write-Warning "Custom mapping source path not found for '$($cfg.Name)': '$([string]$cfg.SourceFolder)'."
            }

            if ([string]::IsNullOrWhiteSpace([string]$customDest)) {
                Write-Warning "Skipping dry diff check for custom mapping '$($cfg.Name)' because no destination is available for the selected target mode."
                continue
            }

            if (-not (Test-Path -LiteralPath ([string]$customDest))) {
                Write-Warning "Custom mapping destination path not found for '$($cfg.Name)': '$([string]$customDest)'. Dry diff will treat destination as empty."
            }

            Write-Host ("Custom dry diff target: '{0}' | Source='{1}' | Dest='{2}'" -f $cfg.Name, $cfg.SourceFolder, $customDest) -ForegroundColor DarkGray
            $compareTargets += [pscustomobject]@{ Name = $cfg.Name; Source = $cfg.SourceFolder; Dest = $customDest }
            continue
        }

        if ($item.Path -like $documentsSelectorPattern) {
            $folderName = $item.Path -replace ("^{0}" -f [regex]::Escape($script:SelectorPaths.DocumentsPrefix)), ''
            $docSource = Join-Path $env:USERPROFILE ("{0}\{1}" -f $script:PathConfig.DocumentsRoot, $folderName)
            $docDest = Join-Path $ActiveBackupRoot $folderName

            $docExclusions = Get-DocumentsDiffExclusionsFromCustomMappings -DocSourceRoot $docSource -DocDestRoot $docDest -CustomMappingsByPath $customMappingLookup
            $docSourceExcludeArray = @($docExclusions.SourceExcludePrefixes)
            $docDestExcludeArray = @($docExclusions.DestExcludePrefixes)
            if ($docSourceExcludeArray.Count -gt 0 -or $docDestExcludeArray.Count -gt 0) {
                Write-Host ("Documents dry diff exclusions for '{0}' | Source: {1} | Dest: {2}" -f $folderName, ($docSourceExcludeArray -join '; '), ($docDestExcludeArray -join '; ')) -ForegroundColor DarkGray
            }

            $compareTargets += [pscustomobject]@{
                Name                  = ("{0}\{1}" -f $script:PathConfig.DocumentsRoot, $folderName)
                Source                = $docSource
                Dest                  = $docDest
                SourceExcludePrefixes = $docSourceExcludeArray
                DestExcludePrefixes   = $docDestExcludeArray
            }
        }
    }

    $totalDiffCount = 0
    foreach ($t in $compareTargets) {
        try {
            $sourceExcludes = @()
            $destExcludes = @()
            if ($t.PSObject.Properties.Name -contains 'SourceExcludePrefixes' -and $null -ne $t.SourceExcludePrefixes) {
                $sourceExcludes = @($t.SourceExcludePrefixes)
            }
            if ($t.PSObject.Properties.Name -contains 'DestExcludePrefixes' -and $null -ne $t.DestExcludePrefixes) {
                $destExcludes = @($t.DestExcludePrefixes)
            }

            $report = Compare-FolderStructureLimited -SourceFolder ([string]$t.Source) -DestFolder ([string]$t.Dest) -DisplayName ([string]$t.Name) -MaxDepth ([int]$DryRunMaxDepth) -TimestampToleranceSeconds ([int]$script:DryDiffTimestampToleranceSeconds) -SourceExcludePrefixes $sourceExcludes -DestExcludePrefixes $destExcludes -ExcludeReparsePointDirectories:$excludeReparseForDiff
            $totalDiffCount += [int]$report.TotalDifferences
            Write-StructureDiffReport -Report $report
        }
        catch {
            Write-Warning ("Dry diff check failed for '{0}'. Source='{1}' Dest='{2}'. Error: {3}" -f $t.Name, $t.Source, $t.Dest, $_.Exception.Message)
        }
    }

    Write-Host "" 
    if ($totalDiffCount -eq 0) {
        Write-Host "Dry diff check complete: no differences found (structure + file timestamps) for selected items." -ForegroundColor Green
    } else {
        Write-Host ("Dry diff check complete: {0} differences found (structure + file timestamps) across selected items." -f $totalDiffCount) -ForegroundColor Yellow
    }
    Write-Host ""
}


#Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert | Select-Object Subject, Thumbprint, NotAfter, Issuer
#$cert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert | Sort-Object NotAfter -Descending | Select-Object -First 1
#Set-AuthenticodeSignature -FilePath "C:\users\lynkcn\downloads\Run-Backup-WithSelector.ps1" -Certificate $cert
#
# SIG # Begin signature block
# MIIFqgYJKoZIhvcNAQcCoIIFmzCCBZcCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCZ7aA9hOCJqolO
# /J27YSKOqvWl5zeg+fchplOp2PNLSaCCAxgwggMUMIIB/KADAgECAhA4a9fNJJqk
# l0Ft/Uys+gnZMA0GCSqGSIb3DQEBCwUAMCIxIDAeBgNVBAMMF1Bvd2VyU2hlbGwg
# Q29kZSBTaWduaW5nMB4XDTI2MDIyMDE3MzkzNVoXDTI3MDIyMDE3NTkzNVowIjEg
# MB4GA1UEAwwXUG93ZXJTaGVsbCBDb2RlIFNpZ25pbmcwggEiMA0GCSqGSIb3DQEB
# AQUAA4IBDwAwggEKAoIBAQDrAwI415l4NH33vetIqPVyZCvGRq1WOUapgxH4X9Gx
# BISguKAX9OgI7IqnsRJ4InnFuAHv6b7k+Y9qfY6rNMEjIc8Gb1h5gI9aaRuE1YWd
# f5su1RChwyO8tLu6YnwYwPny6E3ws2Iw9qSuEcFoTPmE+fH5vw55jJzxfC7ZhzlW
# vzGf5Oxmrd4jG/0eH3gW7ZiSS66WsmWAhL4vaMo5WDWq2b4EW4iYckzdZ8A5BMUg
# t/wf/C5e+bvDCNtJftjcGrNuq7bMBQool+trpAfEK8ZJzxZywCWwsJaXx/xmNHPX
# TgJFdJvwAal3ZWAmYt+K44OLa9JguPiP9iN/lssFMnlRAgMBAAGjRjBEMA4GA1Ud
# DwEB/wQEAwIHgDATBgNVHSUEDDAKBggrBgEFBQcDAzAdBgNVHQ4EFgQUU6I29NDw
# H8H5D01ugSlxDaufTcUwDQYJKoZIhvcNAQELBQADggEBACXraDOhrGphkOhmtD5d
# FAz3Yi7CPglI7Nr2CIZnowMXZK9ta/kBjbGb1usfOxYQGzid7FP4aUCIPGQBd4YA
# WYI42zPc+jnDeVNZs4hjHSWuN7IVEuM5yN73tGLiMUObopU8c+uSk00DZ9mFBc8x
# 7JsEcpeVygp8SthgI3KOnib0CmWEkoGIZrszrMB5YgwJDRwoaSZAZ4348vEzilhx
# dWZOTeHUmOBHxoRF3ex4fzJyFaGbx73IZtGXxhcPwRnwRmFtLhrRqEU7OzodLlJ3
# ziYOm5cXlgIDHRta24dIEQr1QAoI1K7U33dnKNhG/C9/V1tN/A9qfnPt77uV20J4
# 4mcxggHoMIIB5AIBATA2MCIxIDAeBgNVBAMMF1Bvd2VyU2hlbGwgQ29kZSBTaWdu
# aW5nAhA4a9fNJJqkl0Ft/Uys+gnZMA0GCWCGSAFlAwQCAQUAoIGEMBgGCisGAQQB
# gjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYK
# KwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIKSs5dZt
# jr/WB/YcQz/VJlbL8dCmUCWGjAs+LdZyru2UMA0GCSqGSIb3DQEBAQUABIIBAIna
# lKBQXIrjEDU2Vrqea/blK3fOXoNM7qDOyBFftfP/DwFb+t1QniLAjHefAFL9SBde
# Gp+VNKwH2VCqpRN8DBzDWNKV1radi+0leQLC8wM0VxELckYb1nuYZs26U/OkZew+
# LuAc1a9GLrdqL16UPobc2wNzJZc+rKQ7Yf7N7+VdVQOhKNF6HuOF/zx/UOjUJbQB
# Qyq+XQWw6OaEonvCTKt+K4zWRKvl5NWfEysJwRitLDcLedGi6H3bFDYEVe+ExLRc
# Al8xENxKesvT0xPNq1jFFBpBBRCai/s6GyliopuORsasnjcHh4N8PfREmudRyYA0
# o4lyxp1r/2qtttwjX8Q=
# SIG # End signature block
