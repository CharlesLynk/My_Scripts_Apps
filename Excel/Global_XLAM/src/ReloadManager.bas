Attribute VB_Name = "ReloadManager"
Option Explicit
'==========================================
' Module: ReloadManager
' Purpose: Sync workbook VBA standard modules from repo BAS sources.
' Scope: Workbook type tagging, signature checks, and module re-import.
' Notes:
' - XLAM is expected in %APPDATA%\Microsoft\AddIns.
' - Managed from File -> Options -> Add-Ins.
' - Available as VBAProject(Global).ReloadManager.
'==========================================

'================== CONFIGURATION ==================
Private Const REPO_ROOT As String = "D:\Utilities\My_Scripts_Apps\Excel"
Private Const DOC_PROP_TYPE_STRING As Long = 4
Private Const SYNC_DIALOG_TITLE As String = "VBA Sync Manager"
Private Const MSG_SOURCE_UNAVAILABLE As String = "Local VBA source files are not available on this machine."
Private Const MSG_REQUEST_UPDATE As String = "Request updated VBA script from Chase."

' Document properties used in target workbooks
Private Const PROP_TYPE_ID As String = "WB_TYPE_ID"
Private Const PROP_SRC_SIG As String = "WB_SRC_SIGNATURE"
Private Const PROP_SRC_SIG_MAP As String = "WB_SRC_SIGNATURE_MAP"
Private Const PROP_LAST_SYNC As String = "WB_LAST_SYNC_UTC"
Private Const PROP_GLOBAL_SRC_SIG As String = "GLB_SRC_SIGNATURE"
Private Const TYPE_LIST_DELIMITER As String = ";"
Private Const REPO_ADDIN_RELATIVE_PATH As String = "Global_XLAM\Global_CNL.xlam"
Private Const REPO_ADDIN_SRC_RELATIVE_PATH As String = "Global_XLAM\src"
Private Const MODULE_RELOAD_MANAGER As String = "ReloadManager"
Private Const MODULE_RELOAD_MANAGER_BOOTSTRAP As String = "ReloadManagerBootstrap"
Private Const TYPE_PICKER_FORM_NAME As String = "RM_TypePickerForm"
Private Const TYPE_PICKER_LIST_NAME As String = "lstEntries"
Private Const TYPE_PICKER_OK_NAME As String = "btnOK"
Private Const TYPE_PICKER_CANCEL_NAME As String = "btnCancel"
Private Const TYPE_PICKER_SYNC_DRIFT_NAME As String = "btnSyncOnDrift"
Private Const TYPE_PICKER_FORCE_SYNC_NAME As String = "btnForceSync"
Private Const TYPE_PICKER_INFO_NAME As String = "lblInfo"
Private Const TYPE_PICKER_TAG_OK As String = "OK"
Private Const TYPE_PICKER_TAG_CANCEL As String = "0"
Private Const TYPE_PICKER_TAG_SYNC_DRIFT As String = "SYNC_DRIFT"
Private Const TYPE_PICKER_TAG_FORCE_SYNC As String = "FORCE_SYNC"
Private Const MAX_FOLDER_SCAN_DEPTH As Long = 64
Private Const FILE_ATTR_REPARSE_POINT As Long = 1024

' Win32 constants for making the dynamic type picker UserForm resizable.
Private Const GWL_STYLE As Long = -16
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_FRAMECHANGED As Long = &H20

#If VBA7 Then
    Private Declare PtrSafe Function FindWindowW Lib "user32" (ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr) As LongPtr
    Private Declare PtrSafe Function GetWindowLongPtrW Lib "user32" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLongPtrW Lib "user32" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As LongPtr, _
                                                                  ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, _
                                                                  ByVal uFlags As Long) As Long
#Else
    Private Declare Function FindWindowW Lib "user32" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
    Private Declare Function GetWindowLongPtrW Lib "user32" Alias "GetWindowLongW" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLongPtrW Lib "user32" Alias "SetWindowLongW" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
                                                         ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, _
                                                         ByVal uFlags As Long) As Long
#End If

'================== TYPE PICKER UI TUNING ==================
' Checklist indentation and general row rendering.
' - Indentation controls horizontal nesting offset for folder/file rows.
' - Base/Right/Top paddings control drawing space inside the entries frame.
' - Min control width prevents deep rows from becoming too narrow to read.
' - Scroll extra adds a small trailing scroll region so the last row is not clipped.
' - Folder label nudge lets label text align visually with checkbox captions.
Private Const TP_INDENT_PIXELS_PER_LEVEL As Double = 12
Private Const TP_LIST_BASE_LEFT As Double = 4
Private Const TP_LIST_RIGHT_PADDING As Double = 16
Private Const TP_LIST_TOP_PADDING As Double = 3
Private Const TP_LIST_MIN_CONTROL_WIDTH As Double = 200
Private Const TP_LIST_SCROLL_EXTRA As Double = 20
Private Const TP_FOLDER_LABEL_LEFT_NUDGE As Double = 2
Private Const TP_LIST_FRAME_CAPTION As String = ""
' UI font family for picker text. Recommended options: Segoe UI, Tahoma, Verdana, Calibri.
Private Const TP_UI_FONT_NAME As String = "Segoe UI"

' File (checkbox) row styling.
' Row formulas:
' - file row control height = TP_FILE_FONT_HEIGHT + TP_FILE_FONT_ADDITIONAL_HEIGHT
' - file row step = file row control height + TP_FILE_ROW_ADDITIONAL_SPACING
' - checkbox glyph font = TP_FILE_FONT_HEIGHT - TP_FILE_CHECKBOX_FONT_REDUCTION_POINTS
' - checkbox bottom = row bottom - TP_FILE_ROW_ADDITIONAL_SPACING - TP_FILE_CHECKBOX_BOTTOM_OFFSET
' Group gaps provide visual separation when transitioning between folder rows and file groups.
Private Const TP_FILE_FONT_HEIGHT As Double = 9
Private Const TP_FILE_FONT_ADDITIONAL_HEIGHT As Double = 2
Private Const TP_FILE_ROW_ADDITIONAL_SPACING As Double = 3
Private Const TP_FILE_CHECKBOX_FONT_REDUCTION_POINTS As Double = 1
' Positive values lift the checkbox relative to text-bottom alignment; negative values drop it.
Private Const TP_FILE_CHECKBOX_BOTTOM_OFFSET As Double = 1
Private Const TP_GAP_BEFORE_FILE_GROUP As Double = 3
Private Const TP_GAP_AFTER_FILE_GROUP As Double = 3
Private Const TP_GAP_BEFORE_FOLDER_L0 As Double = 6

' Temporary troubleshooting mode: writes rendered row metrics to a plaintext report in %TEMP%.
Private Const TP_DEBUG_ENABLE_ROW_REPORT As Boolean = True

' Folder row styling by hierarchy depth.
' Layer 0 = top-level folder, layer 1 = child folder, etc.
' Folder formulas per layer:
' - folder control height = TP_FOLDER_FONT_HEIGHT_* + TP_FOLDER_FONT_ADDITIONAL_HEIGHT_*
' - folder row step = control height + TP_FOLDER_ROW_ADDITIONAL_SPACING_*
' Use these sets to make parent folders more prominent than deeper levels.
Private Const TP_FOLDER_FONT_HEIGHT_L0 As Double = 11
Private Const TP_FOLDER_FONT_HEIGHT_L1 As Double = 11
Private Const TP_FOLDER_FONT_HEIGHT_L2 As Double = 11
Private Const TP_FOLDER_FONT_HEIGHT_L3 As Double = 11
Private Const TP_FOLDER_FONT_HEIGHT_DEFAULT As Double = 11

Private Const TP_FOLDER_FONT_ADDITIONAL_HEIGHT_L0 As Double = 3
Private Const TP_FOLDER_FONT_ADDITIONAL_HEIGHT_L1 As Double = 3
Private Const TP_FOLDER_FONT_ADDITIONAL_HEIGHT_L2 As Double = 3
Private Const TP_FOLDER_FONT_ADDITIONAL_HEIGHT_L3 As Double = 3
Private Const TP_FOLDER_FONT_ADDITIONAL_HEIGHT_DEFAULT As Double = 3

Private Const TP_FOLDER_ROW_ADDITIONAL_SPACING_L0 As Double = 0
Private Const TP_FOLDER_ROW_ADDITIONAL_SPACING_L1 As Double = 0
Private Const TP_FOLDER_ROW_ADDITIONAL_SPACING_L2 As Double = 0
Private Const TP_FOLDER_ROW_ADDITIONAL_SPACING_L3 As Double = 0
Private Const TP_FOLDER_ROW_ADDITIONAL_SPACING_DEFAULT As Double = 0

' Runtime auto-size and manual resize behavior.
' Layout levers:
' - Margins and button gaps define frame/button padding around the list.
' - Button width/height and gap define the action row footprint.
' - List horizontal/vertical padding enlarges content fit area around measured text/rows.
' - TP_AUTO_ROW_HEIGHT and TP_AUTO_MAX_VISIBLE_ROWS control initial auto-size height estimate.
' - TP_FORM_USABLE_*_RATIO limits the default form size to a fraction of Excel usable area.
' - TP_SEED_LIST_* controls designer-time seed size (replaces hardcoded values).
' - Set TP_FORCE_LIST_* > 0 to bypass autosize and directly control popup/list size.
Private Const TP_LAYOUT_LEFT_MARGIN As Double = 10
Private Const TP_LAYOUT_RIGHT_MARGIN As Double = 10
Private Const TP_LAYOUT_TOP_MARGIN As Double = 8
Private Const TP_LAYOUT_BUTTON_GAP_TOP As Double = 10
Private Const TP_LAYOUT_BUTTON_GAP As Double = 8
Private Const TP_LAYOUT_BUTTON_WIDTH As Double = 48
Private Const TP_LAYOUT_SYNC_BUTTON_WIDTH As Double = 64
Private Const TP_LAYOUT_BUTTON_HEIGHT As Double = 24
Private Const TP_LAYOUT_BOTTOM_MARGIN As Double = TP_LAYOUT_BUTTON_HEIGHT + TP_LAYOUT_BUTTON_GAP + TP_LAYOUT_BUTTON_GAP_TOP
Private Const TP_SEED_LIST_WIDTH As Double = 420
Private Const TP_SEED_LIST_HEIGHT As Double = 720
Private Const TP_LIST_HORIZONTAL_PADDING As Double = 22
Private Const TP_LIST_VERTICAL_PADDING As Double = 8
Private Const TP_AUTO_ROW_HEIGHT As Double = 16
Private Const TP_AUTO_MAX_VISIBLE_ROWS As Long = 24
Private Const TP_FORM_USABLE_WIDTH_RATIO As Double = 0.9
Private Const TP_FORM_USABLE_HEIGHT_RATIO As Double = 0.9
Private Const TP_FORCE_LIST_WIDTH As Double = 0
Private Const TP_FORCE_LIST_HEIGHT As Double = 0

' Internal representation for tree depth in displayLines. Keep fixed to avoid dual user levers.
Private Const TP_INDENT_SPACES_PER_LEVEL_INTERNAL As Long = 2

Private mStartupSyncComplete As Boolean
Private mGlobalAddInPublishedThisSession As Boolean

'================== PUBLIC ENTRY POINTS ==================

' Auto-run when the XLAM loads: check currently open workbooks for source drift.
' Quiet when no type IDs are assigned or when assigned sources are already current.
Public Sub Auto_Open()
    EnsureGlobalAddInStartupSync
    RunOpenTimeDriftCheckAllOpenWorkbooks
End Sub

' Optional cleanup hook when the XLAM unloads.
Public Sub Auto_Close()
    CancelQueuedReloadManagerRefreshSafe

    If Not mGlobalAddInPublishedThisSession Then Exit Sub

    On Error Resume Next
    If Not ThisWorkbook.Saved Then ThisWorkbook.Save
    On Error GoTo 0

    SyncInstalledAddInToRepoBackup False
End Sub

' Manual one-click publisher for Global_CNL maintenance.
' Flow: repo BAS -> installed XLAM, queue ReloadManager refresh, installed XLAM -> repo backup.
Public Sub PublishGlobalAddInNow()
    Dim bootstrapQueued As Boolean
    Dim modulesPublished As Boolean

    On Error GoTo Fail

    modulesPublished = PublishInstalledAddInModulesFromRepo(False, False)

    If modulesPublished Then
        bootstrapQueued = QueueReloadManagerRefreshSafe()
    End If

    If modulesPublished Then
        SyncInstalledAddInToRepoBackup False
    End If
    mStartupSyncComplete = True

    If Not modulesPublished Then
        MsgBox "Global_CNL publish skipped. No module changes were detected.", vbInformation, SYNC_DIALOG_TITLE
    ElseIf bootstrapQueued Then
        MsgBox "Global_CNL publish completed." & vbCrLf & _
               "- Repo BAS published to installed XLAM" & vbCrLf & _
               "- ReloadManager refresh queued" & vbCrLf & _
               "- Installed XLAM synced to repo backup", vbInformation, SYNC_DIALOG_TITLE
    Else
        MsgBox "Global_CNL publish completed with warning." & vbCrLf & _
               "- Repo BAS published to installed XLAM" & vbCrLf & _
               "- Unable to queue ReloadManager refresh" & vbCrLf & _
               "- Installed XLAM synced to repo backup", vbExclamation, SYNC_DIALOG_TITLE
    End If
    Exit Sub

Fail:
    MsgBox "Global_CNL publish failed." & vbCrLf & _
           "Reason: " & Err.Description, vbExclamation, SYNC_DIALOG_TITLE
End Sub

Private Sub EnsureGlobalAddInStartupSync()
    Dim modulesPublished As Boolean

    If mStartupSyncComplete Then Exit Sub
    mStartupSyncComplete = True

    ' Repo BAS files are source of truth for installed add-in modules.
    modulesPublished = PublishInstalledAddInModulesFromRepo(False, False)

    ' Queue deferred ReloadManager refresh only after publish to avoid queue-state loss.
    If modulesPublished Then
        Call QueueReloadManagerRefreshSafe
    End If

    ' Installed XLAM is source of truth for repo backup XLAM.
    If modulesPublished Then
        SyncInstalledAddInToRepoBackup False
    End If
End Sub

Private Function PublishInstalledAddInModulesFromRepo(Optional ByVal Force As Boolean = False, _
                                                      Optional ByVal showStatus As Boolean = False) As Boolean
    Dim srcPath As String
    Dim reason As String
    Dim repoSig As String
    Dim wbSig As String
    Dim fso As Object

    On Error GoTo SafeExit

    srcPath = EnsureTrailingPathSeparator(REPO_ROOT & Application.PathSeparator & _
              Replace$(REPO_ADDIN_SRC_RELATIVE_PATH, "\", Application.PathSeparator))
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Len(srcPath) = 0 Then Exit Function
    If Not fso.FolderExists(srcPath) Then Exit Function

    If Not CanAccessVBProject(ThisWorkbook, reason) Then
        If showStatus Then MsgBox reason, vbExclamation, SYNC_DIALOG_TITLE
        Exit Function
    End If

    repoSig = ComputeFolderSignature(srcPath, "*.bas")
    If Len(repoSig) = 0 Then Exit Function

    wbSig = GetDocProp(ThisWorkbook, PROP_GLOBAL_SRC_SIG)
    If Not Force Then
        If StrComp(repoSig, wbSig, vbBinaryCompare) = 0 Then Exit Function
    End If

    ' Bootstrap can be refreshed directly here; only ReloadManager must be deferred.
    ImportBasFolder ThisWorkbook, srcPath, MODULE_RELOAD_MANAGER

    SetDocProp ThisWorkbook, PROP_GLOBAL_SRC_SIG, repoSig
    SetDocProp ThisWorkbook, PROP_LAST_SYNC, Format$(NowUTC(), "yyyy-mm-dd\THH:NN:SS\Z")
    PublishInstalledAddInModulesFromRepo = True
    mGlobalAddInPublishedThisSession = True

    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo SafeExit

    If showStatus Then
        MsgBox "Global_CNL.xlam modules were refreshed from repo source.", vbInformation, SYNC_DIALOG_TITLE
    End If

SafeExit:
    If Err.Number <> 0 And showStatus Then
        MsgBox "Unable to publish repo modules to Global_CNL.xlam." & vbCrLf & _
               "Reason: " & Err.Description, vbExclamation, SYNC_DIALOG_TITLE
    End If
    On Error GoTo 0
End Function

Private Sub SyncInstalledAddInToRepoBackup(Optional ByVal showStatus As Boolean = False)
    Dim installedPath As String
    Dim repoPath As String
    Dim parentPath As String
    Dim fso As Object

    On Error GoTo SafeExit

    installedPath = ThisWorkbook.FullName
    repoPath = REPO_ROOT & Application.PathSeparator & _
               Replace$(REPO_ADDIN_RELATIVE_PATH, "\", Application.PathSeparator)

    If Len(installedPath) = 0 Or Len(repoPath) = 0 Then Exit Sub
    If StrComp(installedPath, repoPath, vbTextCompare) = 0 Then Exit Sub
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(installedPath) Then Exit Sub

    If Not ShouldCopyInstalledAddInToRepo(installedPath, repoPath, fso) Then Exit Sub

    parentPath = fso.GetParentFolderName(repoPath)
    If Len(parentPath) > 0 Then
        If Not fso.FolderExists(parentPath) Then
            fso.CreateFolder parentPath
        End If
    End If
    fso.CopyFile installedPath, repoPath, True

    If showStatus Then
        MsgBox "Global_CNL.xlam backup refreshed in repo.", vbInformation, SYNC_DIALOG_TITLE
    End If

SafeExit:
    If Err.Number <> 0 And showStatus Then
        MsgBox "Unable to refresh Global_CNL.xlam backup in repo." & vbCrLf & _
               "Reason: " & Err.Description, vbExclamation, SYNC_DIALOG_TITLE
    End If
    On Error GoTo 0
End Sub

Private Function ShouldCopyInstalledAddInToRepo(ByVal installedPath As String, _
                                                ByVal repoPath As String, _
                                                ByVal fso As Object) As Boolean
    Dim installedStamp As Date
    Dim repoStamp As Date
    Dim installedSize As Double
    Dim repoSize As Double

    On Error GoTo SafeExit

    If fso Is Nothing Then Exit Function
    If Not fso.FileExists(installedPath) Then Exit Function
    If Not fso.FileExists(repoPath) Then
        ShouldCopyInstalledAddInToRepo = True
        Exit Function
    End If

    installedStamp = FileDateTime(installedPath)
    repoStamp = FileDateTime(repoPath)
    installedSize = FileLen(installedPath)
    repoSize = FileLen(repoPath)

    ' Installed copy is authoritative for repo backup: refresh whenever metadata differs.
    ShouldCopyInstalledAddInToRepo = (installedStamp <> repoStamp) Or (installedSize <> repoSize)

SafeExit:
    On Error GoTo 0
End Function

Private Function QueueReloadManagerRefreshSafe() As Boolean
    Dim bootstrapModuleName As String

    bootstrapModuleName = ResolveBootstrapModuleName(ThisWorkbook)
    If Len(bootstrapModuleName) = 0 Then Exit Function

    On Error Resume Next
    Application.Run bootstrapModuleName & ".QueueReloadManagerRefresh", _
                    REPO_ROOT, REPO_ADDIN_SRC_RELATIVE_PATH, bootstrapModuleName
    If Err.Number <> 0 Then
        ' Backward compatibility for older bootstrap signatures (2 args).
        Err.Clear
        Application.Run bootstrapModuleName & ".QueueReloadManagerRefresh", REPO_ROOT, REPO_ADDIN_SRC_RELATIVE_PATH
    End If
    QueueReloadManagerRefreshSafe = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Sub CancelQueuedReloadManagerRefreshSafe()
    Dim bootstrapModuleName As String

    bootstrapModuleName = ResolveBootstrapModuleName(ThisWorkbook)
    If Len(bootstrapModuleName) = 0 Then Exit Sub

    On Error Resume Next
    Application.Run bootstrapModuleName & ".CancelQueuedReloadManagerRefresh"
    Err.Clear
    On Error GoTo 0
End Sub

Private Function ResolveBootstrapModuleName(ByVal wb As Workbook) As String
    Dim vbComp As Object
    Dim variantComp As Object

    On Error GoTo Fallback

    For Each vbComp In wb.VBProject.VBComponents
        If vbComp.Type = 1 Then
            If StrComp(vbComp.Name, MODULE_RELOAD_MANAGER_BOOTSTRAP, vbTextCompare) = 0 Then
                ResolveBootstrapModuleName = vbComp.Name
                Exit Function
            End If
        End If
    Next vbComp

    For Each vbComp In wb.VBProject.VBComponents
        If vbComp.Type = 1 Then
            If IsAutoRenamedModuleName(vbComp.Name, MODULE_RELOAD_MANAGER_BOOTSTRAP) Then
                Set variantComp = vbComp
                Exit For
            End If
        End If
    Next vbComp

    If Not variantComp Is Nothing Then
        On Error Resume Next
        variantComp.Name = MODULE_RELOAD_MANAGER_BOOTSTRAP
        If Err.Number = 0 Then
            ResolveBootstrapModuleName = MODULE_RELOAD_MANAGER_BOOTSTRAP
        Else
            Err.Clear
            ResolveBootstrapModuleName = variantComp.Name
        End If
        On Error GoTo Fallback
    End If

Fallback:
    If Len(ResolveBootstrapModuleName) = 0 Then
        ResolveBootstrapModuleName = MODULE_RELOAD_MANAGER_BOOTSTRAP
    End If
End Function

Private Function IsAutoRenamedModuleName(ByVal moduleName As String, ByVal baseName As String) As Boolean
    Dim suffix As String
    Dim i As Long

    If StrComp(moduleName, baseName, vbTextCompare) = 0 Then
        IsAutoRenamedModuleName = True
        Exit Function
    End If

    If Len(moduleName) <= Len(baseName) Then Exit Function
    If StrComp(Left$(moduleName, Len(baseName)), baseName, vbTextCompare) <> 0 Then Exit Function

    suffix = Mid$(moduleName, Len(baseName) + 1)
    If Len(suffix) = 0 Then Exit Function

    For i = 1 To Len(suffix)
        If Mid$(suffix, i, 1) < "0" Or Mid$(suffix, i, 1) > "9" Then
            Exit Function
        End If
    Next i

    IsAutoRenamedModuleName = True
End Function

' Call from ThisWorkbook.App_WorkbookOpen(ByVal Wb As Workbook).
Public Sub HandleWorkbookOpen(ByVal wb As Workbook)
    CheckWorkbookForOpenTimeDrift wb
End Sub

' One-time pass for all currently open non-add-in workbooks.
Public Sub RunOpenTimeDriftCheckAllOpenWorkbooks()
    Dim wb As Workbook

    On Error GoTo SafeExit

    EnsureGlobalAddInStartupSync

    For Each wb In Application.Workbooks
        CheckWorkbookForOpenTimeDrift wb
    Next wb

SafeExit:
    On Error GoTo 0
End Sub

Private Sub CheckWorkbookForOpenTimeDrift(ByVal wb As Workbook)
    Dim typeSummary As String
    Dim srcPaths As Collection
    Dim changedSrcPaths As Collection
    Dim driftTypeSummary As String
    Dim repoSig As String
    Dim wbSig As String
    Dim typeSigMapText As String
    Dim reason As String

    On Error GoTo SafeExit

    If wb Is Nothing Then Exit Sub
    If IsAddInWorkbook(wb) Then Exit Sub

    ' Keep open-time checks quiet unless actionable drift is found.
    If Not PrepareSyncContext(wb, typeSummary, srcPaths, changedSrcPaths, driftTypeSummary, _
                              repoSig, wbSig, typeSigMapText, reason) Then Exit Sub
    If StrComp(repoSig, wbSig, vbBinaryCompare) = 0 Then Exit Sub
    If Len(driftTypeSummary) = 0 Then Exit Sub

    If ConfirmDriftImport(wb, driftTypeSummary) Then
        SyncWorkbook wb, False, False
    End If

SafeExit:
    On Error GoTo 0
End Sub

' Single visible macro entry point for users.
Public Sub ManageActiveWorkbookTypes()
    ShowReloadManagerCommandCenter
End Sub

' UserForm picker preview for visual validation.
Public Sub PreviewWorkbookTypePickerUserForm()
    Dim wb As Workbook
    Dim typeMap As Object
    Dim availableTypeIds As Collection
    Dim displayLines As Collection
    Dim ownerTypeIds As Collection
    Dim selectedIndex As Long
    Dim reason As String

    Set wb = GetForegroundWorkbook()
    If wb Is Nothing Then
        MsgBox "No workbook is currently available.", vbExclamation, SYNC_DIALOG_TITLE
        Exit Sub
    End If

    If IsAddInWorkbook(wb) Then
        MsgBox "Select a regular workbook before running picker preview.", vbExclamation, SYNC_DIALOG_TITLE
        Exit Sub
    End If

    If Not CanAccessVBProject(ThisWorkbook, reason) Then
        MsgBox reason, vbExclamation, SYNC_DIALOG_TITLE
        Exit Sub
    End If

    Set availableTypeIds = GetAvailableTypeIdsCollection()
    If availableTypeIds.Count = 0 Then
        MsgBox "No valid type IDs were found in local VBA sources." & vbCrLf & _
               MSG_REQUEST_UPDATE, vbInformation, SYNC_DIALOG_TITLE
        Exit Sub
    End If

    Set typeMap = GetWorkbookTypeMap(wb)
    Set displayLines = New Collection
    Set ownerTypeIds = New Collection
    BuildTypePickerTreeLines availableTypeIds, typeMap, displayLines, ownerTypeIds

    If displayLines.Count = 0 Then
        MsgBox "No .bas entries were found for available type IDs.", vbInformation, SYNC_DIALOG_TITLE
        Exit Sub
    End If

    selectedIndex = ShowTypePickerPreviewUserForm(displayLines, ownerTypeIds, "Workbook Type Picker")
    If selectedIndex <= 0 Then Exit Sub
    If selectedIndex > ownerTypeIds.Count Then Exit Sub

    MsgBox "Selected preview entry:" & vbCrLf & _
           CStr(displayLines(selectedIndex)) & vbCrLf & vbCrLf & _
           "Owning Type ID:" & vbCrLf & _
           CStr(ownerTypeIds(selectedIndex)), vbInformation, SYNC_DIALOG_TITLE
End Sub

Private Sub ShowReloadManagerCommandCenter()
    Dim reason As String

    If ShowReloadManagerUserFormCommandCenter(reason) Then Exit Sub

    If Len(reason) > 0 Then
        MsgBox "UserForm interface was unavailable. Falling back to legacy menu." & vbCrLf & _
               "Reason: " & reason, vbExclamation, SYNC_DIALOG_TITLE
    End If

    ShowReloadManagerCommandCenterLegacy
End Sub

Private Function ShowReloadManagerUserFormCommandCenter(ByRef reason As String) As Boolean
    Dim wb As Workbook
    Dim availableTypeIds As Collection
    Dim currentTypeMap As Object
    Dim selectedTypeIds As Collection
    Dim selectedTypeMap As Object
    Dim typeId As Variant
    Dim normalizedTypeId As String
    Dim addedCount As Long
    Dim removedCount As Long
    Dim pickerAction As String

    On Error GoTo Fail

    Set wb = GetForegroundWorkbook()
    If wb Is Nothing Then
        reason = "No workbook is currently available."
        Exit Function
    End If

    If IsAddInWorkbook(wb) Then
        reason = "Select a regular workbook before managing workbook types."
        Exit Function
    End If

    Set availableTypeIds = GetAvailableTypeIdsCollection()
    If availableTypeIds.Count = 0 Then
        MsgBox "No valid type IDs were found in local VBA sources." & vbCrLf & _
               MSG_REQUEST_UPDATE, vbInformation, SYNC_DIALOG_TITLE
        ShowReloadManagerUserFormCommandCenter = True
        Exit Function
    End If

    Set currentTypeMap = GetWorkbookTypeMap(wb)
    Set selectedTypeIds = PromptManageWorkbookTypeIdsUserForm(availableTypeIds, currentTypeMap, _
                                                               "ReloadManager - Workbook Types", _
                                                               "Check workbook types, then click Apply, Sync on drift, or Force sync.", _
                                                               pickerAction)

    If selectedTypeIds Is Nothing Then
        ShowReloadManagerUserFormCommandCenter = True
        Exit Function
    End If

    If StrComp(pickerAction, TYPE_PICKER_TAG_CANCEL, vbTextCompare) = 0 Then
        ShowReloadManagerUserFormCommandCenter = True
        Exit Function
    End If

    If selectedTypeIds.Count = 0 And currentTypeMap.count = 0 Then
        If MsgBox("No workbook types were selected. Apply an empty assignment and continue?", _
                  vbYesNo + vbQuestion, SYNC_DIALOG_TITLE) <> vbYes Then
            ShowReloadManagerUserFormCommandCenter = True
            Exit Function
        End If
    End If

    If selectedTypeIds.Count = 0 And currentTypeMap.count > 0 Then
        If MsgBox("No workbook types are checked. Clear all assigned workbook types?", _
                  vbYesNo + vbQuestion, SYNC_DIALOG_TITLE) <> vbYes Then
            ShowReloadManagerUserFormCommandCenter = True
            Exit Function
        End If
    End If

    Set selectedTypeMap = CreateObject("Scripting.Dictionary")
    selectedTypeMap.CompareMode = vbTextCompare
    For Each typeId In selectedTypeIds
        normalizedTypeId = NormalizeTypeId(CStr(typeId))
        If Len(normalizedTypeId) > 0 Then
            selectedTypeMap(normalizedTypeId) = normalizedTypeId
        End If
    Next typeId

    For Each typeId In selectedTypeMap.Keys
        If Not currentTypeMap.Exists(CStr(typeId)) Then addedCount = addedCount + 1
    Next typeId

    For Each typeId In currentTypeMap.Keys
        If Not selectedTypeMap.Exists(CStr(typeId)) Then removedCount = removedCount + 1
    Next typeId

    If addedCount > 0 Or removedCount > 0 Then
        SaveWorkbookTypeMap wb, selectedTypeMap
        MsgBox "Workbook type assignments updated." & vbCrLf & _
               "Added: " & CStr(addedCount) & vbCrLf & _
               "Removed: " & CStr(removedCount), vbInformation, SYNC_DIALOG_TITLE
    Else
        MsgBox "No workbook type changes were applied.", vbInformation, SYNC_DIALOG_TITLE
    End If

    Select Case UCase$(Trim$(pickerAction))
        Case TYPE_PICKER_TAG_SYNC_DRIFT
            SyncWorkbook wb, False, True
        Case TYPE_PICKER_TAG_FORCE_SYNC
            SyncWorkbook wb, True, False
    End Select

    ShowReloadManagerUserFormCommandCenter = True
    Exit Function

Fail:
    reason = "UserForm workflow failed. Error " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Sub ShowReloadManagerCommandCenterLegacy()
    Dim choice As String
    Dim menuText As String

    menuText = "Workbook Type Manager" & vbCrLf & String$(24, "=") & vbCrLf & vbCrLf & _
               "1 - View/Adjust assigned types" & vbCrLf & _
               "2 - Sync workbook (prompt on drift)" & vbCrLf & _
               "3 - Force sync now" & vbCrLf & _
               "0 - Exit"

    Do
        choice = Trim$(InputBox(menuText, "Workbook Type Manager", "1"))
        If Len(choice) = 0 Or choice = "0" Then Exit Do

        Select Case choice
            Case "1"
                ShowActiveWorkbookTypeIds
                ManageActiveWorkbookTypesPrompt
            Case "2"
                RunSyncForForegroundWorkbook False, True
            Case "3"
                RunSyncForForegroundWorkbook True, False
            Case Else
                MsgBox "Invalid selection. Enter 0-3.", vbExclamation, SYNC_DIALOG_TITLE
        End Select
    Loop
End Sub

Private Sub RunSyncForForegroundWorkbook(Optional ByVal Force As Boolean = False, _
                                         Optional ByVal PromptOnDrift As Boolean = True)
    Dim wb As Workbook
    If Not TryGetForegroundWorkbookForTypes(wb, _
                                            "No workbook is currently available for sync.", _
                                            "Select a regular workbook before syncing.") Then Exit Sub

    SyncWorkbook wb, Force, PromptOnDrift
End Sub

Private Function TryGetForegroundWorkbookForTypes(ByRef wb As Workbook, _
                                                  ByVal noWorkbookMessage As String, _
                                                  ByVal addInMessage As String) As Boolean
    Set wb = GetForegroundWorkbook()

    If wb Is Nothing Then
        If Len(noWorkbookMessage) > 0 Then
            MsgBox noWorkbookMessage, vbExclamation, SYNC_DIALOG_TITLE
        End If
        Exit Function
    End If

    If IsAddInWorkbook(wb) Then
        If Len(addInMessage) > 0 Then
            MsgBox addInMessage, vbExclamation, SYNC_DIALOG_TITLE
        End If
        Exit Function
    End If

    TryGetForegroundWorkbookForTypes = True
End Function

Private Sub ShowActiveWorkbookTypeIds()
    Dim wb As Workbook
    Dim typeMap As Object
    Dim summary As String
    Dim sorted() As String

    If Not TryGetForegroundWorkbookForTypes(wb, _
                                            "No workbook is currently available.", _
                                            "Select a regular workbook before viewing assigned types.") Then Exit Sub

    Set typeMap = GetWorkbookTypeMap(wb)
    If Not typeMap Is Nothing Then
        If typeMap.count > 0 Then
            sorted = GetSortedTypeIds(typeMap)
            summary = JoinStringArray(sorted, vbCrLf)
        End If
    End If

    If Len(summary) = 0 Then
        MsgBox "No workbook types are currently assigned.", vbInformation, SYNC_DIALOG_TITLE
    Else
        MsgBox "Assigned workbook types:" & vbCrLf & vbCrLf & summary, vbInformation, SYNC_DIALOG_TITLE
    End If
End Sub

Private Sub ManageActiveWorkbookTypesPrompt()
    Dim choice As String
    Dim menuText As String

    menuText = "Manage Workbook Types" & vbCrLf & String$(24, "=") & vbCrLf & vbCrLf & _
               "1 - Add type(s)" & vbCrLf & _
               "2 - Remove type(s)" & vbCrLf & _
               "3 - Clear all assigned types" & vbCrLf & _
               "0 - Exit"

    Do
        choice = Trim$(InputBox(menuText, "Manage Workbook Types", "1"))
        If Len(choice) = 0 Or choice = "0" Then Exit Do

        Select Case choice
            Case "1"
                AddActiveWorkbookTypeId
            Case "2"
                RemoveActiveWorkbookTypeId
            Case "3"
                ClearActiveWorkbookTypeIds
            Case Else
                MsgBox "Invalid selection. Enter 0-3.", vbExclamation, SYNC_DIALOG_TITLE
        End Select
    Loop
End Sub

Private Sub AddActiveWorkbookTypeId()
    Dim wb As Workbook
    Dim typeMap As Object
    Dim candidateTypeIds As Collection
    Dim selectedTypeIds As Collection
    Dim typeId As Variant
    Dim normalizedTypeId As String
    Dim availableTypeIds As Collection
    Dim addedCount As Long

    If Not TryGetForegroundWorkbookForTypes(wb, _
                                            "No workbook is currently available to update types.", _
                                            "Select a regular workbook before assigning types.") Then Exit Sub

    Set typeMap = GetWorkbookTypeMap(wb)
    Set availableTypeIds = GetAvailableTypeIdsCollection()

    If availableTypeIds.Count = 0 Then
        MsgBox "No valid type IDs were found in local VBA sources." & vbCrLf & _
               MSG_REQUEST_UPDATE, vbInformation, SYNC_DIALOG_TITLE
        Exit Sub
    End If

    Set candidateTypeIds = FilterUnassignedTypeIds(availableTypeIds, typeMap)
    If candidateTypeIds.Count = 0 Then
        MsgBox "All eligible workbook types are already assigned.", vbInformation, SYNC_DIALOG_TITLE
        Exit Sub
    End If

    Set selectedTypeIds = PromptSelectTypeIdsUserForm(candidateTypeIds, _
                                                      "Add Workbook Types", _
                                                      "Check one or more .bas entries to add workbook types.")
    If selectedTypeIds.Count = 0 Then Exit Sub

    For Each typeId In selectedTypeIds
        normalizedTypeId = NormalizeTypeId(CStr(typeId))
        If Len(normalizedTypeId) > 0 Then
            If Not typeMap.Exists(normalizedTypeId) Then
                typeMap(normalizedTypeId) = normalizedTypeId
                addedCount = addedCount + 1
            End If
        End If
    Next typeId

    If addedCount = 0 Then
        MsgBox "No new types were added.", vbInformation, SYNC_DIALOG_TITLE
        Exit Sub
    End If

    SaveWorkbookTypeMap wb, typeMap

    MsgBox CStr(addedCount) & " workbook type(s) added.", vbInformation, SYNC_DIALOG_TITLE
End Sub

Private Sub RemoveActiveWorkbookTypeId()
    Dim wb As Workbook
    Dim typeMap As Object
    Dim selectedTypeIds As Collection
    Dim typeId As Variant
    Dim normalizedTypeId As String
    Dim removedCount As Long

    If Not TryGetForegroundWorkbookForTypes(wb, _
                                            "No workbook is currently available to update types.", _
                                            "Select a regular workbook before updating types.") Then Exit Sub

    Set typeMap = GetWorkbookTypeMap(wb)
    If typeMap.count = 0 Then
        MsgBox "No workbook types are currently assigned.", vbInformation, SYNC_DIALOG_TITLE
        Exit Sub
    End If

    Set selectedTypeIds = PromptSelectTypeIdsUserForm(GetSortedTypeIdCollectionFromMap(typeMap), _
                                                      "Remove Workbook Types", _
                                                      "Check one or more .bas entries to remove workbook types.")
    If selectedTypeIds.Count = 0 Then Exit Sub

    For Each typeId In selectedTypeIds
        normalizedTypeId = NormalizeTypeId(CStr(typeId))
        If Len(normalizedTypeId) > 0 Then
            If typeMap.Exists(normalizedTypeId) Then
                typeMap.Remove normalizedTypeId
                removedCount = removedCount + 1
            End If
        End If
    Next typeId

    If removedCount = 0 Then
        MsgBox "No assigned types were removed.", vbInformation, SYNC_DIALOG_TITLE
        Exit Sub
    End If

    SaveWorkbookTypeMap wb, typeMap
    MsgBox CStr(removedCount) & " workbook type(s) removed.", vbInformation, SYNC_DIALOG_TITLE
End Sub

Private Function GetSortedTypeIdCollectionFromMap(ByVal typeMap As Object) As Collection
    Dim sortedTypes() As String
    Dim i As Long
    Dim results As New Collection

    If typeMap Is Nothing Then
        Set GetSortedTypeIdCollectionFromMap = results
        Exit Function
    End If

    If typeMap.count = 0 Then
        Set GetSortedTypeIdCollectionFromMap = results
        Exit Function
    End If

    sortedTypes = GetSortedTypeIds(typeMap)
    For i = LBound(sortedTypes) To UBound(sortedTypes)
        results.Add CStr(sortedTypes(i))
    Next i

    Set GetSortedTypeIdCollectionFromMap = results
End Function

Private Function FilterUnassignedTypeIds(ByVal availableTypeIds As Collection, ByVal assignedTypeMap As Object) As Collection
    Dim results As New Collection
    Dim typeId As Variant
    Dim normalizedTypeId As String

    For Each typeId In availableTypeIds
        normalizedTypeId = NormalizeTypeId(CStr(typeId))
        If Len(normalizedTypeId) > 0 Then
            If assignedTypeMap Is Nothing Then
                results.Add normalizedTypeId
            ElseIf Not assignedTypeMap.Exists(normalizedTypeId) Then
                results.Add normalizedTypeId
            End If
        End If
    Next typeId

    Set FilterUnassignedTypeIds = results
End Function

Private Sub ClearActiveWorkbookTypeIds()
    Dim wb As Workbook
    Dim typeMap As Object

    If Not TryGetForegroundWorkbookForTypes(wb, _
                                            "No workbook is currently available to clear types.", _
                                            "Select a regular workbook before clearing types.") Then Exit Sub

    Set typeMap = GetWorkbookTypeMap(wb)
    If typeMap.count = 0 Then
        MsgBox "No workbook types are currently assigned.", vbInformation, SYNC_DIALOG_TITLE
        Exit Sub
    End If

    If MsgBox("Clear all assigned workbook types?", vbYesNo + vbQuestion, SYNC_DIALOG_TITLE) <> vbYes Then
        Exit Sub
    End If

    Set typeMap = CreateObject("Scripting.Dictionary")
    typeMap.CompareMode = vbTextCompare
    SaveWorkbookTypeMap wb, typeMap

    MsgBox "All workbook types were cleared.", vbInformation, SYNC_DIALOG_TITLE
End Sub

' Sync a specific workbook instance (safe with multiple workbooks open).
Private Sub SyncWorkbook(ByVal wb As Workbook, _
                         Optional ByVal Force As Boolean = False, _
                         Optional ByVal PromptOnDrift As Boolean = True)
    Dim typeSummary As String
    Dim srcPaths As Collection
    Dim changedSrcPaths As Collection
    Dim driftTypeSummary As String
    Dim repoSig As String
    Dim wbSig As String
    Dim typeSigMapText As String
    Dim reason As String
    Dim previousEnableEvents As Boolean
    Dim previousScreenUpdating As Boolean
    Dim srcPath As Variant

    If Not PrepareSyncContext(wb, typeSummary, srcPaths, changedSrcPaths, driftTypeSummary, _
                              repoSig, wbSig, typeSigMapText, reason) Then
        MsgBox reason, vbExclamation, SYNC_DIALOG_TITLE
        Exit Sub
    End If

    If Not Force Then
        If StrComp(repoSig, wbSig, vbBinaryCompare) = 0 Then
            Exit Sub
        End If

        If changedSrcPaths.Count = 0 Then
            Exit Sub
        End If

        If PromptOnDrift Then
            If Len(driftTypeSummary) = 0 Then
                driftTypeSummary = typeSummary
            End If
            If Not ConfirmDriftImport(wb, driftTypeSummary) Then
                Exit Sub
            End If
        End If
    End If

    On Error Resume Next
    previousEnableEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating
    On Error GoTo SyncFail

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    If Not CanAccessVBProject(wb, reason) Then
        Err.Raise vbObjectError + 7000, "ReloadManager.SyncWorkbook", reason
    End If

    If Force Then
        For Each srcPath In srcPaths
            ImportBasFolder wb, CStr(srcPath)
        Next srcPath
    Else
        For Each srcPath In changedSrcPaths
            ImportBasFolder wb, CStr(srcPath)
        Next srcPath
    End If

    SetDocProp wb, PROP_SRC_SIG, repoSig
    SetDocProp wb, PROP_SRC_SIG_MAP, typeSigMapText
    SetDocProp wb, PROP_LAST_SYNC, Format$(NowUTC(), "yyyy-mm-dd\THH:NN:SS\Z")

SyncExit:
    On Error Resume Next
    Application.EnableEvents = previousEnableEvents
    Application.ScreenUpdating = previousScreenUpdating
    On Error GoTo 0
    Exit Sub

SyncFail:
    MsgBox "VBA sync failed for workbook '" & wb.Name & "'." & vbCrLf & _
           "Reason: " & Err.Description, vbCritical, SYNC_DIALOG_TITLE
    Resume SyncExit
End Sub

Private Function PromptSelectTypeIdsUserForm(ByVal typeIds As Collection, _
                                             ByVal formCaption As String, _
                                             ByVal infoText As String) As Collection
    Set PromptSelectTypeIdsUserForm = PromptTypeIdsUserFormInternal(typeIds, Nothing, formCaption, infoText)
End Function

Private Function PromptManageWorkbookTypeIdsUserForm(ByVal availableTypeIds As Collection, _
                                                    ByVal checkedTypeMap As Object, _
                                                    ByVal formCaption As String, _
                                                    ByVal infoText As String, _
                                                    Optional ByRef selectedAction As String = "") As Collection
    Set PromptManageWorkbookTypeIdsUserForm = PromptTypeIdsUserFormInternal(availableTypeIds, checkedTypeMap, _
                                                                           formCaption, infoText, True, selectedAction)
End Function

Private Function PromptTypeIdsUserFormInternal(ByVal sourceTypeIds As Collection, _
                                               ByVal checkedTypeMap As Object, _
                                               ByVal formCaption As String, _
                                               ByVal infoText As String, _
                                               Optional ByVal showSyncButtons As Boolean = False, _
                                               Optional ByRef selectedAction As String = "") As Collection
    Dim displayLines As Collection
    Dim ownerTypeIds As Collection
    Dim selectedTypeIds As Collection

    Set selectedTypeIds = New Collection
    If sourceTypeIds Is Nothing Then
        Set PromptTypeIdsUserFormInternal = selectedTypeIds
        Exit Function
    End If
    If sourceTypeIds.Count = 0 Then
        Set PromptTypeIdsUserFormInternal = selectedTypeIds
        Exit Function
    End If

    Set displayLines = New Collection
    Set ownerTypeIds = New Collection
    BuildTypePickerTreeLines sourceTypeIds, checkedTypeMap, displayLines, ownerTypeIds

    If displayLines.Count = 0 Then
        Set PromptTypeIdsUserFormInternal = selectedTypeIds
        Exit Function
    End If

    Set selectedTypeIds = ShowTypePickerCheckedTypeIdsUserForm(displayLines, ownerTypeIds, _
                                                               formCaption, infoText, showSyncButtons, selectedAction)
    Set PromptTypeIdsUserFormInternal = selectedTypeIds
End Function

Private Sub BuildTypePickerTreeLines(ByVal typeIds As Collection, _
                                     ByVal checkedTypeMap As Object, _
                                     ByRef displayLines As Collection, _
                                     ByRef ownerTypeIds As Collection)
    Dim i As Long
    Dim j As Long
    Dim typeId As String
    Dim normalizedTypeId As String
    Dim basFiles As Collection
    Dim basFile As Variant
    Dim marker As String
    Dim parts() As String
    Dim previousParts() As String
    Dim hasPrevious As Boolean
    Dim commonDepth As Long
    Dim leafIndex As Long

    For i = 1 To typeIds.Count
        typeId = CStr(typeIds(i))
        normalizedTypeId = NormalizeTypeId(typeId)
        If Len(normalizedTypeId) = 0 Then GoTo NextType

        parts = Split(normalizedTypeId, Application.PathSeparator)
        leafIndex = UBound(parts)

        If hasPrevious Then
            commonDepth = GetCommonFolderDepth(previousParts, parts, True)
        Else
            commonDepth = 0
        End If

        For j = commonDepth To leafIndex
            displayLines.Add String$(j * TP_INDENT_SPACES_PER_LEVEL_INTERNAL, " ") & parts(j)
            ownerTypeIds.Add vbNullString
        Next j

        Set basFiles = GetBasFileNamesInTypeSource(normalizedTypeId)

        If (Not checkedTypeMap Is Nothing) And checkedTypeMap.Exists(normalizedTypeId) Then
            marker = "[x]"
        Else
            marker = "[ ]"
        End If

        For Each basFile In basFiles
            displayLines.Add String$((leafIndex + 1) * TP_INDENT_SPACES_PER_LEVEL_INTERNAL, " ") & marker & " " & CStr(basFile)
            ownerTypeIds.Add normalizedTypeId
        Next basFile

        previousParts = parts
        hasPrevious = True
NextType:
    Next i
End Sub

Private Function ShowTypePickerPreviewUserForm(ByVal displayLines As Collection, _
                                            ByVal ownerTypeIds As Collection, _
                                            ByVal formCaption As String) As Long
    Dim vbComp As Object
    Dim frm As Object
    Dim entriesHost As Object
    Dim lbl As Object
    Dim btnOk As Object
    Dim btnSyncDrift As Object
    Dim btnForceSync As Object
    Dim layoutPlan As Collection
    Dim layoutContentHeight As Double
    Dim reason As String
    Dim stage As String

    On Error GoTo Fail

    stage = "EnsureTypePickerFormComponent"
    Set vbComp = EnsureTypePickerFormComponent(ThisWorkbook, reason)
    If vbComp Is Nothing Then
        MsgBox "Unable to open type picker userform." & vbCrLf & _
                             "Reason: " & reason, vbExclamation, SYNC_DIALOG_TITLE
        Exit Function
    End If

    stage = "UserForms.Add"
    Set frm = VBA.UserForms.Add(vbComp.Name)
    frm.Caption = formCaption
    frm.Tag = "0"

    stage = "Find info label"
    Set lbl = GetFormControlByName(frm, TYPE_PICKER_INFO_NAME)
    If lbl Is Nothing Then Err.Raise vbObjectError + 8101, "ReloadManager.ShowTypePickerPreviewUserForm", _
                                    "Missing control: " & TYPE_PICKER_INFO_NAME
    ResetTypePickerInfoLabel lbl

    stage = "Find entries host"
    Set entriesHost = GetFormControlByName(frm, TYPE_PICKER_LIST_NAME)
    If entriesHost Is Nothing Then Err.Raise vbObjectError + 8102, "ReloadManager.ShowTypePickerPreviewUserForm", _
                                    "Missing control: " & TYPE_PICKER_LIST_NAME
    If StrComp(TypeName(entriesHost), "Frame", vbTextCompare) <> 0 Then _
        Err.Raise vbObjectError + 8103, "ReloadManager.ShowTypePickerPreviewUserForm", _
                  "Control '" & TYPE_PICKER_LIST_NAME & "' must be a Frame for checkbox rendering."
    ApplyTypePickerEntriesHostStyle entriesHost

    stage = "Configure action buttons"
    Set btnOk = GetFormControlByName(frm, TYPE_PICKER_OK_NAME)
    Set btnSyncDrift = GetFormControlByName(frm, TYPE_PICKER_SYNC_DRIFT_NAME)
    Set btnForceSync = GetFormControlByName(frm, TYPE_PICKER_FORCE_SYNC_NAME)
    If Not btnOk Is Nothing Then TrySetObjectProperty btnOk, "Caption", "OK"
    If Not btnSyncDrift Is Nothing Then TrySetObjectProperty btnSyncDrift, "Visible", False
    If Not btnForceSync Is Nothing Then TrySetObjectProperty btnForceSync, "Visible", False

    stage = "Build layout plan"
    Set layoutPlan = BuildTypePickerLayoutPlan(displayLines, ownerTypeIds, 0, layoutContentHeight)

    stage = "AutoSize form"
    AutoSizeTypePickerForm frm, displayLines, layoutPlan, layoutContentHeight

    stage = "Build checkbox entries"
    PopulateTypePickerChecklist entriesHost, layoutPlan, layoutContentHeight

    stage = "Show form"
    frm.Show vbModal

    stage = "Read selection"
    On Error Resume Next
    If StrComp(CStr(frm.Tag), TYPE_PICKER_TAG_OK, vbTextCompare) = 0 Then
        ShowTypePickerPreviewUserForm = GetFirstCheckedChecklistLineIndex(entriesHost, ownerTypeIds)
    ElseIf Len(CStr(frm.Tag)) > 0 And IsNumeric(CStr(frm.Tag)) Then
        ShowTypePickerPreviewUserForm = CLng(Val(CStr(frm.Tag)))
    End If
    Err.Clear
    On Error GoTo Fail

    stage = "Cleanup form"
    On Error Resume Next
    Unload frm
    Set frm = Nothing
    Err.Clear
    On Error GoTo 0
    Exit Function

Fail:
    MsgBox "UserForm failed at stage: " & stage & vbCrLf & _
           "Error " & CStr(Err.Number) & ": " & Err.Description, vbExclamation, SYNC_DIALOG_TITLE
    On Error Resume Next
    If Not frm Is Nothing Then Unload frm
End Function

Private Function ShowTypePickerCheckedTypeIdsUserForm(ByVal displayLines As Collection, _
                                                      ByVal ownerTypeIds As Collection, _
                                                      ByVal formCaption As String, _
                                                      ByVal infoText As String, _
                                                      Optional ByVal showSyncButtons As Boolean = False, _
                                                      Optional ByRef selectedAction As String = "") As Collection
    Dim vbComp As Object
    Dim frm As Object
    Dim entriesHost As Object
    Dim lbl As Object
    Dim btnOk As Object
    Dim btnSyncDrift As Object
    Dim btnForceSync As Object
    Dim layoutPlan As Collection
    Dim layoutContentHeight As Double
    Dim reason As String
    Dim stage As String
    Dim selectedTypeIds As New Collection

    On Error GoTo Fail

    stage = "EnsureTypePickerFormComponent"
    Set vbComp = EnsureTypePickerFormComponent(ThisWorkbook, reason)
    If vbComp Is Nothing Then
        MsgBox "Unable to open type picker." & vbCrLf & _
               "Reason: " & reason, vbExclamation, SYNC_DIALOG_TITLE
        Set ShowTypePickerCheckedTypeIdsUserForm = selectedTypeIds
        Exit Function
    End If

    stage = "UserForms.Add"
    Set frm = VBA.UserForms.Add(vbComp.Name)
    frm.Caption = formCaption
    frm.Tag = TYPE_PICKER_TAG_CANCEL

    stage = "Find info label"
    Set lbl = GetFormControlByName(frm, TYPE_PICKER_INFO_NAME)
    If lbl Is Nothing Then Err.Raise vbObjectError + 8111, "ReloadManager.ShowTypePickerCheckedTypeIdsUserForm", _
                                    "Missing control: " & TYPE_PICKER_INFO_NAME
    ' Keep popup body text empty; title bar already identifies the dialog.
    ResetTypePickerInfoLabel lbl

    stage = "Find entries host"
    Set entriesHost = GetFormControlByName(frm, TYPE_PICKER_LIST_NAME)
    If entriesHost Is Nothing Then Err.Raise vbObjectError + 8112, "ReloadManager.ShowTypePickerCheckedTypeIdsUserForm", _
                                    "Missing control: " & TYPE_PICKER_LIST_NAME
    If StrComp(TypeName(entriesHost), "Frame", vbTextCompare) <> 0 Then _
        Err.Raise vbObjectError + 8113, "ReloadManager.ShowTypePickerCheckedTypeIdsUserForm", _
                  "Control '" & TYPE_PICKER_LIST_NAME & "' must be a Frame for checkbox rendering."
    ApplyTypePickerEntriesHostStyle entriesHost

    stage = "Configure action buttons"
    Set btnOk = GetFormControlByName(frm, TYPE_PICKER_OK_NAME)
    Set btnSyncDrift = GetFormControlByName(frm, TYPE_PICKER_SYNC_DRIFT_NAME)
    Set btnForceSync = GetFormControlByName(frm, TYPE_PICKER_FORCE_SYNC_NAME)
    If btnOk Is Nothing Then Err.Raise vbObjectError + 8114, "ReloadManager.ShowTypePickerCheckedTypeIdsUserForm", _
                                      "Missing control: " & TYPE_PICKER_OK_NAME
    If showSyncButtons Then
        TrySetObjectProperty btnOk, "Caption", "Apply"
    Else
        TrySetObjectProperty btnOk, "Caption", "OK"
    End If
    If Not btnSyncDrift Is Nothing Then
        TrySetObjectProperty btnSyncDrift, "Visible", showSyncButtons
    End If
    If Not btnForceSync Is Nothing Then
        TrySetObjectProperty btnForceSync, "Visible", showSyncButtons
    End If

    stage = "Build layout plan"
    Set layoutPlan = BuildTypePickerLayoutPlan(displayLines, ownerTypeIds, 0, layoutContentHeight)

    stage = "AutoSize form"
    AutoSizeTypePickerForm frm, displayLines, layoutPlan, layoutContentHeight

    stage = "Build checkbox entries"
    PopulateTypePickerChecklist entriesHost, layoutPlan, layoutContentHeight

    stage = "Show form"
    frm.Show vbModal

    selectedAction = CStr(frm.Tag)

    stage = "Collect checked selections"
    If IsTypePickerAcceptTag(CStr(frm.Tag)) Then
        CollectCheckedTypeIdsFromChecklist entriesHost, ownerTypeIds, selectedTypeIds
    End If

    stage = "Cleanup form"
    On Error Resume Next
    Unload frm
    Set frm = Nothing
    Err.Clear
    On Error GoTo 0

    Set ShowTypePickerCheckedTypeIdsUserForm = selectedTypeIds
    Exit Function

Fail:
    MsgBox "Type picker failed at stage: " & stage & vbCrLf & _
           "Error " & CStr(Err.Number) & ": " & Err.Description, vbExclamation, SYNC_DIALOG_TITLE
    On Error Resume Next
    If Not frm Is Nothing Then Unload frm
    Set ShowTypePickerCheckedTypeIdsUserForm = selectedTypeIds
End Function

Private Sub PopulateTypePickerChecklist(ByVal entriesHost As Object, _
                                        ByVal layoutPlan As Collection, _
                                        ByVal layoutContentHeight As Double)
    Dim rowCursor As Long
    Dim rowIndex As Long
    Dim rowInfo As Variant
    Dim lineText As String
    Dim captionText As String
    Dim rowTop As Double
    Dim rowStep As Double
    Dim maxWidth As Double
    Dim offsetLeft As Double
    Dim ctlWidth As Double
    Dim ctl As Object
    Dim lblText As Object
    Dim depthLevel As Long
    Dim fontSize As Double
    Dim rowControlHeight As Double
    Dim checkboxFontHeight As Double
    Dim checkboxWidth As Double
    Dim checkboxHeight As Double
    Dim checkboxRowSize As Double
    Dim checkboxTop As Double
    Dim checkboxBottom As Double
    Dim checkboxLeft As Double
    Dim textTop As Double
    Dim rowBottom As Double
    Dim textControlHeight As Double
    Dim textLeft As Double
    Dim textWidth As Double
    Dim debugEnabled As Boolean
    Dim debugLines As Collection
    Dim actualTextFont As Double
    Dim actualCheckboxFont As Double
    Dim actualTextFontName As String
    Dim actualCheckboxFontName As String
    Dim isCheckRow As Boolean
    Dim textLineForDebug As String
    Dim viewportWidth As Double
    Dim viewportHeight As Double
    Dim scrollHeight As Double
    Dim checkboxHitWidth As Double

    If entriesHost Is Nothing Then Exit Sub

    ClearTypePickerDynamicControls entriesHost

    If layoutContentHeight <= 0 Then layoutContentHeight = TP_LIST_TOP_PADDING + TP_LIST_SCROLL_EXTRA

    viewportWidth = GetControlInsideWidth(entriesHost)
    maxWidth = viewportWidth - TP_LIST_RIGHT_PADDING
    If maxWidth < TP_LIST_MIN_CONTROL_WIDTH Then maxWidth = TP_LIST_MIN_CONTROL_WIDTH
    checkboxFontHeight = GetFileCheckboxFontHeight()
    checkboxHeight = ComputeControlHeightFromFontHeight(checkboxFontHeight, 0)
    If checkboxHeight < 6 Then checkboxHeight = 6
    checkboxWidth = checkboxHeight
    debugEnabled = TP_DEBUG_ENABLE_ROW_REPORT
    If debugEnabled Then
        Set debugLines = New Collection
        debugLines.Add "TypePicker Row Debug Report"
        debugLines.Add "Timestamp: " & Format$(Now, "yyyy-mm-dd HH:NN:SS")
        debugLines.Add "TP_FILE_FONT_HEIGHT=" & CStr(TP_FILE_FONT_HEIGHT) & _
                       " | TP_FILE_FONT_ADDITIONAL_HEIGHT=" & CStr(TP_FILE_FONT_ADDITIONAL_HEIGHT) & _
                       " | TP_FILE_ROW_ADDITIONAL_SPACING=" & CStr(TP_FILE_ROW_ADDITIONAL_SPACING) & _
                       " | TP_FILE_CHECKBOX_FONT_REDUCTION_POINTS=" & CStr(TP_FILE_CHECKBOX_FONT_REDUCTION_POINTS) & _
                       " | TP_FILE_CHECKBOX_BOTTOM_OFFSET=" & CStr(TP_FILE_CHECKBOX_BOTTOM_OFFSET)
        debugLines.Add String$(100, "-")
    End If

    If Not layoutPlan Is Nothing Then
        For rowCursor = 1 To layoutPlan.Count
            rowInfo = layoutPlan.Item(rowCursor)
            rowIndex = CLng(rowInfo(0))
            isCheckRow = CBool(rowInfo(1))
            depthLevel = CLng(rowInfo(2))
            rowTop = CDbl(rowInfo(3))
            rowStep = CDbl(rowInfo(4))
            rowControlHeight = CDbl(rowInfo(5))
            captionText = CStr(rowInfo(6))
            lineText = CStr(rowInfo(7))
            textLineForDebug = captionText

        offsetLeft = TP_LIST_BASE_LEFT + (depthLevel * TP_INDENT_PIXELS_PER_LEVEL)
        ctlWidth = maxWidth - (depthLevel * TP_INDENT_PIXELS_PER_LEVEL)
        If ctlWidth < TP_LIST_MIN_CONTROL_WIDTH Then ctlWidth = TP_LIST_MIN_CONTROL_WIDTH

            If isCheckRow Then
                fontSize = TP_FILE_FONT_HEIGHT
                textControlHeight = rowControlHeight
                textTop = rowTop
                rowBottom = rowTop + rowStep

                checkboxRowSize = checkboxHeight
                If rowStep > 6 Then
                    If checkboxRowSize >= rowStep Then checkboxRowSize = rowStep - 1
                End If
                If checkboxRowSize < 6 Then checkboxRowSize = 6

                checkboxWidth = checkboxRowSize
                checkboxBottom = rowBottom - TP_FILE_ROW_ADDITIONAL_SPACING - TP_FILE_CHECKBOX_BOTTOM_OFFSET
                checkboxTop = checkboxBottom - checkboxRowSize
                checkboxLeft = offsetLeft

                textLeft = checkboxLeft + checkboxWidth + TP_FOLDER_LABEL_LEFT_NUDGE
                textWidth = ctlWidth - (checkboxWidth + TP_FOLDER_LABEL_LEFT_NUDGE)
                If textWidth < 8 Then textWidth = 8
                checkboxHitWidth = ctlWidth
                If checkboxHitWidth < checkboxWidth Then checkboxHitWidth = checkboxWidth

                Set ctl = entriesHost.Controls.Add("Forms.CheckBox.1", "chkEntry_" & CStr(rowIndex), True)
                ApplyUniformControlFontSize ctl, checkboxFontHeight
                TrySetObjectProperty ctl, "Caption", vbNullString
                TrySetObjectProperty ctl, "Value", IsChecklistLineChecked(lineText)
                TrySetObjectProperty ctl, "Tag", CStr(rowIndex)
                TrySetObjectProperty ctl, "Left", checkboxLeft
                TrySetObjectProperty ctl, "Top", checkboxTop
                TrySetObjectProperty ctl, "Width", checkboxWidth
                TrySetObjectProperty ctl, "Height", checkboxRowSize

                Set lblText = entriesHost.Controls.Add("Forms.Label.1", "lblCheckText_" & CStr(rowIndex), True)
                ApplyUniformControlFontSize lblText, fontSize
                TrySetObjectProperty lblText, "Caption", captionText
                TrySetObjectProperty lblText, "Left", textLeft
                TrySetObjectProperty lblText, "Top", textTop
                TrySetObjectProperty lblText, "Width", textWidth
                TrySetObjectProperty lblText, "Height", textControlHeight
                TrySetObjectProperty lblText, "WordWrap", False
                TrySetObjectProperty lblText, "AutoSize", False
                TrySetObjectProperty lblText, "BackStyle", 0

                ' Keep glyph placement unchanged, but make the checkbox clickable across the row text.
                TrySetObjectProperty ctl, "BackStyle", 0
                TrySetObjectProperty ctl, "Width", checkboxHitWidth
                On Error Resume Next
                CallByName ctl, "ZOrder", VbMethod, 0
                Err.Clear
                On Error GoTo 0

                If debugEnabled Then
                    actualTextFont = GetControlFontSize(lblText)
                    actualCheckboxFont = GetControlFontSize(ctl)
                    actualTextFontName = GetControlFontName(lblText)
                    actualCheckboxFontName = GetControlFontName(ctl)
                    debugLines.Add "Row " & CStr(rowIndex) & " | FILE | text='" & Replace(captionText, "'", "''") & _
                                   "' | intendedTextFont=" & CStr(fontSize) & _
                                   " | actualTextFont=" & CStr(actualTextFont) & _
                                   " | actualTextFontName=" & actualTextFontName & _
                                   " | intendedCheckboxFont=" & CStr(checkboxFontHeight) & _
                                   " | actualCheckboxFont=" & CStr(actualCheckboxFont) & _
                                   " | actualCheckboxFontName=" & actualCheckboxFontName & _
                                   " | rowTop=" & CStr(rowTop) & _
                                   " | rowStep=" & CStr(rowStep) & _
                                   " | textTop=" & CStr(textTop) & _
                                   " | textHeight=" & CStr(textControlHeight) & _
                                   " | checkboxTop=" & CStr(checkboxTop) & _
                                   " | checkboxHeight=" & CStr(checkboxRowSize)
                End If
            Else
                fontSize = GetFolderFontHeightByDepth(depthLevel)

                Set ctl = entriesHost.Controls.Add("Forms.Label.1", "lblEntry_" & CStr(rowIndex), True)
                ApplyUniformControlFontSize ctl, fontSize
                TrySetObjectProperty ctl, "Caption", captionText
                TrySetObjectProperty ctl, "Left", offsetLeft + TP_FOLDER_LABEL_LEFT_NUDGE
                TrySetObjectProperty ctl, "Top", rowTop
                TrySetObjectProperty ctl, "Width", ctlWidth
                TrySetObjectProperty ctl, "Height", rowControlHeight
                TrySetObjectProperty ctl, "WordWrap", False
                TrySetObjectProperty ctl, "AutoSize", False

                If debugEnabled Then
                    actualTextFont = GetControlFontSize(ctl)
                    actualTextFontName = GetControlFontName(ctl)
                    debugLines.Add "Row " & CStr(rowIndex) & " | FOLDER | text='" & Replace(textLineForDebug, "'", "''") & _
                                   "' | intendedFont=" & CStr(fontSize) & _
                                   " | actualFont=" & CStr(actualTextFont) & _
                                   " | actualFontName=" & actualTextFontName & _
                                   " | rowTop=" & CStr(rowTop) & _
                                   " | rowStep=" & CStr(rowStep) & _
                                   " | rowHeight=" & CStr(rowControlHeight)
                End If
            End If
        Next rowCursor
    End If

    TrySetObjectProperty entriesHost, "ScrollBars", 2
    TrySetObjectProperty entriesHost, "KeepScrollBarsVisible", 0
    viewportHeight = GetControlInsideHeight(entriesHost)
    scrollHeight = layoutContentHeight
    If scrollHeight < viewportHeight Then scrollHeight = viewportHeight
    TrySetObjectProperty entriesHost, "ScrollHeight", scrollHeight
    TrySetObjectProperty entriesHost, "ScrollWidth", viewportWidth
    On Error Resume Next
    entriesHost.ScrollTop = 0
    On Error GoTo 0

    If debugEnabled Then
        WriteTypePickerRowDebugReport debugLines
    End If
End Sub

Private Sub ApplyUniformControlFontSize(ByVal ctl As Object, ByVal fontSize As Double)
    Dim fnt As Object

    If ctl Is Nothing Then Exit Sub

    On Error Resume Next
    Set fnt = CallByName(ctl, "Font", VbGet)
    If Not fnt Is Nothing Then
        CallByName fnt, "Name", VbLet, TP_UI_FONT_NAME
        CallByName fnt, "Size", VbLet, fontSize
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Private Function GetControlFontSize(ByVal ctl As Object) As Double
    Dim fnt As Object

    If ctl Is Nothing Then
        GetControlFontSize = -1
        Exit Function
    End If

    On Error Resume Next
    Set fnt = CallByName(ctl, "Font", VbGet)
    If Not fnt Is Nothing Then
        GetControlFontSize = CDbl(CallByName(fnt, "Size", VbGet))
    Else
        GetControlFontSize = -1
    End If
    Err.Clear
    On Error GoTo 0
End Function

Private Function GetControlFontName(ByVal ctl As Object) As String
    Dim fnt As Object

    If ctl Is Nothing Then Exit Function

    On Error Resume Next
    Set fnt = CallByName(ctl, "Font", VbGet)
    If Not fnt Is Nothing Then
        GetControlFontName = CStr(CallByName(fnt, "Name", VbGet))
    End If
    Err.Clear
    On Error GoTo 0
End Function

Private Sub WriteTypePickerRowDebugReport(ByVal debugLines As Collection)
    Dim outputPath As String
    Dim lineItem As Variant
    Dim fileHandle As Integer

    If debugLines Is Nothing Then Exit Sub
    If debugLines.Count = 0 Then Exit Sub

    outputPath = GetTypePickerRowDebugReportPath()
    If Len(outputPath) = 0 Then Exit Sub

    On Error GoTo Fail
    fileHandle = FreeFile
    Open outputPath For Output As #fileHandle
    For Each lineItem In debugLines
        Print #fileHandle, CStr(lineItem)
    Next lineItem
    Close #fileHandle

    MsgBox "Type picker debug report written to:" & vbCrLf & outputPath, vbInformation, SYNC_DIALOG_TITLE
    Exit Sub

Fail:
    On Error Resume Next
    If fileHandle <> 0 Then Close #fileHandle
    On Error GoTo 0
End Sub

Private Function GetTypePickerRowDebugReportPath() As String
    Dim tempPath As String

    tempPath = Environ$("TEMP")
    If Len(tempPath) = 0 Then tempPath = ThisWorkbook.Path
    If Len(tempPath) = 0 Then Exit Function

    GetTypePickerRowDebugReportPath = EnsureTrailingPathSeparator(tempPath) & _
                                      "TypePicker_RowDebug_" & Format$(Now, "yyyymmdd_HHNNSS") & ".txt"
End Function

Private Function GetHierarchyDepthFromIndent(ByVal indentChars As Long) As Long
    If TP_INDENT_SPACES_PER_LEVEL_INTERNAL <= 0 Then
        GetHierarchyDepthFromIndent = 0
    Else
        GetHierarchyDepthFromIndent = indentChars \ TP_INDENT_SPACES_PER_LEVEL_INTERNAL
    End If
End Function

Private Function GetFolderFontHeightByDepth(ByVal depthLevel As Long) As Double
    Select Case depthLevel
        Case 0
            GetFolderFontHeightByDepth = TP_FOLDER_FONT_HEIGHT_L0
        Case 1
            GetFolderFontHeightByDepth = TP_FOLDER_FONT_HEIGHT_L1
        Case 2
            GetFolderFontHeightByDepth = TP_FOLDER_FONT_HEIGHT_L2
        Case 3
            GetFolderFontHeightByDepth = TP_FOLDER_FONT_HEIGHT_L3
        Case Else
            GetFolderFontHeightByDepth = TP_FOLDER_FONT_HEIGHT_DEFAULT
    End Select
End Function

Private Function GetFolderFontAdditionalHeightByDepth(ByVal depthLevel As Long) As Double
    Select Case depthLevel
        Case 0
            GetFolderFontAdditionalHeightByDepth = TP_FOLDER_FONT_ADDITIONAL_HEIGHT_L0
        Case 1
            GetFolderFontAdditionalHeightByDepth = TP_FOLDER_FONT_ADDITIONAL_HEIGHT_L1
        Case 2
            GetFolderFontAdditionalHeightByDepth = TP_FOLDER_FONT_ADDITIONAL_HEIGHT_L2
        Case 3
            GetFolderFontAdditionalHeightByDepth = TP_FOLDER_FONT_ADDITIONAL_HEIGHT_L3
        Case Else
            GetFolderFontAdditionalHeightByDepth = TP_FOLDER_FONT_ADDITIONAL_HEIGHT_DEFAULT
    End Select
End Function

Private Function GetFolderRowAdditionalSpacingByDepth(ByVal depthLevel As Long) As Double
    Select Case depthLevel
        Case 0
            GetFolderRowAdditionalSpacingByDepth = TP_FOLDER_ROW_ADDITIONAL_SPACING_L0
        Case 1
            GetFolderRowAdditionalSpacingByDepth = TP_FOLDER_ROW_ADDITIONAL_SPACING_L1
        Case 2
            GetFolderRowAdditionalSpacingByDepth = TP_FOLDER_ROW_ADDITIONAL_SPACING_L2
        Case 3
            GetFolderRowAdditionalSpacingByDepth = TP_FOLDER_ROW_ADDITIONAL_SPACING_L3
        Case Else
            GetFolderRowAdditionalSpacingByDepth = TP_FOLDER_ROW_ADDITIONAL_SPACING_DEFAULT
    End Select
End Function

Private Function ComputeControlHeightFromFontHeight(ByVal fontHeight As Double, ByVal additionalHeight As Double) As Double
    ComputeControlHeightFromFontHeight = fontHeight + additionalHeight
    If ComputeControlHeightFromFontHeight < fontHeight Then
        ComputeControlHeightFromFontHeight = fontHeight
    End If
End Function

Private Function ComputeRowStepFromControlHeight(ByVal controlHeight As Double, _
                                                 ByVal rowStepAdd As Double) As Double
    ComputeRowStepFromControlHeight = controlHeight + rowStepAdd

    ' Keep rows contiguous by default; extra spacing should come from rowStepAdd.
    If ComputeRowStepFromControlHeight < controlHeight Then
        ComputeRowStepFromControlHeight = controlHeight
    End If
End Function

Private Function GetFileCheckboxFontHeight() As Double
    GetFileCheckboxFontHeight = TP_FILE_FONT_HEIGHT - TP_FILE_CHECKBOX_FONT_REDUCTION_POINTS
    If GetFileCheckboxFontHeight < 6 Then
        GetFileCheckboxFontHeight = 6
    End If
End Function

Private Function BuildTypePickerLayoutPlan(ByVal displayLines As Collection, _
                                           ByVal ownerTypeIds As Collection, _
                                           ByVal maxVisibleRows As Long, _
                                           ByRef contentHeight As Double) As Collection
    Dim results As New Collection
    Dim i As Long
    Dim lineText As String
    Dim captionText As String
    Dim indentChars As Long
    Dim depthLevel As Long
    Dim isCheckRow As Boolean
    Dim previousWasCheckRow As Boolean
    Dim rowTop As Double
    Dim rowControlHeight As Double
    Dim rowStep As Double
    Dim rowInfo As Variant

    rowTop = TP_LIST_TOP_PADDING

    If displayLines Is Nothing Then
        contentHeight = rowTop + TP_LIST_SCROLL_EXTRA
        Set BuildTypePickerLayoutPlan = results
        Exit Function
    End If

    For i = 1 To displayLines.Count
        If maxVisibleRows > 0 Then
            If results.Count >= maxVisibleRows Then Exit For
        End If

        lineText = CStr(displayLines(i))
        isCheckRow = ResolveChecklistRowIsCheckable(i, lineText, ownerTypeIds)
        indentChars = CountLeadingSpaces(lineText)
        depthLevel = GetHierarchyDepthFromIndent(indentChars)

        If results.Count > 0 Then
            If previousWasCheckRow And (Not isCheckRow) Then
                rowTop = rowTop + TP_GAP_AFTER_FILE_GROUP
            ElseIf (Not previousWasCheckRow) And isCheckRow Then
                rowTop = rowTop + TP_GAP_BEFORE_FILE_GROUP
            End If

            If (Not isCheckRow) And depthLevel = 0 Then
                rowTop = rowTop + TP_GAP_BEFORE_FOLDER_L0
            End If
        End If

        If isCheckRow Then
            rowControlHeight = ComputeControlHeightFromFontHeight(TP_FILE_FONT_HEIGHT, TP_FILE_FONT_ADDITIONAL_HEIGHT)
            rowStep = ComputeRowStepFromControlHeight(rowControlHeight, TP_FILE_ROW_ADDITIONAL_SPACING)
        Else
            rowControlHeight = ComputeControlHeightFromFontHeight(GetFolderFontHeightByDepth(depthLevel), _
                                                                  GetFolderFontAdditionalHeightByDepth(depthLevel))
            rowStep = ComputeRowStepFromControlHeight(rowControlHeight, GetFolderRowAdditionalSpacingByDepth(depthLevel))
        End If

        captionText = StripChecklistMarker(LTrim$(lineText))
        rowInfo = Array(CLng(i), CBool(isCheckRow), CLng(depthLevel), CDbl(rowTop), CDbl(rowStep), CDbl(rowControlHeight), CStr(captionText), CStr(lineText))
        results.Add rowInfo

        rowTop = rowTop + rowStep
        previousWasCheckRow = isCheckRow
    Next i

    contentHeight = rowTop + TP_LIST_SCROLL_EXTRA
    Set BuildTypePickerLayoutPlan = results
End Function

Private Function ResolveChecklistRowIsCheckable(ByVal rowIndex As Long, _
                                                ByVal lineText As String, _
                                                ByVal ownerTypeIds As Collection) As Boolean
    Dim ownerTypeId As String

    If Not ownerTypeIds Is Nothing Then
        If rowIndex <= ownerTypeIds.Count Then
            ownerTypeId = NormalizeTypeId(CStr(ownerTypeIds(rowIndex)))
        End If
    End If

    If Len(ownerTypeId) > 0 Then
        ResolveChecklistRowIsCheckable = True
    Else
        ResolveChecklistRowIsCheckable = IsChecklistLineCheckable(lineText)
    End If
End Function

Private Function GetTypePickerLayoutHeightForVisibleRows(ByVal layoutPlan As Collection, _
                                                         ByVal visibleRows As Long) As Double
    Dim effectiveRows As Long
    Dim rowInfo As Variant
    Dim rowBottom As Double

    If layoutPlan Is Nothing Then
        GetTypePickerLayoutHeightForVisibleRows = TP_LIST_TOP_PADDING + TP_LIST_SCROLL_EXTRA
        Exit Function
    End If

    effectiveRows = visibleRows
    If effectiveRows <= 0 Or effectiveRows > layoutPlan.Count Then
        effectiveRows = layoutPlan.Count
    End If

    If effectiveRows <= 0 Then
        GetTypePickerLayoutHeightForVisibleRows = TP_LIST_TOP_PADDING + TP_LIST_SCROLL_EXTRA
        Exit Function
    End If

    rowInfo = layoutPlan.Item(effectiveRows)
    rowBottom = CDbl(rowInfo(3)) + CDbl(rowInfo(4))

    GetTypePickerLayoutHeightForVisibleRows = rowBottom + TP_LIST_SCROLL_EXTRA
End Function

Private Sub ClearTypePickerDynamicControls(ByVal entriesHost As Object)
    Dim i As Long
    Dim ctl As Object
    Dim ctlName As String

    If entriesHost Is Nothing Then Exit Sub

    On Error Resume Next
    For i = entriesHost.Controls.Count - 1 To 0 Step -1
        Set ctl = entriesHost.Controls.Item(i)
        ctlName = CStr(ctl.Name)
        If Left$(ctlName, 9) = "chkEntry_" Or Left$(ctlName, 9) = "lblEntry_" Or Left$(ctlName, 13) = "lblCheckText_" Then
            entriesHost.Controls.Remove ctlName
        End If
    Next i
    On Error GoTo 0
End Sub

Private Sub CollectCheckedTypeIdsFromChecklist(ByVal entriesHost As Object, _
                                               ByVal ownerTypeIds As Collection, _
                                               ByRef selectedTypeIds As Collection)
    Dim seen As Object
    Dim i As Long
    Dim ownerTypeId As String
    Dim chk As Object

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    For i = 1 To ownerTypeIds.Count
        ownerTypeId = NormalizeTypeId(CStr(ownerTypeIds(i)))
        If Len(ownerTypeId) > 0 Then
            Set chk = GetFormControlByName(entriesHost, "chkEntry_" & CStr(i))
            If Not chk Is Nothing Then
                On Error Resume Next
                If CBool(chk.Value) Then
                    If Not seen.Exists(ownerTypeId) Then
                        seen(ownerTypeId) = True
                        selectedTypeIds.Add ownerTypeId
                    End If
                End If
                On Error GoTo 0
            End If
        End If
    Next i
End Sub

Private Function GetFirstCheckedChecklistLineIndex(ByVal entriesHost As Object, _
                                                   ByVal ownerTypeIds As Collection) As Long
    Dim i As Long
    Dim ownerTypeId As String
    Dim chk As Object

    If entriesHost Is Nothing Then Exit Function
    If ownerTypeIds Is Nothing Then Exit Function

    For i = 1 To ownerTypeIds.Count
        ownerTypeId = NormalizeTypeId(CStr(ownerTypeIds(i)))
        If Len(ownerTypeId) > 0 Then
            Set chk = GetFormControlByName(entriesHost, "chkEntry_" & CStr(i))
            If Not chk Is Nothing Then
                On Error Resume Next
                If CBool(chk.Value) Then
                    GetFirstCheckedChecklistLineIndex = i
                    Exit Function
                End If
                On Error GoTo 0
            End If
        End If
    Next i
End Function

Private Function CountLeadingSpaces(ByVal textValue As String) As Long
    CountLeadingSpaces = Len(textValue) - Len(LTrim$(textValue))
End Function

Private Function StripChecklistMarker(ByVal trimmedLine As String) As String
    If Left$(trimmedLine, 3) = "[ ]" Or Left$(trimmedLine, 3) = "[x]" Or Left$(trimmedLine, 3) = "[X]" Then
        If Len(trimmedLine) > 4 Then
            StripChecklistMarker = Mid$(trimmedLine, 5)
        Else
            StripChecklistMarker = ""
        End If
    Else
        StripChecklistMarker = trimmedLine
    End If
End Function

Private Function IsChecklistLineCheckable(ByVal lineText As String) As Boolean
    Dim trimmedLine As String

    trimmedLine = LTrim$(lineText)
    IsChecklistLineCheckable = (Left$(trimmedLine, 3) = "[ ]" Or _
                                Left$(trimmedLine, 3) = "[x]" Or _
                                Left$(trimmedLine, 3) = "[X]")
End Function

Private Function IsChecklistLineChecked(ByVal lineText As String) As Boolean
    Dim trimmedLine As String

    trimmedLine = LTrim$(lineText)
    IsChecklistLineChecked = (Left$(trimmedLine, 3) = "[x]" Or Left$(trimmedLine, 3) = "[X]")
End Function

Private Function IsTypePickerAcceptTag(ByVal tagValue As String) As Boolean
    Select Case UCase$(Trim$(tagValue))
        Case TYPE_PICKER_TAG_OK, TYPE_PICKER_TAG_SYNC_DRIFT, TYPE_PICKER_TAG_FORCE_SYNC
            IsTypePickerAcceptTag = True
    End Select
End Function

Private Function EnsureTypePickerFormComponent(ByVal wb As Workbook, ByRef reason As String) As Object
    Dim vbComp As Object
    Dim designer As Object
    Dim stage As String

    On Error GoTo Fail

    stage = "Rebuild component"
    If Not RebuildTypePickerFormComponent(wb, reason, vbComp) Then Exit Function

    If vbComp Is Nothing Then
        stage = "Find component"
        Set vbComp = FindTypePickerFormComponent(wb)
        If vbComp Is Nothing Then
            reason = "Type picker form component was not found after rebuild."
            Exit Function
        End If
    End If

    stage = "Open designer"
    Set designer = vbComp.Designer
    If designer Is Nothing Then
        reason = "Designer is unavailable for component '" & vbComp.Name & "'."
        Exit Function
    End If

    stage = "Configure designer controls"
    ConfigureTypePickerFormDesigner designer

    stage = "Inject form code"
    EnsureTypePickerFormCode vbComp.CodeModule

    Set EnsureTypePickerFormComponent = vbComp
    Exit Function

Fail:
    reason = "Stage: " & stage & " | Error " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Sub ResetTypePickerUserForm()
    Dim reason As String

    If RebuildTypePickerFormComponent(ThisWorkbook, reason) Then
        MsgBox "Type picker form was rebuilt successfully.", vbInformation, SYNC_DIALOG_TITLE
    Else
        MsgBox "Unable to rebuild type picker form." & vbCrLf & _
               "Reason: " & reason, vbExclamation, SYNC_DIALOG_TITLE
    End If
End Sub

Private Function RebuildTypePickerFormComponent(ByVal wb As Workbook, _
                                                    ByRef reason As String, _
                                                    Optional ByRef rebuiltComponent As Object) As Boolean
    Dim vbComp As Object
    Dim createReason As String

    On Error GoTo Fail

    If wb Is Nothing Then
        reason = "Workbook reference is missing."
        Exit Function
    End If

    UnloadTypePickerFormInstances

    If Not ResolveTypePickerNameConflict(wb, reason) Then
        Exit Function
    End If

    RemoveExistingTypePickerFormComponents wb

    Set vbComp = wb.VBProject.VBComponents.Add(3)
    If Not TryAssignTypePickerFormName(vbComp, createReason) Then
        If Len(createReason) > 0 Then
            reason = createReason
        Else
            reason = "Type picker form was created with a fallback component name."
        End If
    End If

    Set rebuiltComponent = vbComp

    RebuildTypePickerFormComponent = True
    Exit Function

Fail:
    reason = "Rebuild failed. Error " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Sub UnloadTypePickerFormInstances()
    Dim frm As Object
    Dim formName As String

    On Error Resume Next
    For Each frm In VBA.UserForms
        formName = vbNullString
        formName = CStr(frm.Name)
        If StrComp(formName, TYPE_PICKER_FORM_NAME, vbTextCompare) = 0 Or _
           StrComp(TypeName(frm), TYPE_PICKER_FORM_NAME, vbTextCompare) = 0 Then
            Unload frm
        End If
    Next frm
    On Error GoTo 0
End Sub

Private Sub RemoveExistingTypePickerFormComponents(ByVal wb As Workbook)
    Dim i As Long
    Dim vbComp As Object

    If wb Is Nothing Then Exit Sub

    On Error Resume Next
    For i = wb.VBProject.VBComponents.Count To 1 Step -1
        Set vbComp = wb.VBProject.VBComponents.Item(i)
        If IsTypePickerFormComponent(vbComp) Then
            wb.VBProject.VBComponents.Remove vbComp
        End If
    Next i
    On Error GoTo 0
End Sub

Private Function IsTypePickerFormComponent(ByVal vbComp As Object) As Boolean
    Dim designer As Object

    If vbComp Is Nothing Then Exit Function
    If vbComp.Type <> 3 Then Exit Function

    If StrComp(vbComp.Name, TYPE_PICKER_FORM_NAME, vbTextCompare) = 0 Then
        IsTypePickerFormComponent = True
        Exit Function
    End If

    If IsAutoRenamedModuleName(vbComp.Name, TYPE_PICKER_FORM_NAME) Then
        IsTypePickerFormComponent = True
        Exit Function
    End If

    On Error Resume Next
    Set designer = vbComp.Designer
    On Error GoTo 0
    If designer Is Nothing Then Exit Function

    If HasRequiredTypePickerControls(designer) Then
        IsTypePickerFormComponent = True
    End If
End Function

Private Function ResolveTypePickerNameConflict(ByVal wb As Workbook, ByRef reason As String) As Boolean
    Dim vbComp As Object
    Dim fallbackName As String

    If wb Is Nothing Then
        reason = "Workbook reference is missing."
        Exit Function
    End If

    For Each vbComp In wb.VBProject.VBComponents
        If StrComp(vbComp.Name, TYPE_PICKER_FORM_NAME, vbTextCompare) = 0 Then
            If vbComp.Type <> 3 Then
                fallbackName = TYPE_PICKER_FORM_NAME & "_Legacy"
                On Error Resume Next
                vbComp.Name = fallbackName
                If Err.Number <> 0 Then
                    Err.Clear
                    fallbackName = TYPE_PICKER_FORM_NAME & "_Legacy_" & CStr(Fix(Timer * 1000))
                    vbComp.Name = fallbackName
                End If
                If Err.Number <> 0 Then
                    reason = "A non-UserForm component owns name '" & TYPE_PICKER_FORM_NAME & "' and could not be renamed."
                    Err.Clear
                    On Error GoTo 0
                    Exit Function
                End If
                Err.Clear
                On Error GoTo 0
            End If
        End If
    Next vbComp

    ResolveTypePickerNameConflict = True
End Function

Private Function FindTypePickerFormComponent(ByVal wb As Workbook) As Object
    Dim vbComp As Object
    Dim candidate As Object

    If wb Is Nothing Then Exit Function

    For Each vbComp In wb.VBProject.VBComponents
        If vbComp.Type = 3 Then
            If StrComp(vbComp.Name, TYPE_PICKER_FORM_NAME, vbTextCompare) = 0 Then
                Set FindTypePickerFormComponent = vbComp
                Exit Function
            End If

            If IsAutoRenamedModuleName(vbComp.Name, TYPE_PICKER_FORM_NAME) Then
                If candidate Is Nothing Then
                    Set candidate = vbComp
                End If
            End If
        End If
    Next vbComp

    If Not candidate Is Nothing Then
        Set FindTypePickerFormComponent = candidate
    End If
End Function

Private Function TryAssignTypePickerFormName(ByVal vbComp As Object, ByRef reason As String) As Boolean
    If vbComp Is Nothing Then Exit Function

    On Error Resume Next
    vbComp.Name = TYPE_PICKER_FORM_NAME
    If Err.Number = 0 Then
        TryAssignTypePickerFormName = True
    Else
        reason = "Type picker form created as '" & vbComp.Name & "' because canonical name was unavailable."
        Err.Clear
        TryAssignTypePickerFormName = False
    End If
    On Error GoTo 0
End Function

Private Function HasRequiredTypePickerControls(ByVal designer As Object) As Boolean
    Dim entriesHost As Object

    Set entriesHost = GetDesignerControlByName(designer, TYPE_PICKER_LIST_NAME)
    If entriesHost Is Nothing Then Exit Function
    If StrComp(TypeName(entriesHost), "Frame", vbTextCompare) <> 0 Then Exit Function
    If GetDesignerControlByName(designer, TYPE_PICKER_OK_NAME) Is Nothing Then Exit Function
    If GetDesignerControlByName(designer, TYPE_PICKER_CANCEL_NAME) Is Nothing Then Exit Function
    If GetDesignerControlByName(designer, TYPE_PICKER_SYNC_DRIFT_NAME) Is Nothing Then Exit Function
    If GetDesignerControlByName(designer, TYPE_PICKER_FORCE_SYNC_NAME) Is Nothing Then Exit Function
    If GetDesignerControlByName(designer, TYPE_PICKER_INFO_NAME) Is Nothing Then Exit Function
    HasRequiredTypePickerControls = True
End Function

Private Sub ConfigureTypePickerFormDesigner(ByVal designer As Object)
    Dim lbl As Object
    Dim lst As Object
    Dim existingEntries As Object
    Dim btnOk As Object
    Dim btnCancel As Object
    Dim btnSyncDrift As Object
    Dim btnForceSync As Object
    Dim buttonGroupLeft As Double
    Dim seedListWidth As Double
    Dim seedListHeight As Double
    Dim seedButtonTop As Double
    Dim seedFormWidth As Double
    Dim seedFormHeight As Double
    Dim stage As String

    On Error GoTo Fail

    seedListWidth = TP_SEED_LIST_WIDTH
    seedListHeight = TP_SEED_LIST_HEIGHT
    seedButtonTop = TP_LAYOUT_TOP_MARGIN + seedListHeight + TP_LAYOUT_BUTTON_GAP_TOP
    seedFormWidth = TP_LAYOUT_LEFT_MARGIN + seedListWidth + TP_LAYOUT_RIGHT_MARGIN
    seedFormHeight = TP_LAYOUT_TOP_MARGIN + seedListHeight + TP_LAYOUT_BOTTOM_MARGIN

    stage = "Form properties"
    TrySetObjectProperty designer, "Caption", "Workbook Type Picker"
    TrySetObjectProperty designer, "Width", seedFormWidth
    TrySetObjectProperty designer, "Height", seedFormHeight

    stage = "Remove stale controls"
    RemoveUnexpectedTypePickerDesignerControls designer

    stage = "Info label"
    Set lbl = GetOrCreateFormControl(designer, "Forms.Label.1", TYPE_PICKER_INFO_NAME)
    ResetTypePickerInfoLabel lbl
    TrySetObjectProperty lbl, "WordWrap", True

    stage = "Entries host"
    Set existingEntries = GetDesignerControlByName(designer, TYPE_PICKER_LIST_NAME)
    If Not existingEntries Is Nothing Then
        If StrComp(TypeName(existingEntries), "Frame", vbTextCompare) <> 0 Then
            On Error Resume Next
            designer.Controls.Remove TYPE_PICKER_LIST_NAME
            Err.Clear
            On Error GoTo Fail

            Set existingEntries = GetDesignerControlByName(designer, TYPE_PICKER_LIST_NAME)
            If Not existingEntries Is Nothing Then
                On Error Resume Next
                CallByName existingEntries, "Name", VbLet, TYPE_PICKER_LIST_NAME & "_Legacy"
                Err.Clear
                On Error GoTo Fail
            End If
        End If
    End If

    Set lst = GetDesignerControlByName(designer, TYPE_PICKER_LIST_NAME)
    If lst Is Nothing Then
        Set lst = GetOrCreateFormControl(designer, "Forms.Frame.1", TYPE_PICKER_LIST_NAME)
    ElseIf StrComp(TypeName(lst), "Frame", vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 8122, "ReloadManager.ConfigureTypePickerFormDesigner", _
                  "Entries host control '" & TYPE_PICKER_LIST_NAME & "' is not a Frame."
    End If

    SetControlBounds lst, TP_LAYOUT_LEFT_MARGIN, TP_LAYOUT_TOP_MARGIN, seedListWidth, seedListHeight
    ApplyTypePickerEntriesHostStyle lst

    buttonGroupLeft = TP_LAYOUT_LEFT_MARGIN + seedListWidth - ((TP_LAYOUT_SYNC_BUTTON_WIDTH * 2) + _
                     (TP_LAYOUT_BUTTON_WIDTH * 2) + (TP_LAYOUT_BUTTON_GAP * 3))
    If buttonGroupLeft < TP_LAYOUT_LEFT_MARGIN Then buttonGroupLeft = TP_LAYOUT_LEFT_MARGIN

    stage = "Sync on drift button"
    Set btnSyncDrift = GetOrCreateFormControl(designer, "Forms.CommandButton.1", TYPE_PICKER_SYNC_DRIFT_NAME)
    TrySetObjectProperty btnSyncDrift, "Caption", "Sync on drift"
    SetTypePickerButtonLayout btnSyncDrift, buttonGroupLeft, seedButtonTop, TP_LAYOUT_SYNC_BUTTON_WIDTH, False

    stage = "Force sync button"
    Set btnForceSync = GetOrCreateFormControl(designer, "Forms.CommandButton.1", TYPE_PICKER_FORCE_SYNC_NAME)
    TrySetObjectProperty btnForceSync, "Caption", "Force sync"
    SetTypePickerButtonLayout btnForceSync, buttonGroupLeft + TP_LAYOUT_SYNC_BUTTON_WIDTH + TP_LAYOUT_BUTTON_GAP, _
                              seedButtonTop, TP_LAYOUT_SYNC_BUTTON_WIDTH, False

    stage = "OK button"
    Set btnOk = GetOrCreateFormControl(designer, "Forms.CommandButton.1", TYPE_PICKER_OK_NAME)
    TrySetObjectProperty btnOk, "Caption", "OK"
    SetTypePickerButtonLayout btnOk, buttonGroupLeft + (TP_LAYOUT_SYNC_BUTTON_WIDTH * 2) + (TP_LAYOUT_BUTTON_GAP * 2), _
                              seedButtonTop, TP_LAYOUT_BUTTON_WIDTH

    stage = "Cancel button"
    Set btnCancel = GetOrCreateFormControl(designer, "Forms.CommandButton.1", TYPE_PICKER_CANCEL_NAME)
    TrySetObjectProperty btnCancel, "Caption", "Cancel"
    SetTypePickerButtonLayout btnCancel, buttonGroupLeft + (TP_LAYOUT_SYNC_BUTTON_WIDTH * 2) + _
                              (TP_LAYOUT_BUTTON_GAP * 2) + TP_LAYOUT_BUTTON_WIDTH + TP_LAYOUT_BUTTON_GAP, _
                              seedButtonTop, TP_LAYOUT_BUTTON_WIDTH
    Exit Sub

Fail:
    Err.Raise vbObjectError + 8120, "ReloadManager.ConfigureTypePickerFormDesigner", _
              "Stage: " & stage & " | Error " & CStr(Err.Number) & ": " & Err.Description
End Sub

Private Sub RemoveUnexpectedTypePickerDesignerControls(ByVal designer As Object)
    Dim i As Long
    Dim ctl As Object
    Dim ctlName As String

    On Error Resume Next
    For i = designer.Controls.Count - 1 To 0 Step -1
        Set ctl = designer.Controls.Item(i)
        ctlName = CStr(ctl.Name)
        If Not IsRequiredTypePickerControlName(ctlName) Then
            designer.Controls.Remove ctlName
        End If
    Next i
    On Error GoTo 0
End Sub

Private Function IsRequiredTypePickerControlName(ByVal controlName As String) As Boolean
    If StrComp(controlName, TYPE_PICKER_INFO_NAME, vbTextCompare) = 0 Then
        IsRequiredTypePickerControlName = True
    ElseIf StrComp(controlName, TYPE_PICKER_LIST_NAME, vbTextCompare) = 0 Then
        IsRequiredTypePickerControlName = True
    ElseIf StrComp(controlName, TYPE_PICKER_OK_NAME, vbTextCompare) = 0 Then
        IsRequiredTypePickerControlName = True
    ElseIf StrComp(controlName, TYPE_PICKER_CANCEL_NAME, vbTextCompare) = 0 Then
        IsRequiredTypePickerControlName = True
    ElseIf StrComp(controlName, TYPE_PICKER_SYNC_DRIFT_NAME, vbTextCompare) = 0 Then
        IsRequiredTypePickerControlName = True
    ElseIf StrComp(controlName, TYPE_PICKER_FORCE_SYNC_NAME, vbTextCompare) = 0 Then
        IsRequiredTypePickerControlName = True
    End If
End Function

Private Function GetOrCreateFormControl(ByVal designer As Object, _
                                        ByVal progId As String, _
                                        ByVal controlName As String) As Object
    On Error Resume Next
    Set GetOrCreateFormControl = GetDesignerControlByName(designer, controlName)

    If GetOrCreateFormControl Is Nothing Then
        Set GetOrCreateFormControl = designer.Controls.Add(progId, controlName, True)
    End If

    If GetOrCreateFormControl Is Nothing Then
        Err.Clear
        Set GetOrCreateFormControl = designer.Controls.Add(progId, controlName)
    End If

    If GetOrCreateFormControl Is Nothing Then
        Err.Clear
        Set GetOrCreateFormControl = designer.Controls.Add(progId)
        If Not GetOrCreateFormControl Is Nothing Then
            CallByName GetOrCreateFormControl, "Name", VbLet, controlName
        End If
    End If
    On Error GoTo 0

    If GetOrCreateFormControl Is Nothing Then
        Err.Raise vbObjectError + 8121, "ReloadManager.GetOrCreateFormControl", _
                  "Unable to add control '" & controlName & "' with ProgID '" & progId & "'."
    End If
End Function

Private Sub TrySetObjectProperty(ByVal target As Object, ByVal propName As String, ByVal propValue As Variant)
    On Error Resume Next
    CallByName target, propName, VbLet, propValue
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub SetControlBounds(ByVal target As Object, _
                             ByVal leftPos As Double, _
                             ByVal topPos As Double, _
                             ByVal widthValue As Double, _
                             ByVal heightValue As Double)
    If target Is Nothing Then Exit Sub
    TrySetObjectProperty target, "Left", leftPos
    TrySetObjectProperty target, "Top", topPos
    TrySetObjectProperty target, "Width", widthValue
    TrySetObjectProperty target, "Height", heightValue
End Sub

Private Sub ResetTypePickerInfoLabel(ByVal lbl As Object)
    If lbl Is Nothing Then Exit Sub
    TrySetObjectProperty lbl, "Caption", vbNullString
    TrySetObjectProperty lbl, "Visible", False
    SetControlBounds lbl, 0, 0, 0, 0
End Sub

Private Sub ApplyTypePickerEntriesHostStyle(ByVal entriesHost As Object)
    If entriesHost Is Nothing Then Exit Sub
    TrySetObjectProperty entriesHost, "Caption", TP_LIST_FRAME_CAPTION
    TrySetObjectProperty entriesHost, "ScrollBars", 2
    TrySetObjectProperty entriesHost, "KeepScrollBarsVisible", 0
End Sub

Private Sub SetTypePickerButtonLayout(ByVal btn As Object, _
                                      ByVal leftPos As Double, _
                                      ByVal topPos As Double, _
                                      ByVal buttonWidth As Double, _
                                      Optional ByVal visibleState As Variant)
    If btn Is Nothing Then Exit Sub
    SetControlBounds btn, leftPos, topPos, buttonWidth, TP_LAYOUT_BUTTON_HEIGHT
    If Not IsMissing(visibleState) Then
        TrySetObjectProperty btn, "Visible", CBool(visibleState)
    End If
End Sub

Private Function GetDesignerControlByName(ByVal designer As Object, ByVal controlName As String) As Object
    Dim ctl As Object

    On Error Resume Next
    Set GetDesignerControlByName = designer.Controls.Item(controlName)
    If GetDesignerControlByName Is Nothing Then
        Set GetDesignerControlByName = designer.Controls(controlName)
    End If
    On Error GoTo 0

    If GetDesignerControlByName Is Nothing Then
        For Each ctl In designer.Controls
            On Error Resume Next
            If StrComp(CStr(ctl.Name), controlName, vbTextCompare) = 0 Then
                Set GetDesignerControlByName = ctl
                On Error GoTo 0
                Exit For
            End If
            On Error GoTo 0
        Next ctl
    End If
End Function

Private Function GetFormControlByName(ByVal frm As Object, ByVal controlName As String) As Object
    Dim ctl As Object

    On Error Resume Next
    Set GetFormControlByName = frm.Controls.Item(controlName)
    If GetFormControlByName Is Nothing Then
        Set GetFormControlByName = frm.Controls(controlName)
    End If
    On Error GoTo 0

    If GetFormControlByName Is Nothing Then
        For Each ctl In frm.Controls
            On Error Resume Next
            If StrComp(CStr(ctl.Name), controlName, vbTextCompare) = 0 Then
                Set GetFormControlByName = ctl
                On Error GoTo 0
                Exit For
            End If
            On Error GoTo 0
        Next ctl
    End If
End Function

Private Sub AutoSizeTypePickerForm(ByVal frm As Object, _
                                       ByVal displayLines As Collection, _
                                       ByVal layoutPlan As Collection, _
                                       ByVal layoutContentHeight As Double)
    Dim lbl As Object
    Dim entriesHost As Object
    Dim btnSyncDrift As Object
    Dim btnForceSync As Object
    Dim btnOk As Object
    Dim btnCancel As Object
    Dim longestLineWidth As Double
    Dim lineCount As Long
    Dim visibleLineCount As Long
    Dim listWidth As Double
    Dim listHeight As Double
    Dim formWidth As Double
    Dim formHeight As Double
    Dim btnTop As Double
    Dim minListHeight As Double
    Dim minListWidth As Double
    Dim minFormWidth As Double
    Dim minFormHeight As Double
    Dim buttonSpan As Double
    Dim reservedHeight As Double
    Dim usableWidth As Double
    Dim usableHeight As Double
    Dim maxFormWidth As Double
    Dim maxFormHeight As Double
    Dim isFixedSizeMode As Boolean
    Dim entriesChromeHeight As Double
    Dim showSyncButtons As Boolean

    Set lbl = GetFormControlByName(frm, TYPE_PICKER_INFO_NAME)
    Set entriesHost = GetFormControlByName(frm, TYPE_PICKER_LIST_NAME)
    Set btnSyncDrift = GetFormControlByName(frm, TYPE_PICKER_SYNC_DRIFT_NAME)
    Set btnForceSync = GetFormControlByName(frm, TYPE_PICKER_FORCE_SYNC_NAME)
    Set btnOk = GetFormControlByName(frm, TYPE_PICKER_OK_NAME)
    Set btnCancel = GetFormControlByName(frm, TYPE_PICKER_CANCEL_NAME)

    If lbl Is Nothing Or entriesHost Is Nothing Or btnOk Is Nothing Or btnCancel Is Nothing Then Exit Sub

    lineCount = 0
    If Not layoutPlan Is Nothing Then lineCount = layoutPlan.Count
    If lineCount < 1 Then lineCount = 1

    longestLineWidth = MeasureLongestCaptionWidth(lbl, displayLines)
    entriesChromeHeight = GetControlHeightChrome(entriesHost)
    minListHeight = GetTypePickerLayoutHeightForVisibleRows(layoutPlan, 1)
    If minListHeight <= 0 Then minListHeight = TP_AUTO_ROW_HEIGHT + TP_LIST_VERTICAL_PADDING
    minListHeight = minListHeight + entriesChromeHeight
    showSyncButtons = AreTypePickerSyncButtonsVisible(btnSyncDrift, btnForceSync)
    buttonSpan = GetTypePickerButtonSpan(showSyncButtons)
    minListWidth = buttonSpan
    If (TP_LIST_MIN_CONTROL_WIDTH + TP_LIST_RIGHT_PADDING) > minListWidth Then
        minListWidth = TP_LIST_MIN_CONTROL_WIDTH + TP_LIST_RIGHT_PADDING
    End If
    reservedHeight = TP_LAYOUT_TOP_MARGIN + TP_LAYOUT_BOTTOM_MARGIN

    On Error Resume Next
    usableWidth = CDbl(Application.UsableWidth)
    usableHeight = CDbl(Application.UsableHeight)
    On Error GoTo 0
    If usableWidth <= 0 Then usableWidth = 1400
    If usableHeight <= 0 Then usableHeight = 900

    maxFormWidth = usableWidth * TP_FORM_USABLE_WIDTH_RATIO
    maxFormHeight = usableHeight * TP_FORM_USABLE_HEIGHT_RATIO

    minFormWidth = TP_LAYOUT_LEFT_MARGIN + minListWidth + TP_LAYOUT_RIGHT_MARGIN
    minFormHeight = reservedHeight + minListHeight
    isFixedSizeMode = (TP_FORCE_LIST_WIDTH > 0 Or TP_FORCE_LIST_HEIGHT > 0)

    If TP_FORCE_LIST_WIDTH > 0 Then
        listWidth = TP_FORCE_LIST_WIDTH
    Else
        listWidth = longestLineWidth + TP_LIST_HORIZONTAL_PADDING
    End If
    If listWidth < minListWidth Then listWidth = minListWidth

    If TP_FORCE_LIST_HEIGHT > 0 Then
        listHeight = TP_FORCE_LIST_HEIGHT
    Else
        visibleLineCount = lineCount
        If visibleLineCount > TP_AUTO_MAX_VISIBLE_ROWS Then visibleLineCount = TP_AUTO_MAX_VISIBLE_ROWS
        If visibleLineCount >= lineCount And layoutContentHeight > 0 Then
            listHeight = layoutContentHeight + entriesChromeHeight
        Else
            listHeight = GetTypePickerLayoutHeightForVisibleRows(layoutPlan, visibleLineCount) + entriesChromeHeight
        End If
    End If
    If listHeight < minListHeight Then listHeight = minListHeight

    btnTop = TP_LAYOUT_TOP_MARGIN + listHeight + TP_LAYOUT_BUTTON_GAP_TOP
    formWidth = TP_LAYOUT_LEFT_MARGIN + listWidth + TP_LAYOUT_RIGHT_MARGIN
    formHeight = TP_LAYOUT_TOP_MARGIN + listHeight + TP_LAYOUT_BOTTOM_MARGIN

    If isFixedSizeMode Then
        If formWidth < minFormWidth Then formWidth = minFormWidth
        If formHeight < minFormHeight Then formHeight = minFormHeight
    Else
        If formWidth < minFormWidth Then formWidth = minFormWidth
        If formWidth > maxFormWidth Then formWidth = maxFormWidth
        If formHeight < minFormHeight Then formHeight = minFormHeight
        If formHeight > maxFormHeight Then formHeight = maxFormHeight
    End If

    listWidth = formWidth - TP_LAYOUT_LEFT_MARGIN - TP_LAYOUT_RIGHT_MARGIN
    If listWidth < minListWidth Then
        listWidth = minListWidth
        formWidth = TP_LAYOUT_LEFT_MARGIN + listWidth + TP_LAYOUT_RIGHT_MARGIN
    End If

    listHeight = formHeight - reservedHeight
    If listHeight < minListHeight Then
        listHeight = minListHeight
        formHeight = reservedHeight + listHeight
    End If
    btnTop = TP_LAYOUT_TOP_MARGIN + listHeight + TP_LAYOUT_BUTTON_GAP_TOP

    SetFormClientSize frm, formWidth, formHeight

    ResetTypePickerInfoLabel lbl
    ApplyTypePickerEntriesHostStyle entriesHost

    HandleTypePickerFormResize frm
End Sub

Public Sub EnableTypePickerWindowResize(ByVal frm As Object)
    Dim hwndForm As LongPtr
    Dim styleValue As LongPtr

    hwndForm = GetUserFormHwnd(frm)
    If hwndForm = 0 Then Exit Sub

    styleValue = GetWindowLongPtrW(hwndForm, GWL_STYLE)
    styleValue = styleValue Or WS_THICKFRAME Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX
    Call SetWindowLongPtrW(hwndForm, GWL_STYLE, styleValue)
    Call SetWindowPos(hwndForm, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED)
End Sub

Public Sub HandleTypePickerFormResize(ByVal frm As Object)
    Dim lbl As Object
    Dim entriesHost As Object
    Dim btnSyncDrift As Object
    Dim btnForceSync As Object
    Dim btnOk As Object
    Dim btnCancel As Object
    Dim buttonSpan As Double
    Dim minListHeight As Double
    Dim minListWidth As Double
    Dim minFormWidth As Double
    Dim minFormHeight As Double
    Dim clientWidth As Double
    Dim clientHeight As Double
    Dim listWidth As Double
    Dim listHeight As Double
    Dim btnTop As Double
    Dim showSyncButtons As Boolean

    Set lbl = GetFormControlByName(frm, TYPE_PICKER_INFO_NAME)
    Set entriesHost = GetFormControlByName(frm, TYPE_PICKER_LIST_NAME)
    Set btnSyncDrift = GetFormControlByName(frm, TYPE_PICKER_SYNC_DRIFT_NAME)
    Set btnForceSync = GetFormControlByName(frm, TYPE_PICKER_FORCE_SYNC_NAME)
    Set btnOk = GetFormControlByName(frm, TYPE_PICKER_OK_NAME)
    Set btnCancel = GetFormControlByName(frm, TYPE_PICKER_CANCEL_NAME)

    If lbl Is Nothing Or entriesHost Is Nothing Or btnOk Is Nothing Or btnCancel Is Nothing Then Exit Sub

    showSyncButtons = AreTypePickerSyncButtonsVisible(btnSyncDrift, btnForceSync)
    buttonSpan = GetTypePickerButtonSpan(showSyncButtons)
    minListHeight = TP_AUTO_ROW_HEIGHT + TP_LIST_VERTICAL_PADDING
    minListWidth = buttonSpan
    If (TP_LIST_MIN_CONTROL_WIDTH + TP_LIST_RIGHT_PADDING) > minListWidth Then
        minListWidth = TP_LIST_MIN_CONTROL_WIDTH + TP_LIST_RIGHT_PADDING
    End If
    minFormWidth = TP_LAYOUT_LEFT_MARGIN + minListWidth + TP_LAYOUT_RIGHT_MARGIN
    minFormHeight = TP_LAYOUT_TOP_MARGIN + minListHeight + TP_LAYOUT_BOTTOM_MARGIN

    clientWidth = GetFormInsideWidth(frm)
    clientHeight = GetFormInsideHeight(frm)

    If clientWidth < minFormWidth Then clientWidth = minFormWidth
    If clientHeight < minFormHeight Then clientHeight = minFormHeight
    SetFormClientSize frm, clientWidth, clientHeight

    clientWidth = GetFormInsideWidth(frm)
    clientHeight = GetFormInsideHeight(frm)

    listWidth = clientWidth - TP_LAYOUT_LEFT_MARGIN - TP_LAYOUT_RIGHT_MARGIN
    If listWidth < minListWidth Then
        listWidth = minListWidth
        clientWidth = TP_LAYOUT_LEFT_MARGIN + listWidth + TP_LAYOUT_RIGHT_MARGIN
        SetFormClientSize frm, clientWidth, clientHeight
        clientWidth = GetFormInsideWidth(frm)
        listWidth = clientWidth - TP_LAYOUT_LEFT_MARGIN - TP_LAYOUT_RIGHT_MARGIN
    End If

    btnTop = clientHeight - TP_LAYOUT_BUTTON_HEIGHT - TP_LAYOUT_BUTTON_GAP
    listHeight = btnTop - TP_LAYOUT_BUTTON_GAP_TOP - TP_LAYOUT_TOP_MARGIN
    If listHeight < minListHeight Then
        listHeight = minListHeight
        clientHeight = TP_LAYOUT_TOP_MARGIN + listHeight + TP_LAYOUT_BOTTOM_MARGIN
        SetFormClientSize frm, clientWidth, clientHeight
        clientHeight = GetFormInsideHeight(frm)
        btnTop = clientHeight - TP_LAYOUT_BUTTON_HEIGHT - TP_LAYOUT_BUTTON_GAP
    End If

    ResetTypePickerInfoLabel lbl
    SetControlBounds entriesHost, TP_LAYOUT_LEFT_MARGIN, TP_LAYOUT_TOP_MARGIN, listWidth, listHeight
    ApplyTypePickerEntriesHostStyle entriesHost

    LayoutTypePickerActionButtons btnSyncDrift, btnForceSync, btnOk, btnCancel, _
                                  TP_LAYOUT_LEFT_MARGIN, listWidth, btnTop, showSyncButtons
End Sub

Private Function AreTypePickerSyncButtonsVisible(ByVal btnSyncDrift As Object, _
                                                 ByVal btnForceSync As Object) As Boolean
    On Error Resume Next
    If btnSyncDrift Is Nothing Or btnForceSync Is Nothing Then Exit Function
    AreTypePickerSyncButtonsVisible = CBool(CallByName(btnSyncDrift, "Visible", VbGet)) Or _
                                     CBool(CallByName(btnForceSync, "Visible", VbGet))
    Err.Clear
    On Error GoTo 0
End Function

Private Function GetTypePickerButtonSpan(ByVal includeSyncButtons As Boolean) As Double
    If includeSyncButtons Then
        GetTypePickerButtonSpan = (TP_LAYOUT_SYNC_BUTTON_WIDTH * 2) + (TP_LAYOUT_BUTTON_WIDTH * 2) + (TP_LAYOUT_BUTTON_GAP * 3)
    Else
        GetTypePickerButtonSpan = (TP_LAYOUT_BUTTON_WIDTH * 2) + TP_LAYOUT_BUTTON_GAP
    End If
End Function

Private Sub LayoutTypePickerActionButtons(ByVal btnSyncDrift As Object, _
                                          ByVal btnForceSync As Object, _
                                          ByVal btnOk As Object, _
                                          ByVal btnCancel As Object, _
                                          ByVal listLeft As Double, _
                                          ByVal listWidth As Double, _
                                          ByVal btnTop As Double, _
                                          ByVal includeSyncButtons As Boolean)
    Dim rightEdge As Double
    Dim okLeft As Double
    Dim cancelLeft As Double
    Dim forceLeft As Double
    Dim driftLeft As Double

    rightEdge = listLeft + listWidth
    cancelLeft = rightEdge - TP_LAYOUT_BUTTON_WIDTH
    okLeft = cancelLeft - TP_LAYOUT_BUTTON_GAP - TP_LAYOUT_BUTTON_WIDTH

    SetTypePickerButtonLayout btnCancel, cancelLeft, btnTop, TP_LAYOUT_BUTTON_WIDTH
    SetTypePickerButtonLayout btnOk, okLeft, btnTop, TP_LAYOUT_BUTTON_WIDTH

    If includeSyncButtons Then
        forceLeft = okLeft - TP_LAYOUT_BUTTON_GAP - TP_LAYOUT_SYNC_BUTTON_WIDTH
        driftLeft = forceLeft - TP_LAYOUT_BUTTON_GAP - TP_LAYOUT_SYNC_BUTTON_WIDTH

        SetTypePickerButtonLayout btnForceSync, forceLeft, btnTop, TP_LAYOUT_SYNC_BUTTON_WIDTH, True
        SetTypePickerButtonLayout btnSyncDrift, driftLeft, btnTop, TP_LAYOUT_SYNC_BUTTON_WIDTH, True
    Else
        If Not btnForceSync Is Nothing Then TrySetObjectProperty btnForceSync, "Visible", False
        If Not btnSyncDrift Is Nothing Then TrySetObjectProperty btnSyncDrift, "Visible", False
    End If
End Sub

Private Function GetFormInsideWidth(ByVal frm As Object) As Double
    On Error Resume Next
    GetFormInsideWidth = CDbl(CallByName(frm, "InsideWidth", VbGet))
    If Err.Number <> 0 Or GetFormInsideWidth <= 0 Then
        Err.Clear
        GetFormInsideWidth = CDbl(CallByName(frm, "Width", VbGet))
    End If
    Err.Clear
    On Error GoTo 0
End Function

Private Function GetFormInsideHeight(ByVal frm As Object) As Double
    On Error Resume Next
    GetFormInsideHeight = CDbl(CallByName(frm, "InsideHeight", VbGet))
    If Err.Number <> 0 Or GetFormInsideHeight <= 0 Then
        Err.Clear
        GetFormInsideHeight = CDbl(CallByName(frm, "Height", VbGet))
    End If
    Err.Clear
    On Error GoTo 0
End Function

Private Function GetControlInsideWidth(ByVal ctl As Object) As Double
    On Error Resume Next
    GetControlInsideWidth = CDbl(CallByName(ctl, "InsideWidth", VbGet))
    If Err.Number <> 0 Or GetControlInsideWidth <= 0 Then
        Err.Clear
        GetControlInsideWidth = CDbl(CallByName(ctl, "Width", VbGet))
    End If
    Err.Clear
    On Error GoTo 0
End Function

Private Function GetControlInsideHeight(ByVal ctl As Object) As Double
    On Error Resume Next
    GetControlInsideHeight = CDbl(CallByName(ctl, "InsideHeight", VbGet))
    If Err.Number <> 0 Or GetControlInsideHeight <= 0 Then
        Err.Clear
        GetControlInsideHeight = CDbl(CallByName(ctl, "Height", VbGet))
    End If
    Err.Clear
    On Error GoTo 0
End Function

Private Function GetControlHeightChrome(ByVal ctl As Object) As Double
    Dim totalHeight As Double
    Dim insideHeight As Double

    On Error Resume Next
    totalHeight = CDbl(CallByName(ctl, "Height", VbGet))
    insideHeight = CDbl(CallByName(ctl, "InsideHeight", VbGet))
    If Err.Number <> 0 Then
        Err.Clear
        GetControlHeightChrome = 0
    Else
        GetControlHeightChrome = totalHeight - insideHeight
        If GetControlHeightChrome < 0 Then GetControlHeightChrome = 0
    End If
    Err.Clear
    On Error GoTo 0
End Function

Private Sub SetFormClientSize(ByVal frm As Object, ByVal insideWidth As Double, ByVal insideHeight As Double)
    Dim chromeWidth As Double
    Dim chromeHeight As Double
    Dim targetWidth As Double
    Dim targetHeight As Double

    If insideWidth < 1 Then insideWidth = 1
    If insideHeight < 1 Then insideHeight = 1

    On Error Resume Next
    chromeWidth = CDbl(CallByName(frm, "Width", VbGet)) - CDbl(CallByName(frm, "InsideWidth", VbGet))
    If Err.Number <> 0 Then
        Err.Clear
        chromeWidth = 0
    End If

    chromeHeight = CDbl(CallByName(frm, "Height", VbGet)) - CDbl(CallByName(frm, "InsideHeight", VbGet))
    If Err.Number <> 0 Then
        Err.Clear
        chromeHeight = 0
    End If

    targetWidth = insideWidth + chromeWidth
    targetHeight = insideHeight + chromeHeight

    If targetWidth < insideWidth Then targetWidth = insideWidth
    If targetHeight < insideHeight Then targetHeight = insideHeight

    CallByName frm, "Width", VbLet, targetWidth
    CallByName frm, "Height", VbLet, targetHeight
    Err.Clear
    On Error GoTo 0
End Sub

Private Function GetUserFormHwnd(ByVal frm As Object) As LongPtr
    Dim captionText As String

    On Error Resume Next
    captionText = CStr(frm.Caption)
    On Error GoTo 0
    If Len(captionText) = 0 Then Exit Function

    GetUserFormHwnd = FindWindowW(StrPtr("ThunderDFrame"), StrPtr(captionText))
    If GetUserFormHwnd = 0 Then
        GetUserFormHwnd = FindWindowW(0, StrPtr(captionText))
    End If
End Function

Private Function MeasureLongestCaptionWidth(ByVal measureLabel As Object, ByVal displayLines As Collection) As Double
    Dim originalCaption As String
    Dim originalAutoSize As Boolean
    Dim originalWordWrap As Boolean
    Dim originalWidth As Double
    Dim originalHeight As Double
    Dim lineText As Variant
    Dim widthNow As Double

    On Error Resume Next
    originalCaption = CStr(measureLabel.Caption)
    originalAutoSize = CBool(measureLabel.AutoSize)
    originalWordWrap = CBool(measureLabel.WordWrap)
    originalWidth = CDbl(measureLabel.Width)
    originalHeight = CDbl(measureLabel.Height)

    measureLabel.WordWrap = False
    measureLabel.AutoSize = True

    For Each lineText In displayLines
        measureLabel.Caption = CStr(lineText)
        widthNow = CDbl(measureLabel.Width)
        If widthNow > MeasureLongestCaptionWidth Then
            MeasureLongestCaptionWidth = widthNow
        End If
    Next lineText

    measureLabel.Caption = originalCaption
    measureLabel.AutoSize = originalAutoSize
    measureLabel.WordWrap = originalWordWrap
    measureLabel.Width = originalWidth
    measureLabel.Height = originalHeight
    On Error GoTo 0
End Function

Private Sub EnsureTypePickerFormCode(ByVal codeModule As Object)
    Dim codeText As String

    codeText = ""
    codeText = codeText & "Option Explicit" & vbCrLf
    codeText = codeText & vbCrLf
    codeText = codeText & "Private mWindowResizeEnabled As Boolean" & vbCrLf
    codeText = codeText & vbCrLf
    codeText = codeText & "Private Sub UserForm_Initialize()" & vbCrLf
    codeText = codeText & "    Me.Tag = """ & TYPE_PICKER_TAG_CANCEL & """" & vbCrLf
    codeText = codeText & "    On Error Resume Next" & vbCrLf
    codeText = codeText & "    Me.lblInfo.Caption = """"" & vbCrLf
    codeText = codeText & "    Me.lblInfo.Visible = False" & vbCrLf
    codeText = codeText & "    Me.lstEntries.Caption = "" """ & vbCrLf
    codeText = codeText & "    On Error GoTo 0" & vbCrLf
    codeText = codeText & "End Sub" & vbCrLf
    codeText = codeText & vbCrLf
    codeText = codeText & "Private Sub UserForm_Activate()" & vbCrLf
    codeText = codeText & "    On Error Resume Next" & vbCrLf
    codeText = codeText & "    If Not mWindowResizeEnabled Then" & vbCrLf
    codeText = codeText & "        ReloadManager.EnableTypePickerWindowResize Me" & vbCrLf
    codeText = codeText & "        mWindowResizeEnabled = True" & vbCrLf
    codeText = codeText & "    End If" & vbCrLf
    codeText = codeText & "    ReloadManager.HandleTypePickerFormResize Me" & vbCrLf
    codeText = codeText & "    On Error GoTo 0" & vbCrLf
    codeText = codeText & "End Sub" & vbCrLf
    codeText = codeText & vbCrLf
    codeText = codeText & "Private Sub UserForm_Resize()" & vbCrLf
    codeText = codeText & "    On Error Resume Next" & vbCrLf
    codeText = codeText & "    ReloadManager.HandleTypePickerFormResize Me" & vbCrLf
    codeText = codeText & "    On Error GoTo 0" & vbCrLf
    codeText = codeText & "End Sub" & vbCrLf
    codeText = codeText & vbCrLf
    codeText = codeText & "Private Sub btnOK_Click()" & vbCrLf
    codeText = codeText & "    Me.Tag = """ & TYPE_PICKER_TAG_OK & """" & vbCrLf
    codeText = codeText & "    Me.Hide" & vbCrLf
    codeText = codeText & "End Sub" & vbCrLf
    codeText = codeText & vbCrLf
    codeText = codeText & "Private Sub " & TYPE_PICKER_SYNC_DRIFT_NAME & "_Click()" & vbCrLf
    codeText = codeText & "    Me.Tag = """ & TYPE_PICKER_TAG_SYNC_DRIFT & """" & vbCrLf
    codeText = codeText & "    Me.Hide" & vbCrLf
    codeText = codeText & "End Sub" & vbCrLf
    codeText = codeText & vbCrLf
    codeText = codeText & "Private Sub " & TYPE_PICKER_FORCE_SYNC_NAME & "_Click()" & vbCrLf
    codeText = codeText & "    Me.Tag = """ & TYPE_PICKER_TAG_FORCE_SYNC & """" & vbCrLf
    codeText = codeText & "    Me.Hide" & vbCrLf
    codeText = codeText & "End Sub" & vbCrLf
    codeText = codeText & vbCrLf
    codeText = codeText & "Private Sub btnCancel_Click()" & vbCrLf
    codeText = codeText & "    Me.Tag = """ & TYPE_PICKER_TAG_CANCEL & """" & vbCrLf
    codeText = codeText & "    Me.Hide" & vbCrLf
    codeText = codeText & "End Sub" & vbCrLf
    codeText = codeText & vbCrLf
    codeText = codeText & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
    codeText = codeText & "    If CloseMode = 0 Then" & vbCrLf
    codeText = codeText & "        Cancel = True" & vbCrLf
    codeText = codeText & "        Me.Tag = """ & TYPE_PICKER_TAG_CANCEL & """" & vbCrLf
    codeText = codeText & "        Me.Hide" & vbCrLf
    codeText = codeText & "    End If" & vbCrLf
    codeText = codeText & "End Sub"

    If codeModule.CountOfLines > 0 Then
        codeModule.DeleteLines 1, codeModule.CountOfLines
    End If
    codeModule.InsertLines 1, codeText
End Sub

Private Function GetBasFileNamesInTypeSource(ByVal typeId As String) As Collection
    Dim srcPath As String
    Dim f As String
    Dim names() As String
    Dim i As Long
    Dim count As Long
    Dim results As New Collection

    srcPath = EnsureTrailingPathSeparator(REPO_ROOT & Application.PathSeparator & _
                                          NormalizeTypeId(typeId) & Application.PathSeparator & "src")

    f = Dir(srcPath & "*.bas")
    Do While Len(f) > 0
        count = count + 1
        ReDim Preserve names(1 To count)
        If Len(f) > 4 And StrComp(Right$(f, 4), ".bas", vbTextCompare) = 0 Then
            names(count) = Left$(f, Len(f) - 4)
        Else
            names(count) = f
        End If
        f = Dir()
    Loop

    If count > 1 Then
        SortStringsCaseInsensitive names
    End If

    For i = 1 To count
        results.Add names(i)
    Next i

    Set GetBasFileNamesInTypeSource = results
End Function

Private Function GetCommonFolderDepth(ByRef leftParts() As String, _
                                      ByRef rightParts() As String, _
                                      Optional ByVal includeLastPart As Boolean = False) As Long
    Dim leftLastFolder As Long
    Dim rightLastFolder As Long
    Dim maxFolderIndex As Long
    Dim i As Long

    If includeLastPart Then
        leftLastFolder = UBound(leftParts)
        rightLastFolder = UBound(rightParts)
    Else
        leftLastFolder = UBound(leftParts) - 1
        rightLastFolder = UBound(rightParts) - 1
    End If

    If leftLastFolder < 0 Or rightLastFolder < 0 Then Exit Function

    maxFolderIndex = leftLastFolder
    If rightLastFolder < maxFolderIndex Then
        maxFolderIndex = rightLastFolder
    End If

    For i = 0 To maxFolderIndex
        If StrComp(leftParts(i), rightParts(i), vbTextCompare) <> 0 Then
            GetCommonFolderDepth = i
            Exit Function
        End If
    Next i

    GetCommonFolderDepth = maxFolderIndex + 1
End Function


'================== CONTEXT & PROMPT HELPERS ==================

Private Function PrepareSyncContext(ByVal wb As Workbook, _
                                    ByRef typeSummary As String, _
                                    ByRef srcPaths As Collection, _
                                    ByRef changedSrcPaths As Collection, _
                                    ByRef driftTypeSummary As String, _
                                    ByRef repoSig As String, _
                                    ByRef wbSig As String, _
                                    ByRef typeSigMapText As String, _
                                    ByRef reason As String) As Boolean
    Dim typeMap As Object
    Dim sortedTypes() As String
    Dim signatureItems() As String
    Dim srcPathMap As Object
    Dim changedSrcPathMap As Object
    Dim currentTypeSigMap As Object
    Dim previousTypeSigMap As Object
    Dim typeToSrcPath As Object
    Dim previousTypeSigMapText As String
    Dim srcPath As String
    Dim srcPathItem As Variant
    Dim folderSig As String
    Dim normalizedTypeId As String
    Dim i As Long

    If wb Is Nothing Then
        reason = "No workbook provided for sync."
        Exit Function
    End If

    If IsAddInWorkbook(wb) Then
        reason = "The selected workbook is an add-in and is excluded from sync."
        Exit Function
    End If

    If wb.ReadOnly Then
        reason = "Workbook '" & wb.Name & "' is read-only and cannot import modules."
        Exit Function
    End If

    Set typeMap = GetWorkbookTypeMap(wb)
    If typeMap Is Nothing Or typeMap.count = 0 Then
        reason = "Workbook '" & wb.Name & "' is missing assigned workbook types." & vbCrLf & _
                 "Use ManageActiveWorkbookTypes to assign at least one type first."
        Exit Function
    End If

    sortedTypes = GetSortedTypeIds(typeMap)
    typeSummary = JoinStringArray(sortedTypes, ", ")

    Set srcPaths = New Collection
    Set changedSrcPaths = New Collection
    Set srcPathMap = CreateObject("Scripting.Dictionary")
    srcPathMap.CompareMode = vbTextCompare
    Set changedSrcPathMap = CreateObject("Scripting.Dictionary")
    changedSrcPathMap.CompareMode = vbTextCompare
    Set currentTypeSigMap = CreateObject("Scripting.Dictionary")
    currentTypeSigMap.CompareMode = vbTextCompare
    Set typeToSrcPath = CreateObject("Scripting.Dictionary")
    typeToSrcPath.CompareMode = vbTextCompare
    ReDim signatureItems(1 To UBound(sortedTypes))

    For i = LBound(sortedTypes) To UBound(sortedTypes)
        If Not ResolveTypeSourcePath(sortedTypes(i), srcPath, reason) Then Exit Function

        folderSig = ComputeFolderSignature(srcPath, "*.bas")
        If Len(folderSig) = 0 Then
            reason = MSG_SOURCE_UNAVAILABLE & vbCrLf & MSG_REQUEST_UPDATE
            Exit Function
        End If

        signatureItems(i) = LCase$(sortedTypes(i)) & "|" & folderSig
        normalizedTypeId = LCase$(sortedTypes(i))
        currentTypeSigMap(normalizedTypeId) = folderSig
        typeToSrcPath(normalizedTypeId) = srcPath

        If Not srcPathMap.Exists(srcPath) Then
            srcPathMap(srcPath) = True
            srcPaths.Add srcPath
        End If
    Next i

    If Not ValidateNoModuleConflicts(srcPaths, reason) Then
        Exit Function
    End If

    repoSig = ComputeSignatureFromItems(signatureItems)

    wbSig = GetDocProp(wb, PROP_SRC_SIG)
    typeSigMapText = JoinStringArray(signatureItems, vbLf)

    previousTypeSigMapText = GetDocProp(wb, PROP_SRC_SIG_MAP)
    If Len(previousTypeSigMapText) = 0 Then
        If StrComp(repoSig, wbSig, vbBinaryCompare) <> 0 Then
            driftTypeSummary = typeSummary
            For Each srcPathItem In srcPaths
                changedSrcPaths.Add CStr(srcPathItem)
            Next srcPathItem
        End If
    Else
        Set previousTypeSigMap = ParseTypeSignatureMap(previousTypeSigMapText)

        For i = LBound(sortedTypes) To UBound(sortedTypes)
            normalizedTypeId = LCase$(sortedTypes(i))
            If (Not previousTypeSigMap.Exists(normalizedTypeId)) Or _
               (StrComp(CStr(previousTypeSigMap(normalizedTypeId)), CStr(currentTypeSigMap(normalizedTypeId)), vbBinaryCompare) <> 0) Then
                If Len(driftTypeSummary) > 0 Then driftTypeSummary = driftTypeSummary & ", "
                driftTypeSummary = driftTypeSummary & sortedTypes(i)

                srcPath = CStr(typeToSrcPath(normalizedTypeId))
                If Not changedSrcPathMap.Exists(srcPath) Then
                    changedSrcPathMap(srcPath) = True
                    changedSrcPaths.Add srcPath
                End If
            End If
        Next i
    End If

    PrepareSyncContext = True
End Function

Private Function ParseTypeSignatureMap(ByVal rawMap As String) As Object
    Dim d As Object
    Dim lines() As String
    Dim lineText As String
    Dim sepPos As Long
    Dim i As Long

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    If Len(rawMap) > 0 Then
        lines = Split(rawMap, vbLf)
        For i = LBound(lines) To UBound(lines)
            lineText = Replace$(Trim$(CStr(lines(i))), vbCr, "")
            If Len(lineText) > 0 Then
                sepPos = InStr(1, lineText, "|", vbBinaryCompare)
                If sepPos > 1 Then
                    d(LCase$(Left$(lineText, sepPos - 1))) = Mid$(lineText, sepPos + 1)
                End If
            End If
        Next i
    End If

    Set ParseTypeSignatureMap = d
End Function

Private Function ConfirmDriftImport(ByVal wb As Workbook, ByVal typeId As String) As Boolean
    Dim msg As String
    msg = "Detected VBA source drift for workbook '" & wb.Name & "'." & vbCrLf & _
          "Type ID(s): " & typeId & vbCrLf & vbCrLf & _
          "Import latest .bas modules from repo now?" & vbCrLf & vbCrLf & _
          "Yes = import now" & vbCrLf & _
          "No = leave workbook unchanged"

    ConfirmDriftImport = (MsgBox(msg, vbYesNo + vbQuestion, SYNC_DIALOG_TITLE) = vbYes)
End Function

Private Function GetWorkbookTypeMap(ByVal wb As Workbook) As Object
    Dim d As Object
    Dim raw As String
    Dim parsed() As String
    Dim i As Long
    Dim normalized As String

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    raw = Trim$(GetDocProp(wb, PROP_TYPE_ID))
    If Len(raw) > 0 Then
        parsed = ParseTypeIdList(raw)
        For i = LBound(parsed) To UBound(parsed)
            normalized = NormalizeTypeId(parsed(i))
            If Len(normalized) > 0 Then
                d(normalized) = normalized
            End If
        Next i
    End If

    Set GetWorkbookTypeMap = d
End Function

Private Sub SaveWorkbookTypeMap(ByVal wb As Workbook, ByVal typeMap As Object)
    Dim sorted() As String

    If typeMap Is Nothing Or typeMap.count = 0 Then
        SetDocProp wb, PROP_TYPE_ID, ""
        Exit Sub
    End If

    sorted = GetSortedTypeIds(typeMap)
    SetDocProp wb, PROP_TYPE_ID, JoinStringArray(sorted, TYPE_LIST_DELIMITER)
End Sub

Private Function GetSortedTypeIds(ByVal typeMap As Object) As String()
    Dim values() As String
    Dim i As Long
    Dim key As Variant

    ReDim values(1 To typeMap.count)
    i = 0
    For Each key In typeMap.Keys
        i = i + 1
        values(i) = CStr(typeMap(key))
    Next key

    If typeMap.count > 1 Then
        QuickSort values, LBound(values), UBound(values)
    End If

    GetSortedTypeIds = values
End Function

Private Function ParseTypeIdList(ByVal raw As String) As String()
    Dim cleaned As String
    Dim parts() As String
    Dim output() As String
    Dim token As String
    Dim i As Long
    Dim count As Long

    cleaned = raw
    cleaned = Replace$(cleaned, vbCrLf, TYPE_LIST_DELIMITER)
    cleaned = Replace$(cleaned, vbCr, TYPE_LIST_DELIMITER)
    cleaned = Replace$(cleaned, vbLf, TYPE_LIST_DELIMITER)
    cleaned = Replace$(cleaned, ",", TYPE_LIST_DELIMITER)
    cleaned = Replace$(cleaned, "|", TYPE_LIST_DELIMITER)

    parts = Split(cleaned, TYPE_LIST_DELIMITER)
    count = 0
    ReDim output(1 To UBound(parts) + 1)

    For i = LBound(parts) To UBound(parts)
        token = Trim$(parts(i))
        If Len(token) > 0 Then
            count = count + 1
            output(count) = token
        End If
    Next i

    If count = 0 Then
        ReDim output(1 To 1)
        output(1) = ""
    Else
        ReDim Preserve output(1 To count)
    End If

    ParseTypeIdList = output
End Function

Private Function JoinStringArray(ByRef arr() As String, ByVal delimiter As String) As String
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If i > LBound(arr) Then
            JoinStringArray = JoinStringArray & delimiter
        End If
        JoinStringArray = JoinStringArray & arr(i)
    Next i
End Function

Private Function ComputeSignatureFromItems(ByRef items() As String) As String
    Dim sig As String
    Dim i As Long

    If UBound(items) > LBound(items) Then
        QuickSort items, LBound(items), UBound(items)
    End If

    For i = LBound(items) To UBound(items)
        sig = sig & items(i) & vbLf
    Next i

    ComputeSignatureFromItems = CStr(SimpleHash32(sig)) & ":" & CStr(UBound(items) - LBound(items) + 1)
End Function

Private Function ValidateNoModuleConflicts(ByVal srcPaths As Collection, ByRef reason As String) As Boolean
    Dim moduleMap As Object
    Dim srcPath As Variant
    Dim f As String
    Dim compName As String
    Dim effectiveName As String
    Dim basPath As String
    Dim key As String

    Set moduleMap = CreateObject("Scripting.Dictionary")
    moduleMap.CompareMode = vbTextCompare

    For Each srcPath In srcPaths
        f = Dir(CStr(srcPath) & "*.bas")
        Do While Len(f) > 0
            compName = Left$(f, Len(f) - 4)
            basPath = CStr(srcPath) & f
            effectiveName = GetBasModuleComponentName(basPath, compName)
            key = LCase$(effectiveName)

            If moduleMap.Exists(key) Then
                reason = "Module name conflict detected: '" & effectiveName & "'." & vbCrLf & _
                         "Found in both:" & vbCrLf & _
                         CStr(moduleMap(key)) & vbCrLf & _
                         basPath
                Exit Function
            End If

            moduleMap(key) = basPath
            f = Dir()
        Loop
    Next srcPath

    ValidateNoModuleConflicts = True
End Function

Private Function CanAccessVBProject(ByVal wb As Workbook, ByRef reason As String) As Boolean
    On Error GoTo Fail
    Dim componentCount As Long
    componentCount = wb.VBProject.VBComponents.count
    CanAccessVBProject = True
    Exit Function

Fail:
    reason = "Cannot access VBA project for workbook '" & wb.Name & "'." & vbCrLf & _
             "Enable 'Trust access to the VBA project object model' in Excel Trust Center."
    CanAccessVBProject = False
End Function

Private Function ResolveTypeSourcePath(ByVal typeId As String, _
                                       ByRef srcPath As String, _
                                       ByRef reason As String) As Boolean
    On Error GoTo Fail

    Dim fso As Object
    Dim rootFolder As Object
    Dim matches As Object
    Dim visited As Object
    Dim key As Variant
    Dim normalizedTypeId As String
    Dim directPath As String
    Dim leafTypeId As String
    Dim pos As Long

    normalizedTypeId = NormalizeTypeId(typeId)
    If Len(normalizedTypeId) = 0 Then
        reason = "Type ID is empty."
        Exit Function
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(REPO_ROOT) Then
        reason = MSG_SOURCE_UNAVAILABLE & vbCrLf & MSG_REQUEST_UPDATE
        Exit Function
    End If

    directPath = EnsureTrailingPathSeparator(REPO_ROOT & Application.PathSeparator & _
                                             normalizedTypeId & Application.PathSeparator & "src")
    If fso.FolderExists(directPath) Then
        srcPath = directPath
        ResolveTypeSourcePath = True
        Exit Function
    End If

    pos = InStrRev(normalizedTypeId, Application.PathSeparator)
    If pos > 0 Then
        leafTypeId = Mid$(normalizedTypeId, pos + 1)
    Else
        leafTypeId = normalizedTypeId
    End If

    Set rootFolder = fso.GetFolder(REPO_ROOT)
    Set matches = CreateObject("Scripting.Dictionary")
    matches.CompareMode = vbTextCompare
    Set visited = CreateObject("Scripting.Dictionary")
    visited.CompareMode = vbTextCompare

    CollectTypeSourceMatches rootFolder, leafTypeId, fso, matches, visited, 0

    If matches.count = 0 Then
        reason = "Type ID '" & typeId & "' was not found in local VBA source folders." & vbCrLf & _
                 "Use ManageActiveWorkbookTypes to assign a valid type."
        Exit Function
    End If

    If matches.count > 1 Then
        reason = "Type ID '" & typeId & "' matches multiple folders." & vbCrLf & _
                 "Use a more specific path-like type ID (example: Assorted_Tools\Formula_Freezer)."
        Exit Function
    End If

    For Each key In matches.Keys
        srcPath = CStr(matches(key))
        Exit For
    Next key

    ResolveTypeSourcePath = True
    Exit Function

Fail:
    reason = MSG_SOURCE_UNAVAILABLE & vbCrLf & MSG_REQUEST_UPDATE
End Function

Private Sub CollectTypeSourceMatches(ByVal currentFolder As Object, _
                                     ByVal leafTypeId As String, _
                                     ByVal fso As Object, _
                                     ByRef matches As Object, _
                                     ByRef visited As Object, _
                                     ByVal depth As Long)
    Dim childFolder As Object
    Dim srcPath As String

    If currentFolder Is Nothing Then Exit Sub
    If depth > MAX_FOLDER_SCAN_DEPTH Then Exit Sub
    If Not MarkFolderSeen(currentFolder, visited) Then Exit Sub

    For Each childFolder In currentFolder.SubFolders
        If StrComp(childFolder.Name, leafTypeId, vbTextCompare) = 0 Then
            srcPath = EnsureTrailingPathSeparator(childFolder.Path & Application.PathSeparator & "src")
            If fso.FolderExists(srcPath) Then
                matches(childFolder.Path) = srcPath
            End If
        End If

        CollectTypeSourceMatches childFolder, leafTypeId, fso, matches, visited, depth + 1
    Next childFolder
End Sub

Private Function NormalizeTypeId(ByVal typeId As String) As String
    Dim t As String
    t = Trim$(typeId)
    t = Replace$(t, "/", Application.PathSeparator)
    t = Replace$(t, "\\", Application.PathSeparator)
    Do While InStr(t, Application.PathSeparator & Application.PathSeparator) > 0
        t = Replace$(t, Application.PathSeparator & Application.PathSeparator, Application.PathSeparator)
    Loop
    If Left$(t, 1) = Application.PathSeparator Then t = Mid$(t, 2)
    If Right$(t, 1) = Application.PathSeparator Then t = Left$(t, Len(t) - 1)
    NormalizeTypeId = t
End Function

Private Function EnsureTrailingPathSeparator(ByVal path As String) As String
    If Len(path) = 0 Then
        EnsureTrailingPathSeparator = path
    ElseIf Right$(path, 1) = Application.PathSeparator Then
        EnsureTrailingPathSeparator = path
    Else
        EnsureTrailingPathSeparator = path & Application.PathSeparator
    End If
End Function

Private Function GetAvailableTypeIdsCollection() As Collection
    On Error GoTo Fail

    Dim fso As Object
    Dim rootFolder As Object
    Dim seen As Object
    Dim visited As Object
    Dim filtered As Object
    Dim values() As String
    Dim i As Long
    Dim key As Variant
    Dim results As New Collection

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(REPO_ROOT) Then
        Set GetAvailableTypeIdsCollection = results
        Exit Function
    End If

    Set rootFolder = fso.GetFolder(REPO_ROOT)
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    Set visited = CreateObject("Scripting.Dictionary")
    visited.CompareMode = vbTextCompare
    Set filtered = CreateObject("Scripting.Dictionary")
    filtered.CompareMode = vbTextCompare

    CollectAvailableTypeIdsRecursive rootFolder, fso, seen, visited, 0

    If seen.count > 0 Then
        For Each key In seen.Keys
            If IsTypeIdAssignableCandidate(CStr(key), fso) Then
                filtered(CStr(key)) = True
            End If
        Next key
    End If

    If filtered.count > 0 Then
        ReDim values(1 To filtered.count)
        i = 0
        For Each key In filtered.Keys
            i = i + 1
            values(i) = CStr(key)
        Next key

        If UBound(values) > LBound(values) Then
            SortStringsCaseInsensitive values
        End If

        For i = LBound(values) To UBound(values)
            results.Add values(i)
        Next i
    End If

    Set GetAvailableTypeIdsCollection = results
    Exit Function

Fail:
    Set GetAvailableTypeIdsCollection = New Collection
End Function

Private Function IsTypeIdAssignableCandidate(ByVal typeId As String, ByVal fso As Object) As Boolean
    Dim normalizedTypeId As String
    Dim addInTypeId As String

    normalizedTypeId = NormalizeTypeId(typeId)
    If Len(normalizedTypeId) = 0 Then Exit Function

    addInTypeId = GetGlobalAddInTypeId()
    If Len(addInTypeId) > 0 Then
        If StrComp(normalizedTypeId, addInTypeId, vbTextCompare) = 0 Then Exit Function
    End If

    If Not TypeSourceContainsBasFiles(normalizedTypeId, fso) Then Exit Function

    IsTypeIdAssignableCandidate = True
End Function

Private Function GetGlobalAddInTypeId() As String
    Dim addInTypeId As String

    addInTypeId = NormalizeTypeId(Replace$(REPO_ADDIN_SRC_RELATIVE_PATH, "\", Application.PathSeparator))
    If Len(addInTypeId) > 3 Then
        If StrComp(Right$(addInTypeId, 3), "src", vbTextCompare) = 0 Then
            addInTypeId = NormalizeTypeId(Left$(addInTypeId, Len(addInTypeId) - 3))
        End If
    End If

    GetGlobalAddInTypeId = addInTypeId
End Function

Private Function TypeSourceContainsBasFiles(ByVal typeId As String, ByVal fso As Object) As Boolean
    Dim srcPath As String

    If fso Is Nothing Then Exit Function

    srcPath = EnsureTrailingPathSeparator(REPO_ROOT & Application.PathSeparator & _
                                          NormalizeTypeId(typeId) & Application.PathSeparator & "src")
    If Not fso.FolderExists(srcPath) Then Exit Function

    TypeSourceContainsBasFiles = (Len(Dir(srcPath & "*.bas")) > 0)
End Function

Private Sub SortStringsCaseInsensitive(ByRef values() As String)
    Dim i As Long
    Dim j As Long
    Dim temp As String

    For i = LBound(values) To UBound(values) - 1
        For j = i + 1 To UBound(values)
            If StrComp(values(i), values(j), vbTextCompare) > 0 Then
                temp = values(i)
                values(i) = values(j)
                values(j) = temp
            End If
        Next j
    Next i
End Sub

Private Sub CollectAvailableTypeIdsRecursive(ByVal currentFolder As Object, _
                                             ByVal fso As Object, _
                                             ByRef seen As Object, _
                                             ByRef visited As Object, _
                                             ByVal depth As Long)
    Dim subFolder As Object
    Dim srcFolderPath As String
    Dim typeId As String
    Dim rootPrefix As String

    If currentFolder Is Nothing Then Exit Sub
    If depth > MAX_FOLDER_SCAN_DEPTH Then Exit Sub
    If Not MarkFolderSeen(currentFolder, visited) Then Exit Sub

    rootPrefix = EnsureTrailingPathSeparator(REPO_ROOT)

    For Each subFolder In currentFolder.SubFolders
        srcFolderPath = subFolder.Path & Application.PathSeparator & "src"
        If fso.FolderExists(srcFolderPath) Then
            typeId = subFolder.Path
            If StrComp(Left$(typeId, Len(rootPrefix)), rootPrefix, vbTextCompare) = 0 Then
                typeId = Mid$(typeId, Len(rootPrefix) + 1)
            End If
            If Len(typeId) > 0 Then
                If Not seen.Exists(typeId) Then seen(typeId) = True
            End If
        End If

        CollectAvailableTypeIdsRecursive subFolder, fso, seen, visited, depth + 1
    Next subFolder
End Sub

Private Function MarkFolderSeen(ByVal folderObj As Object, ByRef visited As Object) As Boolean
    Dim folderPath As String
    Dim pathKey As String
    Dim attrs As Long

    If folderObj Is Nothing Then Exit Function
    If visited Is Nothing Then Exit Function

    On Error Resume Next
    folderPath = CStr(folderObj.Path)
    attrs = CLng(folderObj.Attributes)
    On Error GoTo 0

    If Len(folderPath) = 0 Then Exit Function
    pathKey = LCase$(EnsureTrailingPathSeparator(folderPath))

    If visited.Exists(pathKey) Then Exit Function
    visited(pathKey) = True

    ' Avoid traversing reparse-point folders (junctions/symlinks) to prevent cycles.
    If (attrs And FILE_ATTR_REPARSE_POINT) <> 0 Then Exit Function

    MarkFolderSeen = True
End Function

Private Function GetForegroundWorkbook() As Workbook
    Dim wb As Workbook

    On Error Resume Next
    Set wb = Application.ActiveWorkbook
    On Error GoTo 0
    If Not wb Is Nothing Then
        If Not IsAddInWorkbook(wb) Then
            Set GetForegroundWorkbook = wb
            Exit Function
        End If
    End If

    On Error Resume Next
    If Not Application.ActiveSheet Is Nothing Then
        Set wb = Application.ActiveSheet.Parent
    End If
    On Error GoTo 0
    If Not wb Is Nothing Then
        If Not IsAddInWorkbook(wb) Then
            Set GetForegroundWorkbook = wb
            Exit Function
        End If
    End If

    For Each wb In Application.Workbooks
        If Not IsAddInWorkbook(wb) Then
            Set GetForegroundWorkbook = wb
            Exit Function
        End If
    Next wb
End Function

Private Function IsAddInWorkbook(ByVal wb As Workbook) As Boolean
    On Error Resume Next
    If wb Is Nothing Then
        IsAddInWorkbook = True
    ElseIf wb Is ThisWorkbook Then
        IsAddInWorkbook = True
    ElseIf wb.IsAddin Then
        IsAddInWorkbook = True
    Else
        IsAddInWorkbook = (LCase$(Right$(wb.Name, 5)) = ".xlam")
    End If
    On Error GoTo 0
End Function


'================== CORE IMPORT LOGIC ==================

Private Sub ImportBasFolder(ByVal wb As Workbook, _
                            ByVal folderPath As String, _
                            Optional ByVal skippedModuleName As String = "", _
                            Optional ByVal skippedModuleName2 As String = "")
    Dim f As String
    Dim basPath As String
    Dim fileCompName As String
    Dim targetCompName As String
    Dim importedComp As Object

    f = Dir(folderPath & "*.bas")

    Do While Len(f) > 0
        fileCompName = Left$(f, Len(f) - 4)
        basPath = folderPath & f
        targetCompName = GetBasModuleComponentName(basPath, fileCompName)

        If StrComp(fileCompName, skippedModuleName, vbTextCompare) <> 0 And _
           StrComp(targetCompName, skippedModuleName, vbTextCompare) <> 0 And _
           StrComp(fileCompName, skippedModuleName2, vbTextCompare) <> 0 And _
           StrComp(targetCompName, skippedModuleName2, vbTextCompare) <> 0 Then
            ' Remove standard modules using both filename and declared component name.
            RemoveStdModuleIfExists wb, fileCompName
            If StrComp(targetCompName, fileCompName, vbTextCompare) <> 0 Then
                RemoveStdModuleIfExists wb, targetCompName
            End If

            Set importedComp = wb.VBProject.VBComponents.Import(basPath)
            If Not importedComp Is Nothing Then
                If StrComp(importedComp.Name, targetCompName, vbTextCompare) <> 0 Then
                    On Error Resume Next
                    importedComp.Name = targetCompName
                    On Error GoTo 0
                End If
            End If
        End If

        f = Dir()
    Loop
End Sub

Private Function GetBasModuleComponentName(ByVal basFilePath As String, ByVal fallbackName As String) As String
    Dim ff As Integer
    Dim lineText As String
    Dim i As Long
    Dim q1 As Long
    Dim q2 As Long

    On Error GoTo Fallback

    ff = FreeFile
    Open basFilePath For Input As #ff

    For i = 1 To 40
        If EOF(ff) Then Exit For
        Line Input #ff, lineText

        If InStr(1, lineText, "Attribute VB_Name", vbTextCompare) > 0 Then
            q1 = InStr(1, lineText, Chr$(34))
            If q1 > 0 Then
                q2 = InStr(q1 + 1, lineText, Chr$(34))
                If q2 > q1 Then
                    GetBasModuleComponentName = Mid$(lineText, q1 + 1, q2 - q1 - 1)
                    Exit For
                End If
            End If
        End If
    Next i

    Close #ff
    If Len(GetBasModuleComponentName) = 0 Then
        GetBasModuleComponentName = fallbackName
    End If
    Exit Function

Fallback:
    On Error Resume Next
    If ff > 0 Then Close #ff
    GetBasModuleComponentName = fallbackName
End Function

Private Sub RemoveStdModuleIfExists(ByVal wb As Workbook, ByVal compName As String)
    Dim vbComp As Object

    On Error Resume Next
    Set vbComp = wb.VBProject.VBComponents(compName)
    On Error GoTo 0

    If Not vbComp Is Nothing Then
        ' vbext_ct_StdModule = 1. Document modules are type 100 and should not be removed. [3](https://bettersolutions.com/vba/visual-basic-editor/extensibility-object-model.htm)
        If vbComp.Type = 1 Then
            wb.VBProject.VBComponents.Remove vbComp
        End If
    End If
End Sub


'================== SIGNATURE / VERSION CHECK ==================
' Creates a stable signature based on file name + size + modified time (UTC-ish).
' This avoids re-importing when nothing changed.

Private Function ComputeFolderSignature(ByVal folderPath As String, ByVal pattern As String) As String
    Dim items() As String
    Dim count As Long
    count = 0

    Dim f As String
    f = Dir(folderPath & pattern)
    Do While Len(f) > 0
        count = count + 1
        ReDim Preserve items(1 To count)

        Dim fullPath As String
        fullPath = folderPath & f

        Dim sizeBytes As Double
        sizeBytes = FileLen(fullPath)

        Dim modDate As Date
        modDate = FileDateTime(fullPath)

        ' Normalize record
        items(count) = LCase$(f) & "|" & CStr(sizeBytes) & "|" & Format$(modDate, "yyyy-mm-dd hh:nn:ss")
        f = Dir()
    Loop

    If count = 0 Then
        ComputeFolderSignature = ""
        Exit Function
    End If

    QuickSort items, LBound(items), UBound(items)

    Dim i As Long, sig As String
    sig = ""
    For i = LBound(items) To UBound(items)
        sig = sig & items(i) & vbLf
    Next

    ' Simple folding hash (not cryptographic) to keep property short
    ComputeFolderSignature = CStr(SimpleHash32(sig)) & ":" & CStr(count)
End Function

Private Function SimpleHash32(ByVal s As String) As Long
    Dim i As Long
    Dim h As Long
    Dim hd As Double
    h = -2128831035

    For i = 1 To Len(s)
        h = h Xor AscW(Mid$(s, i, 1))
        hd = CDbl(h) * 16777619#
        hd = hd - 4294967296# * Fix(hd / 4294967296#) ' wrap to unsigned 32-bit
        If hd > 2147483647# Then
            hd = hd - 4294967296#
        ElseIf hd < -2147483648# Then
            hd = hd + 4294967296#
        End If
        h = CLng(hd)
    Next

    If h > 2147483647# Then
        h = h - 4294967296#
    End If

    SimpleHash32 = CLng(h)
End Function

Private Sub QuickSort(arr() As String, ByVal first As Long, ByVal last As Long)
    Dim i As Long, j As Long
    Dim pivot As String, temp As String
    i = first
    j = last
    pivot = arr((first + last) \ 2)

    Do While i <= j
        Do While arr(i) < pivot
            i = i + 1
        Loop
        Do While arr(j) > pivot
            j = j - 1
        Loop
        If i <= j Then
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
            i = i + 1
            j = j - 1
        End If
    Loop

    If first < j Then QuickSort arr, first, j
    If i < last Then QuickSort arr, i, last
End Sub


'================== DOCUMENT PROPERTY HELPERS ==================

Private Function GetDocProp(ByVal wb As Workbook, ByVal propName As String) As String
    On Error Resume Next
    GetDocProp = CStr(wb.CustomDocumentProperties(propName).Value)
    On Error GoTo 0
End Function

Private Sub SetDocProp(ByVal wb As Workbook, ByVal propName As String, ByVal propValue As String)
    Dim currentValue As String

    On Error Resume Next
    currentValue = CStr(wb.CustomDocumentProperties(propName).Value)
    If Err.Number = 0 Then
        If StrComp(currentValue, propValue, vbBinaryCompare) = 0 Then
            On Error GoTo 0
            Exit Sub
        End If
        wb.CustomDocumentProperties(propName).Value = propValue
    Else
        Err.Clear
        wb.CustomDocumentProperties.Add Name:=propName, LinkToContent:=False, _
                                       Type:=DOC_PROP_TYPE_STRING, Value:=propValue
    End If
    On Error GoTo 0
End Sub

Private Function NowUTC() As Date
    ' VBA Date is local; approximate UTC using system bias
    NowUTC = DateAdd("n", -CLng(GetTimeZoneBiasMinutes()), Now)
End Function

Private Function GetTimeZoneBiasMinutes() As Long
    ' Returns minutes to add to local time to get UTC (bias).
    ' Uses WMI via COM; if unavailable, returns 0.
    On Error GoTo Fail
    Dim wmi As Object, col As Object, itm As Object
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Set col = wmi.ExecQuery("Select Bias from Win32_TimeZone")
    For Each itm In col
        GetTimeZoneBiasMinutes = CLng(itm.Bias)
        Exit Function
    Next
Fail:
    GetTimeZoneBiasMinutes = 0
End Function

