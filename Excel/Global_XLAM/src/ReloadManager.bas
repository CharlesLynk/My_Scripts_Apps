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
Private Const TEST_PICKER_FORM_NAME As String = "RM_TypePickerTestForm"
Private Const TEST_PICKER_LIST_NAME As String = "lstEntries"
Private Const TEST_PICKER_OK_NAME As String = "btnOK"
Private Const TEST_PICKER_CANCEL_NAME As String = "btnCancel"
Private Const TEST_PICKER_INFO_NAME As String = "lblInfo"
Private Const MAX_FOLDER_SCAN_DEPTH As Long = 64
Private Const FILE_ATTR_REPARSE_POINT As Long = 1024

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

' Proof-of-concept userform picker to evaluate layout/width before full conversion.
Public Sub TestWorkbookTypePickerUserForm()
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
        MsgBox "Select a regular workbook before running picker test.", vbExclamation, SYNC_DIALOG_TITLE
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
    BuildUserFormPickerPreview availableTypeIds, typeMap, displayLines, ownerTypeIds

    If displayLines.Count = 0 Then
        MsgBox "No .bas entries were found for available type IDs.", vbInformation, SYNC_DIALOG_TITLE
        Exit Sub
    End If

    selectedIndex = ShowTypePickerTestUserForm(displayLines, ownerTypeIds, "Workbook Type Picker Test")
    If selectedIndex <= 0 Then Exit Sub
    If selectedIndex > ownerTypeIds.Count Then Exit Sub

    MsgBox "Selected preview entry:" & vbCrLf & _
           CStr(displayLines(selectedIndex)) & vbCrLf & vbCrLf & _
           "Owning Type ID:" & vbCrLf & _
           CStr(ownerTypeIds(selectedIndex)), vbInformation, SYNC_DIALOG_TITLE
End Sub

Private Sub ShowReloadManagerCommandCenter()
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
    Set wb = GetForegroundWorkbook()

    If wb Is Nothing Then
        MsgBox "No workbook is currently available for sync.", vbExclamation, SYNC_DIALOG_TITLE
        Exit Sub
    End If
    If IsAddInWorkbook(wb) Then
        MsgBox "Select a regular workbook before syncing.", vbExclamation, SYNC_DIALOG_TITLE
        Exit Sub
    End If

    SyncWorkbook wb, Force, PromptOnDrift
End Sub

Private Sub ShowActiveWorkbookTypeIds()
    Dim wb As Workbook
    Dim typeMap As Object
    Dim summary As String
    Dim sorted() As String

    Set wb = GetForegroundWorkbook()
    If wb Is Nothing Then
        MsgBox "No workbook is currently available.", vbExclamation, SYNC_DIALOG_TITLE
        Exit Sub
    End If

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

    Set wb = GetForegroundWorkbook()
    If wb Is Nothing Then
        MsgBox "No workbook is currently available to update types.", vbExclamation, SYNC_DIALOG_TITLE
        Exit Sub
    End If
    If IsAddInWorkbook(wb) Then
        MsgBox "Select a regular workbook before assigning types.", vbExclamation, SYNC_DIALOG_TITLE
        Exit Sub
    End If

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

    Set wb = GetForegroundWorkbook()
    If wb Is Nothing Then
        MsgBox "No workbook is currently available to update types.", vbExclamation, SYNC_DIALOG_TITLE
        Exit Sub
    End If
    If IsAddInWorkbook(wb) Then
        MsgBox "Select a regular workbook before updating types.", vbExclamation, SYNC_DIALOG_TITLE
        Exit Sub
    End If

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

    Set wb = GetForegroundWorkbook()
    If wb Is Nothing Then
        MsgBox "No workbook is currently available to clear types.", vbExclamation, SYNC_DIALOG_TITLE
        Exit Sub
    End If
    If IsAddInWorkbook(wb) Then
        MsgBox "Select a regular workbook before clearing types.", vbExclamation, SYNC_DIALOG_TITLE
        Exit Sub
    End If

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


Private Function PromptSelectAvailableTypeId(ByVal availableTypeIds As Collection, _
                                             ByVal assignedTypeMap As Object) As String
    Dim menuText As String
    Dim userInput As String
    Dim i As Long
    Dim typeId As String
    Dim selectedIndex As Long
    Dim marker As String
    Dim normalizedTypeId As String
    Dim selectableTypeIds As Collection
    Dim selectableEntryPaths As Collection
    Dim basFiles As Collection
    Dim basFile As Variant
    Dim numberWidth As Long
    Dim numberLabel As String
    Dim parts() As String
    Dim previousParts() As String
    Dim hasPrevious As Boolean
    Dim commonDepth As Long
    Dim leafIndex As Long
    Dim j As Long

    Set selectableTypeIds = New Collection
    Set selectableEntryPaths = New Collection

    For i = 1 To availableTypeIds.Count
        typeId = CStr(availableTypeIds(i))
        normalizedTypeId = NormalizeTypeId(typeId)
        Set basFiles = GetBasFileNamesInTypeSource(normalizedTypeId)
        For Each basFile In basFiles
            selectableTypeIds.Add normalizedTypeId
            selectableEntryPaths.Add normalizedTypeId & Application.PathSeparator & CStr(basFile)
        Next basFile
    Next i

    If selectableEntryPaths.Count = 0 Then Exit Function

    menuText = "Add Workbook Type" & vbCrLf & String$(24, "=") & vbCrLf & vbCrLf
    numberWidth = Len(CStr(selectableEntryPaths.Count))
    If numberWidth < 2 Then numberWidth = 2

    For i = 1 To selectableEntryPaths.Count
        parts = Split(CStr(selectableEntryPaths(i)), Application.PathSeparator)
        leafIndex = UBound(parts)

        If hasPrevious Then
            commonDepth = GetCommonFolderDepth(previousParts, parts)
        Else
            commonDepth = 0
        End If

        For j = commonDepth To leafIndex - 1
            menuText = menuText & String$(j * 3, " ") & parts(j) & vbCrLf
        Next j

        If assignedTypeMap.Exists(CStr(selectableTypeIds(i))) Then
            marker = "[Assigned]"
        Else
            marker = "[Available]"
        End If

        numberLabel = "[" & Right$(String$(numberWidth, " ") & CStr(i), numberWidth) & "]"
        menuText = menuText & String$(leafIndex * 3, " ") & numberLabel & " " & parts(leafIndex) & " " & marker & vbCrLf

        previousParts = parts
        hasPrevious = True
    Next i
    menuText = menuText & vbCrLf & "Enter a number (or 0 to cancel)."

    userInput = Trim$(InputBox(menuText, "Add Workbook Type", "0"))
    If Len(userInput) = 0 Or userInput = "0" Then Exit Function
    If Not IsNumeric(userInput) Then
        MsgBox "Please enter a valid number.", vbExclamation, SYNC_DIALOG_TITLE
        Exit Function
    End If

    selectedIndex = CLng(userInput)
    If selectedIndex < 1 Or selectedIndex > selectableEntryPaths.Count Then
        MsgBox "Selection out of range.", vbExclamation, SYNC_DIALOG_TITLE
        Exit Function
    End If

    PromptSelectAvailableTypeId = CStr(selectableTypeIds(selectedIndex))
End Function

Private Function PromptSelectTypeIdsUserForm(ByVal typeIds As Collection, _
                                             ByVal formCaption As String, _
                                             ByVal infoText As String) As Collection
    Dim displayLines As Collection
    Dim ownerTypeIds As Collection
    Dim uncheckedTypeMap As Object
    Dim selectedTypeIds As Collection

    Set selectedTypeIds = New Collection
    If typeIds Is Nothing Then
        Set PromptSelectTypeIdsUserForm = selectedTypeIds
        Exit Function
    End If
    If typeIds.Count = 0 Then
        Set PromptSelectTypeIdsUserForm = selectedTypeIds
        Exit Function
    End If

    Set displayLines = New Collection
    Set ownerTypeIds = New Collection
    Set uncheckedTypeMap = CreateObject("Scripting.Dictionary")
    uncheckedTypeMap.CompareMode = vbTextCompare

    BuildTypePickerTreeLines typeIds, uncheckedTypeMap, displayLines, ownerTypeIds

    If displayLines.Count = 0 Then
        Set PromptSelectTypeIdsUserForm = selectedTypeIds
        Exit Function
    End If

    Set selectedTypeIds = ShowTypePickerCheckedTypeIdsUserForm(displayLines, ownerTypeIds, formCaption, infoText)
    Set PromptSelectTypeIdsUserForm = selectedTypeIds
End Function

Private Sub BuildUserFormPickerPreview(ByVal availableTypeIds As Collection, _
                                       ByVal assignedTypeMap As Object, _
                                       ByRef displayLines As Collection, _
                                       ByRef ownerTypeIds As Collection)
    BuildTypePickerTreeLines availableTypeIds, assignedTypeMap, displayLines, ownerTypeIds
End Sub

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
            commonDepth = GetCommonFolderDepth(previousParts, parts)
        Else
            commonDepth = 0
        End If

        For j = commonDepth To leafIndex
            displayLines.Add String$(j * 2, " ") & parts(j)
            ownerTypeIds.Add vbNullString
        Next j

        Set basFiles = GetBasFileNamesInTypeSource(normalizedTypeId)

        If (Not checkedTypeMap Is Nothing) And checkedTypeMap.Exists(normalizedTypeId) Then
            marker = "[x]"
        Else
            marker = "[ ]"
        End If

        For Each basFile In basFiles
            displayLines.Add String$((leafIndex + 1) * 2, " ") & marker & " " & CStr(basFile)
            ownerTypeIds.Add normalizedTypeId
        Next basFile

        previousParts = parts
        hasPrevious = True
NextType:
    Next i
End Sub

Private Function ShowTypePickerTestUserForm(ByVal displayLines As Collection, _
                                            ByVal ownerTypeIds As Collection, _
                                            ByVal formCaption As String) As Long
    Dim vbComp As Object
    Dim frm As Object
    Dim entriesHost As Object
    Dim lbl As Object
    Dim reason As String
    Dim stage As String

    On Error GoTo Fail

    stage = "EnsureTypePickerTestFormComponent"
    Set vbComp = EnsureTypePickerTestFormComponent(ThisWorkbook, reason)
    If vbComp Is Nothing Then
        MsgBox "Unable to open test userform." & vbCrLf & _
               "Reason: " & reason & vbCrLf & vbCrLf & _
               "Fallback: create a manual UserForm named '" & TEST_PICKER_FORM_NAME & _
               "' in the Global_CNL VBA project with controls:" & vbCrLf & _
               "- " & TEST_PICKER_INFO_NAME & " (Label)" & vbCrLf & _
             "- " & TEST_PICKER_LIST_NAME & " (Frame)" & vbCrLf & _
               "- " & TEST_PICKER_OK_NAME & " (CommandButton)" & vbCrLf & _
               "- " & TEST_PICKER_CANCEL_NAME & " (CommandButton)", vbExclamation, SYNC_DIALOG_TITLE
        Exit Function
    End If

    stage = "UserForms.Add"
    Set frm = VBA.UserForms.Add(vbComp.Name)
    frm.Caption = formCaption
    frm.Tag = "0"

    stage = "Find info label"
    Set lbl = GetFormControlByName(frm, TEST_PICKER_INFO_NAME)
    If lbl Is Nothing Then Err.Raise vbObjectError + 8101, "ReloadManager.ShowTypePickerTestUserForm", _
                                    "Missing control: " & TEST_PICKER_INFO_NAME
    lbl.Caption = "Select .bas entries using checkboxes. Folder rows are context only."

    stage = "Find entries host"
    Set entriesHost = GetFormControlByName(frm, TEST_PICKER_LIST_NAME)
    If entriesHost Is Nothing Then Err.Raise vbObjectError + 8102, "ReloadManager.ShowTypePickerTestUserForm", _
                                    "Missing control: " & TEST_PICKER_LIST_NAME

    stage = "AutoSize form"
    AutoSizeTypePickerTestForm frm, displayLines

    stage = "Build checkbox entries"
    PopulateTypePickerChecklist entriesHost, displayLines, ownerTypeIds

    stage = "Show form"
    frm.Show vbModal

    stage = "Read selection"
    On Error Resume Next
    If StrComp(CStr(frm.Tag), "OK", vbTextCompare) = 0 Then
        ShowTypePickerTestUserForm = GetFirstCheckedChecklistLineIndex(entriesHost, ownerTypeIds)
    ElseIf Len(CStr(frm.Tag)) > 0 And IsNumeric(CStr(frm.Tag)) Then
        ShowTypePickerTestUserForm = CLng(Val(CStr(frm.Tag)))
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
    MsgBox "UserForm test failed at stage: " & stage & vbCrLf & _
           "Error " & CStr(Err.Number) & ": " & Err.Description, vbExclamation, SYNC_DIALOG_TITLE
    On Error Resume Next
    If Not frm Is Nothing Then Unload frm
End Function

Private Function ShowTypePickerCheckedTypeIdsUserForm(ByVal displayLines As Collection, _
                                                      ByVal ownerTypeIds As Collection, _
                                                      ByVal formCaption As String, _
                                                      ByVal infoText As String) As Collection
    Dim vbComp As Object
    Dim frm As Object
    Dim entriesHost As Object
    Dim lbl As Object
    Dim reason As String
    Dim stage As String
    Dim selectedTypeIds As New Collection

    On Error GoTo Fail

    stage = "EnsureTypePickerTestFormComponent"
    Set vbComp = EnsureTypePickerTestFormComponent(ThisWorkbook, reason)
    If vbComp Is Nothing Then
        MsgBox "Unable to open type picker." & vbCrLf & _
               "Reason: " & reason, vbExclamation, SYNC_DIALOG_TITLE
        Set ShowTypePickerCheckedTypeIdsUserForm = selectedTypeIds
        Exit Function
    End If

    stage = "UserForms.Add"
    Set frm = VBA.UserForms.Add(vbComp.Name)
    frm.Caption = formCaption
    frm.Tag = "0"

    stage = "Find info label"
    Set lbl = GetFormControlByName(frm, TEST_PICKER_INFO_NAME)
    If lbl Is Nothing Then Err.Raise vbObjectError + 8111, "ReloadManager.ShowTypePickerCheckedTypeIdsUserForm", _
                                    "Missing control: " & TEST_PICKER_INFO_NAME
    lbl.Caption = infoText

    stage = "Find entries host"
    Set entriesHost = GetFormControlByName(frm, TEST_PICKER_LIST_NAME)
    If entriesHost Is Nothing Then Err.Raise vbObjectError + 8112, "ReloadManager.ShowTypePickerCheckedTypeIdsUserForm", _
                                    "Missing control: " & TEST_PICKER_LIST_NAME

    stage = "AutoSize form"
    AutoSizeTypePickerTestForm frm, displayLines

    stage = "Build checkbox entries"
    PopulateTypePickerChecklist entriesHost, displayLines, ownerTypeIds

    stage = "Show form"
    frm.Show vbModal

    stage = "Collect checked selections"
    If StrComp(CStr(frm.Tag), "OK", vbTextCompare) = 0 Then
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
                                        ByVal displayLines As Collection, _
                                        ByVal ownerTypeIds As Collection)
    Dim i As Long
    Dim lineText As String
    Dim ownerTypeId As String
    Dim indentChars As Long
    Dim captionText As String
    Dim rowTop As Double
    Dim rowHeight As Double
    Dim maxWidth As Double
    Dim offsetLeft As Double
    Dim ctl As Object

    If entriesHost Is Nothing Then Exit Sub

    ClearTypePickerDynamicControls entriesHost

    rowTop = 4
    rowHeight = 18
    maxWidth = CDbl(entriesHost.Width) - 16
    If maxWidth < 120 Then maxWidth = 120

    For i = 1 To displayLines.Count
        lineText = CStr(displayLines(i))
        ownerTypeId = ""
        If Not ownerTypeIds Is Nothing Then
            If i <= ownerTypeIds.Count Then
                ownerTypeId = NormalizeTypeId(CStr(ownerTypeIds(i)))
            End If
        End If

        indentChars = CountLeadingSpaces(lineText)
        offsetLeft = 6 + (indentChars * 3)
        captionText = StripChecklistMarker(LTrim$(lineText))

        If Len(ownerTypeId) > 0 Then
            Set ctl = entriesHost.Controls.Add("Forms.CheckBox.1", "chkEntry_" & CStr(i), True)
            TrySetObjectProperty ctl, "Caption", captionText
            TrySetObjectProperty ctl, "Value", IsChecklistLineChecked(lineText)
            TrySetObjectProperty ctl, "Tag", CStr(i)
            TrySetObjectProperty ctl, "Left", offsetLeft
            TrySetObjectProperty ctl, "Top", rowTop
            TrySetObjectProperty ctl, "Width", maxWidth - (indentChars * 3)
            TrySetObjectProperty ctl, "Height", 16
        Else
            Set ctl = entriesHost.Controls.Add("Forms.Label.1", "lblEntry_" & CStr(i), True)
            TrySetObjectProperty ctl, "Caption", LTrim$(lineText)
            TrySetObjectProperty ctl, "Left", offsetLeft + 2
            TrySetObjectProperty ctl, "Top", rowTop + 1
            TrySetObjectProperty ctl, "Width", maxWidth - (indentChars * 3)
            TrySetObjectProperty ctl, "Height", 15
        End If

        rowTop = rowTop + rowHeight
    Next i

    TrySetObjectProperty entriesHost, "ScrollBars", 2
    TrySetObjectProperty entriesHost, "KeepScrollBarsVisible", 2
    TrySetObjectProperty entriesHost, "ScrollHeight", rowTop + 4
    TrySetObjectProperty entriesHost, "ScrollWidth", CDbl(entriesHost.Width)
    On Error Resume Next
    entriesHost.ScrollTop = 0
    On Error GoTo 0
End Sub

Private Sub ClearTypePickerDynamicControls(ByVal entriesHost As Object)
    Dim i As Long
    Dim ctl As Object
    Dim ctlName As String

    If entriesHost Is Nothing Then Exit Sub

    On Error Resume Next
    For i = entriesHost.Controls.Count - 1 To 0 Step -1
        Set ctl = entriesHost.Controls.Item(i)
        ctlName = CStr(ctl.Name)
        If Left$(ctlName, 9) = "chkEntry_" Or Left$(ctlName, 9) = "lblEntry_" Then
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

Private Function EnsureTypePickerTestFormComponent(ByVal wb As Workbook, ByRef reason As String) As Object
    Dim vbComp As Object
    Dim designer As Object
    Dim stage As String
    Dim createdNow As Boolean

    On Error GoTo Fail

    stage = "Find component"
    Set vbComp = FindVBComponentByName(wb, TEST_PICKER_FORM_NAME)

    If Not vbComp Is Nothing Then
        If vbComp.Type <> 3 Then
            reason = "A non-UserForm component already uses name '" & TEST_PICKER_FORM_NAME & "'." & vbCrLf & _
                     "Rename it or create a UserForm with that name."
            Exit Function
        End If
    End If

    stage = "Create component"
    If vbComp Is Nothing Then
        Set vbComp = wb.VBProject.VBComponents.Add(3)
        vbComp.Name = TEST_PICKER_FORM_NAME
        createdNow = True
    End If

    stage = "Open designer"
    Set designer = vbComp.Designer
    If designer Is Nothing Then
        reason = "Designer is unavailable for component '" & vbComp.Name & "'."
        Exit Function
    End If

    stage = "Configure designer controls"
    If createdNow Or (Not HasRequiredTypePickerControls(designer)) Then
        ConfigureTypePickerTestFormDesigner designer
    End If

    stage = "Inject form code"
    EnsureTypePickerTestFormCode vbComp.CodeModule

    Set EnsureTypePickerTestFormComponent = vbComp
    Exit Function

Fail:
    reason = "Stage: " & stage & " | Error " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function HasRequiredTypePickerControls(ByVal designer As Object) As Boolean
    Dim entriesHost As Object

    Set entriesHost = GetDesignerControlByName(designer, TEST_PICKER_LIST_NAME)
    If entriesHost Is Nothing Then Exit Function
    If StrComp(TypeName(entriesHost), "Frame", vbTextCompare) <> 0 Then Exit Function
    If GetDesignerControlByName(designer, TEST_PICKER_OK_NAME) Is Nothing Then Exit Function
    If GetDesignerControlByName(designer, TEST_PICKER_CANCEL_NAME) Is Nothing Then Exit Function
    If GetDesignerControlByName(designer, TEST_PICKER_INFO_NAME) Is Nothing Then Exit Function
    HasRequiredTypePickerControls = True
End Function

Private Function FindVBComponentByName(ByVal wb As Workbook, ByVal componentName As String) As Object
    Dim vbComp As Object

    For Each vbComp In wb.VBProject.VBComponents
        If StrComp(vbComp.Name, componentName, vbTextCompare) = 0 Then
            Set FindVBComponentByName = vbComp
            Exit Function
        End If
    Next vbComp
End Function

Private Sub ConfigureTypePickerTestFormDesigner(ByVal designer As Object)
    Dim lbl As Object
    Dim lst As Object
    Dim existingEntries As Object
    Dim btnOk As Object
    Dim btnCancel As Object
    Dim stage As String

    On Error GoTo Fail

    stage = "Form properties"
    TrySetObjectProperty designer, "Caption", "Workbook Type Picker Test"
    TrySetObjectProperty designer, "Width", 780
    TrySetObjectProperty designer, "Height", 460

    stage = "Info label"
    Set lbl = GetOrCreateFormControl(designer, "Forms.Label.1", TEST_PICKER_INFO_NAME)
    TrySetObjectProperty lbl, "Left", 12
    TrySetObjectProperty lbl, "Top", 12
    TrySetObjectProperty lbl, "Width", 744
    TrySetObjectProperty lbl, "Height", 24
    TrySetObjectProperty lbl, "Caption", "POC: test readability and width with long paths before full conversion."
    TrySetObjectProperty lbl, "WordWrap", True

    stage = "Entries host"
    Set existingEntries = GetDesignerControlByName(designer, TEST_PICKER_LIST_NAME)
    If Not existingEntries Is Nothing Then
        If StrComp(TypeName(existingEntries), "Frame", vbTextCompare) <> 0 Then
            On Error Resume Next
            designer.Controls.Remove TEST_PICKER_LIST_NAME
            On Error GoTo Fail
        End If
    End If

    Set lst = GetOrCreateFormControl(designer, "Forms.Frame.1", TEST_PICKER_LIST_NAME)
    TrySetObjectProperty lst, "Left", 12
    TrySetObjectProperty lst, "Top", 42
    TrySetObjectProperty lst, "Width", 744
    TrySetObjectProperty lst, "Height", 350
    TrySetObjectProperty lst, "Caption", ""
    TrySetObjectProperty lst, "ScrollBars", 2
    TrySetObjectProperty lst, "KeepScrollBarsVisible", 2

    stage = "OK button"
    Set btnOk = GetOrCreateFormControl(designer, "Forms.CommandButton.1", TEST_PICKER_OK_NAME)
    TrySetObjectProperty btnOk, "Caption", "OK"
    TrySetObjectProperty btnOk, "Left", 588
    TrySetObjectProperty btnOk, "Top", 402
    TrySetObjectProperty btnOk, "Width", 80
    TrySetObjectProperty btnOk, "Height", 24

    stage = "Cancel button"
    Set btnCancel = GetOrCreateFormControl(designer, "Forms.CommandButton.1", TEST_PICKER_CANCEL_NAME)
    TrySetObjectProperty btnCancel, "Caption", "Cancel"
    TrySetObjectProperty btnCancel, "Left", 676
    TrySetObjectProperty btnCancel, "Top", 402
    TrySetObjectProperty btnCancel, "Width", 80
    TrySetObjectProperty btnCancel, "Height", 24
    Exit Sub

Fail:
    Err.Raise vbObjectError + 8120, "ReloadManager.ConfigureTypePickerTestFormDesigner", _
              "Stage: " & stage & " | Error " & CStr(Err.Number) & ": " & Err.Description
End Sub

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

Private Sub AutoSizeTypePickerTestForm(ByVal frm As Object, ByVal displayLines As Collection)
    Dim lbl As Object
    Dim entriesHost As Object
    Dim btnOk As Object
    Dim btnCancel As Object
    Dim lineText As Variant
    Dim longestLineWidth As Double
    Dim measuredLineHeight As Double
    Dim rowHeight As Double
    Dim lineCount As Long
    Dim listWidth As Double
    Dim listHeight As Double
    Dim formWidth As Double
    Dim formHeight As Double
    Dim btnTop As Double
    Dim usableWidth As Double
    Dim usableHeight As Double
    Dim maxListWidth As Double
    Dim maxListHeight As Double
    Const LEFT_MARGIN As Double = 12
    Const TOP_LABEL As Double = 12
    Const LABEL_HEIGHT As Double = 24
    Const LIST_TOP As Double = 42
    Const RIGHT_MARGIN As Double = 12
    Const BUTTON_GAP_TOP As Double = 10
    Const BUTTON_WIDTH As Double = 80
    Const BUTTON_HEIGHT As Double = 24
    Const BUTTON_GAP As Double = 8
    Const BOTTOM_MARGIN As Double = 34
    Const LIST_HORIZONTAL_PADDING As Double = 26
    Const LIST_VERTICAL_PADDING As Double = 8
    Const MIN_VISIBLE_ROWS As Long = 8
    Const MAX_VISIBLE_ROWS As Long = 36

    Set lbl = GetFormControlByName(frm, TEST_PICKER_INFO_NAME)
    Set entriesHost = GetFormControlByName(frm, TEST_PICKER_LIST_NAME)
    Set btnOk = GetFormControlByName(frm, TEST_PICKER_OK_NAME)
    Set btnCancel = GetFormControlByName(frm, TEST_PICKER_CANCEL_NAME)

    If lbl Is Nothing Or entriesHost Is Nothing Or btnOk Is Nothing Or btnCancel Is Nothing Then Exit Sub

    lineCount = displayLines.Count
    If lineCount < 1 Then lineCount = 1

    longestLineWidth = MeasureLongestCaptionWidth(lbl, displayLines)
    measuredLineHeight = MeasureCaptionHeight(lbl)
    If measuredLineHeight <= 0 Then measuredLineHeight = 11
    rowHeight = measuredLineHeight
    If rowHeight < 11 Then rowHeight = 11
    rowHeight = rowHeight + 1

    On Error Resume Next
    usableWidth = CDbl(Application.UsableWidth)
    usableHeight = CDbl(Application.UsableHeight)
    On Error GoTo 0
    If usableWidth <= 0 Then usableWidth = 1400
    If usableHeight <= 0 Then usableHeight = 900

    maxListWidth = (usableWidth * 0.85) - (LEFT_MARGIN + RIGHT_MARGIN)
    If maxListWidth < 420 Then maxListWidth = 420
    maxListHeight = (usableHeight * 0.78) - (LIST_TOP + BUTTON_GAP_TOP + BUTTON_HEIGHT + BOTTOM_MARGIN)
    If maxListHeight < 220 Then maxListHeight = 220

    listWidth = longestLineWidth + LIST_HORIZONTAL_PADDING
    If listWidth < 320 Then listWidth = 320
    If listWidth > maxListWidth Then listWidth = maxListWidth

    If lineCount < MIN_VISIBLE_ROWS Then lineCount = MIN_VISIBLE_ROWS
    If lineCount > MAX_VISIBLE_ROWS Then lineCount = MAX_VISIBLE_ROWS
    listHeight = (lineCount * rowHeight) + LIST_VERTICAL_PADDING
    If listHeight < 200 Then listHeight = 200
    If listHeight > maxListHeight Then listHeight = maxListHeight

    btnTop = LIST_TOP + listHeight + BUTTON_GAP_TOP
    formWidth = LEFT_MARGIN + listWidth + RIGHT_MARGIN
    formHeight = btnTop + BUTTON_HEIGHT + BOTTOM_MARGIN

    If formWidth < 460 Then formWidth = 460
    If formWidth > usableWidth * 0.9 Then formWidth = usableWidth * 0.9
    If formHeight < 320 Then formHeight = 320
    If formHeight > usableHeight * 0.9 Then formHeight = usableHeight * 0.9

    ' Keep margins fixed; let list consume any extra width from form minimums.
    listWidth = formWidth - LEFT_MARGIN - RIGHT_MARGIN
    If listWidth < 320 Then listWidth = 320

    frm.Width = formWidth
    frm.Height = formHeight

    lbl.Left = LEFT_MARGIN
    lbl.Top = TOP_LABEL
    lbl.Width = listWidth
    lbl.Height = LABEL_HEIGHT

    entriesHost.Left = LEFT_MARGIN
    entriesHost.Top = LIST_TOP
    entriesHost.Width = listWidth
    entriesHost.Height = listHeight
    On Error Resume Next
    entriesHost.ScrollBars = 2
    entriesHost.KeepScrollBarsVisible = 2
    On Error GoTo 0

    btnCancel.Width = BUTTON_WIDTH
    btnCancel.Height = BUTTON_HEIGHT
    btnCancel.Top = btnTop
    btnCancel.Left = LEFT_MARGIN + listWidth - BUTTON_WIDTH

    btnOk.Width = BUTTON_WIDTH
    btnOk.Height = BUTTON_HEIGHT
    btnOk.Top = btnTop
    btnOk.Left = LEFT_MARGIN + listWidth - (BUTTON_WIDTH * 2) - BUTTON_GAP
End Sub

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

Private Function MeasureCaptionHeight(ByVal measureLabel As Object) As Double
    Dim originalCaption As String
    Dim originalAutoSize As Boolean
    Dim originalWordWrap As Boolean
    Dim originalWidth As Double
    Dim originalHeight As Double

    On Error Resume Next
    originalCaption = CStr(measureLabel.Caption)
    originalAutoSize = CBool(measureLabel.AutoSize)
    originalWordWrap = CBool(measureLabel.WordWrap)
    originalWidth = CDbl(measureLabel.Width)
    originalHeight = CDbl(measureLabel.Height)

    measureLabel.WordWrap = False
    measureLabel.AutoSize = True
    measureLabel.Caption = "Ag"
    MeasureCaptionHeight = CDbl(measureLabel.Height)

    measureLabel.Caption = originalCaption
    measureLabel.AutoSize = originalAutoSize
    measureLabel.WordWrap = originalWordWrap
    measureLabel.Width = originalWidth
    measureLabel.Height = originalHeight
    On Error GoTo 0
End Function

Private Sub EnsureTypePickerTestFormCode(ByVal codeModule As Object)
    Dim codeText As String

    codeText = ""
    codeText = codeText & "Option Explicit" & vbCrLf
    codeText = codeText & vbCrLf
    codeText = codeText & "Private Sub UserForm_Initialize()" & vbCrLf
    codeText = codeText & "    Me.Tag = ""0""" & vbCrLf
    codeText = codeText & "End Sub" & vbCrLf
    codeText = codeText & vbCrLf
    codeText = codeText & "Private Sub btnOK_Click()" & vbCrLf
    codeText = codeText & "    Me.Tag = ""OK""" & vbCrLf
    codeText = codeText & "    Me.Hide" & vbCrLf
    codeText = codeText & "End Sub" & vbCrLf
    codeText = codeText & vbCrLf
    codeText = codeText & "Private Sub btnCancel_Click()" & vbCrLf
    codeText = codeText & "    Me.Tag = ""0""" & vbCrLf
    codeText = codeText & "    Me.Hide" & vbCrLf
    codeText = codeText & "End Sub" & vbCrLf
    codeText = codeText & vbCrLf
    codeText = codeText & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
    codeText = codeText & "    If CloseMode = 0 Then" & vbCrLf
    codeText = codeText & "        Cancel = True" & vbCrLf
    codeText = codeText & "        Me.Tag = ""0""" & vbCrLf
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
        names(count) = f
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

Private Function GetCommonFolderDepth(ByRef leftParts() As String, ByRef rightParts() As String) As Long
    Dim leftLastFolder As Long
    Dim rightLastFolder As Long
    Dim maxFolderIndex As Long
    Dim i As Long

    leftLastFolder = UBound(leftParts) - 1
    rightLastFolder = UBound(rightParts) - 1

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

