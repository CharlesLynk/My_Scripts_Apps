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
Private Const PROP_LAST_SYNC As String = "WB_LAST_SYNC_UTC"
Private Const PROP_GLOBAL_SRC_SIG As String = "GLB_SRC_SIGNATURE"
Private Const TYPE_LIST_DELIMITER As String = ";"
Private Const REPO_ADDIN_RELATIVE_PATH As String = "Global_XLAM\Global_CNL.xlam"
Private Const REPO_ADDIN_SRC_RELATIVE_PATH As String = "Global_XLAM\src"
Private Const MODULE_RELOAD_MANAGER As String = "ReloadManager"
Private Const MODULE_RELOAD_MANAGER_BOOTSTRAP As String = "ReloadManagerBootstrap"

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
    Dim repoSig As String
    Dim wbSig As String
    Dim reason As String

    On Error GoTo SafeExit

    If wb Is Nothing Then Exit Sub
    If IsAddInWorkbook(wb) Then Exit Sub

    ' Keep open-time checks quiet unless actionable drift is found.
    If Not PrepareSyncContext(wb, typeSummary, srcPaths, repoSig, wbSig, reason) Then Exit Sub
    If StrComp(repoSig, wbSig, vbBinaryCompare) = 0 Then Exit Sub

    If ConfirmDriftImport(wb, typeSummary) Then
        SyncWorkbook wb, True, False
    End If

SafeExit:
    On Error GoTo 0
End Sub

' Single visible macro entry point for users.
Public Sub ManageActiveWorkbookTypes()
    ShowReloadManagerCommandCenter
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
               "1 - Add assigned types" & vbCrLf & _
               "2 - Remove assigned type" & vbCrLf & _
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
    Dim selectedTypeId As String
    Dim normalizedTypeId As String
    Dim availableTypeIds As Collection

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

    selectedTypeId = PromptSelectAvailableTypeId(availableTypeIds, typeMap)
    If Len(selectedTypeId) = 0 Then Exit Sub

    normalizedTypeId = NormalizeTypeId(selectedTypeId)

    If typeMap.Exists(normalizedTypeId) Then
        MsgBox "Type already assigned: " & normalizedTypeId, vbInformation, SYNC_DIALOG_TITLE
        Exit Sub
    End If

    typeMap(normalizedTypeId) = normalizedTypeId
    SaveWorkbookTypeMap wb, typeMap

    MsgBox "Type added: " & normalizedTypeId, vbInformation, SYNC_DIALOG_TITLE
End Sub

Private Sub RemoveActiveWorkbookTypeId()
    Dim wb As Workbook
    Dim typeMap As Object
    Dim typeId As String
    Dim normalizedTypeId As String
    Dim sortedTypes() As String
    Dim selectedIndex As Long
    Dim menuText As String
    Dim i As Long
    Dim userInput As String

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

    sortedTypes = GetSortedTypeIds(typeMap)

    menuText = "Remove Assigned Type" & vbCrLf & String$(24, "=") & vbCrLf & vbCrLf
    For i = LBound(sortedTypes) To UBound(sortedTypes)
        menuText = menuText & CStr(i - LBound(sortedTypes) + 1) & " - " & sortedTypes(i) & vbCrLf
    Next i
    menuText = menuText & vbCrLf & "Enter a number (or 0 to cancel)."

    userInput = Trim$(InputBox(menuText, "Remove Workbook Type", "0"))
    If Len(userInput) = 0 Or userInput = "0" Then Exit Sub
    If Not IsNumeric(userInput) Then
        MsgBox "Please enter a valid number.", vbExclamation, SYNC_DIALOG_TITLE
        Exit Sub
    End If

    selectedIndex = CLng(userInput)
    If selectedIndex < 1 Or selectedIndex > (UBound(sortedTypes) - LBound(sortedTypes) + 1) Then
        MsgBox "Selection out of range.", vbExclamation, SYNC_DIALOG_TITLE
        Exit Sub
    End If

    typeId = sortedTypes(LBound(sortedTypes) + selectedIndex - 1)

    normalizedTypeId = NormalizeTypeId(typeId)
    If Not typeMap.Exists(normalizedTypeId) Then
        MsgBox "Type not assigned: " & normalizedTypeId, vbExclamation, SYNC_DIALOG_TITLE
        Exit Sub
    End If

    typeMap.Remove normalizedTypeId
    SaveWorkbookTypeMap wb, typeMap

    MsgBox "Type removed: " & normalizedTypeId, vbInformation, SYNC_DIALOG_TITLE
End Sub

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
    Dim repoSig As String
    Dim wbSig As String
    Dim reason As String
    Dim previousEnableEvents As Boolean
    Dim previousScreenUpdating As Boolean
    Dim srcPath As Variant

    If Not PrepareSyncContext(wb, typeSummary, srcPaths, repoSig, wbSig, reason) Then
        MsgBox reason, vbExclamation, SYNC_DIALOG_TITLE
        Exit Sub
    End If

    If Not Force Then
        If StrComp(repoSig, wbSig, vbBinaryCompare) = 0 Then
            Exit Sub
        End If

        If PromptOnDrift Then
            If Not ConfirmDriftImport(wb, typeSummary) Then
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

    For Each srcPath In srcPaths
        ImportBasFolder wb, CStr(srcPath)
    Next srcPath

    SetDocProp wb, PROP_SRC_SIG, repoSig
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
    Dim parts() As String
    Dim depth As Long
    Dim leaf As String

    menuText = "Add Workbook Type" & vbCrLf & String$(24, "=") & vbCrLf & vbCrLf
    For i = 1 To availableTypeIds.Count
        typeId = CStr(availableTypeIds(i))
        If assignedTypeMap.Exists(typeId) Then
            marker = "[Assigned]"
        Else
            marker = "[Available]"
        End If

        normalizedTypeId = NormalizeTypeId(typeId)
        parts = Split(normalizedTypeId, Application.PathSeparator)
        depth = UBound(parts)
        leaf = parts(UBound(parts))

        menuText = menuText & CStr(i) & " - " & marker & "  " & String$(depth * 4, " ") & leaf & vbCrLf
    Next i
    menuText = menuText & vbCrLf & "Enter a number (or 0 to cancel)."

    userInput = Trim$(InputBox(menuText, "Add Workbook Type", "0"))
    If Len(userInput) = 0 Or userInput = "0" Then Exit Function
    If Not IsNumeric(userInput) Then
        MsgBox "Please enter a valid number.", vbExclamation, SYNC_DIALOG_TITLE
        Exit Function
    End If

    selectedIndex = CLng(userInput)
    If selectedIndex < 1 Or selectedIndex > availableTypeIds.Count Then
        MsgBox "Selection out of range.", vbExclamation, SYNC_DIALOG_TITLE
        Exit Function
    End If

    PromptSelectAvailableTypeId = CStr(availableTypeIds(selectedIndex))
End Function


'================== CONTEXT & PROMPT HELPERS ==================

Private Function PrepareSyncContext(ByVal wb As Workbook, _
                                    ByRef typeSummary As String, _
                                    ByRef srcPaths As Collection, _
                                    ByRef repoSig As String, _
                                    ByRef wbSig As String, _
                                    ByRef reason As String) As Boolean
    Dim typeMap As Object
    Dim sortedTypes() As String
    Dim signatureItems() As String
    Dim srcPathMap As Object
    Dim srcPath As String
    Dim folderSig As String
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
    Set srcPathMap = CreateObject("Scripting.Dictionary")
    srcPathMap.CompareMode = vbTextCompare
    ReDim signatureItems(1 To UBound(sortedTypes))

    For i = LBound(sortedTypes) To UBound(sortedTypes)
        If Not ResolveTypeSourcePath(sortedTypes(i), srcPath, reason) Then Exit Function

        folderSig = ComputeFolderSignature(srcPath, "*.bas")
        If Len(folderSig) = 0 Then
            reason = MSG_SOURCE_UNAVAILABLE & vbCrLf & MSG_REQUEST_UPDATE
            Exit Function
        End If

        signatureItems(i) = LCase$(sortedTypes(i)) & "|" & folderSig

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

    PrepareSyncContext = True
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

    CollectTypeSourceMatches rootFolder, leafTypeId, fso, matches

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
                                     ByRef matches As Object)
    Dim childFolder As Object
    Dim srcPath As String

    For Each childFolder In currentFolder.SubFolders
        If StrComp(childFolder.Name, leafTypeId, vbTextCompare) = 0 Then
            srcPath = EnsureTrailingPathSeparator(childFolder.Path & Application.PathSeparator & "src")
            If fso.FolderExists(srcPath) Then
                matches(childFolder.Path) = srcPath
            End If
        End If

        CollectTypeSourceMatches childFolder, leafTypeId, fso, matches
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
    Dim values() As String
    Dim i As Long
    Dim results As New Collection

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(REPO_ROOT) Then
        Set GetAvailableTypeIdsCollection = results
        Exit Function
    End If

    Set rootFolder = fso.GetFolder(REPO_ROOT)
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    CollectAvailableTypeIdsRecursive rootFolder, fso, seen

    If seen.count > 0 Then
        ReDim values(1 To seen.count)
        i = 0
        Dim key As Variant
        For Each key In seen.Keys
            i = i + 1
            values(i) = CStr(key)
        Next key

        If UBound(values) > LBound(values) Then
            QuickSort values, LBound(values), UBound(values)
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

Private Sub CollectAvailableTypeIdsRecursive(ByVal currentFolder As Object, _
                                             ByVal fso As Object, _
                                             ByRef seen As Object)
    Dim subFolder As Object
    Dim srcFolderPath As String
    Dim typeId As String
    Dim rootPrefix As String

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

        CollectAvailableTypeIdsRecursive subFolder, fso, seen
    Next subFolder
End Sub

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

