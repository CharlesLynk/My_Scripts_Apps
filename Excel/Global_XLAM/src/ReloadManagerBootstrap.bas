Attribute VB_Name = "ReloadManagerBootstrap"
Option Explicit
'
'==========================================
' Module: ReloadManagerBootstrap
' Purpose: Safely refresh ReloadManager from repo source after startup.
' Scope: Deferred replacement of ReloadManager module in Global_CNL.xlam.
' Notes:
' - This module exists only to update ReloadManager.
' - Replacement runs via Application.OnTime to avoid self-replacement conflicts.
'==========================================

Private Const PROP_RELOADMANAGER_SIG As String = "GLB_RELOADMANAGER_SIG"
Private Const PROP_RELOADMANAGER_REPO_ROOT As String = "GLB_RELOADMANAGER_REPO_ROOT"
Private Const PROP_RELOADMANAGER_SRC_REL As String = "GLB_RELOADMANAGER_SRC_REL"
Private Const RELOADMANAGER_MODULE_NAME As String = "ReloadManager"
Private Const RELOADMANAGER_BOOTSTRAP_MODULE_NAME As String = "ReloadManagerBootstrap"
Private Const RELOADMANAGER_FILE_NAME As String = "ReloadManager.bas"
Private Const DEFERRED_REFRESH_SECONDS As Double = 2#
Private Const RETRY_REFRESH_SECONDS As Double = 5#
Private Const PHASE_ADVANCE_SECONDS As Double = 1#
Private Const DOC_PROP_TYPE_STRING As Long = 4

Private mRepoRoot As String
Private mAddInSrcRelativePath As String
Private mNextRefreshTime As Date
Private mRefreshQueued As Boolean
Private mBootstrapModuleName As String

Public Sub QueueReloadManagerRefresh(ByVal repoRoot As String, _
                                     ByVal addInSrcRelativePath As String, _
                                     Optional ByVal bootstrapModuleName As String = "")
    On Error Resume Next

    mRepoRoot = repoRoot
    mAddInSrcRelativePath = addInSrcRelativePath

    SetDocPropSafe ThisWorkbook, PROP_RELOADMANAGER_REPO_ROOT, mRepoRoot
    SetDocPropSafe ThisWorkbook, PROP_RELOADMANAGER_SRC_REL, mAddInSrcRelativePath

    If Len(bootstrapModuleName) > 0 Then
        mBootstrapModuleName = bootstrapModuleName
    ElseIf Len(mBootstrapModuleName) = 0 Then
        mBootstrapModuleName = ResolveBootstrapModuleName(ThisWorkbook)
    End If

    CancelCurrentQueuedRefresh
    ScheduleQueuedRefresh DEFERRED_REFRESH_SECONDS

    On Error GoTo 0
End Sub

Public Sub CancelQueuedReloadManagerRefresh()
    On Error Resume Next

    CancelCurrentQueuedRefresh

    On Error GoTo 0
End Sub

Public Sub RunQueuedReloadManagerRefresh()
    Dim sourceFile As String
    Dim newSig As String
    Dim currentSig As String
    Dim reason As String
    Dim replacementApplied As Boolean

    On Error GoTo SafeExit

    mRefreshQueued = False
    mNextRefreshTime = 0

    sourceFile = BuildReloadManagerSourcePath()
    If Len(sourceFile) = 0 Then Exit Sub
    If Len(Dir(sourceFile)) = 0 Then Exit Sub

    newSig = BuildFileSignature(sourceFile)
    If Len(newSig) = 0 Then Exit Sub

    currentSig = GetDocPropSafe(ThisWorkbook, PROP_RELOADMANAGER_SIG)
    If StrComp(newSig, currentSig, vbBinaryCompare) = 0 Then Exit Sub

    If Not CanAccessVBProjectSafe(ThisWorkbook, reason) Then Exit Sub

    ' Phase 1: remove canonical module and return so VBIDE fully releases the name.
    If StdModuleExistsByName(ThisWorkbook, RELOADMANAGER_MODULE_NAME) Then
        If Not RemoveStdModuleByName(ThisWorkbook, RELOADMANAGER_MODULE_NAME) Then
            QueueRetryRefresh
            Exit Sub
        End If

        QueuePhaseAdvanceRefresh
        Exit Sub
    End If

    replacementApplied = ReplaceReloadManagerModule(ThisWorkbook, sourceFile)
    If Not replacementApplied Then
        QueueRetryRefresh
        Exit Sub
    End If

    SetDocPropSafe ThisWorkbook, PROP_RELOADMANAGER_SIG, newSig

    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo SafeExit

SafeExit:
    On Error GoTo 0
End Sub

Private Function BuildReloadManagerSourcePath() As String
    Dim srcRoot As String

    If Len(mRepoRoot) = 0 Then
        mRepoRoot = GetDocPropSafe(ThisWorkbook, PROP_RELOADMANAGER_REPO_ROOT)
    End If
    If Len(mAddInSrcRelativePath) = 0 Then
        mAddInSrcRelativePath = GetDocPropSafe(ThisWorkbook, PROP_RELOADMANAGER_SRC_REL)
    End If

    If Len(mRepoRoot) = 0 Then Exit Function
    If Len(mAddInSrcRelativePath) = 0 Then Exit Function

    srcRoot = EnsureTrailingPathSeparator(mRepoRoot) & _
              EnsureTrailingPathSeparator(Replace$(mAddInSrcRelativePath, "\", Application.PathSeparator))

    BuildReloadManagerSourcePath = srcRoot & RELOADMANAGER_FILE_NAME
End Function

Private Function BuildFileSignature(ByVal filePath As String) As String
    On Error GoTo SafeExit

    BuildFileSignature = Format$(FileDateTime(filePath), "yyyy-mm-dd hh:nn:ss") & "|" & CStr(FileLen(filePath))

SafeExit:
    On Error GoTo 0
End Function

Private Function ReplaceReloadManagerModule(ByVal wb As Workbook, ByVal sourceFile As String) As Boolean
    Dim importedComp As Object
    Dim declaredName As String

    ' Always normalize to the canonical module name regardless of BAS header drift.
    declaredName = RELOADMANAGER_MODULE_NAME

    If Not RemoveStdModuleVariants(wb, declaredName) Then Exit Function
    If HasConflictingComponentName(wb, declaredName) Then Exit Function

    Set importedComp = wb.VBProject.VBComponents.Import(sourceFile)
    If importedComp Is Nothing Then Exit Function

    If Not NormalizeImportedStdModuleName(wb, importedComp, declaredName) Then Exit Function

    ReplaceReloadManagerModule = IsCanonicalStdModuleState(wb, declaredName)
End Function

Private Function RemoveStdModuleVariants(ByVal wb As Workbook, ByVal baseName As String) As Boolean
    Dim vbComp As Object
    Dim toRemove As Collection
    Dim itemName As Variant
    Dim removeErr As Long

    Set toRemove = New Collection

    For Each vbComp In wb.VBProject.VBComponents
        If vbComp.Type = 1 Then
            If IsAutoRenamedModuleName(vbComp.Name, baseName) Then
                If StrComp(vbComp.Name, baseName, vbTextCompare) <> 0 Then
                    toRemove.Add vbComp.Name
                End If
            End If
        End If
    Next vbComp

    For Each itemName In toRemove
        On Error Resume Next
        wb.VBProject.VBComponents.Remove wb.VBProject.VBComponents(CStr(itemName))
        removeErr = Err.Number
        Err.Clear
        On Error GoTo 0

        If removeErr <> 0 Then
            RemoveStdModuleVariants = False
            Exit Function
        End If
    Next itemName

    RemoveStdModuleVariants = True
End Function

Private Sub QueuePhaseAdvanceRefresh()
    ScheduleQueuedRefresh PHASE_ADVANCE_SECONDS
End Sub

Private Function HasConflictingComponentName(ByVal wb As Workbook, ByVal componentName As String) As Boolean
    Dim vbComp As Object

    For Each vbComp In wb.VBProject.VBComponents
        If StrComp(vbComp.Name, componentName, vbTextCompare) = 0 Then
            If vbComp.Type <> 1 Then
                HasConflictingComponentName = True
                Exit Function
            End If
        End If
    Next vbComp
End Function

Private Function NormalizeImportedStdModuleName(ByVal wb As Workbook, _
                                                ByVal importedComp As Object, _
                                                ByVal targetName As String) As Boolean
    Dim renameErr As Long

    If importedComp Is Nothing Then Exit Function

    If StrComp(importedComp.Name, targetName, vbTextCompare) = 0 Then
        NormalizeImportedStdModuleName = True
        Exit Function
    End If

    If StdModuleExistsByName(wb, targetName) Then
        If Not RemoveStdModuleByName(wb, targetName) Then Exit Function
    End If

    On Error Resume Next
    importedComp.Name = targetName
    renameErr = Err.Number
    Err.Clear
    On Error GoTo 0

    If renameErr <> 0 Then Exit Function
    NormalizeImportedStdModuleName = (StrComp(importedComp.Name, targetName, vbTextCompare) = 0)
End Function

Private Function StdModuleExistsByName(ByVal wb As Workbook, ByVal moduleName As String) As Boolean
    Dim vbComp As Object

    On Error Resume Next
    Set vbComp = wb.VBProject.VBComponents(moduleName)
    On Error GoTo 0

    If Not vbComp Is Nothing Then
        StdModuleExistsByName = (vbComp.Type = 1)
    End If
End Function

Private Function RemoveStdModuleByName(ByVal wb As Workbook, ByVal moduleName As String) As Boolean
    Dim vbComp As Object
    Dim removeErr As Long

    On Error Resume Next
    Set vbComp = wb.VBProject.VBComponents(moduleName)
    On Error GoTo 0

    If vbComp Is Nothing Then
        RemoveStdModuleByName = True
        Exit Function
    End If
    If vbComp.Type <> 1 Then Exit Function

    On Error Resume Next
    wb.VBProject.VBComponents.Remove vbComp
    removeErr = Err.Number
    Err.Clear
    On Error GoTo 0

    RemoveStdModuleByName = (removeErr = 0)
End Function

Private Function IsCanonicalStdModuleState(ByVal wb As Workbook, ByVal moduleName As String) As Boolean
    Dim vbComp As Object
    Dim variantCount As Long
    Dim canonicalCount As Long

    For Each vbComp In wb.VBProject.VBComponents
        If vbComp.Type = 1 Then
            If IsAutoRenamedModuleName(vbComp.Name, moduleName) Then
                variantCount = variantCount + 1
                If StrComp(vbComp.Name, moduleName, vbTextCompare) = 0 Then
                    canonicalCount = canonicalCount + 1
                End If
            End If
        End If
    Next vbComp

    IsCanonicalStdModuleState = (canonicalCount = 1 And variantCount = 1)
End Function

Private Sub QueueRetryRefresh()
    ScheduleQueuedRefresh RETRY_REFRESH_SECONDS
End Sub

Private Sub ScheduleQueuedRefresh(ByVal delaySeconds As Double)
    mNextRefreshTime = Now + (delaySeconds / 86400#)
    Application.OnTime EarliestTime:=mNextRefreshTime, _
                       Procedure:=GetScheduledProcedureName(), _
                       Schedule:=True
    mRefreshQueued = True
End Sub

Private Sub CancelCurrentQueuedRefresh()
    If mRefreshQueued And mNextRefreshTime > 0 Then
        Application.OnTime EarliestTime:=mNextRefreshTime, _
                           Procedure:=GetScheduledProcedureName(), _
                           Schedule:=False
    End If

    mRefreshQueued = False
    mNextRefreshTime = 0
End Sub

Private Function ResolveBootstrapModuleName(ByVal wb As Workbook) As String
    Dim vbComp As Object

    For Each vbComp In wb.VBProject.VBComponents
        If vbComp.Type = 1 Then
            If IsAutoRenamedModuleName(vbComp.Name, RELOADMANAGER_BOOTSTRAP_MODULE_NAME) Then
                ResolveBootstrapModuleName = vbComp.Name
                Exit Function
            End If
        End If
    Next vbComp

    ResolveBootstrapModuleName = RELOADMANAGER_BOOTSTRAP_MODULE_NAME
End Function

Private Function GetScheduledProcedureName() As String
    Dim moduleName As String

    moduleName = mBootstrapModuleName
    If Len(moduleName) = 0 Then
        moduleName = ResolveBootstrapModuleName(ThisWorkbook)
        mBootstrapModuleName = moduleName
    End If

    If Len(moduleName) = 0 Then moduleName = RELOADMANAGER_BOOTSTRAP_MODULE_NAME
    GetScheduledProcedureName = moduleName & ".RunQueuedReloadManagerRefresh"
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

Private Function CanAccessVBProjectSafe(ByVal wb As Workbook, ByRef reason As String) As Boolean
    On Error GoTo Fail
    Dim componentCount As Long
    componentCount = wb.VBProject.VBComponents.count
    CanAccessVBProjectSafe = True
    Exit Function

Fail:
    reason = "Cannot access VBA project for workbook '" & wb.Name & "'." & vbCrLf & _
             "Enable 'Trust access to the VBA project object model' in Excel Trust Center."
    CanAccessVBProjectSafe = False
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

Private Function GetDocPropSafe(ByVal wb As Workbook, ByVal propName As String) As String
    On Error Resume Next
    GetDocPropSafe = CStr(wb.CustomDocumentProperties(propName).Value)
    On Error GoTo 0
End Function

Private Sub SetDocPropSafe(ByVal wb As Workbook, ByVal propName As String, ByVal propValue As String)
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
