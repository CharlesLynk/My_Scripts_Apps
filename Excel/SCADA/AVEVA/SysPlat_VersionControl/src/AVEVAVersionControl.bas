Attribute VB_Name = "VersionControl"
Option Explicit

'==========================================
' Module: VersionControl
' Purpose: Manage manifest imports, diff generation, and backup handling.
' Scope: Primary/Secondary/Diff sheet maintenance for AVEVA workflows.
'==========================================

'================== USER CONFIGURATION ==================

' Standard sheet names (change here if you use different names)
Private Const SHEET_PRIMARY   As String = "Primary Galaxy"
Private Const SHEET_SECONDARY As String = "Secondary Galaxy"
Private Const SHEET_DIFF      As String = "Diff"
Private Const SHEET_MODE      As String = "_ManifestMode"
Private Const SHEET_DEBUG     As String = "_Debug"

' Optional abbreviations for header/status labels (leave blank for full sheet names)
Private Const PRIMARY_ABBREVIATED   As String = ""
Private Const SECONDARY_ABBREVIATED As String = ""

' Backup retention limit (full backups)
Private Const BACKUP_MAX_COUNT        As Long = 50

' Allow slight timestamp drift (minutes) before marking stale
Private Const STALE_TOLERANCE_MINUTES As Long = 5

'================== INTERNAL CONSTANTS ==================

Private Const DEBUG_AAPKG As Boolean = False

' Managed column locations (fixed positions)
Private Const COL_TAG_IDX     As Long = 1
Private Const COL_CFG_IDX     As Long = 2
Private Const COL_PARENT_IDX  As Long = 3
Private Const COL_DEPTH_IDX   As Long = 4
Private Const COL_LASTIMP_IDX As Long = 5

' Diff column locations (fixed positions)
Private Const DIFF_COL_TAG            As Long = 1
Private Const DIFF_COL_PRIMARY_CFG    As Long = 2
Private Const DIFF_COL_PRIMARY_LAST   As Long = 3
Private Const DIFF_COL_SECONDARY_CFG  As Long = 4
Private Const DIFF_COL_SECONDARY_LAST As Long = 5
Private Const DIFF_COL_STATUS         As Long = 6
Private Const DIFF_COL_DEPTH          As Long = 7
Private Const DIFF_COL_PARENT         As Long = 8
Private Const DIFF_COL_SORTKEY        As Long = 9

' Backup sheet naming/metadata
Private Const DIFF_BACKUP_PREFIX      As String = "Diff_"
Private Const DIFF_BACKUP_MANUAL_FLAG As String = "_MANUAL"
Private Const BACKUP_META_LABEL       As String = "__FULL_BACKUP__"
Private Const BACKUP_META_CELL        As String = "A1"
Private Const BACKUP_TIMESTAMP_CELL   As String = "B1"
Private Const BACKUP_CHECKSUM_CELL    As String = "D1"
Private Const BACKUP_PRIMARY_ROWS_CELL   As String = "B2"
Private Const BACKUP_SECONDARY_ROWS_CELL As String = "D2"
Private Const BACKUP_DIFF_ROWS_CELL      As String = "F2"
Private Const BACKUP_LABEL_PRIMARY_CELL  As String = "A2"
Private Const BACKUP_LABEL_SECONDARY_CELL As String = "C2"
Private Const BACKUP_LABEL_DIFF_CELL      As String = "E2"
Private Const BACKUP_SECTION_START_ROW    As Long = 4
Private Const BACKUP_SECTION_GAP_ROWS     As Long = 2
Private Const BACKUP_LABEL_PRIMARY_TEXT   As String = "PrimaryRows"
Private Const BACKUP_LABEL_SECONDARY_TEXT As String = "SecondaryRows"
Private Const BACKUP_LABEL_DIFF_TEXT      As String = "DiffRows"

' Status labels
Private Const STATUS_MATCH           As String = "MATCH"
Private Const STATUS_NA              As String = "N/A"
Private Const STATUS_DRIFT           As String = "DRIFT"

' Misc constants
Private Const MANUAL_STAMP            As String = "Manual"
Private Const MODE_TEMPLATE           As String = "TEMPLATE"
Private Const MODE_INSTANCE           As String = "INSTANCE"
Private Const MODE_GRAPHIC            As String = "GRAPHIC"
Private Const MANIFEST_XML_NAME       As String = "Manifest.xml"

'================== ENTRY POINTS ==================

' Prompts for XML and target sheet; performs upsert + dating
Public Sub ManifestUpsert_Run()
    Dim paths As Collection
    Set paths = PickXmlFiles()
    If paths Is Nothing Or paths.Count = 0 Then Exit Sub

    Dim targetSheetName As String
    targetSheetName = GetTargetSheetFromActive()
    If Len(targetSheetName) = 0 Then Exit Sub

    Dim confirmText As String
    confirmText = BuildImportConfirmText(paths, targetSheetName)
    If MsgBox(confirmText, vbYesNo + vbQuestion, "Confirm import") <> vbYes Then Exit Sub

    Dim resolvedImports As Collection
    Dim precheckError As String
    If Not BuildResolvedImportItems(paths, resolvedImports, precheckError) Then
        MsgBox precheckError, vbExclamation
        Exit Sub
    End If

    Dim expectedByTag As Object
    Dim mismatchItems As Collection
    Dim mismatchDetails As Object
    If Not ResolveBatchTagVersionChoice(resolvedImports, expectedByTag, mismatchItems, mismatchDetails) Then Exit Sub

    BuildDiffSheet
    Dim wsPrimary As Worksheet, wsSecondary As Worksheet, wsDiff As Worksheet
    Set wsPrimary = GetWorksheetByName(SHEET_PRIMARY)
    Set wsSecondary = GetWorksheetByName(SHEET_SECONDARY)
    Set wsDiff = GetWorksheetByName(SHEET_DIFF)
    If wsPrimary Is Nothing Or wsSecondary Is Nothing Or wsDiff Is Nothing Then
        MsgBox "Backup failed: Primary/Secondary/Diff sheets must exist.", vbExclamation
        Exit Sub
    End If

    Dim wsBackup As Worksheet
    Dim createdNew As Boolean
    Dim skipMessage As String
    Dim showSkipMessage As Boolean
    showSkipMessage = (paths.Count = 1)
    Set wsBackup = CreateHiddenFullBackup(wsPrimary, wsSecondary, wsDiff, createdNew, skipMessage)
    If wsBackup Is Nothing Then
        MsgBox "Backup failed: Unable to create backup sheet.", vbExclamation
        Exit Sub
    End If
    If showSkipMessage And Not createdNew And Len(skipMessage) > 0 Then MsgBox skipMessage, vbInformation

    Dim i As Long
    For i = 1 To resolvedImports.Count
        Dim importItem As Object
        Dim sourcePath As String, xmlPath As String
        Dim importTime As Date
        Dim errMessage As String
        Set importItem = resolvedImports(i)
        sourcePath = CStr(importItem("SourcePath"))
        xmlPath = CStr(importItem("XmlPath"))
        importTime = CDate(importItem("ImportTime"))

        If Not ManifestUpsert_IntoSheet(xmlPath, targetSheetName, importTime, errMessage, False, False) Then
            HandleImportFailure wsBackup, createdNew, "Import failed for " & SourceDisplayName(sourcePath) & ": " & errMessage
            Exit Sub
        End If
    Next i

    BuildDiffSheet
    If paths.Count = 1 Then
        MsgBox "Upsert complete - " & targetSheetName, vbInformation
    Else
        MsgBox "Batch import complete - " & CStr(paths.Count) & " files", vbInformation
        If Not createdNew And Len(skipMessage) > 0 Then MsgBox skipMessage, vbInformation
    End If

    If Not mismatchItems Is Nothing Then
        If mismatchItems.Count > 0 Then
            MsgBox BuildTagVersionMismatchSummary(mismatchItems, mismatchDetails), vbExclamation, "Version mismatch summary"
        End If
    End If

    If createdNew Then
        EnforceBackupCap BACKUP_MAX_COUNT
    End If

    EnsureEventsOn
End Sub

' Refreshes the Diff sheet (macro entry point)
Public Sub BuildDiffSheet_Run()
    BuildDiffSheet
    MsgBox "Diff sheet refreshed.", vbInformation
End Sub

' Button action: reapply status/highlights after manual edits
Public Sub UpdateDiffStatus_Run()
    EnsureEventsOn
    ReapplyDiffSheetFormatting
    MsgBox "Diff status updated.", vbInformation
End Sub

' Button action: review/delete very hidden sheets
Public Sub ManageHiddenSheets_Run()
    Dim names As Collection
    Dim listText As String
    Set names = New Collection
    listText = ListVeryHiddenSheets(names)

    If Len(listText) = 0 Then
        MsgBox "No very hidden sheets found.", vbInformation
        Exit Sub
    End If

    Dim s As String
    s = InputBox("Very hidden sheets:" & vbCrLf & listText & vbCrLf & _
                 "Enter numbers or ranges to delete (e.g., 1,3-5), or leave blank to cancel.", _
                 "Manage hidden sheets")

    If Len(Trim$(s)) = 0 Then Exit Sub
    DeleteVeryHiddenSheetsByIndex names, s
End Sub

' Button action: create/update manifest mode sheet (Template/Instance/Graphic)
Public Sub SetManifestMode_Run()
    Dim ws As Worksheet
    Set ws = EnsureManifestModeSheet()

    Dim current As String
    current = UCase$(Trim$(CStr(ws.Range("A1").Value)))
    If Len(current) = 0 Then current = MODE_TEMPLATE

    Dim resp As String
    resp = InputBox("Enter manifest mode: " & MODE_TEMPLATE & ", " & MODE_INSTANCE & ", or " & MODE_GRAPHIC, _
                    "Set manifest mode", current)
    resp = UCase$(Trim$(resp))
    If Len(resp) = 0 Then Exit Sub

    If resp <> MODE_TEMPLATE And resp <> MODE_INSTANCE And resp <> MODE_GRAPHIC Then
        MsgBox "Invalid mode. Use " & MODE_TEMPLATE & ", " & MODE_INSTANCE & ", or " & MODE_GRAPHIC & ".", vbExclamation
        Exit Sub
    End If

    ws.Range("A1").Value = resp
    ws.Visible = xlSheetHidden
    MsgBox "Manifest mode set to: " & resp, vbInformation
End Sub


' Button action: rebuild Diff from a backup sheet
Public Sub RebuildDiffFromBackup_Run()
    Dim names As Collection
    Dim listText As String
    Set names = New Collection
    listText = ListBackupSheets(names)

    If Len(listText) = 0 Then
        MsgBox "No backup sheets found.", vbExclamation
        Exit Sub
    End If

    Dim s As String
    s = InputBox("Backup sheets:" & vbCrLf & listText & vbCrLf & _
                 "Enter the number to restore from.", _
                 "Rebuild Diff from backup")

    If Len(Trim$(s)) = 0 Then Exit Sub
    Dim idx As Long
    idx = val(Trim$(s))
    If idx < 1 Or idx > names.Count Then
        MsgBox "Invalid selection.", vbExclamation
        Exit Sub
    End If

    Dim wsBackup As Worksheet
    Set wsBackup = GetWorksheetByName(CStr(names(idx)))
    If wsBackup Is Nothing Then
        MsgBox "Backup sheet not found.", vbExclamation
        Exit Sub
    End If
    If Not IsFullBackupSheet(wsBackup) Then
        MsgBox "Selected backup is not a full backup with a valid checksum.", vbExclamation
        Exit Sub
    End If

    BuildDiffSheet
    Dim wsPrimary As Worksheet, wsSecondary As Worksheet, wsDiff As Worksheet
    Set wsPrimary = GetWorksheetByName(SHEET_PRIMARY)
    Set wsSecondary = GetWorksheetByName(SHEET_SECONDARY)
    Set wsDiff = GetWorksheetByName(SHEET_DIFF)
    If wsPrimary Is Nothing Or wsSecondary Is Nothing Or wsDiff Is Nothing Then
        MsgBox "Restore failed: Primary/Secondary/Diff sheets must exist.", vbExclamation
        Exit Sub
    End If

    Dim preBackup As Worksheet
    Dim createdNew As Boolean
    Dim skipMessage As String
    Set preBackup = CreateHiddenFullBackup(wsPrimary, wsSecondary, wsDiff, createdNew, skipMessage)
    If preBackup Is Nothing Then
        MsgBox "Restore failed: Unable to create pre-restore backup.", vbExclamation
        Exit Sub
    End If
    If Not createdNew And Len(skipMessage) > 0 Then MsgBox skipMessage, vbInformation

    Dim prevEvents As Boolean, prevCalc As XlCalculation, prevScreen As Boolean
    prevEvents = Application.EnableEvents
    prevCalc = Application.Calculation
    prevScreen = Application.ScreenUpdating
    On Error GoTo CleanExit
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    If Not RestoreDiffOnlyFromBackup(wsBackup) Then
        MsgBox "Rebuild failed: backup format not recognized.", vbExclamation
        GoTo CleanExit
    End If

CleanExit:
    Application.EnableEvents = prevEvents
    Application.Calculation = prevCalc
    Application.ScreenUpdating = prevScreen
    If Err.Number <> 0 Then
        MsgBox "Rebuild failed: " & Err.Description, vbExclamation
        Err.Clear
        Exit Sub
    End If

    MsgBox "Diff rebuilt from backup.", vbInformation
End Sub

' Button action: restore Primary/Secondary + Diff from a Diff backup sheet
Public Sub RestoreAllFromDiffBackup_Run()
    Dim names As Collection
    Dim listText As String
    Set names = New Collection
    listText = ListBackupSheets(names)

    If Len(listText) = 0 Then
        MsgBox "No backup sheets found.", vbExclamation
        Exit Sub
    End If

    Dim s As String
    s = InputBox("Backup sheets:" & vbCrLf & listText & vbCrLf & _
                 "Enter the number to restore from.", _
                 "Restore all from backup")

    If Len(Trim$(s)) = 0 Then Exit Sub
    Dim idx As Long
    idx = val(Trim$(s))
    If idx < 1 Or idx > names.Count Then
        MsgBox "Invalid selection.", vbExclamation
        Exit Sub
    End If

    Dim wsBackup As Worksheet
    Set wsBackup = GetWorksheetByName(CStr(names(idx)))
    If wsBackup Is Nothing Then
        MsgBox "Backup sheet not found.", vbExclamation
        Exit Sub
    End If
    If Not IsFullBackupSheet(wsBackup) Then
        MsgBox "Selected backup is not a full backup with a valid checksum.", vbExclamation
        Exit Sub
    End If

    BuildDiffSheet
    Dim wsPrimary As Worksheet, wsSecondary As Worksheet, wsDiff As Worksheet
    Set wsPrimary = GetWorksheetByName(SHEET_PRIMARY)
    Set wsSecondary = GetWorksheetByName(SHEET_SECONDARY)
    Set wsDiff = GetWorksheetByName(SHEET_DIFF)
    If wsPrimary Is Nothing Or wsSecondary Is Nothing Or wsDiff Is Nothing Then
        MsgBox "Restore failed: Primary/Secondary/Diff sheets must exist.", vbExclamation
        Exit Sub
    End If

    Dim preBackup As Worksheet
    Dim createdNew As Boolean
    Dim skipMessage As String
    Set preBackup = CreateHiddenFullBackup(wsPrimary, wsSecondary, wsDiff, createdNew, skipMessage)
    If preBackup Is Nothing Then
        MsgBox "Restore failed: Unable to create pre-restore backup.", vbExclamation
        Exit Sub
    End If
    If Not createdNew And Len(skipMessage) > 0 Then MsgBox skipMessage, vbInformation

    Dim prevEvents As Boolean, prevCalc As XlCalculation, prevScreen As Boolean
    prevEvents = Application.EnableEvents
    prevCalc = Application.Calculation
    prevScreen = Application.ScreenUpdating
    On Error GoTo CleanExit
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    If Not RestoreFromDiffBackup(wsBackup) Then
        MsgBox "Restore failed: backup format not recognized.", vbExclamation
        GoTo CleanExit
    End If

CleanExit:
    Application.EnableEvents = prevEvents
    Application.Calculation = prevCalc
    Application.ScreenUpdating = prevScreen
    If Err.Number <> 0 Then
        MsgBox "Restore failed: " & Err.Description, vbExclamation
        Err.Clear
        Exit Sub
    End If

    MsgBox "Primary/Secondary/Diff restored from backup.", vbInformation

    If createdNew Then
        EnforceBackupCap BACKUP_MAX_COUNT
    End If
End Sub

' Builds/refreshes the Diff sheet
Public Sub BuildDiffSheet()
    Dim prevCalc As XlCalculation
    Dim prevEvents As Boolean
    Dim prevScreen As Boolean
    prevCalc = Application.Calculation
    prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating
    On Error GoTo CleanExit

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim wsPrimary As Worksheet, wsSecondary As Worksheet
    On Error Resume Next
    Set wsPrimary = ThisWorkbook.Worksheets(SHEET_PRIMARY)
    Set wsSecondary = ThisWorkbook.Worksheets(SHEET_SECONDARY)
    On Error GoTo 0

    If wsPrimary Is Nothing Or wsSecondary Is Nothing Then
        MsgBox "Both '" & SHEET_PRIMARY & "' and '" & SHEET_SECONDARY & "' must exist.", vbExclamation
        GoTo CleanExit
    End If

    Dim primaryData As Object, secondaryData As Object
    Set primaryData = ReadSheetToDict(wsPrimary)   ' key=Tag -> (Cfg, LastImp, Depth, Parent)
    Set secondaryData = ReadSheetToDict(wsSecondary)

    Dim latestPrimary As Date, latestSecondary As Date
    latestPrimary = LatestImportOnSheet(wsPrimary)
    latestSecondary = LatestImportOnSheet(wsSecondary)

    ' Prepare Diff sheet
    Dim wsDiff As Worksheet
    Set wsDiff = GetWorksheetByName(SHEET_DIFF)
    If wsDiff Is Nothing Then
        Set wsDiff = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsDiff.name = SHEET_DIFF
    Else
        wsDiff.Visible = xlSheetVisible
        wsDiff.Cells.Clear
    End If

    ' Headers (helper columns for sort)
    Dim hdr As Variant
    hdr = BuildDiffHeaders(wsPrimary, wsSecondary)
    Dim c As Long
    For c = LBound(hdr) To UBound(hdr)
        wsDiff.Cells(1, c + 1).Value = hdr(c)
    Next c

    ' Union of tags
    Dim tag As Variant, tags As Object
    Set tags = CreateObject("Scripting.Dictionary"): tags.CompareMode = 1
    For Each tag In primaryData.Keys: tags(tag) = True: Next
    For Each tag In secondaryData.Keys: tags(tag) = True: Next

    Dim instanceMode As Boolean
    instanceMode = IsInstanceMode()

    ' Build parent map for sorting (templates only)
    Dim parentMap As Object
    If Not instanceMode Then
        Set parentMap = CreateObject("Scripting.Dictionary"): parentMap.CompareMode = 1
    End If

    ' Write rows
    Dim r As Long: r = 2
    For Each tag In tags.Keys
        Dim primaryCfg As Variant, primaryImp As Variant, secondaryCfg As Variant, secondaryImp As Variant
        Dim primaryDepth As Variant, secondaryDepth As Variant, primaryParent As Variant, secondaryParent As Variant
        primaryCfg = Empty: primaryImp = Empty: secondaryCfg = Empty: secondaryImp = Empty
        primaryDepth = Empty: secondaryDepth = Empty
        primaryParent = Empty: secondaryParent = Empty

        If primaryData.Exists(tag) Then
            primaryCfg = primaryData(tag)(0)
            primaryImp = primaryData(tag)(1)
            primaryDepth = primaryData(tag)(2)
            primaryParent = primaryData(tag)(3)
        End If
        If secondaryData.Exists(tag) Then
            secondaryCfg = secondaryData(tag)(0)
            secondaryImp = secondaryData(tag)(1)
            secondaryDepth = secondaryData(tag)(2)
            secondaryParent = secondaryData(tag)(3)
        End If

        wsDiff.Cells(r, DIFF_COL_TAG).Value = tag
        wsDiff.Cells(r, DIFF_COL_PRIMARY_CFG).Value = primaryCfg
        wsDiff.Cells(r, DIFF_COL_PRIMARY_LAST).Value = primaryImp
        wsDiff.Cells(r, DIFF_COL_SECONDARY_CFG).Value = secondaryCfg
        wsDiff.Cells(r, DIFF_COL_SECONDARY_LAST).Value = secondaryImp

        ApplyDiffRowFormatting wsDiff, r, primaryCfg, primaryImp, secondaryCfg, secondaryImp, latestPrimary, latestSecondary, primaryDepth, secondaryDepth

        Dim depthVal As Long
        depthVal = 0
        If Not IsEmptyOrNull(primaryDepth) Then depthVal = CLng(primaryDepth)
        If depthVal = 0 And Not IsEmptyOrNull(secondaryDepth) Then depthVal = CLng(secondaryDepth)
        wsDiff.Cells(r, DIFF_COL_DEPTH).Value = depthVal

        Dim parentVal As String
        parentVal = ""
        If Not IsEmptyOrNull(primaryParent) Then parentVal = CStr(primaryParent)
        If Len(parentVal) = 0 And Not IsEmptyOrNull(secondaryParent) Then parentVal = CStr(secondaryParent)
        wsDiff.Cells(r, DIFF_COL_PARENT).Value = parentVal
        If Not instanceMode Then
            parentMap(tag) = parentVal
        End If

        r = r + 1
    Next tag

    ' Compute hierarchical sort keys
    Dim rr As Long
    If instanceMode Then
        For rr = 2 To LastUsedRow(wsDiff)
            wsDiff.Cells(rr, DIFF_COL_SORTKEY).Value = BuildToolsetSortKey(wsDiff.Cells(rr, DIFF_COL_PARENT).Value, wsDiff.Cells(rr, DIFF_COL_TAG).Value)
        Next rr
    Else
        Dim sortCache As Object
        Dim childMap As Object
        Set childMap = BuildChildPresenceMap(parentMap)
        Set sortCache = CreateObject("Scripting.Dictionary"): sortCache.CompareMode = 1
        For rr = 2 To LastUsedRow(wsDiff)
            tag = CStr(wsDiff.Cells(rr, DIFF_COL_TAG).Value)
            wsDiff.Cells(rr, DIFF_COL_SORTKEY).Value = BuildSortKey(tag, parentMap, childMap, sortCache)
        Next rr
    End If

    ' Sort by depth, parent, tag
    Dim lr As Long
    lr = LastUsedRow(wsDiff)
    If lr >= 2 Then
        With wsDiff.Sort
            .SortFields.Clear
            .SortFields.Add key:=wsDiff.Cells(1, DIFF_COL_SORTKEY).Resize(lr), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange wsDiff.Range(wsDiff.Cells(1, DIFF_COL_TAG), wsDiff.Cells(lr, DIFF_COL_SORTKEY))
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .Apply
        End With
    End If

    wsDiff.Columns.AutoFit
    HideDiffHelperColumns wsDiff

    ' Freeze header row (safe)
    FreezeHeaderRow wsDiff

    ' Full backups are handled by explicit operations; Diff rebuilds do not create backups here.

CleanExit:
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    If Err.Number <> 0 Then
        MsgBox "Diff build failed: " & Err.Description, vbExclamation
        Err.Clear
    End If
End Sub

' Reapply status/highlighting on demand (useful after manual edits)
Public Sub ReapplyDiffSheetFormatting()
    Dim wsDiff As Worksheet
    Set wsDiff = GetWorksheetByName(SHEET_DIFF)
    If wsDiff Is Nothing Then Exit Sub

    Dim wsPrimary As Worksheet, wsSecondary As Worksheet
    On Error Resume Next
    Set wsPrimary = ThisWorkbook.Worksheets(SHEET_PRIMARY)
    Set wsSecondary = ThisWorkbook.Worksheets(SHEET_SECONDARY)
    On Error GoTo 0
    If wsPrimary Is Nothing Or wsSecondary Is Nothing Then Exit Sub

    Dim primaryData As Object, secondaryData As Object
    Set primaryData = ReadSheetToDict(wsPrimary)
    Set secondaryData = ReadSheetToDict(wsSecondary)

    Dim latestPrimary As Date, latestSecondary As Date
    latestPrimary = LatestImportOnSheet(wsPrimary)
    latestSecondary = LatestImportOnSheet(wsSecondary)

    Dim lr As Long, r As Long, tag As String
    lr = LastUsedRow(wsDiff)
    Dim instanceMode As Boolean
    instanceMode = IsInstanceMode()
    Dim parentMap As Object
    If Not instanceMode Then
        Set parentMap = CreateObject("Scripting.Dictionary"): parentMap.CompareMode = 1
    End If

    For r = 2 To lr
        tag = CStr(wsDiff.Cells(r, DIFF_COL_TAG).Value)
        Dim primaryCfg As Variant, primaryImp As Variant, secondaryCfg As Variant, secondaryImp As Variant
        Dim primaryDepth As Variant, secondaryDepth As Variant, primaryParent As Variant, secondaryParent As Variant
        primaryCfg = wsDiff.Cells(r, DIFF_COL_PRIMARY_CFG).Value
        primaryImp = wsDiff.Cells(r, DIFF_COL_PRIMARY_LAST).Value
        secondaryCfg = wsDiff.Cells(r, DIFF_COL_SECONDARY_CFG).Value
        secondaryImp = wsDiff.Cells(r, DIFF_COL_SECONDARY_LAST).Value

        primaryDepth = Empty: secondaryDepth = Empty
        primaryParent = Empty: secondaryParent = Empty
        If primaryData.Exists(tag) Then
            primaryDepth = primaryData(tag)(2)
            primaryParent = primaryData(tag)(3)
        End If
        If secondaryData.Exists(tag) Then
            secondaryDepth = secondaryData(tag)(2)
            secondaryParent = secondaryData(tag)(3)
        End If

        ApplyDiffRowFormatting wsDiff, r, primaryCfg, primaryImp, secondaryCfg, secondaryImp, latestPrimary, latestSecondary, primaryDepth, secondaryDepth

        Dim depthVal As Long
        depthVal = 0
        If Not IsEmptyOrNull(primaryDepth) Then depthVal = CLng(primaryDepth)
        If depthVal = 0 And Not IsEmptyOrNull(secondaryDepth) Then depthVal = CLng(secondaryDepth)
        wsDiff.Cells(r, DIFF_COL_DEPTH).Value = depthVal

        Dim parentVal As String
        parentVal = ""
        If Not IsEmptyOrNull(primaryParent) Then parentVal = CStr(primaryParent)
        If Len(parentVal) = 0 And Not IsEmptyOrNull(secondaryParent) Then parentVal = CStr(secondaryParent)
        wsDiff.Cells(r, DIFF_COL_PARENT).Value = parentVal
        If Not instanceMode Then
            parentMap(tag) = parentVal
        End If
    Next r

    Dim rr As Long
    If instanceMode Then
        For rr = 2 To lr
            wsDiff.Cells(rr, DIFF_COL_SORTKEY).Value = BuildToolsetSortKey(wsDiff.Cells(rr, DIFF_COL_PARENT).Value, wsDiff.Cells(rr, DIFF_COL_TAG).Value)
        Next rr
    Else
        Dim sortCache As Object
        Dim childMap As Object
        Set childMap = BuildChildPresenceMap(parentMap)
        Set sortCache = CreateObject("Scripting.Dictionary"): sortCache.CompareMode = 1
        For rr = 2 To lr
            tag = CStr(wsDiff.Cells(rr, DIFF_COL_TAG).Value)
            wsDiff.Cells(rr, DIFF_COL_SORTKEY).Value = BuildSortKey(tag, parentMap, childMap, sortCache)
        Next rr
    End If

    If lr >= 2 Then
        With wsDiff.Sort
            .SortFields.Clear
            .SortFields.Add key:=wsDiff.Cells(1, DIFF_COL_SORTKEY).Resize(lr), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange wsDiff.Range(wsDiff.Cells(1, DIFF_COL_TAG), wsDiff.Cells(lr, DIFF_COL_SORTKEY))
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .Apply
        End With
    End If

    HideDiffHelperColumns wsDiff
End Sub


' Hook this from the Primary/Secondary worksheets' Worksheet_Change event
' Stamps LastImport as "Manual" when Config_Version is edited directly.
Public Sub PrimarySecondarySheet_OnChange(ByVal Target As Range)
    On Error GoTo SafeExit
    If Target Is Nothing Then Exit Sub

    Dim ws As Worksheet
    Set ws = Target.Worksheet
    If ws Is Nothing Then Exit Sub

    If ws.name <> SHEET_PRIMARY And ws.name <> SHEET_SECONDARY Then Exit Sub

    Dim cCfg As Long, cLast As Long
    cCfg = COL_CFG_IDX
    cLast = COL_LASTIMP_IDX

    Dim hit As Range
    Set hit = Intersect(Target, ws.Columns(cCfg))
    If hit Is Nothing Then Exit Sub

    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    Dim c As Range
    For Each c In hit.Cells
        If c.Row > 1 Then ws.Cells(c.Row, cLast).Value = MANUAL_STAMP
    Next c

SafeExit:
    Application.EnableEvents = prevEvents
End Sub

' ========= Upsert Core =========

' Upserts a manifest XML into the given sheet, updating Config_Version & LastImport; adds new rows; preserves existing others.
Public Function ManifestUpsert_IntoSheet(ByVal xmlPath As String, ByVal targetSheetName As String, Optional ByVal importTime As Variant, _
                                         Optional ByRef errMessage As String, Optional ByVal showMessage As Boolean = True, _
                                         Optional ByVal rebuildDiff As Boolean = True) As Boolean
    Dim rowsData As Variant
    rowsData = ParseManifestToRows(xmlPath) ' -> (Tag, Config_Version, ParentTag, Depth)

    If IsEmpty(rowsData) Then
        If IsInstanceMode() Then
            errMessage = "No <instance> nodes found in manifest."
        Else
            errMessage = "No <template> nodes found in manifest."
        End If
        ManifestUpsert_IntoSheet = False
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = EnsureTargetSheet(targetSheetName)
    EnsureHeaders ws

    Dim prevCalc As XlCalculation
    Dim prevEvents As Boolean
    Dim prevScreen As Boolean
    prevCalc = Application.Calculation
    prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating
    On Error GoTo CleanExit

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    UpsertRows ws, rowsData         ' writes Tag/Config/Parent/Depth
    ' Set LastImport = Now() for the rows just updated/inserted
    StampLastImport ws, rowsData, importTime

    SortByHierarchy ws
    ws.Columns.AutoFit

CleanExit:
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    If Err.Number <> 0 Then
        errMessage = Err.Description
        Err.Clear
        ManifestUpsert_IntoSheet = False
        Exit Function
    End If

    If rebuildDiff Then
        BuildDiffSheet
    End If
    If showMessage Then
        MsgBox "Upsert complete - " & ws.name, vbInformation
    End If
    ManifestUpsert_IntoSheet = True
End Function

' ========= UI Helpers =========

Private Function PickXmlFiles() As Collection
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = True
        .Title = "Select AVEVA Manifest XML or AAPKG"
        .Filters.Clear
        .Filters.Add "XML or AAPKG", "*.xml;*.aapkg"
        If .Show <> -1 Then
            Set PickXmlFiles = New Collection
        Else
            Dim items As Collection
            Set items = New Collection
            Dim i As Long
            For i = 1 To .SelectedItems.Count
                items.Add .SelectedItems(i)
            Next i
            Set PickXmlFiles = items
        End If
    End With
End Function

Private Function BuildImportConfirmText(ByVal paths As Collection, ByVal targetSheetName As String) As String
    Dim s As String
    Dim count As Long
    count = paths.Count

    Dim maxList As Long
    maxList = count

    Dim i As Long
    For i = 1 To count
        If i > maxList Then Exit For
        s = s & "- " & SourceDisplayName(CStr(paths(i))) & vbCrLf
    Next i
    If count > maxList Then
        s = s & "- ... (" & CStr(count - maxList) & " more)" & vbCrLf
    End If

    Dim header As String
    If count = 1 Then
        header = "Import file:" & vbCrLf
    Else
        header = "Import files (" & CStr(count) & "):" & vbCrLf
    End If

    BuildImportConfirmText = header & s & vbCrLf & "Destination: " & targetSheetName & vbCrLf & vbCrLf & "Proceed?"
End Function

Private Function BuildResolvedImportItems(ByVal paths As Collection, ByRef imports As Collection, ByRef errMessage As String) As Boolean
    Set imports = New Collection

    Dim i As Long
    For i = 1 To paths.Count
        Dim sourcePath As String
        Dim xmlPath As String
        Dim importTime As Date
        Dim importItem As Object
        Dim detectedVersion As String
        Dim rowsData As Variant
        Dim tagVersions As Object

        sourcePath = CStr(paths(i))
        xmlPath = ResolveSelectedManifest(sourcePath, importTime)
        If Len(xmlPath) = 0 Then
            errMessage = "Failed to resolve manifest from: " & SourceDisplayName(sourcePath)
            Exit Function
        End If

        rowsData = ParseManifestToRows(xmlPath)
        Set tagVersions = BuildTagVersionMapFromRows(rowsData)

        If tagVersions Is Nothing Then
            errMessage = "Failed to parse tag/version data from: " & SourceDisplayName(sourcePath)
            Exit Function
        End If

        If IsEmpty(rowsData) Then
            detectedVersion = "(unknown)"
        Else
            detectedVersion = MostCommonVersionInRows(rowsData)
        End If

        Set importItem = CreateObject("Scripting.Dictionary")
        importItem.CompareMode = 1
        importItem("SourcePath") = sourcePath
        importItem("DisplayName") = SourceDisplayName(sourcePath)
        importItem("XmlPath") = xmlPath
        importItem("ImportTime") = importTime
        importItem("Version") = detectedVersion
        Set importItem("TagVersions") = tagVersions
        imports.Add importItem
    Next i

    BuildResolvedImportItems = True
End Function

Private Function ResolveBatchTagVersionChoice(ByVal imports As Collection, ByRef expectedByTag As Object, ByRef mismatchItems As Collection, ByRef mismatchDetails As Object) As Boolean
    Set expectedByTag = CreateObject("Scripting.Dictionary")
    expectedByTag.CompareMode = 1
    Set mismatchItems = New Collection
    Set mismatchDetails = CreateObject("Scripting.Dictionary")
    mismatchDetails.CompareMode = 1

    If imports Is Nothing Then
        ResolveBatchTagVersionChoice = True
        Exit Function
    End If

    Dim tagOrder As Collection
    Set tagOrder = New Collection

    Dim knownTags As Object
    Set knownTags = CreateObject("Scripting.Dictionary")
    knownTags.CompareMode = 1

    Dim i As Long
    For i = 1 To imports.Count
        Dim tagVersions As Object
        Set tagVersions = imports(i)("TagVersions")

        Dim tagKey As Variant
        For Each tagKey In tagVersions.Keys
            Dim tagName As String
            tagName = CStr(tagKey)
            If Not knownTags.Exists(tagName) Then
                knownTags(tagName) = True
                tagOrder.Add tagName
            End If
        Next tagKey
    Next i

    If tagOrder.Count = 0 Then
        ResolveBatchTagVersionChoice = True
        Exit Function
    End If

    For i = 1 To tagOrder.Count
        Dim versionOrder As Collection
        Dim versionCounts As Object
        Dim currentTag As String
        Dim j As Long

        currentTag = CStr(tagOrder(i))
        Set versionOrder = New Collection
        Set versionCounts = CreateObject("Scripting.Dictionary")
        versionCounts.CompareMode = 1

        For j = 1 To imports.Count
            If ImportContainsTag(imports(j), currentTag) Then
                Dim versionText As String
                versionText = GetImportTagVersion(imports(j), currentTag)
                If Not versionCounts.Exists(versionText) Then
                    versionCounts(versionText) = 0
                    versionOrder.Add versionText
                End If
                versionCounts(versionText) = CLng(versionCounts(versionText)) + 1
            End If
        Next j

        If versionOrder.Count = 0 Then
            GoTo NextTag
        End If

        If versionCounts.Count <= 1 Then
            expectedByTag(currentTag) = CStr(versionOrder(1))
        Else
            Dim defaultIdx As Long
            defaultIdx = 1
            Dim maxCount As Long
            maxCount = -1

            For j = 1 To versionOrder.Count
                Dim countForVersion As Long
                countForVersion = CLng(versionCounts(CStr(versionOrder(j))))
                If countForVersion > maxCount Then
                    maxCount = countForVersion
                    defaultIdx = j
                End If
            Next j

            Dim prompt As String
            prompt = BuildTagVersionSelectionPrompt(currentTag, imports, versionOrder, versionCounts, defaultIdx)

            Dim choice As String
            Dim chosenIdx As Long
            Do
                choice = InputBox(prompt, "Version mismatch detected", CStr(defaultIdx))
                choice = Trim$(choice)
                If Len(choice) = 0 Then Exit Function

                If IsNumeric(choice) Then
                    chosenIdx = CLng(choice)
                    If chosenIdx >= 1 And chosenIdx <= versionOrder.Count Then Exit Do
                End If

                MsgBox "Invalid selection. Enter one number from the list.", vbExclamation
            Loop

            expectedByTag(currentTag) = CStr(versionOrder(chosenIdx))
        End If
NextTag:
    Next i

    For i = 1 To imports.Count
        Dim importItem As Object
        Dim fileDetails As String
        Set importItem = imports(i)

        Dim expectedTag As Variant
        For Each expectedTag In expectedByTag.Keys
            Dim expectedVersion As String
            Dim foundVersion As String

            expectedVersion = CStr(expectedByTag(CStr(expectedTag)))
            If ImportContainsTag(importItem, CStr(expectedTag)) Then
                foundVersion = GetImportTagVersion(importItem, CStr(expectedTag))
                If StrComp(foundVersion, expectedVersion, vbTextCompare) <> 0 Then
                    If Len(fileDetails) > 0 Then fileDetails = fileDetails & "; "
                    fileDetails = fileDetails & CStr(expectedTag) & " [expected: " & expectedVersion & ", found: " & foundVersion & "]"
                End If
            End If
        Next expectedTag

        If Len(fileDetails) > 0 Then
            mismatchItems.Add importItem
            mismatchDetails(CStr(importItem("DisplayName"))) = fileDetails
        End If
    Next i

    ResolveBatchTagVersionChoice = True
End Function

Private Function BuildTagVersionSelectionPrompt(ByVal tagName As String, ByVal imports As Collection, ByVal versionOrder As Collection, ByVal versionCounts As Object, ByVal defaultIdx As Long) As String
    Dim s As String
    Dim i As Long

    s = "Version disagreement detected for tag: " & tagName & vbCrLf & vbCrLf
    s = s & "Choose the expected (correct) config_version for this tag:" & vbCrLf

    For i = 1 To versionOrder.Count
        Dim versionText As String
        versionText = CStr(versionOrder(i))
        s = s & CStr(i) & ") " & versionText & "  [" & CStr(versionCounts(versionText)) & " file(s)]"
        If i = defaultIdx Then s = s & "  <-- majority"
        s = s & vbCrLf
    Next i

    s = s & vbCrLf & "Enter the number to continue, or leave blank to cancel."

    Dim details As String
    details = BuildTagVersionBreakdown(tagName, imports, versionOrder)
    If Len(details) > 0 Then
        s = s & vbCrLf & vbCrLf & details
    End If

    BuildTagVersionSelectionPrompt = s
End Function

Private Function BuildTagVersionBreakdown(ByVal tagName As String, ByVal imports As Collection, ByVal versionOrder As Collection) As String
    Dim s As String
    Dim i As Long, j As Long

    For i = 1 To versionOrder.Count
        Dim versionText As String
        versionText = CStr(versionOrder(i))
        s = s & "config_version " & versionText & ":" & vbCrLf

        For j = 1 To imports.Count
            If StrComp(GetImportTagVersion(imports(j), tagName), versionText, vbTextCompare) = 0 Then
                s = s & "- " & CStr(imports(j)("DisplayName")) & vbCrLf
            End If
        Next j
    Next i

    BuildTagVersionBreakdown = s
End Function

Private Function BuildTagVersionMismatchSummary(ByVal mismatchItems As Collection, ByVal mismatchDetails As Object) As String
    Dim s As String
    Dim i As Long

    s = "Files with non-matching tag config_version values:" & vbCrLf

    For i = 1 To mismatchItems.Count
        Dim displayName As String
        displayName = CStr(mismatchItems(i)("DisplayName"))
        s = s & "- " & displayName
        If Not mismatchDetails Is Nothing Then
            If mismatchDetails.Exists(displayName) Then
                s = s & " -> " & CStr(mismatchDetails(displayName))
            End If
        End If
        s = s & vbCrLf
    Next i

    BuildTagVersionMismatchSummary = s
End Function

Private Function GetImportTagVersion(ByVal importItem As Object, ByVal tagName As String) As String
    Dim tagVersions As Object
    Set tagVersions = importItem("TagVersions")

    If tagVersions.Exists(tagName) Then
        GetImportTagVersion = CStr(tagVersions(tagName))
    Else
        GetImportTagVersion = "(missing)"
    End If
End Function

Private Function ImportContainsTag(ByVal importItem As Object, ByVal tagName As String) As Boolean
    Dim tagVersions As Object
    Set tagVersions = importItem("TagVersions")

    ImportContainsTag = tagVersions.Exists(tagName)
End Function

Private Function BuildTagVersionMapFromRows(ByVal rowsData As Variant) As Object
    Dim tagVersions As Object
    Set tagVersions = CreateObject("Scripting.Dictionary")
    tagVersions.CompareMode = 1

    If IsEmpty(rowsData) Then
        Set BuildTagVersionMapFromRows = tagVersions
        Exit Function
    End If

    Dim i As Long
    For i = LBound(rowsData, 1) To UBound(rowsData, 1)
        Dim tagName As String
        Dim versionText As String

        tagName = Trim$(CStr(rowsData(i, 1)))
        If Len(tagName) = 0 Then GoTo NextRow

        versionText = NormalizeVersionText(rowsData(i, 2))
        tagVersions(tagName) = versionText
NextRow:
    Next i

    Set BuildTagVersionMapFromRows = tagVersions
End Function

Private Function NormalizeVersionText(ByVal rawValue As Variant) As String
    NormalizeVersionText = Trim$(CStr(rawValue))
    If Len(NormalizeVersionText) = 0 Then NormalizeVersionText = "0"
End Function

Private Function DetectManifestVersion(ByVal xmlPath As String) As String
    Dim rowsData As Variant
    rowsData = ParseManifestToRows(xmlPath)

    If IsEmpty(rowsData) Then
        DetectManifestVersion = "(unknown)"
        Exit Function
    End If

    DetectManifestVersion = MostCommonVersionInRows(rowsData)
End Function

Private Function MostCommonVersionInRows(ByVal rowsData As Variant) As String
    Dim counts As Object
    Set counts = CreateObject("Scripting.Dictionary")
    counts.CompareMode = 1

    Dim order As Collection
    Set order = New Collection

    Dim i As Long
    For i = LBound(rowsData, 1) To UBound(rowsData, 1)
        Dim versionText As String
        versionText = Trim$(CStr(rowsData(i, 2)))
        If Len(versionText) = 0 Then versionText = "(blank)"

        If Not counts.Exists(versionText) Then
            counts(versionText) = 0
            order.Add versionText
        End If
        counts(versionText) = CLng(counts(versionText)) + 1
    Next i

    Dim bestVersion As String
    Dim bestCount As Long
    bestCount = -1

    For i = 1 To order.Count
        Dim currentVersion As String
        Dim currentCount As Long
        currentVersion = CStr(order(i))
        currentCount = CLng(counts(currentVersion))
        If currentCount > bestCount Then
            bestCount = currentCount
            bestVersion = currentVersion
        End If
    Next i

    If Len(bestVersion) = 0 Then bestVersion = "(unknown)"
    MostCommonVersionInRows = bestVersion
End Function

Private Function SourceDisplayName(ByVal sourcePath As String) As String
    SourceDisplayName = XmlFileName(sourcePath)
End Function

Private Sub HandleImportFailure(ByVal wsBackup As Worksheet, ByVal createdNew As Boolean, ByVal message As String)
    Dim restored As Boolean
    restored = RestoreFromDiffBackup(wsBackup)

    Dim restoreNote As String
    If restored Then
        restoreNote = "Workbook restored from backup."
        If createdNew Then DeleteWorksheetQuietly wsBackup
    Else
        restoreNote = "Restore from backup failed."
    End If

    MsgBox message & vbCrLf & restoreNote, vbExclamation
End Sub

Private Function ResolveSelectedManifest(ByVal sourcePath As String, ByRef importTime As Date) As String
    If LCase$(Right$(sourcePath, 6)) = ".aapkg" Then
        ResolveSelectedManifest = ResolveManifestFromAapkg(sourcePath, importTime)
    Else
        importTime = Now
        ResolveSelectedManifest = sourcePath
    End If
End Function


' ========= Sheet/Header Utilities =========

Private Function EnsureTargetSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.name = sheetName
    End If
    Set EnsureTargetSheet = ws
End Function

Private Sub EnsureHeaders(ByVal ws As Worksheet)
    Dim c As Long
    For c = COL_TAG_IDX To COL_LASTIMP_IDX
        If Len(Trim$(CStr(ws.Cells(1, c).Value))) = 0 Then
            ws.Cells(1, c).Value = DefaultHeaderText(c)
        End If
    Next c
End Sub

Private Sub ResetSheetWithHeaders(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Cells(1, COL_TAG_IDX).Value = DefaultHeaderText(COL_TAG_IDX)
    ws.Cells(1, COL_CFG_IDX).Value = DefaultHeaderText(COL_CFG_IDX)
    ws.Cells(1, COL_PARENT_IDX).Value = DefaultHeaderText(COL_PARENT_IDX)
    ws.Cells(1, COL_DEPTH_IDX).Value = DefaultHeaderText(COL_DEPTH_IDX)
    ws.Cells(1, COL_LASTIMP_IDX).Value = DefaultHeaderText(COL_LASTIMP_IDX)
End Sub

Private Function DefaultHeaderText(ByVal colIndex As Long) As String
    Select Case colIndex
        Case COL_TAG_IDX
            DefaultHeaderText = "Tag"
        Case COL_CFG_IDX
            DefaultHeaderText = "Version"
        Case COL_PARENT_IDX
            DefaultHeaderText = "ParentTag"
        Case COL_DEPTH_IDX
            DefaultHeaderText = "Depth"
        Case COL_LASTIMP_IDX
            DefaultHeaderText = "LastImport"
        Case Else
            DefaultHeaderText = ""
    End Select
End Function

Private Function HeaderTextAt(ByVal ws As Worksheet, ByVal colIndex As Long) As String
    Dim v As String
    v = Trim$(CStr(ws.Cells(1, colIndex).Value))
    If Len(v) = 0 Then v = DefaultHeaderText(colIndex)
    HeaderTextAt = v
End Function

Private Function BuildDiffHeaders(ByVal wsPrimary As Worksheet, ByVal wsSecondary As Worksheet) As Variant
    Dim primaryLabel As String, secondaryLabel As String
    primaryLabel = SheetLabel(wsPrimary, PRIMARY_ABBREVIATED)
    secondaryLabel = SheetLabel(wsSecondary, SECONDARY_ABBREVIATED)

    BuildDiffHeaders = Array( _
        HeaderTextAt(wsPrimary, COL_TAG_IDX), _
        primaryLabel & "_" & HeaderTextAt(wsPrimary, COL_CFG_IDX), _
        primaryLabel & "_" & HeaderTextAt(wsPrimary, COL_LASTIMP_IDX), _
        secondaryLabel & "_" & HeaderTextAt(wsSecondary, COL_CFG_IDX), _
        secondaryLabel & "_" & HeaderTextAt(wsSecondary, COL_LASTIMP_IDX), _
        "Status", _
        "_Depth", _
        "_Parent", _
        "_SortKey" _
    )
End Function


Private Function SheetLabel(ByVal ws As Worksheet, ByVal abbrev As String) As String
    Dim labelText As String
    labelText = Trim$(abbrev)
    If Len(labelText) = 0 Then labelText = ws.name
    SheetLabel = labelText
End Function

Private Function PrimaryOnlyLabel() As String
    Dim labelText As String
    labelText = Trim$(PRIMARY_ABBREVIATED)
    If Len(labelText) = 0 Then
        PrimaryOnlyLabel = "PRIMARY ONLY"
    Else
        PrimaryOnlyLabel = UCase$(labelText) & " ONLY"
    End If
End Function

Private Function SecondaryOnlyLabel() As String
    Dim labelText As String
    labelText = Trim$(SECONDARY_ABBREVIATED)
    If Len(labelText) = 0 Then
        SecondaryOnlyLabel = "SECONDARY ONLY"
    Else
        SecondaryOnlyLabel = UCase$(labelText) & " ONLY"
    End If
End Function

' ========= XML Parsing =========

' Returns a 2D array of [Tag, Config_Version, ParentTag, Depth]
Private Function ParseManifestToRows(ByVal xmlPath As String) As Variant
    Dim dom As Object: Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    dom.async = False: dom.validateOnParse = False
    If Not dom.Load(xmlPath) Then
        MsgBox "Failed to load XML: " & xmlPath, vbCritical
        Exit Function
    End If

    If IsInstanceMode() Then
        ParseManifestToRows = ParseInstancesToRows(dom)
    Else
        ParseManifestToRows = ParseTemplatesToRows(dom)
    End If
End Function

Private Function ParseTemplatesToRows(ByVal dom As Object) As Variant
    Dim rows As Collection: Set rows = New Collection

    Dim node As Object
    For Each node In dom.DocumentElement.ChildNodes
        If node.NodeType = 1 And LCase$(node.nodeName) = "template" Then
            WalkTemplate node, rows, New Collection
        End If
    Next node

    If rows.Count = 0 Then Exit Function

    Dim arr() As Variant, i As Long, r As Variant
    ReDim arr(1 To rows.Count, 1 To 4)
    For i = 1 To rows.Count
        r = rows(i)
        arr(i, 1) = r(1) ' Tag
        arr(i, 2) = r(2) ' Config
        arr(i, 3) = r(3) ' Parent
        arr(i, 4) = r(4) ' Depth
    Next i

    ParseTemplatesToRows = arr
End Function

Private Function ParseInstancesToRows(ByVal dom As Object) As Variant
    Dim nodes As Object
    Set nodes = dom.getElementsByTagName("instance")
    If nodes Is Nothing Then Exit Function
    If nodes.Length = 0 Then Exit Function

    Dim rows As Collection: Set rows = New Collection
    Dim node As Object
    For Each node In nodes
        If node.NodeType = 1 Then
            Dim tagName As String, cfg As String, cfgVal As Variant
            Dim toolsetName As String, toolsetStore As String, depth As Long
            tagName = GetAttrSafe(node, "tag_name")
            cfg = GetAttrSafe(node, "config_version")
            toolsetName = GetAttrSafe(node, "toolset_name")
            toolsetStore = NormalizeToolsetForStorage(toolsetName)
            If IsNumeric(cfg) Then cfgVal = CLng(cfg) Else cfgVal = cfg
            depth = ToolsetDepth(toolsetName)

            Dim rowArr(1 To 4) As Variant
            rowArr(1) = tagName
            rowArr(2) = cfgVal
            rowArr(3) = toolsetStore
            rowArr(4) = depth
            rows.Add rowArr
        End If
    Next node

    If rows.Count = 0 Then Exit Function

    Dim arr() As Variant, i As Long, r As Variant
    ReDim arr(1 To rows.Count, 1 To 4)
    For i = 1 To rows.Count
        r = rows(i)
        arr(i, 1) = r(1) ' Tag
        arr(i, 2) = r(2) ' Config
        arr(i, 3) = r(3) ' Toolset (stored in ParentTag column)
        arr(i, 4) = r(4) ' Depth
    Next i

    ParseInstancesToRows = arr
End Function

Private Sub WalkTemplate(ByVal tplNode As Object, ByRef rows As Collection, ByVal parents As Collection)
    If Not (tplNode.NodeType = 1 And LCase$(tplNode.nodeName) = "template") Then Exit Sub

    Dim tagName As String, cfg As String, cfgVal As Variant
    tagName = GetAttrSafe(tplNode, "tag_name")
    cfg = GetAttrSafe(tplNode, "config_version")
    If IsNumeric(cfg) Then cfgVal = CLng(cfg) Else cfgVal = cfg

    Dim parentTag As String
    If parents.Count > 0 Then
        parentTag = parents(parents.Count)
    Else
        parentTag = ""
    End If
    Dim depth As Long: depth = parents.Count
    Dim rowArr(1 To 4) As Variant
    rowArr(1) = tagName
    rowArr(2) = cfgVal
    rowArr(3) = parentTag
    rowArr(4) = depth
    rows.Add rowArr

    ' Recurse into children
    parents.Add tagName
    Dim child As Object, grand As Object
    For Each child In tplNode.ChildNodes
        If child.NodeType = 1 Then
            Select Case LCase$(child.nodeName)
                Case "derived_templates"
                    For Each grand In child.ChildNodes
                        If grand.NodeType = 1 And LCase$(grand.nodeName) = "template" Then
                            WalkTemplate grand, rows, parents
                        End If
                    Next grand
                Case "template"
                    WalkTemplate child, rows, parents
            End Select
        End If
    Next child
    parents.Remove parents.Count
End Sub

Private Function GetAttrSafe(ByVal node As Object, ByVal attrName As String) As String
    On Error Resume Next
    Dim val As String: val = node.Attributes.getNamedItem(attrName).Text
    If IsEmpty(val) Or Len(val) = 0 Then val = ""
    GetAttrSafe = val
    On Error GoTo 0
End Function

' ========= Upsert & Dating =========

Private Sub UpsertRows(ByVal ws As Worksheet, ByVal arr As Variant)
    Dim cTag As Long, cCfg As Long, cPar As Long, cDep As Long
    cTag = COL_TAG_IDX
    cCfg = COL_CFG_IDX
    cPar = COL_PARENT_IDX
    cDep = COL_DEPTH_IDX

    Dim lastRow As Long: lastRow = LastUsedRow(ws)

    ' Index existing by Tag
    Dim idx As Object: Set idx = CreateObject("Scripting.Dictionary"): idx.CompareMode = 1
    Dim r As Long
    For r = 2 To IIf(lastRow < 2, 1, lastRow)
        If Len(ws.Cells(r, cTag).Value) > 0 Then idx(ws.Cells(r, cTag).Value) = r
    Next r

    ' Update or append
    Dim i As Long, tag As String, targetRow As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        tag = CStr(arr(i, 1))
        If Len(tag) = 0 Then GoTo NextI

        If idx.Exists(tag) Then
            targetRow = idx(tag)
        Else
            lastRow = LastUsedRow(ws) + 1
            targetRow = IIf(lastRow < 2, 2, lastRow)
            idx(tag) = targetRow
        End If

        ws.Cells(targetRow, cTag).Value = arr(i, 1)
        ws.Cells(targetRow, cCfg).Value = arr(i, 2)
        ws.Cells(targetRow, cPar).Value = arr(i, 3)
        ws.Cells(targetRow, cDep).Value = arr(i, 4)
        ' indent tag for visual hierarchy (max 15)
        ws.Cells(targetRow, cTag).IndentLevel = WorksheetFunction.Min(15, CLng(arr(i, 4)))
NextI:
    Next i

    ' Any rows not touched will still have their prior values
End Sub

Private Sub StampLastImport(ByVal ws As Worksheet, ByVal arr As Variant, Optional ByVal importTime As Variant)
    Dim cTag As Long, cLast As Long
    cTag = COL_TAG_IDX
    cLast = COL_LASTIMP_IDX

    ' Index by Tag
    Dim idx As Object: Set idx = CreateObject("Scripting.Dictionary"): idx.CompareMode = 1
    Dim lastRow As Long: lastRow = LastUsedRow(ws)
    Dim r As Long
    For r = 2 To IIf(lastRow < 2, 1, lastRow)
        If Len(ws.Cells(r, cTag).Value) > 0 Then idx(ws.Cells(r, cTag).Value) = r
    Next r

    ' For each imported tag, stamp time
    Dim stamp As Date
    If IsDate(importTime) Then
        stamp = CDate(importTime)
    Else
        stamp = Now
    End If
    Dim i As Long, tag As String
    For i = LBound(arr, 1) To UBound(arr, 1)
        tag = CStr(arr(i, 1))
        If Len(tag) = 0 Then GoTo NextI
        If idx.Exists(tag) Then
            ws.Cells(idx(tag), cLast).Value = stamp
        End If
NextI:
    Next i

    ' Optional: format column
    ws.Columns(cLast).NumberFormat = "yyyy-mm-dd hh:mm:ss"
End Sub

Private Sub SortByHierarchy(ByVal ws As Worksheet)
    Dim cPar As Long, cTag As Long
    cPar = COL_PARENT_IDX
    cTag = COL_TAG_IDX

    Dim lr As Long, lc As Long
    lr = LastUsedRow(ws)
    lc = LastUsedCol(ws)
    If lr < 2 Then Exit Sub

    Dim sortCol As Long
    sortCol = lc + 1
    ws.Cells(1, sortCol).Value = "_SortKey"

    Dim instanceMode As Boolean
    instanceMode = IsInstanceMode()

    Dim r As Long, t As String
    If instanceMode Then
        For r = 2 To lr
            t = CStr(ws.Cells(r, cTag).Value)
            If Len(t) > 0 Then
                ws.Cells(r, sortCol).Value = BuildToolsetSortKey(ws.Cells(r, cPar).Value, t)
            End If
        Next r
    Else
        Dim parentMap As Object
        Set parentMap = CreateObject("Scripting.Dictionary"): parentMap.CompareMode = 1
        For r = 2 To lr
            t = CStr(ws.Cells(r, cTag).Value)
            If Len(t) > 0 Then parentMap(t) = CStr(ws.Cells(r, cPar).Value)
        Next r

        Dim sortCache As Object
        Dim childMap As Object
        Set childMap = BuildChildPresenceMap(parentMap)
        Set sortCache = CreateObject("Scripting.Dictionary"): sortCache.CompareMode = 1
        For r = 2 To lr
            t = CStr(ws.Cells(r, cTag).Value)
            If Len(t) > 0 Then
                ws.Cells(r, sortCol).Value = BuildSortKey(t, parentMap, childMap, sortCache)
            End If
        Next r
    End If

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.Cells(1, sortCol).Resize(lr), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range(ws.Cells(1, 1), ws.Cells(lr, sortCol))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With

    ws.Columns(sortCol).Delete
End Sub

' ========= Diff Helpers =========

Private Function ReadSheetToDict(ByVal ws As Worksheet) As Object
    Dim cTag As Long, cCfg As Long, cLast As Long, cDepth As Long, cParent As Long
    cTag = COL_TAG_IDX
    cCfg = COL_CFG_IDX
    cLast = COL_LASTIMP_IDX
    cDepth = COL_DEPTH_IDX
    cParent = COL_PARENT_IDX

    Dim lr As Long: lr = LastUsedRow(ws)
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary"): dict.CompareMode = 1

    Dim r As Long, t As String
    For r = 2 To lr
        t = CStr(ws.Cells(r, cTag).Value)
        If Len(t) > 0 Then
            dict(t) = Array(ws.Cells(r, cCfg).Value, ws.Cells(r, cLast).Value, ws.Cells(r, cDepth).Value, ws.Cells(r, cParent).Value)
        End If
    Next r
    Set ReadSheetToDict = dict
End Function

Private Function CreateHiddenFullBackup(ByVal wsPrimary As Worksheet, ByVal wsSecondary As Worksheet, ByVal wsDiff As Worksheet, _
                                        ByRef createdNew As Boolean, Optional ByRef skipMessage As String) As Worksheet
    createdNew = False
    skipMessage = ""
    If wsPrimary Is Nothing Or wsSecondary Is Nothing Or wsDiff Is Nothing Then Exit Function

    Dim checksum As String
    checksum = ComputeBackupChecksum(wsPrimary, wsSecondary, wsDiff)

    Dim existing As Worksheet
    Set existing = FindBackupByChecksum(checksum)
    If Not existing Is Nothing Then
        Dim ts As String
        ts = BackupTimestampText(existing)
        If Len(ts) > 0 Then
            skipMessage = "Backup not created: state is identical to backup created on " & ts & "."
        Else
            skipMessage = "Backup not created: state is identical to an existing backup."
        End If
        Set CreateHiddenFullBackup = existing
        Exit Function
    End If

    Dim wsB As Worksheet
    Set wsB = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsB.name = GetUniqueDiffBackupName(wsDiff)

    wsB.Range(BACKUP_META_CELL).Value = BACKUP_META_LABEL
    wsB.Range(BACKUP_TIMESTAMP_CELL).Value = Now
    wsB.Range(BACKUP_CHECKSUM_CELL).Value = checksum

    wsB.Range(BACKUP_LABEL_PRIMARY_CELL).Value = BACKUP_LABEL_PRIMARY_TEXT
    wsB.Range(BACKUP_LABEL_SECONDARY_CELL).Value = BACKUP_LABEL_SECONDARY_TEXT
    wsB.Range(BACKUP_LABEL_DIFF_CELL).Value = BACKUP_LABEL_DIFF_TEXT

    Dim primaryRows As Long, secondaryRows As Long, diffRows As Long
    primaryRows = LastUsedRow(wsPrimary)
    secondaryRows = LastUsedRow(wsSecondary)
    diffRows = LastUsedRow(wsDiff)

    wsB.Range(BACKUP_PRIMARY_ROWS_CELL).Value = primaryRows
    wsB.Range(BACKUP_SECONDARY_ROWS_CELL).Value = secondaryRows
    wsB.Range(BACKUP_DIFF_ROWS_CELL).Value = diffRows

    Dim startRow As Long
    startRow = BACKUP_SECTION_START_ROW
    CopySheetRange wsPrimary, primaryRows, COL_LASTIMP_IDX, wsB, startRow
    startRow = startRow + primaryRows + BACKUP_SECTION_GAP_ROWS
    CopySheetRange wsSecondary, secondaryRows, COL_LASTIMP_IDX, wsB, startRow
    startRow = startRow + secondaryRows + BACKUP_SECTION_GAP_ROWS
    CopySheetRange wsDiff, diffRows, DIFF_COL_SORTKEY, wsB, startRow

    wsB.Visible = xlSheetVeryHidden
    createdNew = True
    Set CreateHiddenFullBackup = wsB
End Function

Private Sub CopySheetRange(ByVal wsSource As Worksheet, ByVal rowCount As Long, ByVal colCount As Long, _
                           ByVal wsTarget As Worksheet, ByVal targetRow As Long)
    If rowCount < 1 Or colCount < 1 Then Exit Sub

    Dim src As Range
    Dim dst As Range
    Set src = wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(rowCount, colCount))
    Set dst = wsTarget.Cells(targetRow, 1)

    src.Copy
    dst.PasteSpecial xlPasteValues
    dst.PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
End Sub

Private Function ComputeBackupChecksum(ByVal wsPrimary As Worksheet, ByVal wsSecondary As Worksheet, ByVal wsDiff As Worksheet) As String
    Dim h As Double
    h = 2166136261#

    h = HashRange(wsPrimary, LastUsedRow(wsPrimary), COL_LASTIMP_IDX, h)
    h = HashRange(wsSecondary, LastUsedRow(wsSecondary), COL_LASTIMP_IDX, h)
    h = HashRange(wsDiff, LastUsedRow(wsDiff), DIFF_COL_SORTKEY, h)

    ComputeBackupChecksum = ChecksumToHex(h)
End Function

Private Function HashRange(ByVal ws As Worksheet, ByVal rowCount As Long, ByVal colCount As Long, ByVal seed As Double) As Double
    Dim h As Double
    h = seed
    If rowCount < 1 Or colCount < 1 Then
        HashRange = h
        Exit Function
    End If

    Dim arr As Variant
    arr = ws.Range(ws.Cells(1, 1), ws.Cells(rowCount, colCount)).Value

    Dim r As Long, c As Long
    For r = 1 To rowCount
        For c = 1 To colCount
            h = HashString(h, CStr(arr(r, c)))
            h = HashNumber(h, r)
            h = HashNumber(h, c)
        Next c
    Next r

    HashRange = h
End Function

Private Function HashString(ByVal seed As Double, ByVal text As String) As Double
    Dim h As Double
    h = seed
    Dim i As Long
    For i = 1 To Len(text)
        h = HashNumber(h, AscW(Mid$(text, i, 1)))
    Next i
    HashString = h
End Function

Private Function HashNumber(ByVal seed As Double, ByVal value As Long) As Double
    Dim h As Double
    h = seed * 131# + value
    h = h - 2147483647# * Int(h / 2147483647#)
    HashNumber = h
End Function

Private Function ChecksumToHex(ByVal hashValue As Double) As String
    Dim v As Long
    v = CLng(hashValue)
    If v < 0 Then v = v And &H7FFFFFFF
    ChecksumToHex = Hex$(v)
End Function

Private Function FindBackupByChecksum(ByVal checksum As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If Left$(ws.name, Len(DIFF_BACKUP_PREFIX)) = DIFF_BACKUP_PREFIX Then
            If IsFullBackupSheet(ws) Then
                If CStr(ws.Range(BACKUP_CHECKSUM_CELL).Value) = checksum Then
                    Set FindBackupByChecksum = ws
                    Exit Function
                End If
            End If
        End If
    Next ws
    Set FindBackupByChecksum = Nothing
End Function

Private Function BackupTimestampText(ByVal wsBackup As Worksheet) As String
    If wsBackup Is Nothing Then Exit Function
    Dim v As Variant
    v = wsBackup.Range(BACKUP_TIMESTAMP_CELL).Value
    If IsDate(v) Then
        BackupTimestampText = Format$(CDate(v), "yyyy-mm-dd hh:nn:ss")
    Else
        BackupTimestampText = ""
    End If
End Function

Private Function IsFullBackupSheet(ByVal wsBackup As Worksheet) As Boolean
    If wsBackup Is Nothing Then Exit Function
    IsFullBackupSheet = (CStr(wsBackup.Range(BACKUP_META_CELL).Value) = BACKUP_META_LABEL)
End Function

Private Sub GetFullBackupLayout(ByVal wsBackup As Worksheet, ByRef primaryStart As Long, ByRef primaryRows As Long, _
                                ByRef secondaryStart As Long, ByRef secondaryRows As Long, ByRef diffStart As Long, ByRef diffRows As Long)
    primaryRows = CLng(Val(wsBackup.Range(BACKUP_PRIMARY_ROWS_CELL).Value))
    secondaryRows = CLng(Val(wsBackup.Range(BACKUP_SECONDARY_ROWS_CELL).Value))
    diffRows = CLng(Val(wsBackup.Range(BACKUP_DIFF_ROWS_CELL).Value))

    primaryStart = BACKUP_SECTION_START_ROW
    secondaryStart = primaryStart + primaryRows + BACKUP_SECTION_GAP_ROWS
    diffStart = secondaryStart + secondaryRows + BACKUP_SECTION_GAP_ROWS
End Sub

Private Function RestoreDiffOnlyFromBackup(ByVal wsBackup As Worksheet) As Boolean
    On Error GoTo SafeExit
    If wsBackup Is Nothing Then Exit Function
    If Not IsFullBackupSheet(wsBackup) Then Exit Function

    Dim wsDiff As Worksheet
    Set wsDiff = GetWorksheetByName(SHEET_DIFF)
    If wsDiff Is Nothing Then
        Set wsDiff = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsDiff.name = SHEET_DIFF
    End If

    Dim primaryStart As Long, primaryRows As Long
    Dim secondaryStart As Long, secondaryRows As Long
    Dim diffStart As Long, diffRows As Long
    GetFullBackupLayout wsBackup, primaryStart, primaryRows, secondaryStart, secondaryRows, diffStart, diffRows
    If diffRows < 1 Then Exit Function

    wsDiff.Cells.Clear
    wsBackup.Range(wsBackup.Cells(diffStart, 1), wsBackup.Cells(diffStart + diffRows - 1, DIFF_COL_SORTKEY)).Copy
    wsDiff.Cells(1, 1).PasteSpecial xlPasteValues
    wsDiff.Cells(1, 1).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False

    HideDiffHelperColumns wsDiff
    FreezeHeaderRow wsDiff
    wsDiff.Columns.AutoFit

    RestoreDiffOnlyFromBackup = True
    Exit Function

SafeExit:
    If Err.Number <> 0 Then Err.Clear
End Function

Private Function RestoreFullBackup(ByVal wsBackup As Worksheet) As Boolean
    On Error GoTo SafeExit
    If wsBackup Is Nothing Then Exit Function
    If Not IsFullBackupSheet(wsBackup) Then Exit Function

    Dim wsPrimary As Worksheet, wsSecondary As Worksheet, wsDiff As Worksheet
    Set wsPrimary = EnsureTargetSheet(SHEET_PRIMARY)
    Set wsSecondary = EnsureTargetSheet(SHEET_SECONDARY)
    Set wsDiff = GetWorksheetByName(SHEET_DIFF)
    If wsDiff Is Nothing Then
        Set wsDiff = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsDiff.name = SHEET_DIFF
    End If

    Dim primaryStart As Long, primaryRows As Long
    Dim secondaryStart As Long, secondaryRows As Long
    Dim diffStart As Long, diffRows As Long
    GetFullBackupLayout wsBackup, primaryStart, primaryRows, secondaryStart, secondaryRows, diffStart, diffRows
    If primaryRows < 1 Or secondaryRows < 1 Or diffRows < 1 Then Exit Function

    wsPrimary.Cells.Clear
    wsSecondary.Cells.Clear
    wsDiff.Cells.Clear

    wsBackup.Range(wsBackup.Cells(primaryStart, 1), wsBackup.Cells(primaryStart + primaryRows - 1, COL_LASTIMP_IDX)).Copy
    wsPrimary.Cells(1, 1).PasteSpecial xlPasteValues
    wsPrimary.Cells(1, 1).PasteSpecial xlPasteFormats

    wsBackup.Range(wsBackup.Cells(secondaryStart, 1), wsBackup.Cells(secondaryStart + secondaryRows - 1, COL_LASTIMP_IDX)).Copy
    wsSecondary.Cells(1, 1).PasteSpecial xlPasteValues
    wsSecondary.Cells(1, 1).PasteSpecial xlPasteFormats

    wsBackup.Range(wsBackup.Cells(diffStart, 1), wsBackup.Cells(diffStart + diffRows - 1, DIFF_COL_SORTKEY)).Copy
    wsDiff.Cells(1, 1).PasteSpecial xlPasteValues
    wsDiff.Cells(1, 1).PasteSpecial xlPasteFormats

    Application.CutCopyMode = False

    HideDiffHelperColumns wsDiff
    FreezeHeaderRow wsDiff
    wsPrimary.Columns.AutoFit
    wsSecondary.Columns.AutoFit
    wsDiff.Columns.AutoFit

    RestoreFullBackup = True
    Exit Function

SafeExit:
    If Err.Number <> 0 Then Err.Clear
End Function

Private Sub EnforceBackupCap(ByVal maxCount As Long)
    Dim names As Collection
    Dim unused As String
    Set names = New Collection
    unused = ListBackupSheets(names)

    If names.Count <= maxCount Then Exit Sub

    Dim arr() As String
    ReDim arr(1 To names.Count)
    Dim i As Long
    For i = 1 To names.Count
        arr(i) = CStr(names(i))
    Next i

    Dim j As Long
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(j) < arr(i) Then
                Dim tmp As String
                tmp = arr(i)
                arr(i) = arr(j)
                arr(j) = tmp
            End If
        Next j
    Next i

    Dim deleteCount As Long
    deleteCount = UBound(arr) - maxCount
    For i = 1 To deleteCount
        Dim ws As Worksheet
        Set ws = GetWorksheetByName(arr(i))
        DeleteWorksheetQuietly ws
    Next i
End Sub

Private Function GetUniqueDiffBackupName(ByVal wsDiff As Worksheet) As String
    Dim baseName As String
    baseName = DIFF_BACKUP_PREFIX & Format(Now, "yyyy-mm-dd_hhmmss")
    If DiffHasManualEdits(wsDiff) Then
        baseName = baseName & DIFF_BACKUP_MANUAL_FLAG
    End If

    Dim candidate As String, i As Long
    For i = 0 To 99
        candidate = baseName
        If i > 0 Then candidate = Left$(baseName, 28) & "_" & CStr(i)
        If GetWorksheetByName(candidate) Is Nothing Then
            GetUniqueDiffBackupName = candidate
            Exit Function
        End If
    Next i

    GetUniqueDiffBackupName = Left$(baseName, 20) & "_" & Format(Now, "hhnnss")
End Function

Private Function DiffHasManualEdits(ByVal wsDiff As Worksheet) As Boolean
    On Error GoTo SafeExit
    If wsDiff Is Nothing Then Exit Function

    Dim lr As Long, r As Long
    lr = LastUsedRow(wsDiff)
    For r = 2 To lr
        If IsManualStamp(wsDiff.Cells(r, DIFF_COL_PRIMARY_LAST).Value) Or IsManualStamp(wsDiff.Cells(r, DIFF_COL_SECONDARY_LAST).Value) Then
            DiffHasManualEdits = True
            Exit Function
        End If
    Next r

SafeExit:
    If Err.Number <> 0 Then Err.Clear
End Function


Private Function LatestImportOnSheet(ByVal ws As Worksheet) As Date
    Dim cLast As Long: cLast = COL_LASTIMP_IDX
    Dim lr As Long: lr = LastUsedRow(ws)
    Dim r As Long, maxD As Date, v As Variant
    maxD = 0
    For r = 2 To lr
        v = ws.Cells(r, cLast).Value
        If IsDate(v) Then If v > maxD Then maxD = v
    Next r
    LatestImportOnSheet = maxD
End Function

Private Function DiffStatus(ByVal primaryCfg As Variant, ByVal secondaryCfg As Variant) As String
    If IsEmptyOrNull(primaryCfg) And IsEmptyOrNull(secondaryCfg) Then
        DiffStatus = STATUS_NA
    ElseIf IsEmptyOrNull(primaryCfg) And Not IsEmptyOrNull(secondaryCfg) Then
        DiffStatus = SecondaryOnlyLabel()
    ElseIf Not IsEmptyOrNull(primaryCfg) And IsEmptyOrNull(secondaryCfg) Then
        DiffStatus = PrimaryOnlyLabel()
    ElseIf CStr(primaryCfg) = CStr(secondaryCfg) Then
        DiffStatus = STATUS_MATCH
    Else
        DiffStatus = STATUS_DRIFT
    End If
End Function

Private Sub ColorRow(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal rgbColor As Long)
    ws.Range(ws.Cells(rowNum, DIFF_COL_TAG), ws.Cells(rowNum, DIFF_COL_STATUS)).Interior.Color = rgbColor
End Sub

Private Function IsEmptyOrNull(ByVal v As Variant) As Boolean
    If IsEmpty(v) Then
        IsEmptyOrNull = True
    ElseIf IsNull(v) Then
        IsEmptyOrNull = True
    Else
        IsEmptyOrNull = (Len(Trim$(CStr(v))) = 0)
    End If
End Function

' ========= Small Utilities =========

Private Function RestoreFromDiffBackup(ByVal wsBackup As Worksheet) As Boolean
    On Error GoTo SafeExit
    If wsBackup Is Nothing Then Exit Function
    If Not IsFullBackupSheet(wsBackup) Then Exit Function

    Dim prevEvents As Boolean
    Dim prevCalc As XlCalculation
    Dim prevScreen As Boolean
    prevEvents = Application.EnableEvents
    prevCalc = Application.Calculation
    prevScreen = Application.ScreenUpdating

    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    If Not RestoreFullBackup(wsBackup) Then GoTo SafeExit

    Application.EnableEvents = prevEvents
    Application.Calculation = prevCalc
    Application.ScreenUpdating = prevScreen

    RestoreFromDiffBackup = True
    Exit Function

SafeExit:
    Application.EnableEvents = prevEvents
    Application.Calculation = prevCalc
    Application.ScreenUpdating = prevScreen
    If Err.Number <> 0 Then Err.Clear
End Function

Private Sub DeleteWorksheetQuietly(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    Dim prevAlerts As Boolean
    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = prevAlerts
End Sub

Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    Dim lr As Long: lr = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    If lr < 1 Then lr = 1
    LastUsedRow = lr
End Function

Private Function LastUsedCol(ByVal ws As Worksheet) As Long
    LastUsedCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
End Function

Private Sub ApplyDiffRowFormatting(ByVal wsDiff As Worksheet, ByVal r As Long, _
                                   ByVal primaryCfg As Variant, ByVal primaryImp As Variant, _
                                   ByVal secondaryCfg As Variant, ByVal secondaryImp As Variant, _
                                   ByVal latestPrimary As Date, ByVal latestSecondary As Date, _
                                   ByVal primaryDepth As Variant, ByVal secondaryDepth As Variant)
    Dim status As String
    status = DiffStatus(primaryCfg, secondaryCfg)
    wsDiff.Cells(r, DIFF_COL_STATUS).Value = status

    Dim depthVal As Long
    depthVal = 0
    If Not IsEmptyOrNull(primaryDepth) Then depthVal = CLng(primaryDepth)
    If depthVal = 0 And Not IsEmptyOrNull(secondaryDepth) Then depthVal = CLng(secondaryDepth)
    wsDiff.Cells(r, DIFF_COL_TAG).IndentLevel = WorksheetFunction.Min(15, depthVal)

    Dim primaryStale As Boolean, secondaryStale As Boolean, anyStale As Boolean
    primaryStale = False
    secondaryStale = False
    Dim tolDays As Double
    tolDays = STALE_TOLERANCE_MINUTES / 1440#

    If Not IsManualStamp(primaryImp) Then
        If Not IsEmptyOrNull(primaryImp) And IsDate(primaryImp) And latestPrimary <> 0 Then
            primaryStale = (CDate(primaryImp) < (latestPrimary - tolDays))
        End If
    End If
    If Not IsManualStamp(secondaryImp) Then
        If Not IsEmptyOrNull(secondaryImp) And IsDate(secondaryImp) And latestSecondary <> 0 Then
            secondaryStale = (CDate(secondaryImp) < (latestSecondary - tolDays))
        End If
    End If
    anyStale = (primaryStale Or secondaryStale)

    ' Row coloring priority:
    ' 1) Mismatch/missing => red
    ' 2) Match but stale   => yellow
    ' 3) Match & up-to-date => green
    ' 4) N/A => light gray
    Select Case status
        Case STATUS_MATCH
            If anyStale Then
                ColorRow wsDiff, r, RGB(255, 235, 156) ' yellow
            Else
                ColorRow wsDiff, r, RGB(226, 239, 218) ' light green
            End If
        Case STATUS_NA
            ColorRow wsDiff, r, RGB(242, 242, 242) ' light gray
        Case Else
            ColorRow wsDiff, r, RGB(255, 199, 206) ' light red
    End Select
End Sub

Private Function IsManualStamp(ByVal v As Variant) As Boolean
    If IsEmpty(v) Or IsNull(v) Then
        IsManualStamp = False
    Else
        IsManualStamp = (LCase$(Trim$(CStr(v))) = LCase$(MANUAL_STAMP))
    End If
End Function

Private Function BuildSortKey(ByVal tag As String, ByVal parentMap As Object, ByVal childMap As Object, ByVal cache As Object) As String
    If cache.Exists(tag) Then
        BuildSortKey = cache(tag)
        Exit Function
    End If

    Dim path As String
    Dim cur As String
    Dim safety As Long
    cur = tag
    safety = 0

    Do While Len(cur) > 0 And safety < 50
        path = "/" & SortSegment(cur, childMap.Exists(cur)) & path
        If parentMap.Exists(cur) Then
            cur = CStr(parentMap(cur))
        Else
            cur = ""
        End If
        safety = safety + 1
    Loop

    If Len(path) = 0 Then path = "/" & SortSegment(tag, False)
    cache(tag) = path
    BuildSortKey = path
End Function

Private Function BuildToolsetSortKey(ByVal toolsetName As Variant, ByVal tag As Variant) As String
    Dim p As String
    p = NormalizeToolsetPathForSort(CStr(toolsetName))
    Dim key As String
    key = ""

    If Len(p) > 0 Then
        Dim parts() As String
        parts = Split(p, "/")

        Dim i As Long
        For i = LBound(parts) To UBound(parts)
            If Len(parts(i)) > 0 Then
                key = key & "/" & SortSegment(parts(i), True)
            End If
        Next i
    End If

    key = key & "/" & SortSegment(CStr(tag), False)
    BuildToolsetSortKey = key
End Function

Private Function BuildChildPresenceMap(ByVal parentMap As Object) As Object
    Dim childMap As Object
    Set childMap = CreateObject("Scripting.Dictionary")
    childMap.CompareMode = 1

    Dim k As Variant
    For Each k In parentMap.Keys
        Dim p As String
        p = CStr(parentMap(k))
        If Len(p) > 0 Then
            childMap(p) = True
        End If
    Next k

    Set BuildChildPresenceMap = childMap
End Function

Private Function SortSegment(ByVal name As String, ByVal isBranch As Boolean) As String
    If isBranch Then
        SortSegment = "1|" & name
    Else
        SortSegment = "0|" & name
    End If
End Function

Private Function NormalizeToolsetPath(ByVal toolsetName As String) As String
    Dim p As String
    p = Trim$(toolsetName)
    If Len(p) = 0 Then
        NormalizeToolsetPath = ""
        Exit Function
    End If

    Dim parts() As String
    parts = Split(p, "$")
    If UBound(parts) >= 0 Then
        Dim head As String
        head = Trim$(parts(0))
        If LCase$(head) <> "industrial graphic library" And _
           LCase$(head) <> "situational awareness library" And _
           LCase$(head) <> "_default content" Then
            If InStr(head, " ") = 0 Then
                Dim i As Long, rebuilt As String
                For i = 1 To UBound(parts)
                    If Len(rebuilt) = 0 Then
                        rebuilt = parts(i)
                    Else
                        rebuilt = rebuilt & "$" & parts(i)
                    End If
                Next i
                If Len(rebuilt) > 0 Then p = rebuilt
            End If
        End If
    End If

    p = Replace(p, "$", "/")
    Do While InStr(p, "//") > 0
        p = Replace(p, "//", "/")
    Loop
    If Left$(p, 1) = "/" Then p = Mid$(p, 2)
    If Right$(p, 1) = "/" Then p = Left$(p, Len(p) - 1)
    NormalizeToolsetPath = p
End Function

Private Function NormalizeToolsetPathForSort(ByVal toolsetName As String) As String
    Dim p As String
    p = Trim$(toolsetName)
    If Len(p) = 0 Then
        NormalizeToolsetPathForSort = ""
        Exit Function
    End If

    p = Replace(p, "$", "/")
    Do While InStr(p, "//") > 0
        p = Replace(p, "//", "/")
    Loop
    If Left$(p, 1) = "/" Then p = Mid$(p, 2)
    If Right$(p, 1) = "/" Then p = Left$(p, Len(p) - 1)
    NormalizeToolsetPathForSort = p
End Function

Private Function NormalizeToolsetForStorage(ByVal toolsetName As String) As String
    Dim p As String
    p = NormalizeToolsetPath(toolsetName)
    If Len(p) = 0 Then
        NormalizeToolsetForStorage = ""
    Else
        NormalizeToolsetForStorage = Replace(p, "/", "$")
    End If
End Function

Private Function ToolsetDepth(ByVal toolsetName As String) As Long
    Dim p As String
    p = NormalizeToolsetPath(toolsetName)
    If Len(p) = 0 Then
        ToolsetDepth = 0
        Exit Function
    End If
    ToolsetDepth = UBound(Split(p, "/")) + 1
End Function

Private Function GetManifestMode() As String
    Dim ws As Worksheet
    Set ws = GetWorksheetByName(SHEET_MODE)
    If Not ws Is Nothing Then
        Dim v As String
        v = UCase$(Trim$(CStr(ws.Range("A1").Value)))
        If v = MODE_INSTANCE Or v = MODE_TEMPLATE Then
            GetManifestMode = v
            Exit Function
        End If
        If InStr(1, v, MODE_GRAPHIC, vbTextCompare) > 0 Then
            GetManifestMode = MODE_INSTANCE
            Exit Function
        End If
    End If

    Dim nm As String
    nm = LCase$(ThisWorkbook.name)
    If InStr(1, nm, "instance", vbTextCompare) > 0 Or InStr(1, nm, "graphic", vbTextCompare) > 0 Then
        GetManifestMode = MODE_INSTANCE
    Else
        GetManifestMode = MODE_TEMPLATE
    End If
End Function

Private Function IsInstanceMode() As Boolean
    IsInstanceMode = (GetManifestMode() = MODE_INSTANCE)
End Function

Private Function GetTargetSheetFromActive() As String
    Dim ws As Worksheet
    Set ws = ActiveSheet

    If ws Is Nothing Then
        GetTargetSheetFromActive = ""
        Exit Function
    End If

    If ws.name = SHEET_PRIMARY Or ws.name = SHEET_SECONDARY Then
        GetTargetSheetFromActive = ws.name
    Else
        MsgBox "Please run the Import XML button from the " & SHEET_PRIMARY & " or " & SHEET_SECONDARY & " sheet.", vbExclamation
        GetTargetSheetFromActive = ""
    End If
End Function

Private Function XmlFileName(ByVal fullPath As String) As String
    Dim p As Long
    p = InStrRev(fullPath, "\")
    If p > 0 Then
        XmlFileName = Mid$(fullPath, p + 1)
    Else
        XmlFileName = fullPath
    End If
End Function

Private Sub FreezeHeaderRow(ByVal ws As Worksheet)
    On Error GoTo SafeExit
    If ws Is Nothing Then Exit Sub
    If Application.Windows.Count = 0 Then Exit Sub

    Application.GoTo ws.Range("A2"), True
    ActiveWindow.FreezePanes = True
SafeExit:
    On Error GoTo 0
End Sub

Private Sub HideDiffHelperColumns(ByVal wsDiff As Worksheet)
    On Error Resume Next
    If wsDiff Is Nothing Then Exit Sub
    wsDiff.Columns(DIFF_COL_DEPTH).Resize(, DIFF_COL_SORTKEY - DIFF_COL_DEPTH + 1).Hidden = True
    On Error GoTo 0
End Sub

Private Function ListVeryHiddenSheets(ByRef names As Collection) As String
    Dim ws As Worksheet, s As String, i As Long
    i = 1
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVeryHidden Then
            names.Add ws.name
            s = s & i & ") " & ws.name & vbCrLf
            i = i + 1
        End If
    Next ws
    ListVeryHiddenSheets = s
End Function

Private Sub DeleteVeryHiddenSheetsByIndex(ByVal names As Collection, ByVal indicesCsv As String)
    Dim indices As Collection
    Set indices = ParseIndexList(indicesCsv, names.Count)
    If indices.Count = 0 Then
        MsgBox "No valid sheet numbers provided.", vbExclamation
        Exit Sub
    End If

    Dim i As Long, idx As Long, ws As Worksheet
    Dim deleted As Long
    Dim prevAlerts As Boolean

    prevAlerts = Application.DisplayAlerts
    On Error GoTo CleanFail
    Application.DisplayAlerts = False

    For i = 1 To indices.Count
        idx = CLng(indices(i))
        If idx > 0 And idx <= names.Count Then
            Set ws = GetWorksheetByName(CStr(names(idx)))
            If Not ws Is Nothing Then
                If ws.Visible = xlSheetVeryHidden Then
                    ws.Visible = xlSheetVisible
                    ws.Delete
                    deleted = deleted + 1
                End If
            End If
        End If
    Next i

CleanExit:
    Application.DisplayAlerts = prevAlerts
    MsgBox "Deleted: " & deleted, vbInformation
    Exit Sub

CleanFail:
    Application.DisplayAlerts = prevAlerts
    MsgBox "Delete failed: " & Err.Description, vbExclamation
End Sub

Private Function ParseIndexList(ByVal indicesCsv As String, ByVal maxIndex As Long) As Collection
    Dim result As New Collection
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1

    Dim parts() As String
    parts = Split(indicesCsv, ",")

    Dim i As Long, token As String
    For i = LBound(parts) To UBound(parts)
        token = Trim$(parts(i))
        If Len(token) = 0 Then GoTo NextPart

        Dim dashPos As Long
        dashPos = InStr(1, token, "-")
        If dashPos > 0 Then
            Dim leftPart As String, rightPart As String
            leftPart = Trim$(Left$(token, dashPos - 1))
            rightPart = Trim$(Mid$(token, dashPos + 1))

            Dim startIdx As Long, endIdx As Long, tmp As Long
            startIdx = val(leftPart)
            endIdx = val(rightPart)
            If startIdx > 0 And endIdx > 0 Then
                If startIdx > endIdx Then
                    tmp = startIdx
                    startIdx = endIdx
                    endIdx = tmp
                End If
                Dim j As Long
                For j = startIdx To endIdx
                    If j >= 1 And j <= maxIndex Then
                        If Not seen.Exists(CStr(j)) Then
                            seen(CStr(j)) = True
                            result.Add j
                        End If
                    End If
                Next j
            End If
        Else
            Dim idx As Long
            idx = val(token)
            If idx >= 1 And idx <= maxIndex Then
                If Not seen.Exists(CStr(idx)) Then
                    seen(CStr(idx)) = True
                    result.Add idx
                End If
            End If
        End If
NextPart:
    Next i

    Set ParseIndexList = result
End Function


Private Function ListBackupSheets(ByRef names As Collection) As String
    Dim ws As Worksheet, s As String, i As Long
    i = 1
    For Each ws In ThisWorkbook.Worksheets
        If Left$(ws.name, Len(DIFF_BACKUP_PREFIX)) = DIFF_BACKUP_PREFIX Then
            If IsFullBackupSheet(ws) Then
                names.Add ws.name
                s = s & i & ") " & ws.name & vbCrLf
                i = i + 1
            End If
        End If
    Next ws
    ListBackupSheets = s
End Function

Private Function GetWorksheetByName(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.name, sheetName, vbTextCompare) = 0 Then
            Set GetWorksheetByName = ws
            Exit Function
        End If
    Next ws
    Set GetWorksheetByName = Nothing
End Function

Public Sub EnsureEventsOn()
    If Application.EnableEvents = False Then Application.EnableEvents = True
End Sub

Private Function EnsureManifestModeSheet() As Worksheet
    Dim ws As Worksheet
    Set ws = GetWorksheetByName(SHEET_MODE)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.name = SHEET_MODE
    End If
    Set EnsureManifestModeSheet = ws
End Function

Private Function ResolveManifestFromAapkg(ByVal aapkgPath As String, ByRef importTime As Date) As String
    On Error GoTo SafeExit

    importTime = FileModifiedDate(aapkgPath)
    DebugWrite "AAPKG: " & aapkgPath
    DebugWrite "Modified: " & CStr(importTime)

    Dim tempRoot7 As String
    tempRoot7 = Environ$("TEMP") & "\aapkg_extract_" & Format(Now, "yyyymmdd_hhnnss")
    If Dir(tempRoot7, vbDirectory) = "" Then MkDir tempRoot7

    Dim cabPath7 As String
    Dim manifest7 As String

    Dim sevenZip As String
    sevenZip = Find7Zip()
    DebugWrite "7-Zip path: " & sevenZip
    If Len(sevenZip) = 0 Then
        MsgBox "7-Zip is required to import .aapkg files. Please install 7-Zip and try again.", vbExclamation
        Exit Function
    End If

    cabPath7 = ExtractCabWith7Zip(sevenZip, aapkgPath, tempRoot7)
    DebugWrite "CAB from 7-Zip: " & cabPath7
    If Len(cabPath7) > 0 Then
        manifest7 = ExtractManifestFromCabWith7Zip(sevenZip, cabPath7, tempRoot7 & "\cab")
        DebugWrite "Manifest from 7-Zip: " & manifest7
        If Len(manifest7) > 0 Then
            ResolveManifestFromAapkg = manifest7
            Exit Function
        End If
    End If
    DebugWrite "7-Zip flow failed"

SafeExit:
    If Err.Number <> 0 Then Err.Clear
End Function


Private Function Find7Zip() As String
    Dim paths As Variant
    paths = Array( _
        Environ$("ProgramFiles") & "\7-Zip\7z.exe", _
        Environ$("ProgramFiles(x86)") & "\7-Zip\7z.exe", _
        "C:\Program Files\7-Zip\7z.exe", _
        "C:\Program Files (x86)\7-Zip\7z.exe" _
    )

    Dim i As Long
    For i = LBound(paths) To UBound(paths)
        If Len(Dir(paths(i))) > 0 Then
            Find7Zip = paths(i)
            Exit Function
        End If
    Next i
End Function

Private Function ExtractCabWith7Zip(ByVal sevenZip As String, ByVal aapkgPath As String, ByVal tempRoot As String) As String
    On Error GoTo SafeExit
    Dim cabOut As String
    cabOut = tempRoot & "\cab"
    If Dir(cabOut, vbDirectory) = "" Then MkDir cabOut

    Dim cmd As String
    cmd = """" & sevenZip & """ e -y -o""" & cabOut & """ """ & aapkgPath & """ *.cab"
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    Dim exitCode As Long
    exitCode = wsh.Run(cmd, 0, True)
    DebugWrite "7z cab exit: " & CStr(exitCode)

    Dim cabFile As String
    cabFile = Dir(cabOut & "\*.cab")
    If Len(cabFile) > 0 Then ExtractCabWith7Zip = cabOut & "\" & cabFile
SafeExit:
    If Err.Number <> 0 Then Err.Clear
End Function

Private Function ExtractManifestFromCabWith7Zip(ByVal sevenZip As String, ByVal cabPath As String, ByVal outFolder As String) As String
    On Error GoTo SafeExit
    If Dir(outFolder, vbDirectory) = "" Then MkDir outFolder

    Dim cmd As String
    cmd = """" & sevenZip & """ e -y -o""" & outFolder & """ """ & cabPath & """ " & MANIFEST_XML_NAME
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    Dim exitCode As Long
    exitCode = wsh.Run(cmd, 0, True)
    DebugWrite "7z manifest exit: " & CStr(exitCode)

    Dim manifestPath As String
    manifestPath = outFolder & "\" & MANIFEST_XML_NAME
    If Len(Dir(manifestPath)) > 0 Then ExtractManifestFromCabWith7Zip = manifestPath
SafeExit:
    If Err.Number <> 0 Then Err.Clear
End Function

Private Sub DebugWrite(ByVal message As String)
    If Not DEBUG_AAPKG Then Exit Sub

    Dim ws As Worksheet
    Set ws = GetWorksheetByName(SHEET_DEBUG)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.name = SHEET_DEBUG
        ws.Range("A1:B1").Value = Array("Timestamp", "Message")
    End If

    Dim r As Long
    r = ws.Cells(ws.rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = message
End Sub

Private Function FileModifiedDate(ByVal path As String) As Date
    On Error GoTo SafeExit
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileModifiedDate = fso.GetFile(path).DateLastModified
SafeExit:
    If Err.Number <> 0 Then
        FileModifiedDate = Now
        Err.Clear
    End If
End Function





