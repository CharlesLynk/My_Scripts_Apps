Attribute VB_Name = "func4_PLCCore"
Option Explicit

'==========================================
' PLC DATA MANAGEMENT CORE MODULE
' Purpose: PLC filtering, summarization, and standardization
' Dependencies: func4_Utilities
'
' MULTI-USER PROTECTION:
' ----------------------
' All destructive operations (cleanup, standardization) now detect:
' 1. Multiple active users (Excel co-authoring via OneDrive/SharePoint)
' 2. Traditional shared workbook mode
' 3. Display critical warning if other users detected
' 4. Require explicit user confirmation to proceed
'
' Health check is READ-ONLY and safe to run anytime:
' - Scans for blank rows, inconsistent naming, character issues
' - Reports active user count and safe-to-clean status
' - No modifications made during health check
'
'==========================================

'================== CONFIG ==================
Private Const SUMMARY_SHEET          As String = "PLC-X Summary"
Private Const PLC_DROPDOWN_ADDR      As String = "C5"
Private Const SUMMARY_OUT_START      As String = "H3"
Private Const SUMMARY_HELPER_RANGE   As String = "CZ10:CZ1000"  ' Safe in Summary sheet
Private Const SCRATCH_SHEET_NAME     As String = "__TEMP_SCRATCH__"
Private Const HEADER_TEXT            As String = "SCADA PLC"
Private Const MAX_HEADER_SCAN_ROWS   As Long = 10
Private Const MAX_HEADER_SCAN_COLS   As Long = 20
Private Const TABLE_STYLE            As String = "TableStyleMedium15"
Private Const INTELLIGENT_FORMAT_SAMPLE_SIZE As Long = 50  ' Samples for column formatting

' Module-level scratch sheet cache (singleton pattern)
Private m_ScratchSheet As Worksheet

' ---- Column selection modes ----
Private Const CSM_CONTIGUOUS As Long = 1
Private Const CSM_NONCONTIGUOUS As Long = 2
Private Const COLUMN_SELECTION_MODE As Long = CSM_CONTIGUOUS
Private Const CONTIG_START           As Variant = "C"
Private Const CONTIG_END             As Variant = "P"

' Filter matching mode
Private Const USE_NORMALIZED_MATCHING As Boolean = True  ' True = normalize keys (remove spaces/hyphens), False = exact match
Private Const ENABLE_FILTER_DEBUG     As Boolean = True  ' True = output filter debug info to Immediate window

'================== ARRAY FUNCTIONS ==================
Private Function NONCONTIG_COLUMNS() As Variant
    NONCONTIG_COLUMNS = Array("Site", "Pkg", "Location Code", "Area", "Equip. Code", "ISA Code", _
                              "New /Existing", "PCD Link")
End Function

Private Function BLANK_ROW_CHECK_COLUMNS() As Variant
    BLANK_ROW_CHECK_COLUMNS = Array("SCADA PLC", "Description", "Site", "Location Code", _
                                    "Equip. Code", "ISA Code", "Loop")
End Function

'================== COLUMN CONFIG VALIDATION ==================

Private Function ValidateColumnConfig() As Boolean
    ' Validates column selection configuration at setup time
    If COLUMN_SELECTION_MODE = CSM_CONTIGUOUS Then
        ' Verify contiguous column range is valid
        If ColRefToNum(CONTIG_START) = 0 Or ColRefToNum(CONTIG_END) = 0 Then
            MsgBox "Invalid contiguous column configuration!" & vbCrLf & _
                   "CONTIG_START: " & CONTIG_START & ", CONTIG_END: " & CONTIG_END, vbCritical
            ValidateColumnConfig = False
            Exit Function
        End If
        If ColRefToNum(CONTIG_START) > ColRefToNum(CONTIG_END) Then
            MsgBox "Invalid contiguous column range: START must be <= END!", vbCritical
            ValidateColumnConfig = False
            Exit Function
        End If
    ElseIf COLUMN_SELECTION_MODE = CSM_NONCONTIGUOUS Then
        ' Verify non-contiguous columns array has items
        Dim arr As Variant: arr = NONCONTIG_COLUMNS()
        If UBound(arr) < LBound(arr) Then
            MsgBox "Non-contiguous columns array is empty!", vbCritical
            ValidateColumnConfig = False
            Exit Function
        End If
    Else
        MsgBox "Invalid COLUMN_SELECTION_MODE value: " & COLUMN_SELECTION_MODE & vbCrLf & _
               "Must be CSM_CONTIGUOUS (1) or CSM_NONCONTIGUOUS (2)", vbCritical
        ValidateColumnConfig = False
        Exit Function
    End If
    
    ValidateColumnConfig = True
End Function

'================== PUBLIC ENTRY POINTS ==================

' Run once to create dropdown and populate list
Public Sub PLC_FirstTimeSetup()
    Dim state As CalcState: state = PushAppState()
    On Error GoTo CleanFail

    ' Validate column configuration first
    If Not ValidateColumnConfig() Then
        GoTo CleanExit
    End If

    Dim wsSum As Worksheet, dd As Range
    Set wsSum = SheetByName(SUMMARY_SHEET)
    If wsSum Is Nothing Then
        MsgBox "Cannot find '" & SUMMARY_SHEET & "' sheet!", vbExclamation
        GoTo CleanExit
    End If

    Set dd = wsSum.Range(PLC_DROPDOWN_ADDR)
    On Error Resume Next
    ThisWorkbook.Names("selPLC").Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:="selPLC", RefersTo:=dd

    PLC_RebuildPLCList
    ApplyDropdownValidation dd

    MsgBox "Setup complete. Choose a PLC in " & SUMMARY_SHEET & "!" & PLC_DROPDOWN_ADDR & ".", vbInformation
CleanExit:
    PopAppState state
    Exit Sub
CleanFail:
    MsgBox "Error in PLC_FirstTimeSetup: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

Public Sub PLC_RebuildPLCList()
    Dim state As CalcState: state = PushAppState()
    On Error GoTo CleanFail

    Dim wsSum As Worksheet: Set wsSum = SheetByName(SUMMARY_SHEET)
    If wsSum Is Nothing Then GoTo CleanExit
    
    Dim arr() As String: arr = CollectUniquePLCs()
    Dim n As Long: n = ArrayCount(arr)

    ' Clear the helper range first
    wsSum.Range(SUMMARY_HELPER_RANGE).ClearContents
    
    ' Write the PLC list
    If n > 0 And Len(arr(1)) > 0 Then
        ' Write as 2D array for better compatibility
        Dim i As Long
        For i = 1 To n
            wsSum.Cells(i, wsSum.Range("CZ1").Column).value = arr(i)
        Next i
    End If

    ' Apply validation
    ApplyDropdownValidation wsSum.Range(PLC_DROPDOWN_ADDR)
    
CleanExit:
    PopAppState state
    Exit Sub
CleanFail:
    MsgBox "Error in PLC_RebuildPLCList: " & Err.Description & vbCrLf & _
           "Error Number: " & Err.Number, vbExclamation
    Resume CleanExit
End Sub

Public Sub PLC_SyncSelection(Optional ByVal runHealthCheckAfter As Boolean = False)
    ' ENSURES DATA FRESHNESS: This function rebuilds both the summary table AND the PLC dropdown list
    ' Called by all cleanup operations to guarantee UI reflects current data state
    ' runHealthCheckAfter: If True, prompts user for interactive health check after sync
    Dim state As CalcState: state = PushAppState()
    On Error GoTo CleanFail

    ' Clear all filters first
    PLC_ClearAllFilters
    
    ' Proceed with sync
    Dim sel As String: sel = GetSelectedPLC()
    
    BuildSummary sel
    UpdateAllPrintAreas
    
    ' Rebuild PLC list (BuildSummary clears the helper range area)
    PLC_RebuildPLCList
    
    ' If requested, prompt user for health check
    If runHealthCheckAfter Then
        Dim msg As String
        Dim userResp As VbMsgBoxResult
        
        msg = "Sync complete!" & vbCrLf & vbCrLf
        msg = msg & "Would you like to run a health check now?" & vbCrLf & vbCrLf
        msg = msg & "Health check will scan for:" & vbCrLf
        msg = msg & "  - Remaining blank rows" & vbCrLf
        msg = msg & "  - Inconsistent PLC naming" & vbCrLf
        msg = msg & "  - Character issues (Alt+0160, etc.)" & vbCrLf & vbCrLf
        msg = msg & "Run health check?"
        
        userResp = MsgBox(msg, vbYesNo + vbQuestion, "Run Health Check?")
        
        If userResp = vbYes Then
            ' Run the interactive health check
            PLC_CheckHealth
        End If
    End If
    
CleanExit:
    KillScratchSheetIfExists  ' Final cleanup at end of public operation
    PopAppState state
    Exit Sub
CleanFail:
    MsgBox "Error in PLC_SyncSelection: " & Err.Description, vbExclamation
    KillScratchSheetIfExists  ' Cleanup on error
    Resume CleanExit
End Sub

' MASTER CLEANUP: All operations (blank rows, PLC names, text characters, sync)
' This is the most comprehensive cleanup - runs all destructive operations
Public Sub PLC_FullCleanup()
    Dim state As CalcState: state = PushAppState()
    On Error GoTo CleanFail

    ' Safety gate: check for multiple users
    If Not ConfirmDestructiveOperation("full cleanup (blank rows, PLC names, text standardization)") Then
        GoTo CleanExit
    End If

    ' Step 1: Remove blank rows from data tables
    CleanBlankRowsAcrossSheets
    
    ' Step 2: Standardize PLC naming (prompts user for each change)
    StandardizePLCNames
    
    ' Step 3: Standardize all text characters (Alt+0160, fullwidth, etc.)
    ' NOTE: This calls StandardizeAllWorkbookText from func4_Utilities
    StandardizeAllWorkbookText
    
    ' Step 4: Refresh dropdown and summary
    ' NOTE: PLC_SyncSelection internally rebuilds PLC list after BuildSummary,
    '       so no need to call PLC_RebuildPLCList separately (avoids duplicate work)
    PLC_SyncSelection runHealthCheckAfter:=True  ' Rebuilds summary AND PLC list, prompts for health check
    
CleanExit:
    PopAppState state
    Exit Sub
CleanFail:
    MsgBox "Error in PLC_FullCleanup: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

' Standalone cleanup (no standardization)
Public Sub PLC_CleanupAndSync()
    Dim state As CalcState: state = PushAppState()
    On Error GoTo CleanFail

    ' Safety gate: check for multiple users
    If Not ConfirmDestructiveOperation("cleanup") Then
        GoTo CleanExit
    End If

    CleanBlankRowsAcrossSheets
    ' NOTE: PLC_SyncSelection rebuilds PLC list internally, no separate call needed
    PLC_SyncSelection runHealthCheckAfter:=True  ' Rebuilds summary AND PLC list, prompts for health check
    
CleanExit:
    PopAppState state
    Exit Sub
CleanFail:
    MsgBox "Error in PLC_CleanupAndSync: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

' Standalone standardization (prompts user for each change)
Public Sub PLC_StandardizeNames()
    Dim state As CalcState: state = PushAppState()
    On Error GoTo CleanFail

    ' Safety gate: check for multiple users
    If Not ConfirmDestructiveOperation("standardize PLC names") Then
        GoTo CleanExit
    End If

    StandardizePLCNames
    ' NOTE: PLC_SyncSelection rebuilds PLC list internally, no separate call needed
    PLC_SyncSelection runHealthCheckAfter:=True  ' Rebuilds summary AND PLC list to show standardized names, prompts for health check
    
CleanExit:
    PopAppState state
    Exit Sub
CleanFail:
    MsgBox "Error in PLC_StandardizeNames: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

'================== SIMPLE FILTER OPERATIONS ==================

' Quick filter all sheets by selected PLC (no cleanup, no validation)
Public Sub PLC_QuickFilter()
    Dim state As CalcState: state = PushAppState()
    On Error GoTo CleanFail

    Dim sel As String: sel = GetSelectedPLC()
    
    If Len(sel) = 0 Then
        MsgBox "Please select a PLC first.", vbExclamation
        GoTo CleanExit
    End If
    
    FilterAllSheets sel
    
CleanExit:
    PopAppState state
    Exit Sub
CleanFail:
    MsgBox "Error in PLC_QuickFilter: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

' Remove all filters from all sheets
Public Sub PLC_ClearAllFilters()
    Dim state As CalcState: state = PushAppState()
    On Error GoTo CleanFail

    ' Use centralized clearing with row unhiding
    ClearAllFilters unhideRows:=True
    
    ' Rebuild print areas for full data
    UpdateAllPrintAreas
    
CleanExit:
    PopAppState state
    Exit Sub
CleanFail:
    MsgBox "Error in PLC_ClearAllFilters: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

'================== HEALTH CHECK FUNCTION ==================

' Interactive health check - scans workbook and provides recommendations
Public Sub PLC_CheckHealth()
    Dim state As CalcState: state = PushAppState()
    On Error GoTo CleanFail

    Dim hasBlankRows As Boolean
    Dim hasInconsistentNames As Boolean
    Dim characterIssueCount As Long
    Dim userCount As Long
    Dim isShared As Boolean
    Dim isSafeToClean As Boolean
    Dim msg As String
    Dim scanDepth As Double
    
    ' Prompt user for scan depth (for StandardizeAllWorkbookText thoroughness only)
    ' Health checks are comprehensive by design (scan all target columns)
    scanDepth = PromptHealthCheckDepth()
    
    If scanDepth = 0 Then
        ' User cancelled
        GoTo CleanExit
    End If
    
    ' Run silent detection (all read-only operations)
    ' Note: DetectBlankRows and DetectInconsistentPLCNames always run comprehensive
    ' (they only check specific columns, so scanning all rows is safe)
    hasBlankRows = DetectBlankRows(1#)
    hasInconsistentNames = DetectInconsistentPLCNames(1#)
    characterIssueCount = QuickScanCharacterIssues(scanDepth)
    
    ' Check multi-user status
    userCount = GetActiveUserCount()
    isShared = IsWorkbookShared()
    isSafeToClean = IsSafeToModifyWorkbook()
    
    ' Build comprehensive health report
    Dim hasAnyIssues As Boolean
    hasAnyIssues = hasBlankRows Or hasInconsistentNames Or (characterIssueCount > 0)
    
    ' Build scan depth description (only applies to character issue scanning)
    Dim scanDesc As String
    If scanDepth >= 0.99 Then
        scanDesc = "Comprehensive (100%)"
    ElseIf scanDepth >= 0.45 And scanDepth <= 0.55 Then
        scanDesc = "Medium (50%)"
    Else
        scanDesc = "Quick (" & Format$(scanDepth * 100, "0") & "%)"
    End If
    
    If Not hasAnyIssues Then
        ' All good!
        msg = "[OK] Workbook Health Check Complete" & vbCrLf & vbCrLf
        msg = msg & "Character scan: " & scanDesc & vbCrLf
        msg = msg & "Data integrity: Comprehensive (all rows checked)" & vbCrLf
        msg = msg & "No issues detected. Your data is clean and ready to use." & vbCrLf & vbCrLf
        msg = msg & "Data quality: Excellent" & vbCrLf & vbCrLf
        
        ' Add user status even when healthy
        If userCount > 1 Or isShared Then
            msg = msg & "[!] Note: " & userCount & " user(s) currently editing"
            If isShared Then msg = msg & " (Shared mode)"
        Else
            msg = msg & "[OK] Safe to run cleanup operations"
        End If
        
        MsgBox msg, vbInformation, "Health Check - All Clear"
    Else
        ' Issues found
        msg = "Workbook Health Check Results" & vbCrLf & vbCrLf
        msg = msg & "Character scan: " & scanDesc & vbCrLf
        msg = msg & "Data integrity: Comprehensive (all rows checked)" & vbCrLf & vbCrLf
        msg = msg & "Issues detected that may affect data accuracy:" & vbCrLf & vbCrLf
        
        If hasBlankRows Then
            msg = msg & "[!] BLANK ROWS DETECTED" & vbCrLf
            msg = msg & "   Impact: Summary tables may be incomplete" & vbCrLf
            msg = msg & "   Recommended action: Run 'Cleanup And Sync'" & vbCrLf
            msg = msg & "   This will remove blank rows from all sheets." & vbCrLf & vbCrLf
        End If
        
        If hasInconsistentNames Then
            msg = msg & "[!] INCONSISTENT PLC NAMING" & vbCrLf
            msg = msg & "   Examples: 'PLC 1A' vs 'PLC-1A' or 'PLC1A'" & vbCrLf
            msg = msg & "   Impact: Same PLC may appear multiple times in dropdown" & vbCrLf
            msg = msg & "   Recommended action: Run 'Standardize PLC Names'" & vbCrLf
            msg = msg & "   You will be prompted to approve each change." & vbCrLf & vbCrLf
        End If
        
        If characterIssueCount > 0 Then
            msg = msg & "[!] CHARACTER ISSUES DETECTED" & vbCrLf
            msg = msg & "   Cells with problematic characters: " & characterIssueCount
            If scanDepth < 0.99 Then msg = msg & " (sampled)"
            msg = msg & vbCrLf
            msg = msg & "   Examples: Alt+0160 spaces, fullwidth dashes, Cyrillic letters" & vbCrLf
            msg = msg & "   Impact: Sorting and matching may behave unexpectedly" & vbCrLf
            msg = msg & "   Recommended action: Run 'Standardize All Workbook Text'" & vbCrLf & vbCrLf
        End If
        
        ' Add cleanup safety status
        msg = msg & String$(50, "-") & vbCrLf
        If isSafeToClean Then
            msg = msg & "[OK] SAFE TO RUN CLEANUP: YES" & vbCrLf
            msg = msg & "   Only you are editing this workbook."
        Else
            msg = msg & "[!] SAFE TO RUN CLEANUP: NO" & vbCrLf
            msg = msg & "   Active users: " & userCount
            If isShared Then msg = msg & " (Shared mode enabled)"
            msg = msg & vbCrLf
            msg = msg & "   Ask other users to close the file before running cleanup."
        End If
        
        If hasBlankRows And hasInconsistentNames Then
            msg = msg & vbCrLf & vbCrLf & "TIP: Use 'Cleanup, Standardize & Sync' to fix multiple issues at once."
        End If
        
        MsgBox msg, vbInformation, "Health Check - Issues Found"
    End If
    
CleanExit:
    PopAppState state
    Exit Sub
CleanFail:
    MsgBox "Error in PLC_CheckHealth: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

'================== SMART DETECTION FUNCTIONS ==================

' Scan to detect if any sheets have blank rows in data area
' Comprehensive by default (checks all rows in specific columns only)
' scanDepth: Reserved for future use, currently ignored (always scans 100%)
Private Function DetectBlankRows(Optional ByVal scanDepth As Double = 1#) As Boolean
    Dim ws As Worksheet, hdrRow As Long, plcCol As Long
    Dim colsToCheck As Variant: colsToCheck = BLANK_ROW_CHECK_COLUMNS()
    Dim dataStart As Long, lastRow As Long, r As Long
    Dim colMap As Object, checkCols() As Long, numCheckCols As Long
    Dim i As Long, isBlank As Boolean, c As Long
    
    DetectBlankRows = False
    
    For Each ws In ThisWorkbook.Worksheets
        If Not IsTargetSheet(ws) Then GoTo NextWsDetect
        
        If Not FindHeaderFlexible(ws, HEADER_TEXT, hdrRow, plcCol) Then GoTo NextWsDetect
        
        dataStart = DataStartRow(ws, hdrRow)
        lastRow = LastUsedRow(ws)
        If lastRow < dataStart Then GoTo NextWsDetect
        
        ' Build column map
        Set colMap = HeaderMapFromRow(ws, hdrRow)
        ReDim checkCols(1 To UBound(colsToCheck) - LBound(colsToCheck) + 1)
        numCheckCols = 0
        
        For i = LBound(colsToCheck) To UBound(colsToCheck)
            c = LookupHeader(colMap, CStr(colsToCheck(i)))
            If c > 0 Then
                numCheckCols = numCheckCols + 1
                checkCols(numCheckCols) = c
            End If
        Next i
        
        If numCheckCols = 0 Then GoTo NextWsDetect
        
        ' Comprehensive scan: check all rows (scanDepth parameter ignored)
        ' Safe because we only check specific columns, not entire workbook
        For r = dataStart To lastRow
            isBlank = True
            For i = 1 To numCheckCols
                If Not IsBlankAfterNormalize(ws.Cells(r, checkCols(i)).value2) Then
                    isBlank = False
                    Exit For
                End If
            Next i
            
            If isBlank Then
                DetectBlankRows = True
                Exit Function
            End If
        Next r
        
NextWsDetect:
    Next ws
End Function

' Detect if there are inconsistent PLC name variants (e.g., "PLC 1A" vs "PLC-1A")
' Comprehensive by default (checks all rows in PLC column only)
' scanDepth: Reserved for future use, currently ignored (always scans 100%)
Private Function DetectInconsistentPLCNames(Optional ByVal scanDepth As Double = 1#) As Boolean
    Dim dictOriginal As Object, dictNormalized As Object
    Dim ws As Worksheet, hdrRow As Long, plcCol As Long
    Dim dataStart As Long, lastRow As Long, r As Long
    Dim origVal As String, normVal As String, cleanOrigVal As String
    
    Set dictOriginal = CreateObject("Scripting.Dictionary")
    dictOriginal.CompareMode = vbTextCompare
    Set dictNormalized = CreateObject("Scripting.Dictionary")
    dictNormalized.CompareMode = vbTextCompare
    
    DetectInconsistentPLCNames = False
    
    For Each ws In ThisWorkbook.Worksheets
        If Not IsTargetSheet(ws) Then GoTo NextWsNorm
        
        If Not FindHeaderFlexible(ws, HEADER_TEXT, hdrRow, plcCol) Then GoTo NextWsNorm
        
        dataStart = DataStartRow(ws, hdrRow)
        lastRow = LastUsedRow(ws)
        If lastRow < dataStart Then GoTo NextWsNorm
        
        ' Comprehensive scan: check all rows (scanDepth parameter ignored)
        ' Safe because we only check PLC column, not entire workbook
        For r = dataStart To lastRow
            origVal = Trim$(CStr(ws.Cells(r, plcCol).value2))
            If Len(origVal) > 0 Then
                ' Clean the original value first (apply StandardizeSymbols)
                cleanOrigVal = StandardizeSymbols(origVal)
                normVal = NormalizePLCName(cleanOrigVal)
                
                If Len(normVal) > 0 Then
                    If Not dictNormalized.Exists(normVal) Then
                        dictNormalized(normVal) = cleanOrigVal
                    Else
                        ' Check if cleaned values differ (case-insensitive but symbol-sensitive)
                        If StrComp(dictNormalized(normVal), cleanOrigVal, vbBinaryCompare) <> 0 Then
                            DetectInconsistentPLCNames = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        Next r
        
NextWsNorm:
    Next ws
End Function

'================== PLC NAME STANDARDIZATION ==================

' Interactive batch rename with user approval
Private Sub StandardizePLCNames()
    Dim ws As Worksheet, hdrRow As Long, plcCol As Long
    Dim dictChanges As Object, dictApproved As Object
    Dim lastRow As Long, r As Long, origVal As String, newVal As String
    Dim userResp As VbMsgBoxResult
    Dim key As Variant
    Dim changeCount As Long
    
    Set dictChanges = CreateObject("Scripting.Dictionary")
    dictChanges.CompareMode = vbTextCompare
    Set dictApproved = CreateObject("Scripting.Dictionary")
    dictApproved.CompareMode = vbTextCompare
    
    ' Pass 1: collect unique proposed changes
    For Each ws In ThisWorkbook.Worksheets
        If Not IsTargetSheet(ws) Then GoTo NextWsCollect
        
        If Not FindHeaderFlexible(ws, HEADER_TEXT, hdrRow, plcCol) Then GoTo NextWsCollect
        
        lastRow = LastUsedRow(ws)
        
        ' Performance: read entire column at once to reduce COM calls
        Dim colValues As Variant
        If lastRow > hdrRow Then
            colValues = ws.Range(ws.Cells(hdrRow + 1, plcCol), ws.Cells(lastRow, plcCol)).value2
        Else
            GoTo NextWsCollect
        End If
        
        ' Process in-memory array
        For r = 1 To UBound(colValues, 1)
            On Error Resume Next
            origVal = Trim$(CStr(colValues(r, 1)))
            If Err.Number <> 0 Then
                Err.Clear
                origVal = ""
            End If
            On Error GoTo 0
            
            If Len(origVal) > 0 Then
                newVal = NormalizePLCName(origVal)
                If Len(newVal) > 0 And StrComp(origVal, newVal, vbBinaryCompare) <> 0 Then
                    If Not dictChanges.Exists(origVal) Then
                        dictChanges(origVal) = newVal
                    End If
                End If
            End If
        Next r
        
NextWsCollect:
    Next ws
    
    If dictChanges.Count = 0 Then
        MsgBox "No PLC names need standardization.", vbInformation
        Exit Sub
    End If
    
    ' Pass 2: prompt approvals (one prompt per unique change)
    For Each key In dictChanges.Keys
        newVal = dictChanges(key)
        userResp = MsgBox("Standardize '" & key & "' to '" & newVal & "'?", _
                         vbYesNoCancel + vbQuestion, "Confirm Change")
        
        If userResp = vbCancel Then
            Exit Sub
        ElseIf userResp = vbYes Then
            dictApproved(key) = newVal
        End If
    Next key
    
    If dictApproved.Count = 0 Then
        MsgBox "No changes approved.", vbInformation
        Exit Sub
    End If
    
    ' Pass 3: apply approved changes
    changeCount = 0
    For Each ws In ThisWorkbook.Worksheets
        If Not IsTargetSheet(ws) Then GoTo NextWsApply
        
        If Not FindHeaderFlexible(ws, HEADER_TEXT, hdrRow, plcCol) Then GoTo NextWsApply
        
        lastRow = LastUsedRow(ws)
        For r = hdrRow + 1 To lastRow
            origVal = Trim$(CStr(ws.Cells(r, plcCol).value2))
            If dictApproved.Exists(origVal) Then
                ws.Cells(r, plcCol).value = dictApproved(origVal)
                changeCount = changeCount + 1
            End If
        Next r
        
NextWsApply:
    Next ws
    
    MsgBox "Standardization complete. " & changeCount & " cells updated.", vbInformation
End Sub

' Normalize PLC name to canonical format "PLC-ALPHANUM"
Private Function NormalizePLCName(ByVal s As String) As String
    Dim t As String
    t = UCase$(s)
    t = CleanText(t)
    t = StandardizeSymbols(t)     ' CRITICAL: Standardize before cleaning whitespace
    t = CleanWhitespace(t)
    
    ' Remove leading "PLC" and common separators only at start
    t = RxReplace(t, "^\s*PLC[\s\-_]*", "", False)
    
    ' Strip everything that is not A-Z or 0-9
    t = RxReplace(t, "[^A-Z0-9]+", "", True)
    
    If LenB(t) > 0 Then
        NormalizePLCName = "PLC-" & t
    Else
        NormalizePLCName = ""
    End If
End Function

'================== CORE WORK - BUILD SUMMARY ==================
Private Sub BuildSummary(ByVal selectedPLC As String)
    Dim wsSum As Worksheet: Set wsSum = SheetByName(SUMMARY_SHEET)
    If wsSum Is Nothing Then Exit Sub
    
    Dim outTopLeft As Range: Set outTopLeft = wsSum.Range(SUMMARY_OUT_START)

    ' Clear previous output
    With outTopLeft.Resize(100000, 200)
        .ClearContents
        .ClearFormats
    End With
    On Error Resume Next
    wsSum.ListObjects("tblPLC_Summary").Delete
    On Error GoTo 0
    
    If Len(selectedPLC) = 0 Then Exit Sub

    Dim selectedKey As String: selectedKey = NormalizeKey(selectedPLC)
    Dim outRow As Long: outRow = outTopLeft.row
    Dim outColStart As Long: outColStart = outTopLeft.Column
    Dim headerWritten As Boolean: headerWritten = False

    Dim ws As Worksheet, gMissing As String
    Dim selectedCols() As Long, selectedColsCount As Long
    Dim hdrRow As Long, plcCol As Long, dataStart As Long, lastRow As Long
    Dim r As Long, c As Long, srcVal As Variant, key As String
    Dim outColIdx As Long
    Dim totalColsWithSource As Long

    For Each ws In ThisWorkbook.Worksheets
        If Not IsTargetSheet(ws) Then GoTo NextSheet
        
        If Not FindHeaderFlexible(ws, HEADER_TEXT, hdrRow, plcCol) Then GoTo NextSheet
        
        selectedColsCount = BuildSelectedColumns(ws, hdrRow, plcCol, selectedCols, gMissing)
        If selectedColsCount = 0 Then GoTo NextSheet
        
        dataStart = DataStartRow(ws, hdrRow)
        lastRow = LastUsedRow(ws)
        If lastRow < dataStart Then GoTo NextSheet
        
        ' Write header once (with Source Sheet as first column)
        If Not headerWritten Then
            wsSum.Cells(outRow, outColStart).value = "Source Sheet"
            For c = 1 To selectedColsCount
                wsSum.Cells(outRow, outColStart + c).value = _
                    ws.Cells(hdrRow, selectedCols(c)).value2
            Next c
            totalColsWithSource = selectedColsCount + 1
            outRow = outRow + 1
            headerWritten = True
        End If
        
        ' Copy matching rows (with source sheet name as first column)
        For r = dataStart To lastRow
            srcVal = ws.Cells(r, plcCol).value2
            key = NormalizeKey(srcVal)
            
            If key = selectedKey Then
                ' Write source sheet name in first column
                wsSum.Cells(outRow, outColStart).value = ws.Name
                
                ' Write data columns - preserve text formatting including leading zeros
                For c = 1 To selectedColsCount
                    ' Format destination cell as text FIRST to preserve leading zeros
                    wsSum.Cells(outRow, outColStart + c).NumberFormat = "@"
                    
                    ' Use .Text property to preserve displayed format (including leading zeros)
                    ' Fall back to .value if .Text is empty
                    Dim srcCell As Range
                    Set srcCell = ws.Cells(r, selectedCols(c))
                    
                    If Len(srcCell.Text) > 0 Then
                        wsSum.Cells(outRow, outColStart + c).value = srcCell.Text
                    Else
                        wsSum.Cells(outRow, outColStart + c).value = srcCell.value
                    End If
                Next c
                outRow = outRow + 1
            End If
        Next r
        
NextSheet:
    Next ws

    If Len(gMissing) > 0 Then
        MsgBox "These requested headers were not found on at least one sheet and were skipped:" & vbCrLf & gMissing, vbInformation
    End If

    ' Format as table and apply intelligent formatting
    If headerWritten And outRow > outTopLeft.row + 1 Then
        Dim tblRange As Range
        Set tblRange = wsSum.Range(outTopLeft, wsSum.Cells(outRow - 1, outColStart + totalColsWithSource - 1))
        
        On Error Resume Next
        wsSum.ListObjects.Add(xlSrcRange, tblRange, , xlYes).Name = "tblPLC_Summary"
        wsSum.ListObjects("tblPLC_Summary").TableStyle = TABLE_STYLE
        On Error GoTo 0
        
        ' Apply intelligent column formatting
        ApplyIntelligentFormatting wsSum, outTopLeft, outRow - 1, totalColsWithSource
        
        ' Suppress "Number Stored as Text" errors after data is in place
        SuppressNumberAsTextErrors tblRange
        
        ' Auto-fit columns to content
        tblRange.Columns.AutoFit
    End If
End Sub

'================== INTELLIGENT FORMATTING ==================
Private Sub ApplyIntelligentFormatting(ws As Worksheet, ByVal topLeft As Range, _
                                       ByVal lastDataRow As Long, ByVal numCols As Long)
    Dim col As Long, row As Long
    Dim totalLength As Long, sampleCount As Long, avgLength As Double
    Dim val As Variant, cellText As String
    Dim DataStartRow As Long, colRange As Range
    Dim sampleSize As Long, sampleInterval As Long
    
    DataStartRow = topLeft.row + 1  ' Skip header
    
    ' Determine intelligent sampling strategy
    Dim totalRows As Long
    totalRows = lastDataRow - DataStartRow + 1
    
    If totalRows <= INTELLIGENT_FORMAT_SAMPLE_SIZE Then
        ' Small dataset - check all rows
        sampleSize = totalRows
        sampleInterval = 1
    Else
        ' Large dataset - sample every Nth row
        sampleSize = INTELLIGENT_FORMAT_SAMPLE_SIZE
        sampleInterval = totalRows \ sampleSize
        If sampleInterval < 1 Then sampleInterval = 1
    End If
    
    ' Analyze and format each column
    For col = 0 To numCols - 1
        totalLength = 0
        sampleCount = 0
        
        ' Sample rows to calculate average text length
        For row = DataStartRow To lastDataRow Step sampleInterval
            val = ws.Cells(row, topLeft.Column + col).value2
            
            If Not IsEmpty(val) And Not IsBlankAfterNormalize(val) Then
                cellText = CStr(val)
                totalLength = totalLength + Len(cellText)
                sampleCount = sampleCount + 1
            End If
            
            ' Stop if we've collected enough samples
            If sampleCount >= sampleSize Then Exit For
        Next row
        
        ' Calculate average length
        If sampleCount > 0 Then
            avgLength = totalLength / sampleCount
        Else
            avgLength = 0
        End If
        
        ' Get the column range (data only, excluding header)
        Set colRange = ws.Range(ws.Cells(DataStartRow, topLeft.Column + col), _
                                ws.Cells(lastDataRow, topLeft.Column + col))
        
        ' Apply text formatting to all columns
        With colRange
            .NumberFormat = "@"  ' Text format for all columns
            .VerticalAlignment = xlCenter
            
            ' Center-align if average length <= 16 characters, otherwise left-align
            If avgLength > 0 And avgLength <= 16 Then
                .HorizontalAlignment = xlCenter
            Else
                .HorizontalAlignment = xlLeft
            End If
        End With
    Next col
    
    ' Format header row
    With ws.Range(topLeft, ws.Cells(topLeft.row, topLeft.Column + numCols - 1))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

'================== SUPPRESS ERROR INDICATORS ==================
Private Sub SuppressNumberAsTextErrors(ByVal targetRange As Range)
    ' Bulk suppress "Number Stored as Text" error indicators (green triangles)
    ' This is more reliable than setting .Ignore property during formatting
    
    On Error Resume Next
    
    ' Method 1: Set Ignore property on the entire range
    targetRange.Errors(xlNumberAsText).Ignore = True
    
    ' Method 2: If Method 1 doesn't work, iterate through cells with errors
    ' This is slower but more compatible across Excel versions
    If Err.Number <> 0 Or targetRange.Errors(xlNumberAsText).Count > 0 Then
        Err.Clear
        
        Dim cell As Range
        Dim processedCount As Long
        processedCount = 0
        
        ' Only process cells that actually have the error to improve performance
        For Each cell In targetRange.Cells
            ' Check if cell has the "Number Stored as Text" error
            If cell.Errors.Item(xlNumberAsText).value Then
                ' Acknowledge/ignore the error
                cell.Errors.Item(xlNumberAsText).Ignore = True
                processedCount = processedCount + 1
                
                ' Process in batches to avoid hanging on very large ranges
                If processedCount Mod 1000 = 0 Then
                    DoEvents  ' Allow Excel to remain responsive
                End If
            End If
        Next cell
    End If
    
    On Error GoTo 0
End Sub

'================== FILTER & PRINT AREA ==================

Private Sub FilterAllSheets(ByVal selectedPLC As String)
    Dim selectedKey As String: selectedKey = NormalizeKey(selectedPLC)
    Dim ws As Worksheet, hdrRow As Long, plcCol As Long
    Dim dataStart As Long, lastRow As Long, lastCol As Long
    Dim rng As Range
    Dim r As Long, cellVal As String, cellKey As String
    Dim matchingValues As Object  ' Dictionary to collect matching values
    Dim criteriaArr As Variant, criteriaCount As Long, i As Long

    For Each ws In ThisWorkbook.Worksheets
        If Not IsTargetSheet(ws) Then GoTo NextSheet
        
        If Len(selectedPLC) = 0 Then GoTo NextSheet
        
        If Not FindHeaderFlexible(ws, HEADER_TEXT, hdrRow, plcCol) Then GoTo NextSheet
        
        dataStart = DataStartRow(ws, hdrRow)
        lastRow = LastUsedRow(ws)
        lastCol = LastUsedCol(ws)
        
        If lastRow < dataStart Or lastCol < plcCol Then GoTo NextSheet
        
        ' Define filter range: from column 1 (A) to last used column, header row to last row
        ' This matches how print areas are defined
        Set rng = ws.Range(ws.Cells(hdrRow, 1), ws.Cells(lastRow, lastCol))
        
        ' Step 1: Ensure AutoFilter range is correctly set up
        On Error Resume Next
        If ws.AutoFilterMode Then
            ' Clear all filter criteria (show all data)
            If ws.FilterMode Then
                ws.ShowAllData
            End If
            
            ' Verify AutoFilter range matches our desired range
            Dim currentRange As Range
            Set currentRange = ws.AutoFilter.Range
            
            If Not RangesMatch(currentRange, rng) Then
                ' Range doesn't match - turn off and reapply with correct range
                ws.AutoFilterMode = False
                rng.AutoFilter
            End If
        Else
            ' AutoFilter not enabled - set it up
            rng.AutoFilter
        End If
        On Error GoTo 0
        
        ' Step 2: Collect all unique values that match the selected PLC (with normalization)
        Set matchingValues = CreateObject("Scripting.Dictionary")
        matchingValues.CompareMode = vbTextCompare
        
        For r = dataStart To lastRow
            cellVal = Trim$(CStr(ws.Cells(r, plcCol).value2))
            If Len(cellVal) > 0 Then
                cellKey = NormalizeKey(cellVal)
                If cellKey = selectedKey Then
                    ' Store the original value (not the normalized key) for filtering
                    matchingValues(cellVal) = True
                End If
            End If
        Next r
        
        ' Debug output
        If ENABLE_FILTER_DEBUG Then
            Debug.Print "=== Filter Debug: " & ws.Name & " ==="
            Debug.Print "  Header Row: " & hdrRow & ", PLC Col: " & plcCol
            Debug.Print "  Filter Range: " & rng.Address
            Debug.Print "  Selected PLC: " & selectedPLC & " -> Key: " & selectedKey
            Debug.Print "  Matching values found: " & matchingValues.Count
        End If
        
        ' Step 3: Apply AutoFilter
        ' Field number must be relative to the filter range (column 1 = first column of range)
        Dim fieldNum As Long
        fieldNum = plcCol  ' Since range starts at column 1, plcCol is already the correct field number
        
        If matchingValues.Count = 0 Then
            ' No matches - apply filter that hides all data rows (but shows header)
            On Error Resume Next
            rng.AutoFilter Field:=fieldNum, Criteria1:="<>#N/A", Operator:=xlAnd, Criteria2:="=#N/A"
            On Error GoTo 0
            
            If ENABLE_FILTER_DEBUG Then Debug.Print "  No matches - hiding all rows"
        Else
            ' Convert dictionary keys to array for filtering
            criteriaCount = matchingValues.Count
            ReDim criteriaArr(1 To criteriaCount)
            i = 1
            Dim key As Variant
            For Each key In matchingValues.Keys
                criteriaArr(i) = CStr(key)
                i = i + 1
            Next key
            
            ' Apply AutoFilter with the matching values
            On Error Resume Next
            If criteriaCount = 1 Then
                ' Single criterion
                rng.AutoFilter Field:=fieldNum, Criteria1:=criteriaArr(1)
            ElseIf criteriaCount <= 2 Then
                ' Two criteria (use Criteria1 and Criteria2 with OR)
                rng.AutoFilter Field:=fieldNum, Criteria1:=criteriaArr(1), _
                              Operator:=xlOr, Criteria2:=criteriaArr(2)
            Else
                ' Multiple criteria (use array)
                rng.AutoFilter Field:=fieldNum, Criteria1:=criteriaArr, Operator:=xlFilterValues
            End If
            
            If Err.Number <> 0 Then
                If ENABLE_FILTER_DEBUG Then Debug.Print "  Filter ERROR: " & Err.Description
                Err.Clear
            Else
                If ENABLE_FILTER_DEBUG Then Debug.Print "  Filter applied with " & criteriaCount & " criteria"
            End If
            On Error GoTo 0
        End If
        
NextSheet:
    Next ws
End Sub

Private Sub UpdateAllPrintAreas()
    Dim ws As Worksheet, hdrRow As Long, plcCol As Long
    Dim lastRow As Long, lastCol As Long, visRows As Long, r As Long

    For Each ws In ThisWorkbook.Worksheets
        If Not IsTargetSheet(ws) Then GoTo NextSheet
        
        If Not FindHeaderFlexible(ws, HEADER_TEXT, hdrRow, plcCol) Then GoTo NextSheet
        
        lastRow = LastUsedRow(ws)
        lastCol = LastUsedCol(ws)
        
        If lastRow < hdrRow Or lastCol < 1 Then
            ws.PageSetup.PrintArea = ""
            GoTo NextSheet
        End If
        
        ' Count visible rows
        visRows = 0
        For r = hdrRow To lastRow
            If Not ws.Rows(r).Hidden Then visRows = visRows + 1
        Next r
        
        If visRows > 1 Then
            ws.PageSetup.PrintArea = ws.Range(ws.Cells(hdrRow, 1), ws.Cells(lastRow, lastCol)).Address
        Else
            ws.PageSetup.PrintArea = ""
        End If
        
NextSheet:
    Next ws
End Sub

'================== FILTER MANAGEMENT HELPERS ==================

' Centralized filter clearing for a single sheet
Private Sub ClearSheetFilter(ws As Worksheet, Optional ByVal unhideRows As Boolean = False)
    Dim hdrRow As Long, plcCol As Long
    Dim lastRow As Long, lastCol As Long
    Dim rng As Range
    
    On Error Resume Next
    
    ' Clear filter criteria but preserve AutoFilter setup
    If ws.AutoFilterMode Then
        If ws.FilterMode Then
            ' Show all data (clears criteria but keeps AutoFilter dropdown arrows)
            ws.ShowAllData
        End If
    End If
    
    ' Ensure AutoFilter range is correctly set up (matching filter range definition)
    If FindHeaderFlexible(ws, HEADER_TEXT, hdrRow, plcCol) Then
        lastRow = LastUsedRow(ws)
        lastCol = LastUsedCol(ws)
        
        If lastRow >= hdrRow And lastCol >= 1 Then
            Set rng = ws.Range(ws.Cells(hdrRow, 1), ws.Cells(lastRow, lastCol))
            
            ' Set up or update AutoFilter range
            If ws.AutoFilterMode Then
                ' Check if current AutoFilter range matches our desired range
                Dim currentRange As Range
                Set currentRange = ws.AutoFilter.Range
                
                If Not RangesMatch(currentRange, rng) Then
                    ' Range doesn't match - turn off and reapply with correct range
                    ws.AutoFilterMode = False
                    rng.AutoFilter
                End If
            Else
                ' AutoFilter not enabled - set it up
                rng.AutoFilter
            End If
        End If
    End If
    
    ' Optionally unhide all rows (useful for cleanup operations)
    If unhideRows Then
        ws.Rows.Hidden = False
    End If
    
    On Error GoTo 0
End Sub

' Helper function to check if two ranges are identical
' Note: RangesMatch moved to func4_Utilities (general-purpose)

' Clear filters from all target sheets
Private Sub ClearAllFilters(Optional ByVal unhideRows As Boolean = False)
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        If IsTargetSheet(ws) Then
            ClearSheetFilter ws, unhideRows
        End If
    Next ws
End Sub
'================== BLANK ROW CLEANUP ==================
Private Sub CleanBlankRowsAcrossSheets()
    Dim ws As Worksheet, hdrRow As Long, plcCol As Long
    Dim colsToCheck As Variant: colsToCheck = BLANK_ROW_CHECK_COLUMNS()
    Dim dataStart As Long, lastRow As Long, r As Long
    Dim colMap As Object, checkCols() As Long, numCheckCols As Long
    Dim i As Long, isBlank As Boolean, c As Long
    Dim deletedCount As Long
    
    deletedCount = 0

    For Each ws In ThisWorkbook.Worksheets
        If Not IsTargetSheet(ws) Then GoTo NextWs
        
        ' Use centralized filter clearing
        ClearSheetFilter ws, unhideRows:=True
        
        If Not FindHeaderFlexible(ws, HEADER_TEXT, hdrRow, plcCol) Then GoTo NextWs
        
        dataStart = DataStartRow(ws, hdrRow)
        lastRow = LastUsedRow(ws)
        If lastRow < dataStart Then GoTo NextWs
        
        ' Build column map
        Set colMap = HeaderMapFromRow(ws, hdrRow)
        ReDim checkCols(1 To UBound(colsToCheck) - LBound(colsToCheck) + 1)
        numCheckCols = 0
        
        For i = LBound(colsToCheck) To UBound(colsToCheck)
            c = LookupHeader(colMap, CStr(colsToCheck(i)))
            If c > 0 Then
                numCheckCols = numCheckCols + 1
                checkCols(numCheckCols) = c
            End If
        Next i
        
        If numCheckCols = 0 Then GoTo NextWs
        
        ' Delete blank rows (iterate backwards)
        For r = lastRow To dataStart Step -1
            isBlank = True
            For i = 1 To numCheckCols
                If Not IsBlankAfterNormalize(ws.Cells(r, checkCols(i)).value2) Then
                    isBlank = False
                    Exit For
                End If
            Next i
            
            If isBlank Then
                ws.Rows(r).Delete
                deletedCount = deletedCount + 1
            End If
        Next r
        
NextWs:
    Next ws
    
    If deletedCount > 0 Then
        MsgBox "Removed " & deletedCount & " blank rows.", vbInformation
    Else
        MsgBox "No blank rows found.", vbInformation
    End If
End Sub

'================== HEADER & COLUMN HELPERS ==================

' Project-specific wrapper for FindHeaderFlexible (uses config constants)
Private Function FindHeaderFlexible(ws As Worksheet, ByVal headerText As String, _
                                    ByRef headerRow As Long, ByRef headerCol As Long) As Boolean
    ' Delegate to utility function with project-specific config
    FindHeaderFlexible = func4_Utilities.FindHeaderFlexible(ws, headerText, headerRow, headerCol, _
                                                             MAX_HEADER_SCAN_ROWS, MAX_HEADER_SCAN_COLS)
End Function

Private Function BuildSelectedColumns(ws As Worksheet, ByVal hdrRow As Long, ByVal plcCol As Long, _
                                      ByRef cols() As Long, ByRef missList As String) As Long
    Dim i As Long, s As Long, e As Long, idx As Long
    
    If COLUMN_SELECTION_MODE = CSM_CONTIGUOUS Then
        s = ColRefToNum(CONTIG_START)
        e = ColRefToNum(CONTIG_END)
        
        If s <= 0 Or e <= 0 Or s > e Then
            BuildSelectedColumns = 0
            Exit Function
        End If
        
        ReDim cols(1 To (e - s + 1) + 1)
        cols(1) = plcCol
        idx = 2
        For i = s To e
            If i <> plcCol Then
                cols(idx) = i
                idx = idx + 1
            End If
        Next i
        
        BuildSelectedColumns = idx - 1
        Exit Function
    End If

    ' Non-contiguous mode
    Dim map As Object: Set map = HeaderMapFromRow(ws, hdrRow)
    Dim wanted As Variant: wanted = NONCONTIG_COLUMNS()
    Dim tmp() As Long
    ReDim tmp(1 To 1 + UBound(wanted) - LBound(wanted))
    idx = 1
    tmp(idx) = plcCol: idx = idx + 1

    Dim k As Long, nm As String, colIdx As Long
    For k = LBound(wanted) To UBound(wanted)
        nm = CStr(wanted(k))
        colIdx = LookupHeader(map, nm)
        
        If colIdx > 0 And colIdx <> plcCol Then
            tmp(idx) = colIdx
            idx = idx + 1
        Else
            If Len(missList) > 0 Then missList = missList & ", "
            missList = missList & nm
        End If
    Next k

    If idx = 1 Then
        BuildSelectedColumns = 0
    Else
        ReDim cols(1 To idx - 1)
        For i = 1 To idx - 1
            cols(i) = tmp(i)
        Next i
        BuildSelectedColumns = UBound(cols)
    End If
End Function

' Note: HeaderMapFromRow, LookupHeader, DataStartRow moved to func4_Utilities (general-purpose)

'================== SCRATCH SHEET MANAGEMENT ==================
Private Function GetScratchSheet() As Worksheet
    ' Singleton pattern - create once and reuse
    If m_ScratchSheet Is Nothing Then
        On Error Resume Next
        Set m_ScratchSheet = ThisWorkbook.Worksheets(SCRATCH_SHEET_NAME)
        On Error GoTo 0
        
        If m_ScratchSheet Is Nothing Then
            Set m_ScratchSheet = ThisWorkbook.Worksheets.Add
            m_ScratchSheet.Name = SCRATCH_SHEET_NAME
            m_ScratchSheet.Visible = xlSheetVeryHidden
        End If
    End If
    
    Set GetScratchSheet = m_ScratchSheet
End Function

Private Sub ClearScratchContent()
    ' Clear scratch sheet content without destroying the sheet
    If Not m_ScratchSheet Is Nothing Then
        On Error Resume Next
        m_ScratchSheet.Cells.Clear
        On Error GoTo 0
    End If
End Sub

Private Sub KillScratchSheetIfExists()
    On Error Resume Next
    If Not m_ScratchSheet Is Nothing Then
        Application.DisplayAlerts = False
        m_ScratchSheet.Delete
        Application.DisplayAlerts = True
        Set m_ScratchSheet = Nothing  ' Reset module variable
    End If
    On Error GoTo 0
End Sub

' Call from Workbook_Open to clean up scratch sheet on startup
Public Sub AutoCleanScratchOnOpen()
    KillScratchSheetIfExists
End Sub

'================== DIAGNOSTIC HELPERS ==================

' Public helper function for testing normalization in worksheet
' Usage in Excel: =TestPLCMatch("PLC-1A", "PLC 1A") returns TRUE if they match
Private Function TestPLCMatch(ByVal value1 As Variant, ByVal value2 As Variant) As Boolean
    Dim key1 As String, key2 As String
    key1 = NormalizeKey(value1)
    key2 = NormalizeKey(value2)
    TestPLCMatch = (key1 = key2)
End Function

' Public helper to see normalized key in worksheet
' Usage in Excel: =GetNormalizedPLC("PLC-1A") returns "PLC1A"
Private Function GetNormalizedPLC(ByVal value As Variant) As String
    GetNormalizedPLC = NormalizeKey(value)
End Function

' Test filter criteria formula in immediate window
Private Sub TestFilterCriteria(ByVal sheetName As String, ByVal selectedPLC As String)
    Dim ws As Worksheet: Set ws = SheetByName(sheetName)
    If ws Is Nothing Then
        Debug.Print "Sheet not found: " & sheetName
        Exit Sub
    End If
    
    Dim hdrRow As Long, plcCol As Long, dataStart As Long
    If Not FindHeaderFlexible(ws, HEADER_TEXT, hdrRow, plcCol) Then
        Debug.Print "Header not found on sheet: " & sheetName
        Exit Sub
    End If
    
    dataStart = DataStartRow(ws, hdrRow)
    Dim selectedKey As String: selectedKey = NormalizeKey(selectedPLC)
    
    Debug.Print "=== Filter Criteria Test ==="
    Debug.Print "Sheet: " & sheetName
    Debug.Print "Header Row: " & hdrRow & ", PLC Column: " & plcCol
    Debug.Print "Data Start Row: " & dataStart
    Debug.Print "Selected PLC: " & selectedPLC
    Debug.Print "Normalized Key: " & selectedKey
    Debug.Print ""
    Debug.Print "Sample Data Rows:"
    
    Dim r As Long, cellVal As String, normVal As String
    For r = dataStart To Min2(dataStart + 5, LastUsedRow(ws))
        cellVal = CStr(ws.Cells(r, plcCol).value2)
        normVal = NormalizeKey(cellVal)
        Debug.Print "  Row " & r & ": '" & cellVal & "' -> '" & normVal & "' | Match: " & (normVal = selectedKey)
    Next r
End Sub

'================== HELPERS ==================

' Prompt user to choose health check depth with clear button labels
' Returns: percentage as decimal (0.25 = 25%, 1.0 = 100%), or 0 (cancelled)
Private Function PromptHealthCheckDepth() As Double
    Dim msg As String
    Dim userResp As VbMsgBoxResult
    Dim userInput As String
    Dim percentage As Double
    
    ' Step 1: Offer comprehensive scan with warning
    msg = "Run Comprehensive Health Check?" & vbCrLf & vbCrLf
    msg = msg & "This will scan 100% of data for:" & vbCrLf
    msg = msg & "  - Blank rows in data tables" & vbCrLf
    msg = msg & "  - Inconsistent PLC naming" & vbCrLf
    msg = msg & "  - Character issues (Alt+0160, etc.)" & vbCrLf & vbCrLf
    msg = msg & "[!] Note: May take longer on large workbooks" & vbCrLf & vbCrLf
    msg = msg & "YES = Run comprehensive scan (100%)" & vbCrLf
    msg = msg & "NO = Choose a custom scan percentage" & vbCrLf
    msg = msg & "CANCEL = Skip health check"
    
    userResp = MsgBox(msg, vbYesNoCancel + vbQuestion, "Health Check")
    
    Select Case userResp
        Case vbYes
            ' Run comprehensive scan
            PromptHealthCheckDepth = 1#
            Exit Function
            
        Case vbCancel
            ' User cancelled
            PromptHealthCheckDepth = 0
            Exit Function
            
        Case vbNo
            ' Prompt for custom percentage
            userInput = InputBox( _
                "Enter scan percentage (1-100):" & vbCrLf & vbCrLf & _
                "Examples:" & vbCrLf & _
                "  25 = Quick scan (1 in 4 rows)" & vbCrLf & _
                "  50 = Medium scan (half of rows)" & vbCrLf & _
                "  75 = Thorough scan" & vbCrLf & vbCrLf & _
                "Value will be rounded to nearest 5%.", _
                "Custom Scan Percentage", _
                "50")
            
            ' Handle cancellation or empty input
            If Len(Trim$(userInput)) = 0 Then
                PromptHealthCheckDepth = 0  ' User cancelled
                Exit Function
            End If
            
            ' Parse and validate input
            On Error Resume Next
            percentage = CDbl(userInput)
            On Error GoTo 0
            
            ' Bound to 0-100%
            If percentage < 0 Then percentage = 0
            If percentage > 100 Then percentage = 100
            
            ' Round to nearest 5%
            percentage = Round(percentage / 5) * 5
            
            ' Ensure minimum of 5% if user entered something > 0
            If percentage > 0 And percentage < 5 Then percentage = 5
            
            ' Convert to decimal (25% -> 0.25)
            PromptHealthCheckDepth = percentage / 100#
    End Select
End Function

' Confirm destructive operation with multi-user detection
Private Function ConfirmDestructiveOperation(ByVal operationName As String) As Boolean
    Dim userCount As Long: userCount = GetActiveUserCount()
    Dim isShared As Boolean: isShared = IsWorkbookShared()
    Dim msg As String
    Dim userConfirm As VbMsgBoxResult
    
    If userCount > 1 Or isShared Then
        ' Multiple users detected - show critical warning
        msg = "[!] WARNING: Multiple Users Detected!" & vbCrLf & vbCrLf
        msg = msg & "Active users: " & userCount & vbCrLf
        If isShared Then msg = msg & "Shared workbook mode: ENABLED" & vbCrLf
        msg = msg & vbCrLf & "This operation will " & operationName & "." & vbCrLf
        msg = msg & "Running this now may cause conflicts or data loss." & vbCrLf & vbCrLf
        msg = msg & "RECOMMENDED: Ask other users to close the file first." & vbCrLf & vbCrLf
        msg = msg & "Do you want to continue anyway?"
        
        userConfirm = MsgBox(msg, vbYesNo + vbCritical, "Multiple Users Detected")
    Else
        ' Solo user - standard confirmation
        msg = "This will " & operationName & "." & vbCrLf & _
              "Do you want to continue?"
        
        userConfirm = MsgBox(msg, vbYesNo + vbQuestion, "Confirm Operation")
    End If
    
    ConfirmDestructiveOperation = (userConfirm = vbYes)
End Function

Private Function IsTargetSheet(ws As Worksheet) As Boolean
    IsTargetSheet = (ws.Visible = xlSheetVisible) _
                    And (LCase$(Left$(ws.Name, 5)) <> "sheet") _
                    And (ws.Name <> SUMMARY_SHEET)
End Function

Private Function GetSelectedPLC() As String
    Dim ws As Worksheet: Set ws = SheetByName(SUMMARY_SHEET)
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    GetSelectedPLC = Trim$(CStr(ThisWorkbook.Names("selPLC").RefersToRange.value2))
    If Len(GetSelectedPLC) = 0 Then GetSelectedPLC = Trim$(CStr(ws.Range(PLC_DROPDOWN_ADDR).value2))
    On Error GoTo 0
End Function

Private Function CollectUniquePLCs() As String()
    ' Collects raw PLC values from sheets (NO cleaning applied here)
    ' This preserves user's original data in the dropdown list
    ' Matching is done via NormalizeKey() at filter time
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim ws As Worksheet, hdrRow As Long, plcCol As Long
    Dim dataStart As Long, lastRow As Long, r As Long, v As Variant, s As String
    
    For Each ws In ThisWorkbook.Worksheets
        If Not IsTargetSheet(ws) Then GoTo NextWsPLC
        
        If Not FindHeaderFlexible(ws, HEADER_TEXT, hdrRow, plcCol) Then GoTo NextWsPLC
        
        dataStart = DataStartRow(ws, hdrRow)
        lastRow = LastUsedRow(ws)
        
        For r = dataStart To lastRow
            v = ws.Cells(r, plcCol).value2
            
            ' Only basic trim - no cleaning of symbols
            ' Users who want clean data should run PLC_CleanupStandardizeAndSync()
            s = Trim$(CStr(v))
            If Len(s) > 0 And Not dict.Exists(s) Then
                dict(s) = True
            End If
        Next r
        
NextWsPLC:
    Next ws

    Dim arr() As String, i As Long, k As Variant
    If dict.Count = 0 Then
        ReDim arr(1 To 1)
        arr(1) = ""
    Else
        ReDim arr(1 To dict.Count)
        i = 1
        For Each k In dict.Keys
            arr(i) = CStr(k)
            i = i + 1
        Next k
        QuickSortStrings arr, LBound(arr), UBound(arr)
    End If

    CollectUniquePLCs = arr
End Function

Private Sub ApplyDropdownValidation(ByVal targetCell As Range)
    Dim wsSum As Worksheet: Set wsSum = SheetByName(SUMMARY_SHEET)
    If wsSum Is Nothing Then Exit Sub
    
    Dim firstCell As Range: Set firstCell = wsSum.Range("CZ1")
    Dim lastRow As Long
    Dim formulaString As String
    
    ' Find the last non-empty cell in the list
    On Error Resume Next
    If Not IsEmpty(firstCell.value2) Then
        lastRow = wsSum.Range("CZ:CZ").Find(What:="*", After:=firstCell, _
                                            LookIn:=xlValues, SearchDirection:=xlPrevious).row
    Else
        lastRow = firstCell.row
    End If
    If Err.Number <> 0 Then lastRow = firstCell.row
    On Error GoTo 0
    
    ' If no data, exit without applying validation
    If lastRow < firstCell.row Or IsEmpty(firstCell.value2) Then
        On Error Resume Next
        targetCell.Validation.Delete
        On Error GoTo 0
        Exit Sub
    End If
    
    ' Build validation formula (use range reference instead of string)
    Dim validationRange As Range
    Set validationRange = wsSum.Range("CZ1:CZ" & lastRow)
    
    On Error Resume Next
    With targetCell.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:="=" & validationRange.Address(True, True, xlA1, True)
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = True
        .ErrorMessage = "Select a PLC from the dropdown list."
    End With
    
    ' If validation failed, report it
    If Err.Number <> 0 Then
        Debug.Print "Validation Error: " & Err.Description & " | Range: " & validationRange.Address(True, True, xlA1, True)
        Err.Clear
    End If
    On Error GoTo 0
End Sub

