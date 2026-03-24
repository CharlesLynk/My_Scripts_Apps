Attribute VB_Name = "func3_MultiSheetConsolidation"
Option Explicit
' =====================================================================================
' Author: Chase Lynk
' =====================================================================================

' -----------------------------
' GLOBALS & CONSTANTS
' -----------------------------
Private gWB As Workbook                ' This workbook handle
Private gWSArch As Worksheet           ' Destination worksheet: SystemArchData
Private gTbl As ListObject             ' Destination table object

' Destination sheet name
Private Const ARCH_SHEET As String = "SystemArchData"

' Header names used repeatedly (keep names in one place)
Private Const H_SITE As String = "Site Name"
Private Const H_MUNI As String = "Municipality"
Private Const H_TYPE As String = "Site Type"
Private Const H_TIER As String = "SCADA Tier"
Private Const H_BUILD As String = "Building Name"
Private Const H_ROOM As String = "Room Name"
Private Const H_PANEL As String = "Panel Name"
Private Const H_COMM_MFG As String = "Comm Manufacturer Name"
Private Const H_COMM_MODEL As String = "Comm Model Number"
Private Const H_INST_TAG As String = "Instrument Tag"
Private Const H_INST_CAT As String = "Instrument Category"
Private Const H_INST_TYPE As String = "Instrument Type"
Private Const H_INST_MFG As String = "Instrument Manufacturer"
Private Const H_INST_MODEL As String = "Instrument Model"
Private Const H_INST_HEALTH As String = "Instrument Health"
Private Const H_INST_COND As String = "Instrument Condition Rating"
Private Const H_INST_NAME As String = "Field Instrument Name"

' Unified table headers (order defines the table)
Public gHeaders As Variant
' Source worksheets to ingest (in order)
Public gSrcSheets As Variant

' =====================================================================================
' ENTRY POINT
' =====================================================================================
Public Sub UnifyData()
    Dim prevCalc As XlCalculation
    On Error GoTo CleanFail
    
    ' --- Init globals, headers, and settings ---
    InitGlobals

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        prevCalc = .Calculation
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
    End With

    ' --- Create or clear destination sheet & table ---
    PrepareDestination

    ' --- (Optional) global data hygiene in sources ---
    'GeneralCleanup  ' uncomment if you want broad cleanup across the book

    ' --- Normalize special cases in sources before ingest ---
    FixSitesInSourceSheets

    ' --- Pull data from each source sheet into unified table ---
    Dim s As Long
    For s = LBound(gSrcSheets) To UBound(gSrcSheets)
        AppendSheetData gWB.sheets(CStr(gSrcSheets(s)))
    Next s

    ' --- Final cleanups & shaping ---
    FixMismatchedSiteNames
    FixSystemArchDataTypos
    SortAndFill_Muni_Type_Tier
    SortAndConsolidatePanels

    ' --- Finish ---
CleanExit:
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = prevCalc
        .DisplayAlerts = True
    End With

    If Err.Number = 0 Then
        If Not gWSArch Is Nothing Then gWSArch.Columns.AutoFit
        MsgBox "System architecture data unified successfully.", vbInformation
    End If
    Exit Sub

CleanFail:
    MsgBox "UnifyData failed: " & Err.Number & " - " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

' =====================================================================================
' INITIALIZATION & DESTINATION PREP
' =====================================================================================

Private Sub InitGlobals()
    Set gWB = ThisWorkbook

    ' Keep these in one place
    gHeaders = Array( _
        H_SITE, H_MUNI, H_TYPE, H_TIER, _
        H_BUILD, H_ROOM, H_PANEL, "Panel Function", _
        "Major Panel Components", "Other Components", _
        "PLC Name", "PLC Manufacturer", "PLC Model", _
        "OIT Manufacturer", "OIT Model", _
        H_COMM_MFG, H_COMM_MODEL, _
        H_INST_CAT, H_INST_TYPE, _
        H_INST_TAG, H_INST_MFG, H_INST_MODEL, _
        H_INST_HEALTH, H_INST_COND, H_INST_NAME _
    )

    gSrcSheets = Array("Site", "Building", "Room", "Control Panel", "PLC", "OIT", "Communication Equipment", "Field Instruments")
End Sub

Private Sub PrepareDestination()
    Dim ws As Worksheet

    ' Create or reset destination sheet
    On Error Resume Next
    Set ws = gWB.sheets(ARCH_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = gWB.sheets.Add
        ws.Name = ARCH_SHEET
    Else
        ws.Cells.Clear
    End If
    Set gWSArch = ws

    ' Write headers
    gWSArch.Range("A1").Resize(1, UBound(gHeaders) - LBound(gHeaders) + 1).value = gHeaders

    ' Create an empty table with one data row (table will auto-expand as we write)
    Set gTbl = gWSArch.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        Source:=gWSArch.Range("A1").Resize(2, UBound(gHeaders) - LBound(gHeaders) + 1), _
        XlListObjectHasHeaders:=xlYes)
    gTbl.DisplayName = "SystemArchTable"
End Sub

' =====================================================================================
' DATA INGEST
' =====================================================================================

' Append one source sheet's data into the unified table
Private Sub AppendSheetData(ByVal wsSrc As Worksheet)
    Dim hdrDict As Object             ' header name -> column index (1-based)
    Dim lastRow As Long, lastCol As Long
    Dim rowsToAdd As Long, rStart As Long, rEnd As Long
    Dim colName As Variant
    Dim destCol As Long
    Dim hasSiteCol As Boolean
    Dim tempCol() As Variant

    If wsSrc Is Nothing Then Exit Sub

    NormalizeSourceHeaders wsSrc

    ' Find used header row, last col & row (guard if empty)
    lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then Exit Sub

    Set hdrDict = HeaderIndexDict(wsSrc)
    If hdrDict Is Nothing Or hdrDict.Count = 0 Then Exit Sub

    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).row
    If lastRow < 2 Then Exit Sub  ' no data rows

    rowsToAdd = lastRow - 1
    If rowsToAdd <= 0 Then Exit Sub

    ' Determine destination start row (first row below table header)
    rStart = gWSArch.Cells(gWSArch.Rows.Count, 1).End(xlUp).row + IIf(gWSArch.Cells(1, 1).value = H_SITE, 1, 0)
    If rStart < 2 Then rStart = 2
    rEnd = rStart + rowsToAdd - 1

    ' Ensure the table expands by first writing a blank vector to Site Name
    destCol = GetColumnIndex(gWSArch, H_SITE)
    WriteBlankVector gWSArch.Cells(rStart, destCol).Resize(rowsToAdd, 1)

    ' If source provides Site Name, overwrite with real values after table expansion
    hasSiteCol = hdrDict.Exists(UCase$(H_SITE))
    If hasSiteCol Then
        tempCol = ColumnAsArray(wsSrc, 2, hdrDict(UCase$(H_SITE)), lastRow)
        gWSArch.Cells(rStart, destCol).Resize(rowsToAdd, 1).value = ToColumn(tempCol)
    End If

    ' For each unified header that exists in source, copy over the values
    For Each colName In gHeaders
        If hdrDict.Exists(UCase$(CStr(colName))) Then
            destCol = GetColumnIndex(gWSArch, CStr(colName))
            If destCol > 0 Then
                tempCol = ColumnAsArray(wsSrc, 2, hdrDict(UCase$(CStr(colName))), lastRow)
                gWSArch.Cells(rStart, destCol).Resize(rowsToAdd, 1).value = ToColumn(tempCol)
            End If
        End If
    Next colName
End Sub

' Normalize/sync specific header names in source sheets
Private Sub NormalizeSourceHeaders(ByVal ws As Worksheet)
    Dim map As Object, cell As Range, key As Variant
    Set map = CreateObject("Scripting.Dictionary")
    map.CompareMode = vbTextCompare

    ' Specific remaps requested in original code
    If ws.Name = "Communication Equipment" Then
        map("Manufacturer Name") = H_COMM_MFG
        map("Model Number") = H_COMM_MODEL
    ElseIf ws.Name = "Field Instruments" Then
        map("Manufacturer") = H_INST_MFG
        map("Model") = H_INST_MODEL
        map("Condition Health") = H_INST_HEALTH
        map("Condition Rating") = H_INST_COND
    End If

    If map.Count = 0 Then Exit Sub

    For Each cell In ws.Range(ws.Cells(1, 1), ws.Cells(1, ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column)).Cells
        If map.Exists(CStr(cell.value)) Then
            cell.value = map(CStr(cell.value))
        End If
    Next cell
End Sub

' =====================================================================================
' CLEANUP & FIX-UPS
' =====================================================================================

Public Sub GeneralCleanup()
    ' Light-weight cleanup across all sheets: normalize quotes, hyphen, line breaks, trim, collapse spaces.
    Dim ws As Worksheet, cell As Range
    Dim unicodeHyphen As String
    unicodeHyphen = ChrW(&H2010)

    For Each ws In gWB.Worksheets
        With ws.Cells
            .Replace What:=unicodeHyphen, replacement:="-", LookAt:=xlPart
            .Replace What:="’", replacement:="'", LookAt:=xlPart
            ' Consolidate ampersands (in case "&&" was used)
            .Replace What:="&&", replacement:="&", LookAt:=xlPart

            ' Normalize line breaks & tabs to spaces
            .Replace What:=vbCrLf, replacement:=" ", LookAt:=xlPart
            .Replace What:=vbCr, replacement:=" ", LookAt:=xlPart
            .Replace What:=vbLf, replacement:=" ", LookAt:=xlPart
            .Replace What:=vbTab, replacement:=" ", LookAt:=xlPart
        End With

        ' Trim and collapse multiple spaces per cell using Application.Trim
        For Each cell In ws.usedRange
            If VarType(cell.value) = vbString Then cell.value = Application.WorksheetFunction.Trim(cell.value)
        Next cell
    Next ws
End Sub

Public Sub FixSitesInSourceSheets()
    ' Split ambiguous site names based on ObjectID thresholds in specific sheets
    Dim sheets As Variant, splitObjectID As Variant
    Dim searchString As Variant, columnOfInterest As Variant
    Dim lowerReplacement As Variant, upperReplacement As Variant
    Dim s As Long
    Dim ws As Worksheet

    sheets = Array("Building", "Room", "Room")
    splitObjectID = Array(381, 596, 0)
    searchString = Array("Wadi Arab (II) PS2 old", "Wadi Arab Pumping Station 2 New-Old", "T3 Miyahuna Kazyah Well")
    columnOfInterest = Array(H_SITE, H_SITE, H_SITE)
    lowerReplacement = Array("Wadi Arab (II) PS2 old", "Wadi Arab (II) PS2 old", "")
    upperReplacement = Array("Wadi Arab (II) PS2 new", "Wadi Arab (II) PS2 new", "T3 Gas Station Well")

    For s = LBound(sheets) To UBound(sheets)
        Set ws = gWB.sheets(CStr(sheets(s)))
        ReplaceBasedOnObjectID ws, CLng(splitObjectID(s)), CStr(columnOfInterest(s)), CStr(searchString(s)), CStr(lowerReplacement(s)), CStr(upperReplacement(s))
    Next s
End Sub

Public Sub FixMismatchedSiteNames()
    Dim incorrectSiteNames As Variant, correctSiteNames As Variant
    incorrectSiteNames = Array("1-044 Kufr Asad PS1 & Well 1", "Abu Basal Pumping Station and Well 2", "Abu Basal Well 1", "Al Bwaydah ps", "Al Hdoud PS", "Al Za'tary Well 6", "Al-safi additional PS", "Al-Safi Main PS", "Amman Ps", "Aqeb 94.5 well", "Aqeb k 102 well", "Aqeb k 102.5 well", "Aqeb k 95 well", "Aqeb k104 well", _
    "Aqeb k109 well", "Aqeb k110 well", "Aqeb k114 well", "Aqeb k122 well", "Aqeb k91.5 well", "As-sammar old", "Bushra Pumping Station and Well 4", "Cg 25", "Cg 26", "Cg 27", "Cg 29", "Cg 30", "Cg31", "Check gate 10", "Check gate 12", "Check gate 14", "Check gate 15", "Check gate 17", "Check gate 18", "Check gate 19", _
    "Check gate 2", "Check gate 21", "Check gate 22", "Check gate 24 (CG)", "check gate 3", "Check gate 4", "Check gate 5", "Check gate 7", "Coridoor 17 well", "corridor 17", "Dabouq Reservoir", "Economic 2 well", "Economic 5 well", "Economic 6 well", "Edraa pump station", "H", "Hadeetheh Pump Station", _
    "Hakama well 6", "Hakama Well 8", "Hakama Well 9", "Hofa pump station", "Jaber well 1", "Juhfiyya pumping station , well 1 and well A1", "Kufr Asad well 7", "Kufr Assad well 6", "Kufr Assad well 8", "Mahasi 8b (new berkeh)", "Mindah PS4 & Well 4", "Mindah Pumping Station 1 and Well 1", "Mindah well 2", "Mindah well 3", _
    "Miyahuna - Main SCADA Center", "MIZFFER", "Mudawwara Balad PS & RSV", "Mudawwara Border PS & RSV", "Muogaer sarhan1", "Muogaer sarhan2", "New Zatary PS", "No'aymeh pumping station 2 and well", "No'aymeh pumping station1 & well1", "No'aymeh well 3", "No'aymeh well A1", "Nu'aymeh pumping station 4 and well 4", _
    "Nu'aymeh well 5", "Nu'aymeh well 6", "Nu'aymeh well 8", "Old Za'tary", "Old Za'tary PS", "Oyoon al hamman well 1", "Oyoon al hamman well 2", "Oyoon al hamman well 4", "Oyoon al hamman well 5", "Oyoon al hamman well 6", "Oyoun Al Hamam PS", "Pump Station 10", "Pump Station 11 and check gate 3", _
    "Reservoir um njash", "Sumaya 5", "Sumaya 6", "Suwelmeh ps", "T3 10 Well", "T3 17 Well", "T3 Abu Siah Booster", "T3 Booster 11", "T3 Mazra A well11", "T3 Miyahuna Ain Ghazal 7 Well", "T3 Miyahuna Nader pumping station", "T3 miyahuna sebaq well 27", "Tlal althahab", "To 10", "To 18", "To 21 and Check Gate 8", "TO 22.2", _
    "TO 24", "TO 28", "TO 33", "TO 36", "TO 41", "TO 50", "To 53", "TO 62", "To 65", "To 72", "To 78", "To 81", "To 86", "To 91", "To 95", "UmEnjasha Reservoir", "Wadi Arab Pumping Station 1 New", "Wadi Arab Pumping Station 1 New-Old", "Wadi Arab Pumping Station 2 New", "Wadi Arab Pumping Station 2 old", "Wadi Arab Well-1", _
    "Wadi Arab Well-13", "Wadi Arab Well-5", "Wadi Arab Well-6", "Wadi Arab Well-9", "Zai pump station", "Zara mae'en ps", "Zara Ma'in wtp", "Zara Ma'in wtp + pump station 1", "Zarqa mae'en PS", "Zatary Old PS", "Znaieh well 10", "Znaieh well 3", "Znaieh well 4", "Znaieh well 5", "Znaieh well 6", "Weir abu zeghab", _
    "Al’ Aqeb K104C", "Aqeb 94.5***", "T3 Miyahuna DZ13 F*", "T3 Miyahuna DZ13C*", "Weir Abu Zeghab")
    
    correctSiteNames = Array("Kufur Assad PS1 & Kufur Assad 1", "Abu Albassal PS and Abu Albassal 2", "Abu Albassal 1", "Al Bwaydah PS", "Hodoud Pumping Station and Well", "Al Za'tary 6a", "Al-Safi Additional Pump Station", "Al-Safi Main Pump Station", "Amman Pump Station", "Aqeb 94.5", "Aqeb k 102", "Al' Aqeb K102.5", _
    "Al' Aqeb K95b", "Al' Aqeb K104c", "Al' Aqeb K109b", "Al' Aqeb K110", "Al' Aqeb K114", "Al' Aqeb K122", "Aqeb 91.5", "As-Sammar Old Pump Station", "Bushra PS and hakama 4a", "Check Gate 25", "Check Gate 26", "Check Gate 27", "Check Gate 29", "Check Gate 30", "Check Gate 31", "Check Gate 10", "Check Gate 12", _
    "Check Gate 14", "Check Gate 15", "Check Gate 17", "Check Gate 18", "Check Gate 19", "Check Gate 2", "Check Gate 21", "Check Gate 22", "Check Gate 24", "Check Gate 3", "Check Gate 4", "Check Gate 5", "Check Gate 7", "Corodor 17 Well", "Corodor 17 Well", "T2 - Miyahuna Dabouq Reservoir", "Economic 2", "Economic 5", _
    "Economic 6", "Edra'a Pump Station", "Ain Hamed Spring", "Hadeetheh Pump Station, Ain Maghara and Wadi Hamad Sources", "Hakama 6", "Hakama 8", "Hakama 9", "Hofa new PS", "Jaber 1", "Juhfia PS, Juhfiyah 1 and Juhfiyah 2", "Kufur Assad 7", "Kufur Assad 6", "Kufur Assad 8", "Mahasi Pumping Station, Well 5, Well 6 and Well 6B", _
    "Mandah PS 4 and Mandah 4", "Mandah PS 1 and Mandah 1", "Mandah 2", "Mandah 3", "T1 - Miyahuna Main SCADA Center", "Mizfer PS", "Mudawwara Balad RO, PS & RSV", "Mudawwara Border RO, PS & RSV", "Muogaer sarhan 1", "Muogaer sarhan 2", "New Za'tary", "Noa'meh PS 2 and No'aymeh 2", "Noa'meh PS 1", _
    "No'aymeh 3", "No'aymeh 1", "Noa'meh PS 4 and No'aymeh 4", "No'aymeh 5", "No'aymeh 6", "No'aymeh 8", "Old Za'tary PS", "Old Za'tary PS", "Oyoon Al Hammam 1", "Oyoon Al Hammam 2", "Oyoon Al Hammam 4", "Oyoon Al Hammam 5", "Oyoon Al Hammam 6", "Oyoon al hamman pump station", "Pump Station 10 New", "Pump Station 11", _
    "Um Enjasha Reservoir", "Sumaya 5a", "Sumaya 6b", "Swailemah PS", "T3 Ruseifa Well 10", "T3 Ruseifa Well 17", "T3 Abu Sayah Booster", "T3 Rusaifa 11 booster", "T3 Mazra'a well", "T3 Miyahuna Ain Ghazal Well 7", "T3 Miyahuna Naser pumping station", "T3 miyahuna swaqa well 27", "Tlal Azzahab", "Pump Station 10 New", _
    "Pump Station 18", "Pump Station 21", "Pump Station 22", "Pump Station 24", "Pump Station 28", "Pump Station 33 (Al-Mashare')", "Pump Station 36 (Wadi Al-Rayan)", "Pump Station 41 (Abu Habeel)", "Pump Station 50", "Pump Station 53", "Pump Station 62", "Pump Station 65", "Pump Station 72", "Pump Station 78", _
    "Pump Station 81", "Pump Station 86", "Pump Station 91", "Pump Station 95", "Um Enjasha Reservoir", "Wadi Arab (I) PS1 new", "Wadi Arab (I) PS1 new", "Wadi Arab (II) PS2 new", "Wadi Arab (II) PS2 old", "Wadi Al Arab 1", "Wadi Al Arab 13a", "Wadi Al Arab 5", "Wadi El Arab 6", "Wadi Al Arab 9", _
    "Zai WTP and Pump Station 5", "Zara Ma'een Pump Station", "Zara ma'een pumping station 1", "Zara ma'een pumping station 1", "Zarqa Ma'een Pump Station", "Old Za'tary PS", "Znaieh 10", "Znaieh 3", "Znaieh 4", "Znaieh 5", "Znaieh 6", "Abu Zeghan Canal Inlet", "Al' Aqeb K104C", "Aqeb 94.5", "T3 Miyahuna DZ13 F", "T3 Miyahuna DZ13C", _
    "Abu Zeghan Canal Inlet")
    ReplaceSiteNames ARCH_SHEET, incorrectSiteNames, correctSiteNames
End Sub

Public Sub SortAndConsolidatePanels()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim key As String
    Dim panelDict As Object
    Dim i As Long, j As Long
    Dim siteName As String, buildingName As String, roomName As String, panelName As String
    Dim skipRow As Boolean
    Dim commManufCol As Long, commModelCol As Long, instCatCol As Long, instTypeCol As Long
    Dim targetRow As Long

    Set ws = gWSArch
    If ws Is Nothing Then Exit Sub

    ' Sort by core hierarchy
    SortTableByHeaders gTbl, Array(H_SITE, H_BUILD, H_ROOM, H_PANEL)

    lastRow = ws.Cells(ws.Rows.Count, GetColumnIndex(ws, H_SITE)).End(xlUp).row
    If lastRow < 2 Then Exit Sub

    commManufCol = GetColumnIndex(ws, H_COMM_MFG)
    commModelCol = GetColumnIndex(ws, H_COMM_MODEL)
    instCatCol = GetColumnIndex(ws, H_INST_CAT)
    instTypeCol = GetColumnIndex(ws, H_INST_TYPE)

    Set panelDict = CreateObject("Scripting.Dictionary")
    i = 2
    Do While i <= lastRow
        siteName = CStr(ws.Cells(i, GetColumnIndex(ws, H_SITE)).value)
        buildingName = CStr(ws.Cells(i, GetColumnIndex(ws, H_BUILD)).value)
        roomName = CStr(ws.Cells(i, GetColumnIndex(ws, H_ROOM)).value)
        panelName = CStr(ws.Cells(i, GetColumnIndex(ws, H_PANEL)).value)

        skipRow = (Nz(ws.Cells(i, commManufCol).value) <> "" Or _
                   Nz(ws.Cells(i, commModelCol).value) <> "" Or _
                   Nz(ws.Cells(i, instCatCol).value) <> "" Or _
                   Nz(ws.Cells(i, instTypeCol).value) <> "")

        key = siteName & vbLf & buildingName & vbLf & roomName & vbLf & panelName

        If key <> vbLf & vbLf & vbLf And Not panelDict.Exists(key) Then
            panelDict.Add key, i
            i = i + 1
        ElseIf skipRow Then
            i = i + 1
        Else
            ' Merge non-key fields into the first instance, then delete duplicate row
            targetRow = CLng(panelDict(key))
            For j = GetColumnIndex(ws, H_PANEL) + 1 To commManufCol - 1
                If Nz(ws.Cells(i, j).value) <> "" Then
                    ws.Cells(targetRow, j).value = ws.Cells(i, j).value
                End If
            Next j
            ws.Rows(i).Delete
            lastRow = lastRow - 1
            ' do not increment i; check the new row that moved up
        End If
    Loop
End Sub

Public Sub SortAndFill_Muni_Type_Tier()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long

    Set ws = gWSArch
    If ws Is Nothing Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, GetColumnIndex(ws, H_SITE)).End(xlUp).row
    If lastRow < 2 Then Exit Sub

    ' Initial sort by Site Name, Municipality
    SortTableByHeaders gTbl, Array(H_SITE, H_MUNI)

    ' Normalize capitalization differences by harmonizing consecutive equal site names (case-insensitive)
    For i = 3 To lastRow
        If StrComp(ws.Cells(i, GetColumnIndex(ws, H_SITE)).value, ws.Cells(i - 1, GetColumnIndex(ws, H_SITE)).value, vbTextCompare) = 0 Then
            ws.Cells(i, GetColumnIndex(ws, H_SITE)).value = ws.Cells(i - 1, GetColumnIndex(ws, H_SITE)).value
        End If
    Next i

    ' Resort by Site Name only
    SortTableByHeaders gTbl, Array(H_SITE)

    ' Fill down Municipality, Site Type, SCADA Tier within same Site
    For i = 3 To lastRow
        If StrComp(ws.Cells(i, GetColumnIndex(ws, H_SITE)).value, ws.Cells(i - 1, GetColumnIndex(ws, H_SITE)).value, vbTextCompare) = 0 Then
            If Nz(ws.Cells(i, GetColumnIndex(ws, H_MUNI)).value) = "" Then ws.Cells(i, GetColumnIndex(ws, H_MUNI)).value = ws.Cells(i - 1, GetColumnIndex(ws, H_MUNI)).value
            If Nz(ws.Cells(i, GetColumnIndex(ws, H_TYPE)).value) = "" Then ws.Cells(i, GetColumnIndex(ws, H_TYPE)).value = ws.Cells(i - 1, GetColumnIndex(ws, H_TYPE)).value
            If Nz(ws.Cells(i, GetColumnIndex(ws, H_TIER)).value) = "" Then ws.Cells(i, GetColumnIndex(ws, H_TIER)).value = ws.Cells(i - 1, GetColumnIndex(ws, H_TIER)).value
        End If
    Next i

    ' Final sort by Site > Building > Room > Panel
    SortTableByHeaders gTbl, Array(H_SITE, H_BUILD, H_ROOM, H_PANEL)
End Sub

Public Sub FixSystemArchDataTypos()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long
    Dim cell As Range, txt As String, words() As String, newText As String
    Dim capitalizationException As Variant, upperCaseList As Variant
    Dim removalList As Variant, spellCheckOriginal As Variant, spellCheckFixed As Variant
    Dim k As Long, m As Long, n As Long
    Dim w As String, spellChecked As Boolean

    Set ws = gWSArch
    If ws Is Nothing Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, GetColumnIndex(ws, H_SITE)).End(xlUp).row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastRow < 1 Or lastCol < 1 Then Exit Sub

    capitalizationException = Array("and", "the", "in")
    upperCaseList = Array("ARUBA", "GPRS", "GTO", "GTU", "IPC", "KAC", "NVR", "SCADA", "TP", "WTP", "RO", _
                          "ID", "CPU", "RAM", "SSD", "HDD", "PSU", "UPS", "GPU", "PLC", "OIT", "MCC", "PS", "HQ", "RSV")
    removalList = Array("T3 - ", "T3 ", " T3 ", "T2 - ", "T2 ", "T1 - ")
    spellCheckOriginal = Array("(adaseya)", "(alakhdar)", "(Allen-Bradlley)", "(Modawwarah)", ".Old", "Adasyieh", "alandaa", "Alandaa(AA2", "al-arab", "Al-Deek", "ALJafer", "AL-Riyadi", _
        "Al-safi", "arab", "ashrafeh", "Ashrafeyeh", "As-sammar", "Ba'oon", "Birmathkoor", "Chlor", "Chlore", "Chlorification", "Chlorinatetion", "Chlorinatoin", "Chlorintion", "Cholrine", "Collecter", _
        "Coridoor", "Dabbet", "dusconnector", "Filteration", "Gantarah", "Hanoot", "Harfosshie", "Hofa", "Irainbeh", "Iranbeh", "jmail", "Jmail", "joba", "Judtta", "Kofranjeh", "Kraimeh", "Kraymeh", _
        "Kufranja", "Kufranjeh", "lonk", "ma'en", "Ma'in", "mae'en", "Maglis", "Manshiyyah", "Mazra", "Mazra'a", "miyahuma", "Miyahuna-", "miyauhna", "Mughaier", "Muwaqer", "Mwaqar", "Mwaqar-WELL", "Nader", _
        "No'aymeh", "Nu'aymeh", "Numatic", "PANL", "Pupms", "qalaa", "Qantara", "Qanter", "Qweisma", "Rahemh", "Rashdieh", "relayes,", "Reservior", "Reservoird", "Reshah", "Resheh", "resrvoir", "rOOM", "Rusaifa", _
        "Rusaiffeh", "russaifeh", "Russaiffeh", "Russeifa", "Scadapack", "Shammout", "Shmesani", "Shmisani", "Shujairat", "Softstarter", "Starl", "Sttion", "Suweileh", "Swich", "Tabaqat", "termenals", _
        "Unstractured", "Uthainah", "Utheineh", "wier", "Yadoudhe", "Za'atary", "Zatary", "Zmailah", "Pumping", "PS", "PS1", "PS2", "PS3", "PS4", "PS5", "PS6", "PS7", "PS8", "PS9", "PS10", "Well1", "Well2", _
        "Well3", "Well4", "Well5", "Well6", "Well7", "Well8", "Well9", "Well10", "Well11", "Well12", "Well13", "Well14", "Well15", "Well16", "Well17", "Well18", "Well19", "Well20", "Well21", "Well22", "Well23", _
        "Well24", "Well25", "Well26", "Well27", "Well28", "Well29", "Well30", "Well31", "Well32", "Well33", "Well34", "Well35", "Juhfia", "Juhfiyya", "Mindah", "Noa'meh", "Swailemah")
    spellCheckFixed = Array("(Adasiya)", "(Alakhdar)", "(Allen-Bradley)", "(Mudawwara)", "Old", "Adasiya", "Alanda", "Alanda (AA2", "Al-Arab", "Al Deek", "Al Jafer", "Al-Riyadi", _
        "Al-Safi", "Arab", "Ashrafiyeh", "Ashrafiyeh", "As-Sammar", "Ba'oon", "Bir Mathkoor", "Chlorine", "Chlorine", "Chlorination", "Chlorination", "Chlorination", "Chlorination", "Chlorine", "Collector", _
        "Corodor", "Dabet", "disconnector", "Filtration", "Qantarah", "Hanut", "Harfosheya", "Hoffa", "Ireinbeh", "Ireinbeh", "Jamil", "Jamil", "Jobba", "Judetta", "Kfrnga", "Krayymeh", "Krayymeh", "Kfrnga", _
        "Kfrnga", "Link", "Mae'en", "Mae'en", "Mae'en", "Magelis", "Manshyeh", "Mazra'a", "Mazra'a", "Miyahuna", "Miyahuna", "Miyahuna", "Muogaer", "Muwaqqar", "Muwaqqar", "Muwaqqar-WELL", "Naser", "No'aymeh", _
        "No'aymeh", "Pneumatic", "PANEL", "Pumps", "Qala'a", "Qantarah", "Qantarah", "Qwaismeh", "Rahmeh", "Rashdiah", "relays,", "Reservoir", "Reservoir", "Risha", "Risha", "Reservoir", "Room", _
        "Ruseifa", "Ruseifa", "Ruseifa", "Ruseifa", "Ruseifa", "SCADAPack", "Shamout", "Shmeisani", "Shmeisani", "Shujeirat", "Soft Starter", "Starlink", "Station", "Swaileh", "Switch", "Tabagat", "Terminals", _
        "Unstructured", "Uthaina", "Uthaina", "Weir", "Yadoudeh", "Za'tary", "Za'tary", "Zmaileh", "Pump", "Pump Station", "Pump Station 1", "Pump Station 2", "Pump Station 3", "Pump Station 4", _
        "Pump Station 5", "Pump Station 6", "Pump Station 7", "Pump Station 8", "Pump Station 9", "Pump Station 10", "Well 1", "Well 2", "Well 3", "Well 4", "Well 5", "Well 6", "Well 7", "Well 8", "Well 9", _
        "Well 10", "Well 11", "Well 12", "Well 13", "Well 14", "Well 15", "Well 16", "Well 17", "Well 18", "Well 19", "Well 20", "Well 21", "Well 22", "Well 23", "Well 24", "Well 25", "Well 26", "Well 27", _
        "Well 28", "Well 29", "Well 30", "Well 31", "Well 32", "Well 33", "Well 34", "Well 35", "Juhfiyah", "Juhfiyah", "Mandah", "No'aymeh", "Suwelmeh")

    ' Remove unwanted prefixes broadly
    Dim rr As Long
    For rr = LBound(removalList) To UBound(removalList)
        ws.Cells.Replace What:=CStr(removalList(rr)), replacement:="", LookAt:=xlPart
    Next rr

    ' Process cell-by-cell string normalization
    For i = 1 To lastRow
        For j = 1 To lastCol
            Set cell = ws.Cells(i, j)
            If VarType(cell.value) = vbString Then
                txt = CStr(cell.value)
                words = Split(txt, " ")

                For k = LBound(words) To UBound(words)
                    w = words(k)
                    spellChecked = False

                    ' Light spell correction
                    For m = LBound(spellCheckOriginal) To UBound(spellCheckOriginal)
                        If StrComp(w, spellCheckOriginal(m), vbTextCompare) = 0 Then
                            words(k) = spellCheckFixed(m)
                            spellChecked = True
                            Exit For
                        End If
                    Next m
                    If spellChecked Then GoTo NextWord

                    ' Uppercase known codes or tokens containing digits
                    If IsInArray(UCase$(w), upperCaseList) Then
                        words(k) = UCase$(w)
                    ElseIf w Like "*[0-9]*" Then
                        words(k) = UCase$(w)
                    ElseIf IsInArray(LCase$(w), capitalizationException) Then
                        words(k) = LCase$(w)
                    ElseIf Not (Left$(w, 1) Like "[A-Z]") Then
                        ' Title-case first alphabetic character
                        n = 1
                        Do While n <= Len(w) And Not (Mid$(w, n, 1) Like "[A-Za-z]")
                            n = n + 1
                        Loop
                        If n <= Len(w) Then
                            words(k) = Left$(w, n - 1) & UCase$(Mid$(w, n, 1)) & Mid$(w, n + 1)
                        End If
                    End If
NextWord:
                Next k

                newText = Join(words, " ")
                If newText <> txt Then cell.value = newText
            End If
        Next j
    Next i

    ws.Columns.AutoFit
    ws.Rows.AutoFit
End Sub

' =====================================================================================
' REPLACEMENT UTILITIES
' =====================================================================================

Public Sub ReplaceBasedOnObjectID(ws As Worksheet, _
                                  ByVal splitObjectID As Long, _
                                  ByVal columnOfInterest As String, _
                                  ByVal searchString As String, _
                                  ByVal lowerReplacement As String, _
                                  ByVal upperReplacement As String)
    Dim colIndex As Long, objectIdCol As Long, lastRow As Long, i As Long
    Dim objectID As Variant

    colIndex = GetColumnIndex(ws, columnOfInterest)
    objectIdCol = GetColumnIndex(ws, "ObjectID")
    If colIndex = 0 Or objectIdCol = 0 Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, colIndex).End(xlUp).row
    For i = 2 To lastRow
        If InStr(1, CStr(ws.Cells(i, colIndex).value), searchString, vbTextCompare) > 0 Then
            objectID = ws.Cells(i, objectIdCol).value
            If IsNumeric(objectID) Then
                If CLng(objectID) <= splitObjectID Then
                    ws.Cells(i, colIndex).value = lowerReplacement
                ElseIf CLng(objectID) >= splitObjectID Then
                    ws.Cells(i, colIndex).value = upperReplacement
                End If
            End If
        End If
    Next i
End Sub  ' <<<< Fixed missing End Sub

Public Sub ReplaceSiteNames(ByVal sheetName As String, ByVal incorrectNames As Variant, ByVal correctNames As Variant)
    Dim ws As Worksheet
    Dim siteCol As Long
    Dim i As Long

    Set ws = gWB.sheets(sheetName)
    siteCol = GetColumnIndex(ws, H_SITE)
    If siteCol = 0 Then
        MsgBox "Column '" & H_SITE & "' not found in sheet '" & sheetName & "'", vbExclamation
        Exit Sub
    End If

    With ws.Columns(siteCol)
        For i = LBound(incorrectNames) To UBound(incorrectNames)
            .Replace What:=CStr(incorrectNames(i)), replacement:=CStr(correctNames(i)), LookAt:=xlWhole, MatchCase:=False
        Next i
    End With
End Sub

' =====================================================================================
' HELPER FUNCTIONS
' =====================================================================================

' Build a case-insensitive dictionary of header -> column index from row 1
Private Function HeaderIndexDict(ByVal ws As Worksheet) As Object
    Dim d As Object, cell As Range
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    For Each cell In ws.Range(ws.Cells(1, 1), ws.Cells(1, ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column))
        If Nz(cell.value) <> "" Then d(UCase$(CStr(cell.value))) = cell.Column
    Next cell
    Set HeaderIndexDict = d
End Function

' Get column index by header name in a worksheet (row 1)
Public Function GetColumnIndex(ByVal ws As Worksheet, ByVal headerName As String) As Long
    Dim cell As Range, headerStr As String
    headerStr = Trim$(headerName)
    If Len(headerStr) = 0 Then GetColumnIndex = 0: Exit Function

    For Each cell In ws.Range("A1", ws.Cells(1, ws.Columns.Count).End(xlToLeft))
        If StrComp(Trim$(CStr(cell.value)), headerStr, vbTextCompare) = 0 Then
            GetColumnIndex = cell.Column
            Exit Function
        End If
    Next cell
    GetColumnIndex = 0
End Function

' Case-insensitive membership test for small arrays
Public Function IsInArray(ByVal val As String, ByVal arr As Variant) As Boolean
    Dim element As Variant
    For Each element In arr
        If StrComp(val, CStr(element), vbTextCompare) = 0 Then IsInArray = True: Exit Function
    Next element
    IsInArray = False
End Function

' Null/empty-safe string
Private Function Nz(ByVal v As Variant, Optional ByVal fallback As String = "") As String
    If IsError(v) Then
        Nz = fallback
    ElseIf IsNull(v) Then
        Nz = fallback
    ElseIf VarType(v) = vbString And v = "" Then
        Nz = fallback
    Else
        Nz = CStr(v)
    End If
End Function

' Read a sheet column (from rowStart to lastRow) into a 1-D variant (1-based)
Private Function ColumnAsArray(ByVal ws As Worksheet, ByVal rowStart As Long, ByVal colIndex As Long, ByVal lastRow As Long) As Variant
    Dim n As Long, i As Long
    Dim src As Variant, outArr() As Variant

    If lastRow < rowStart Then
        ReDim outArr(1 To 0) ' empty
        ColumnAsArray = outArr
        Exit Function
    End If

    src = ws.Cells(rowStart, colIndex).Resize(lastRow - rowStart + 1, 1).value
    n = UBound(src, 1)
    ReDim outArr(1 To n)
    For i = 1 To n
        If Not IsError(src(i, 1)) Then
            outArr(i) = Trim$(CStr(src(i, 1)))
        Else
            outArr(i) = "#FormError#"
        End If
    Next i
    ColumnAsArray = outArr
End Function

' Convert a 1-D (1-based) array to a 2-D column vector safe for .Value assignment (no Transpose)
Private Function ToColumn(ByVal v As Variant) As Variant
    Dim n As Long, i As Long, lb As Long, ub As Long
    lb = LBound(v): ub = UBound(v)
    n = ub - lb + 1
    Dim arr() As Variant
    ReDim arr(1 To n, 1 To 1)
    For i = 1 To n
        arr(i, 1) = v(lb + i - 1)
    Next i
    ToColumn = arr
End Function

' Quickly write blanks (to force table growth before overwriting with real data)
Private Sub WriteBlankVector(ByVal rng As Range)
    Dim n As Long, i As Long
    n = rng.Rows.Count
    Dim arr() As Variant
    ReDim arr(1 To n, 1 To 1)
    For i = 1 To n
        arr(i, 1) = vbNullString
    Next i
    rng.value = arr
End Sub

' Sort a ListObject by an array of header names (ascending)
Private Sub SortTableByHeaders(ByVal lo As ListObject, ByVal headerNames As Variant)
    Dim sf As SortFields
    Dim i As Long
    With lo.Parent.Sort
        .SortFields.Clear
        For i = LBound(headerNames) To UBound(headerNames)
            .SortFields.Add key:=lo.ListColumns(CStr(headerNames(i))).Range, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        Next i
        .SetRange lo.Range
        .Header = xlYes
        .MatchCase = False
        .Apply
    End With
End Sub



