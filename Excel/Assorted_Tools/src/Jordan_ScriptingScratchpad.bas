Private Sub UnifyData()
    Dim ws As Worksheet
    Dim wsSystemArchData As Worksheet
    Dim tblSystemArch As ListObject
    Dim rangeString As String
    Dim lastRow As Long
    Dim nextRow As Long
    Dim nextColumn As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim newRowCreated As Long
    Dim srcSheets As Variant
    Dim srcAllColumnLetters As Variant
    Dim columnArray As Variant
    Dim dataArray As Variant
    Dim headerRange As Range
    
    Dim headerArray As Variant
    Dim headerMap As Object
    Dim sheetIndex As Long
    Dim colIndex As Long
    Dim cell As Range
    Dim foundCol As Long
    Dim headerName As String
    
    Dim header As Variant
    Dim srcHeaderRow As Range
    Dim srcHeaders As Variant
    
    Dim columnIndex As Long
    Dim workingColumn As Long

    
    ThisWorkbook.Activate

    ' Create or clear the "SystemArchData" sheet
    On Error Resume Next
    Set wsSystemArchData = sheets("SystemArchData")
    On Error GoTo 0
    If wsSystemArchData Is Nothing Then
        Set wsSystemArchData = sheets.Add
        wsSystemArchData.Name = "SystemArchData"
    Else
        wsSystemArchData.Cells.Clear
    End If

    ' Set headers for the "SystemArchData" sheet
    headerArray = Array("Site Name", "Municipality", "Site Type", "SCADA Tier", "Building Name", "Room Name", "Panel Name", "Panel Function", "Major Panel Components", "Other Components", "PLC Name", "PLC Manufacturer", "PLC Model", "OIT Manufacturer", "OIT Model", "Comm Manufacturer Name", "Comm Model Number", "Instrument Category", "Instrument Type", "Instrument Tag", "Instrument Manufacturer", "Instrument Model", "Instrument Health", "Instrument Condition Rating", "Field Instrument Name") 'ObjectID becomes irrelevant after ConsolidatePanels
 
    'Sets the first row of the SystemArchData sheet to the values found in the headerArray
    wsSystemArchData.Range("A1").Resize(1, ArrayDimSize(headerArray)).Value = headerArray
    
    'Set initTableRange to contain 2 rows (one for the header and one for data), and as many columns as necessary given the headers in headerArray
    Set initTableRange = wsSystemArchData.Range(wsSystemArchData.Cells(1, 1), wsSystemArchData.Cells(2, ArrayDimSize(headerArray)))

    Set tblSystemArch = wsSystemArchData.ListObjects.Add(SourceType:=xlSrcRange, Source:=initTableRange, XlListObjectHasHeaders:=xlYes)
    tblSystemArch.DisplayName = "SystemArchTable"

    ' Define sheet names and relevant columns
    srcSheets = Array("Site", "Building", "Room", "Control Panel", "PLC", "OIT", "Communication Equipment", "Field Instruments")
    
    'Replace characters, double spaces, line breaks, etc that will mess up string parsing and display. This is sheet-wide
    'Call GeneralCleanup
    
    'Fix site names that were typed in incorrectly in various sheets. Ideally this will only implement small chnages to the source data.
    Call FixSitesInSourceSheets
    
    'Sets the first row of the SystemArchData sheet to the values found in the headerArray
    For i = LBound(headerArray, 1) To UBound(headerArray, 1)
        wsSystemArchData.Cells(1, i + 1).Value = headerArray(i)
    Next i

    nextRow = 2
    nextColumn = 1
    
    ' Loop through each sheet
    For sheetIndex = LBound(srcSheets) To UBound(srcSheets)
        Set ws = ThisWorkbook.sheets(srcSheets(sheetIndex))
        Set srcHeaderRow = ws.Range("A1", ws.Cells(1, ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column))
        foundHeaderCount = 0
        ReDim srcHeaders(0)

        ' Create a dictionary to map header names to column numbers
        Set headerMap = CreateObject("Scripting.Dictionary")

        ' Replace two headers in Comm sheet for functionality
        If ws.Name = "Communication Equipment" Then
            ws.Range("1:1").Replace What:="Manufacturer Name", Replacement:="Comm Manufacturer Name", LookAt:=xlWhole
            ws.Range("1:1").Replace What:="Model Number", Replacement:="Comm Model Number", LookAt:=xlWhole
        End If
        
        ' Replace two headers in Field Instrument sheet for functionality
        If ws.Name = "Field Instruments" Then
            ws.Range("1:1").Replace What:="Manufacturer", Replacement:="Instrument Manufacturer", LookAt:=xlWhole
            ws.Range("1:1").Replace What:="Model", Replacement:="Instrument Model", LookAt:=xlWhole
            ws.Range("1:1").Replace What:="Condition Health", Replacement:="Instrument Health", LookAt:=xlWhole
            ws.Range("1:1").Replace What:="Condition Rating", Replacement:="Instrument Condition Rating", LookAt:=xlWhole
        End If
        
        ' Scan the header row and populate the dictionary
        For Each cell In srcHeaderRow.Cells
            For i = LBound(headerArray) To UBound(headerArray)
                If StrComp(Trim(cell.Value2), headerArray(i), vbTextCompare) = 0 Then
                    headerMap(Trim(cell.Value2)) = cell.Column
                    headerMap(cell.Column) = Trim(cell.Value2)
                    
                    ReDim Preserve srcHeaders(1 To UBound(srcHeaders) + 1)
                    srcHeaders(UBound(srcHeaders)) = headerArray(i)
                    Exit For
                End If
            Next i
        Next cell
        
        lastRow = ws.Cells(ws.Rows.Count, GetColumnIndex(ws, "Site Name")).End(xlUp).Row
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        
        ReDim dataArray(LBound(srcHeaders) To UBound(srcHeaders), 2 To lastRow) ' Dimmed from Column 1, Row 2 to skip header row
        workingColumn = LBound(dataArray, 1)
        
        For i = 1 To lastCol
            If headerMap.exists(i) Then
                For j = LBound(dataArray, 2) To UBound(dataArray, 2)
                    If Not IsError(ws.Cells(j, i).Value2) Then
                        dataArray(workingColumn, j) = Trim(ws.Cells(j, i).Value2)
                    Else
                        dataArray(workingColumn, j) = "#FormError#"
                    End If
                Next j
                workingColumn = workingColumn + 1
            End If
        Next i

        lastCol = wsSystemArchData.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        For i = LBound(dataArray, 1) To UBound(dataArray, 1)
            columnIndex = GetColumnIndex(wsSystemArchData, srcHeaders(i))
            wsSystemArchData.Range(wsSystemArchData.Cells(nextRow, columnIndex), wsSystemArchData.Cells(nextRow + UBound(dataArray, 2) - 2, columnIndex)).Value = Application.Transpose(Application.Index(dataArray, i, 0))
        Next i
        
        'Assess next row based on last filled row in "Site Name" column
        nextRow = wsSystemArchData.Cells(wsSystemArchData.Rows.Count, GetColumnIndex(wsSystemArchData, "Site Name")).End(xlUp).Row + 1

    Next sheetIndex
        
    'Fix whole site names that don't match as they should in the SystemArchData sheet
    Call FixMismatchedSiteNames
    
    'Fix individual words that were typed incorrectly in the SystemArchData sheet
    Call FixSystemArchDataTypos
    
    'Because Municipality, Site Type, and Site Tier are only present in the Site sheet, we need to add this info to rows derived from other sheets
    'Depends on the site names, etc matching
    Call SortAndFill_Muni_Type_Tier
    
    'Combine all panel equipment into one row from multiple
    Call SortAndConsolidatePanels

    ' Autofit columns
    wsSystemArchData.Columns.AutoFit
End Sub

Private Sub GeneralCleanup()
    'Some survey results use a non-ASCII hyphen, which interferes with search results and prevents a string's use in VBA. This sub replaces them with a more typical hyphen
    ThisWorkbook.Activate
    Dim ws As Worksheet
    
    Dim unicodeChar As String
    Dim replaceChar As String
    
    'Replaces all instances of hyphen with hyphen minus (standard on US keyboards)
    unicodeChar = ChrW(&H2010)
    replaceChar = "-"
        
    For Each ws In ThisWorkbook.Worksheets
    'If ws.Name = "Field Instruments" Then
        'ws.UsedRange.Replace What:=unicodeChar, Replacement:=replaceChar
        ws.Cells.Replace What:=unicodeChar, Replacement:=replaceChar
        ws.Cells.Replace What:="’", Replacement:="'"
        ws.Cells.Replace What:="&amp;", Replacement:="&"
        ws.Cells.Replace What:="&amp;&amp;", Replacement:="&"
        ws.Cells.Replace What:="’", Replacement:="'"
        
        
        'Remove all double spaces and line breaks
        ws.Cells.Replace What:=vbCrLf, Replacement:=" "
        ws.Cells.Replace What:=vbCr, Replacement:=" "
        ws.Cells.Replace What:=vbLf, Replacement:=" "
        ws.Cells.Replace What:=vbTab, Replacement:=" "
        Do
            ws.Cells.Replace What:="  ", Replacement:=" ", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
            Set cell = ws.Cells.Find(What:="  ", LookIn:=xlValues)
        Loop Until cell Is Nothing
        
        ' Remove leading and trailing spaces from each cell
        For Each cell In ws.UsedRange
            If Not IsEmpty(cell.Value) Then
                If VarType(cell.Value) = vbString Then
                    cell.Value = Trim(cell.Value)
                End If
            End If
        Next cell
    'End If
    Next ws
End Sub

Private Sub FixSitesInSourceSheets()
    ThisWorkbook.Activate
    Dim ws As Worksheet
    Dim sheets As Variant, splitObjectID As Variant, searchString As Variant, columnOfInterest As Variant, lowerReplacement As Variant, upperReplacement As Variant
    Dim i As Long, s As Long
    Dim siteNameCol As Long, objectIdCol As Long, buildingNameCol As Long, lastRow As Long
    
    'Define the info for the replacements. Use 0 in splitObjectID and a blank in lowerReplacement to replace strings that don't have two options
    sheets = Array("Building", "Room", "Room")
    splitObjectID = Array(381, 596, 0)
    searchString = Array("Wadi Arab (II) PS2 old", "Wadi Arab Pumping Station 2 New-Old", "T3 Miyahuna Kazyah Well")
    columnOfInterest = Array("Site Name", "Site Name", "Site Name")
    lowerReplacement = Array("Wadi Arab (II) PS2 old", "Wadi Arab (II) PS2 old", "")
    upperReplacement = Array("Wadi Arab (II) PS2 new", "Wadi Arab (II) PS2 new", "T3 Gas Station Well")
    
    For s = LBound(sheets) To UBound(sheets)
        Set ws = ThisWorkbook.sheets(sheets(s))
        siteNameCol = GetColumnIndex(ws, "Site Name")
        objectIdCol = GetColumnIndex(ws, "ObjectID")
        buildingNameCol = GetColumnIndex(ws, "Building Name")
        lastRow = ws.Cells(ws.Rows.Count, siteNameCol).End(xlUp).Row
        
        Call ReplaceBasedOnObjectID(ws, CLng(splitObjectID(s)), CStr(columnOfInterest(s)), CStr(searchString(s)), CStr(lowerReplacement(s)), CStr(upperReplacement(s)))
        
    Next s
End Sub


Private Sub FixMismatchedSiteNames()
    ThisWorkbook.Activate
    Dim ws As Worksheet
    Dim incorrectSiteNames As Variant, correctSiteNames As Variant, columnOfInterest As Variant, sheets As Variant, splitObjectID As Variant, objectID As Variant
    Dim i As Long, s As Long
    Dim siteNameCol As Long, objectIdCol As Long, buildingNameCol As Long, lastRow As Long
                
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
    
    'Fix discrepancies between Site Names that were pulled from different sheets. This does not change the source data
    Call ReplaceSiteNames("SystemArchData", incorrectSiteNames, correctSiteNames)
End Sub


Private Sub SortAndConsolidatePanels()
    ThisWorkbook.Activate
    Dim ws As Worksheet
    Set ws = sheets("SystemArchData")
    
    Dim lastRow As Long
    'Assess last row based on last filled row in "Site Name" column
    lastRow = ws.Cells(ws.Rows.Count, GetColumnIndex(ws, "Site Name")).End(xlUp).Row
    
    Dim siteColumn As Long, buildingColumn As Long, roomColumn As Long, panelColumn As Long, commManufacturerColumn As Long, commModelColumn As Long
    siteColumn = GetColumnIndex(ws, "Site Name")
    buildingColumn = GetColumnIndex(ws, "Building Name")
    roomColumn = GetColumnIndex(ws, "Room Name")
    panelColumn = GetColumnIndex(ws, "Panel Name")
    commManufacturerColumn = GetColumnIndex(ws, "Comm Manufacturer Name")
    commModelColumn = GetColumnIndex(ws, "Comm Model Number")
    instCategoryColumn = GetColumnIndex(ws, "Instrument Category")
    instTypeColumn = GetColumnIndex(ws, "Instrument Type")
    
    
    'Sort created table by Site Name, Building Name, Room Name, Panel Name
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.ListObjects("SystemArchTable").ListColumns(siteColumn).Range, Order:=xlAscending
        .SortFields.Add key:=ws.ListObjects("SystemArchTable").ListColumns(buildingColumn).Range, Order:=xlAscending
        .SortFields.Add key:=ws.ListObjects("SystemArchTable").ListColumns(roomColumn).Range, Order:=xlAscending
        .SortFields.Add key:=ws.ListObjects("SystemArchTable").ListColumns(panelColumn).Range, Order:=xlAscending
        .SetRange ws.ListObjects("SystemArchTable").Range
        .header = xlYes
        .Apply
    End With
        
    Dim panelDict As Object
    Set panelDict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    i = 2
    While i <= lastRow ' Assuming the first row is headers
        Dim siteName As String
        Dim buildingName As String
        Dim roomName As String
        Dim panelName As String
        Dim skipThisRow As Boolean
        siteName = ws.Cells(i, siteColumn).Value ' Column A for Site Name
        buildingName = ws.Cells(i, buildingColumn).Value ' Column E for Building Name
        roomName = ws.Cells(i, roomColumn).Value ' Column G for Room Name
        panelName = ws.Cells(i, panelColumn).Value ' Column G for Panel Name
        skipThisRow = (ws.Cells(i, commManufacturerColumn).Value <> "" Or ws.Cells(i, commModelColumn).Value <> "" Or ws.Cells(i, instCategoryColumn).Value <> "" Or ws.Cells(i, instTypeColumn).Value <> "") ' Columns for Comm Equipment and Instruments should be skipped, since they might have duplicates
        
        Dim key As String
        key = siteName & "|" & buildingName & "|" & roomName & "|" & panelName
        
        If key <> "|||" And Not panelDict.exists(key) Then
            panelDict.Add key, i
            i = i + 1
        ElseIf skipThisRow Then
            i = i + 1
        Else
            Dim rowToUpdate As Long
            rowToUpdate = panelDict(key)
            
            Dim j As Long
            For j = panelColumn + 1 To commManufacturerColumn - 1 ' Start from column I (Panel Function) to the last panel column (OIT Model)
                If ws.Cells(i, j).Value <> "" Then
                    ws.Cells(rowToUpdate, j).Value = ws.Cells(i, j).Value
                End If
            Next j
            
            ws.Rows(i).EntireRow.Delete
            lastRow = lastRow - 1
        End If
    Wend
    
End Sub

Private Sub SortAndFill_Muni_Type_Tier()
    ThisWorkbook.Activate
    Dim ws As Worksheet
    Set ws = sheets("SystemArchData")
    
    Dim lastRow As Long
    'Assess last row based on last filled row in "Site Name" column
    lastRow = ws.Cells(ws.Rows.Count, GetColumnIndex(ws, "Site Name")).End(xlUp).Row
    
    Dim siteColumn As Long, muniColumn As Long, typeColumn As Long, tierColumn As Long
    siteColumn = GetColumnIndex(ws, "Site Name")
    muniColumn = GetColumnIndex(ws, "Municipality")
    typeColumn = GetColumnIndex(ws, "Site Name")
    tierColumn = GetColumnIndex(ws, "SCADA Tier")
    
    'Sort created table by Site Name, then fill in redundant blank fields
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.ListObjects("SystemArchTable").ListColumns(siteColumn).Range, Order:=xlAscending
        .SortFields.Add key:=ws.ListObjects("SystemArchTable").ListColumns(muniColumn).Range, Order:=xlAscending
        .SetRange ws.ListObjects("SystemArchTable").Range
        .header = xlYes
        .Apply
    End With
    
    'Fill site names whose only difference is capitalization so they can be populated, capitalization will be fixed later
    For i = 2 To lastRow
        If StrComp(ws.Cells(i, 1).Value, ws.Cells(i - 1, 1).Value, vbTextCompare) = 0 Then
            ws.Cells(i, 1).Value = ws.Cells(i - 1, 1).Value
        End If
    Next i
    
    'Sort created table by Site Name, then fill in redundant blank fields
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.ListObjects("SystemArchTable").ListColumns(siteColumn).Range, Order:=xlAscending
        .SetRange ws.ListObjects("SystemArchTable").Range
        .header = xlYes
        .Apply
    End With
    
    'Fill Municipality, Site Type, and SCADA Tier for all blank rows. Data MUST BE SORTED so that info can be pulled DOWN
    For i = 2 To lastRow
        If ws.Cells(i, 2).Value = "" Then
            If StrComp(ws.Cells(i, 1).Value, ws.Cells(i - 1, 1).Value, vbTextCompare) = 0 Then
                ws.Cells(i, 2).Value = ws.Cells(i - 1, 2).Value
            End If
        End If
        If ws.Cells(i, 3).Value = "" Then
            If StrComp(ws.Cells(i, 1).Value, ws.Cells(i - 1, 1).Value, vbTextCompare) = 0 Then
                ws.Cells(i, 3).Value = ws.Cells(i - 1, 3).Value
            End If
        End If
        If ws.Cells(i, 4).Value = "" Then
            If StrComp(ws.Cells(i, 1).Value, ws.Cells(i - 1, 1).Value, vbTextCompare) = 0 Then
                ws.Cells(i, 4).Value = ws.Cells(i - 1, 4).Value
            End If
        End If
    Next i
    
    'Sort created table by Site Name, Building Name, Room Name, Panel Name
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.ListObjects("SystemArchTable").ListColumns(1).Range, Order:=xlAscending
        .SortFields.Add key:=ws.ListObjects("SystemArchTable").ListColumns(5).Range, Order:=xlAscending
        .SortFields.Add key:=ws.ListObjects("SystemArchTable").ListColumns(6).Range, Order:=xlAscending
        .SortFields.Add key:=ws.ListObjects("SystemArchTable").ListColumns(7).Range, Order:=xlAscending
        .SetRange ws.ListObjects("SystemArchTable").Range
        .header = xlYes
        .Apply
    End With
    
End Sub

Private Sub FixSystemArchDataTypos()
    ThisWorkbook.Activate
    Dim ws As Worksheet
    Set ws = ThisWorkbook.sheets("SystemArchData")
    
    Dim lastRow As Long, lastCol As Long
        lastRow = ws.Cells(ws.Rows.Count, GetColumnIndex(ws, "Site Name")).End(xlUp).Row
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Define your arrays
    Dim capitalizationException As Variant, upperCaseList As Variant, spellCheck_Formatting_Original As Variant, spellCheck_Formatting_Fixed As Variant
    
    capitalizationException = Array("and", "the", "in")
    upperCaseList = Array("ARUBA", "GPRS", "GTO", "GTU", "IPC", "KAC", "NVR", "SCADA", "TP", "WTP", "RO", "ID", "CPU", "RAM", "SSD", "HDD", "PSU", "UPS", "GPU", "PLC", "OIT", "MCC", "PS", "HQ", "RSV")
    removalList = Array("T3 - ", "T3 ", " T3 ", "T2 - ", "T2 ", "T1 - ")
    
    spellCheck_Formatting_Original = Array("(adaseya)", "(alakhdar)", "(Allen-Bradlley)", "(Modawwarah)", ".Old", "Adasyieh", "alandaa", "Alandaa(AA2", "al-arab", "Al-Deek", "ALJafer", "AL-Riyadi", _
        "Al-safi", "arab", "ashrafeh", "Ashrafeyeh", "As-sammar", "Ba'oon", "Birmathkoor", "Chlor", "Chlore", "Chlorification", "Chlorinatetion", "Chlorinatoin", "Chlorintion", "Cholrine", "Collecter", _
        "Coridoor", "Dabbet", "dusconnector", "Filteration", "Gantarah", "Hanoot", "Harfosshie", "Hofa", "Irainbeh", "Iranbeh", "jmail", "Jmail", "joba", "Judtta", "Kofranjeh", "Kraimeh", "Kraymeh", _
        "Kufranja", "Kufranjeh", "lonk", "ma'en", "Ma'in", "mae'en", "Maglis", "Manshiyyah", "Mazra", "Mazra'a", "miyahuma", "Miyahuna-", "miyauhna", "Mughaier", "Muwaqer", "Mwaqar", "Mwaqar-WELL", "Nader", _
        "No'aymeh", "Nu'aymeh", "Numatic", "PANL", "Pupms", "qalaa", "Qantara", "Qanter", "Qweisma", "Rahemh", "Rashdieh", "relayes,", "Reservior", "Reservoird", "Reshah", "Resheh", "resrvoir", "rOOM", "Rusaifa", _
        "Rusaiffeh", "russaifeh", "Russaiffeh", "Russeifa", "Scadapack", "Shammout", "Shmesani", "Shmisani", "Shujairat", "Softstarter", "Starl", "Sttion", "Suweileh", "Swich", "Tabaqat", "termenals", _
        "Unstractured", "Uthainah", "Utheineh", "wier", "Yadoudhe", "Za'atary", "Zatary", "Zmailah", "Pumping", "PS", "PS1", "PS2", "PS3", "PS4", "PS5", "PS6", "PS7", "PS8", "PS9", "PS10", "Well1", "Well2", _
        "Well3", "Well4", "Well5", "Well6", "Well7", "Well8", "Well9", "Well10", "Well11", "Well12", "Well13", "Well14", "Well15", "Well16", "Well17", "Well18", "Well19", "Well20", "Well21", "Well22", "Well23", _
        "Well24", "Well25", "Well26", "Well27", "Well28", "Well29", "Well30", "Well31", "Well32", "Well33", "Well34", "Well35", "Juhfia", "Juhfiyya", "Mindah", "Noa'meh", "Swailemah")
            
    spellCheck_Formatting_Fixed = Array("(Adasiya)", "(Alakhdar)", "(Allen-Bradley)", "(Mudawwara)", "Old", "Adasiya", "Alanda", "Alanda (AA2", "Al-Arab", "Al Deek", "Al Jafer", "Al-Riyadi", _
        "Al-Safi", "Arab", "Ashrafiyeh", "Ashrafiyeh", "As-Sammar", "Ba'oon", "Bir Mathkoor", "Chlorine", "Chlorine", "Chlorination", "Chlorination", "Chlorination", "Chlorination", "Chlorine", "Collector", _
        "Corodor", "Dabet", "disconnector", "Filtration", "Qantarah", "Hanut", "Harfosheya", "Hoffa", "Ireinbeh", "Ireinbeh", "Jamil", "Jamil", "Jobba", "Judetta", "Kfrnga", "Krayymeh", "Krayymeh", "Kfrnga", _
        "Kfrnga", "Link", "Mae'en", "Mae'en", "Mae'en", "Magelis", "Manshyeh", "Mazra'a", "Mazra'a", "Miyahuna", "Miyahuna", "Miyahuna", "Muogaer", "Muwaqqar", "Muwaqqar", "Muwaqqar-WELL", "Naser", "No'aymeh", _
        "No'aymeh", "Pneumatic", "PANEL", "Pumps", "Qala'a", "Qantarah", "Qantarah", "Qwaismeh", "Rahmeh", "Rashdiah", "relays,", "Reservoir", "Reservoir", "Risha", "Risha", "Reservoir", "Room", _
        "Ruseifa", "Ruseifa", "Ruseifa", "Ruseifa", "Ruseifa", "SCADAPack", "Shamout", "Shmeisani", "Shmeisani", "Shujeirat", "Soft Starter", "Starlink", "Station", "Swaileh", "Switch", "Tabagat", "Terminals", _
        "Unstructured", "Uthaina", "Uthaina", "Weir", "Yadoudeh", "Za'tary", "Za'tary", "Zmaileh", "Pump", "Pump Station", "Pump Station 1", "Pump Station 2", "Pump Station 3", "Pump Station 4", _
        "Pump Station 5", "Pump Station 6", "Pump Station 7", "Pump Station 8", "Pump Station 9", "Pump Station 10", "Well 1", "Well 2", "Well 3", "Well 4", "Well 5", "Well 6", "Well 7", "Well 8", "Well 9", _
        "Well 10", "Well 11", "Well 12", "Well 13", "Well 14", "Well 15", "Well 16", "Well 17", "Well 18", "Well 19", "Well 20", "Well 21", "Well 22", "Well 23", "Well 24", "Well 25", "Well 26", "Well 27", "Well 28", "Well 29", "Well 30", "Well 31", "Well 32", "Well 33", "Well 34", "Well 35", "Juhfiyah", "Juhfiyah", "Mandah", "No'aymeh", "Suwelmeh")
            
    Dim cell As Range
    Dim txt As String
    Dim i As Long, j As Long, k As Long, n As Long, m As Long, spellChecked As Long
    Dim words() As String, newText As String
        
    For i = LBound(removalList) To UBound(removalList)
        ws.Cells.Replace What:=removalList(i), Replacement:=""
    Next i
    
    For i = 1 To lastRow
        For j = 1 To lastCol
            Set cell = ws.Cells(i, j)
            If Not IsEmpty(cell.Value) Then
                
                'txt = Trim(cell.Value)
                txt = cell.Value

                ' Split into words and process
                words = Split(txt, " ")
                For k = LBound(words) To UBound(words)
                    Dim w As String
                    w = words(k)
                    
                    spellChecked = 0
                    
                    ' Spell check replacements
                    For m = LBound(spellCheck_Formatting_Original) To UBound(spellCheck_Formatting_Original)
                        If StrComp(w, spellCheck_Formatting_Original(m), vbTextCompare) = 0 Then
                            words(k) = spellCheck_Formatting_Fixed(m)
                            
                            spellChecked = 1
                            Exit For
                        End If
                    Next m

                    If spellChecked <> 1 Then
                        ' Uppercase specific words
                        If IsInArray(UCase(w), upperCaseList) Then
                            words(k) = UCase(w)
                        
                        ' Uppercase entire word if it contains a number
                        ElseIf w Like "*[0-9]*" Then
                            words(k) = UCase(w)
                        
                        ' Lowercase entire word if listed
                        ElseIf IsInArray(LCase(w), capitalizationException) Then
                            words(k) = LCase(w)
                        
                        ' Skip any remaining word that already has a capitalized first letter. It may have been spell-checked already
                        ElseIf Not (Left(w, 1) Like "[A-Z]") Then
                        
                            ' Uppercase first letter with exceptions
                            n = 1
                            Do While n <= Len(w) And Not (Mid(w, n, 1) Like "[A-Za-z]")
                                n = n + 1
                            Loop
                            If n <= Len(w) Then
                                words(k) = Left(w, n - 1) & UCase(Mid(w, n, 1)) & Mid(w, n + 1)
                            End If
                        End If
                    End If
                Next k
                
                newText = Join(words, " ")
                cell.Value = newText
                'cell.Interior.Color = RGB(51, 153, 255) ' Light blue highlight
            End If
        Next j
    Next i

    ws.Columns.AutoFit
    ws.Rows.AutoFit
End Sub
















'These are helper subs and functions that don't do anything on their own
Private Sub ReplaceBasedOnObjectID(ws As Worksheet, splitObjectID As Long, columnOfInterest As String, searchString As String, lowerReplacement As String, upperReplacement As String)
    Dim colIndex As Long, objectIdCol As Long, lastRow As Long, i As Long
    
    colIndex = GetColumnIndex(ws, columnOfInterest)
    objectIdCol = GetColumnIndex(ws, "ObjectID")
    lastRow = ws.Cells(ws.Rows.Count, colIndex).End(xlUp).Row

    For i = 2 To lastRow
        If InStr(1, ws.Cells(i, colIndex).Value, searchString, vbTextCompare) > 0 Then
            objectID = ws.Cells(i, objectIdCol).Value
            If IsNumeric(objectID) Then
                If objectID <= splitObjectID Then
                    ws.Cells(i, colIndex).Value = lowerReplacement
                ElseIf objectID >= splitObjectID Then
                    ws.Cells(i, colIndex).Value = upperReplacement
                End If
            End If
        End If
    Next i
End Sub

Private Sub ReplaceSiteNames(sheetName As String, incorrectNames As Variant, correctNames As Variant)
    Dim ws As Worksheet
    Dim siteCol As Long

    Set ws = sheets(sheetName)
    siteCol = GetColumnIndex(ws, "Site Name")

    If siteCol = 0 Then
        MsgBox "Column 'Site Name' not found in sheet '" & sheetName & "'", vbExclamation
        Exit Sub
    End If

    Dim i As Long
    With ws.Columns(siteCol)
        For i = LBound(incorrectNames) To UBound(incorrectNames)
            .Replace What:=incorrectNames(i), Replacement:=correctNames(i), LookAt:=xlWhole, MatchCase:=False
        Next i
    End With
End Sub

Private Function IsInArray(val As String, arr As Variant) As Boolean
    Dim element As Variant
    For Each element In arr
        If StrComp(val, element, vbTextCompare) = 0 Then
            IsInArray = True
            Exit Function
        End If
    Next element
    IsInArray = False
End Function

Private Function GetColumnIndex(ws As Worksheet, headerName As Variant) As Long
    Dim cell As Range
    Dim headerStr As String
    ' Safely convert headerName to string
    If IsError(headerName) Or IsMissing(headerName) Then
        GetColumnIndex = 0
        Exit Function
    End If

    On Error Resume Next
    headerStr = Trim(CStr(headerName))
    On Error GoTo 0

    If Len(headerStr) = 0 Then
        GetColumnIndex = 0
        Exit Function
    End If

    ' Search through the first row
    For Each cell In ws.Range("A1", ws.Cells(1, ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column))
        If StrComp(Trim(cell.Value), headerStr, vbTextCompare) = 0 Then
            GetColumnIndex = cell.Column
            Exit Function
        End If
    Next cell

    GetColumnIndex = 0 ' Not found
End Function

Private Function ArrayDimSize(arr As Variant, Optional dimension As Variant) As Long
    On Error GoTo ErrorHandler

    ' Default to dimension 1 if not provided or invalid
    If IsMissing(dimension) Or Not IsNumeric(dimension) Then
        dimension = 1
    End If

    ' Check if the input is an array
    If IsArray(arr) Then
        ArrayDimSize = UBound(arr, dimension) - LBound(arr, dimension) + 1
    Else
        ArrayDimSize = -1 ' Not an array
    End If
    Exit Function

ErrorHandler:
    ArrayDimSize = -1 ' Error occurred (e.g., invalid dimension)
End Function




























'These are development assistance tools to help identify problems in the code
Private Sub HighlightMatchingCells()
    ThisWorkbook.Activate
    Dim ws As Worksheet
    Dim searchStrings As Variant
    Set ws = ThisWorkbook.sheets("SystemArchData")
    

    ' Define the array of strings to search for
    searchStrings = Array("(adaseya)", "(alakhdar)", "(Allen-Bradlley)", "(Modawwarah)", "Adasyieh", "alandaa", "Alandaa(AA2", "al-arab", "Al-Deek", "AL-Riyadi", "Al-safi", "arab", _
        "ashrafeh", "Ashrafeyeh", "As-sammar", "Ba’oon", "Birmathkoor", "Chlore", "Chlor", "Chlorification", "Chlorinatetion", "Chlorinatoin", "Chlorintion", "Cholrine", "Collecter", "Coridoor", "Dabbet", _
        "dusconnector", "Filteration", "Gantarah", "Hanoot", "Harfosshie", "Irainbeh", "Iranbeh", "jmail", "joba", "Judtta", "Kofranjeh", "Kraimeh", "Kraymeh", "Kufranja", "Kufranjeh", "lonk", "ma’en", _
        "Ma’in", "mae’en", "Maglis", "Manshiyyah", "Mazra", "Mazra’a", "miyahuma", "Miyahuna-", "miyauhna", "Mughaier", "Muwaqer", "Mwaqar", "Mwaqar-WELL", "No’aymeh", "Nu’aymeh", "Numatic", "PANL", "Pupms", _
        "qalaa", "Qantara", "Qanter", "Qweisma", "Rahemh", "Rashdieh", "relayes,", "Reservior", "Reservoird", "Reshah", "Resheh", "resrvoir", "rOOM", "Rusaifa", "Rusaiffeh", "russaifeh", "Russaiffeh", _
        "Russeifa", "Scadapack", "Shammout", "Shmesani", "Shmisani", "Shujairat", "Softstarter", "Starl", "Sttion", "Suweileh", "Swich", "Tabaqat", "termenals", "Unstractured", "Uthainah", "Utheineh", "wier", _
        "Yadoudhe", "Za’atary", "Zatary", "Zmailah", ".Old")


    Dim lastRow As Long, lastCol As Long
        'Assess last row based on last filled row in "Site Name" column
        lastRow = ws.Cells(ws.Rows.Count, GetColumnIndex(ws, "Site Name")).End(xlUp).Row
        'Assess last column based on last filled row in first row
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    Dim cell As Range
    Dim i As Long, j As Long, k As Long, n As Long, m As Long
    Dim words() As String, newText As String

    For i = 1 To lastRow
        For j = 1 To lastCol
            Set cell = ws.Cells(i, j)
            If Not IsEmpty(cell.Value) Then
                Dim txt As String
                txt = Trim(cell.Value)
                
                'Replace line breaks with spaces
                txt = Replace(txt, vbCrLf, " ")
                txt = Replace(txt, vbCr, " ")
                txt = Replace(txt, vbLf, " ")

                ' Remove double spaces
                Do While InStr(txt, "  ") > 0
                    txt = Replace(txt, "  ", " ")
                Loop

                ' Split into words and process
                words = Split(txt, " ")
                For k = LBound(words) To UBound(words)
                    n = 0
                    Dim w As String
                    w = words(k)

                    ' Spell check
                    For m = LBound(searchStrings) To UBound(searchStrings)
                        If StrComp(w, searchStrings(m), vbTextCompare) = 0 Then
                            'cell.Interior.Color = RGB(255, 255, 0) ' Highlight in yellow
                            n = 1 ' Set n to 1 to indicate this cell is done
                            
                            If StrComp(w, searchStrings(m), vbBinaryCompare) = 0 Then
                                cell.Interior.Color = RGB(255, 0, 0) ' Highlight in red
                                n = 1 ' Set n to 1 to indicate this cell is done
                                Exit For ' Exit inner loop once a match is found
                            End If

                            Exit For ' Exit inner loop once a match is found
                        End If
                    Next m
                    If n = 1 Then
                        Exit For ' Exit outer loop if n has been set
                    End If
                Next k
            End If
        Next j
    Next i
    
    'Application.ScreenUpdating = True
End Sub

Private Sub HighlightSpellingErrors()
    'This is not for repeated use. It is intended to help assemble a dictionary of misspelled words and check the function of the other parts of the script
    ThisWorkbook.Activate
    Dim ws As Worksheet
    Set ws = ThisWorkbook.sheets("SystemArchData")

    Dim typoSheet As Worksheet
    Dim typoList As Collection
    Set typoList = New Collection

    Dim cell As Range
    Dim rng As Range
    Dim token As Variant
    Dim tokens() As String
    Dim cleanedToken As String
    Dim i As Long

    ' Define the range to check
    Set rng = ws.UsedRange

    For Each cell In rng
        If Not IsEmpty(cell.Value) And VarType(cell.Value) = vbString Then
            tokens = Split(cell.Value, " ")
            Dim hasError As Boolean
            hasError = False

            For Each token In tokens
                If Not Application.CheckSpelling(word:=token) Then
                    hasError = True
                    On Error Resume Next
                    typoList.Add token, CStr(token)
                    On Error GoTo 0
                End If
            Next token

            If hasError Then
                cell.Interior.Color = RGB(255, 200, 200) ' Light red highlight
            Else
                cell.Interior.ColorIndex = xlNone ' Clear previous highlights
            End If
        End If
    Next cell

    ' Create a new sheet for typos
    'Set typoSheet = ThisWorkbook.Sheets("Spelling Errors")
    If typoSheet Is Nothing Then
        Set typoSheet = ThisWorkbook.sheets.Add(After:=ThisWorkbook.sheets("SystemArchData"))
        typoSheet.Name = "Spelling Errors"
    Else
        typoSheet.Cells.Clear
    End If

    typoSheet.Cells(1, 1).Value = "Misspelled Words"
    For i = 1 To typoList.Count
        typoSheet.Cells(i + 1, 1).Value = typoList(i)
    Next i

    MsgBox "Spellcheck complete. Misspelled words highlighted and listed on 'Spelling Errors' sheet.", vbInformation
End Sub


Private Sub HighlightMatchingSiteNames()
    Dim siteSheet As Worksheet
    Dim systemSheet As Worksheet
    Dim siteDict As Object
    Dim lastRow As Long
    Dim siteCol As Long
    Dim systemCol As Long
    Dim cell As Range
    Dim headerCell As Range

    ' Set references to the sheets
    Set siteSheet = ThisWorkbook.sheets("Site2")
    Set systemSheet = ThisWorkbook.sheets("SystemArchData2")
    
    ' Create a dictionary to store site names
    Set siteDict = CreateObject("Scripting.Dictionary")
    
    ' Find the "Site Name" column in the "Site" sheet
    For Each headerCell In siteSheet.Rows(1).Cells
        If Trim(headerCell.Value) = "Site Name" Then
            siteCol = headerCell.Column
            Exit For
        End If
    Next headerCell
    
    If siteCol = 0 Then
        MsgBox "Site Name column not found in 'Site' sheet.", vbExclamation
        Exit Sub
    End If
    
    ' Build the site name library
    lastRow = siteSheet.Cells(siteSheet.Rows.Count, siteCol).End(xlUp).Row
    For Each cell In siteSheet.Range(siteSheet.Cells(2, siteCol), siteSheet.Cells(lastRow, siteCol))
        If Not siteDict.exists(cell.Value) And cell.Value <> "" Then
            siteDict.Add cell.Value, True
        End If
    Next cell
    
    ' Find the "Site Name" column in the "SystemArchData" sheet
    For Each headerCell In systemSheet.Rows(1).Cells
        If Trim(headerCell.Value) = "Site Name" Then
            systemCol = headerCell.Column
            Exit For
        End If
    Next headerCell
    
    If systemCol = 0 Then
        MsgBox "Site Name column not found in 'SystemArchData' sheet.", vbExclamation
        Exit Sub
    End If
    
    ' Highlight matching cells
    lastRow = systemSheet.Cells(systemSheet.Rows.Count, systemCol).End(xlUp).Row
    For Each cell In systemSheet.Range(systemSheet.Cells(2, systemCol), systemSheet.Cells(lastRow, systemCol))
        If siteDict.exists(cell.Value) Then
            cell.Interior.Color = RGB(144, 238, 144) ' Light green
        End If
    Next cell
End Sub

Private Sub HighlightDifferencesBetweenSheets()
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim dict1 As Object, dict2 As Object
    Dim lastRow1 As Long, lastRow2 As Long
    Dim lastCol1 As Long, lastCol2 As Long
    Dim i As Long, j As Long
    Dim rowKey As String
    Dim key As Variant
    Dim delimiter As String: delimiter = "|"

    Set ws1 = ThisWorkbook.sheets("SystemArchData")
    Set ws2 = ThisWorkbook.sheets("SystemArchData2")
    Set dict1 = CreateObject("Scripting.Dictionary")
    Set dict2 = CreateObject("Scripting.Dictionary")

    ' Find last used row and column in both sheets
    lastRow1 = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
    lastCol1 = ws1.Cells(1, ws1.Columns.Count).End(xlToLeft).Column
    lastRow2 = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row
    lastCol2 = ws2.Cells(1, ws2.Columns.Count).End(xlToLeft).Column

    ' Build dictionary for Sheet1
    For i = 1 To lastRow1
        rowKey = ""
        For j = 1 To 8 '1 To lastCol1
            rowKey = rowKey & ws1.Cells(i, j).Text & delimiter
        Next j
        dict1(rowKey) = i ' Store row number for highlighting
    Next i

    ' Build dictionary for Sheet2
    For i = 1 To lastRow2
        rowKey = ""
        For j = 2 To 9 '1 To lastCol2
            rowKey = rowKey & ws2.Cells(i, j).Text & delimiter
        Next j
        dict2(rowKey) = i
    Next i

    ' Highlight rows in Sheet1 not in Sheet2
    For Each key In dict1.Keys
        If Not dict2.exists(key) Then
            ws1.Rows(dict1(key)).Interior.Color = RGB(255, 255, 0) ' Yellow
        End If
    Next key

    ' Highlight rows in Sheet2 not in Sheet1
    For Each key In dict2.Keys
        If Not dict1.exists(key) Then
            ws2.Rows(dict2(key)).Interior.Color = RGB(255, 255, 0) ' Yellow
        End If
    Next key

    MsgBox "Comparison complete. Differences highlighted in yellow.", vbInformation
End Sub


Private Sub HighlightUnmatchedSiteNames()
    ThisWorkbook.Activate
    Dim sheetNames As Variant
    'sheetNames = Array("Site", "Building", "Room", "Control Panel", "PLC", "OIT", "Communication Equipment")
    sheetNames = Array("RoomFixes")

    Dim siteSheet As Worksheet
    Set siteSheet = ThisWorkbook.sheets("Site")

    Dim siteColIndex As Long
    siteColIndex = GetColumnIndex(siteSheet, "Site Name")
    If siteColIndex = 0 Then
        MsgBox "Column 'Site Name' not found in 'Site' sheet."
        Exit Sub
    End If

    ' Create a dictionary of valid site names
    Dim siteDict As Object
    Set siteDict = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long
    lastRow = siteSheet.Cells(siteSheet.Rows.Count, siteColIndex).End(xlUp).Row

    Dim i As Long
    For i = 2 To lastRow ' Assuming row 1 is header
        Dim siteValue As String
        siteValue = Trim(siteSheet.Cells(i, siteColIndex).Value)
        If Len(siteValue) > 0 Then
            siteDict(siteValue) = True
        End If
    Next i

    ' Loop through each sheet in the array
    Dim ws As Worksheet
    Dim sheetName As Variant
    For Each sheetName In sheetNames
        Set ws = ThisWorkbook.sheets(sheetName)
        Dim colIndex As Long
        colIndex = GetColumnIndex(ws, "Site Name")
        If colIndex = 0 Then GoTo NextSheet

        lastRow = ws.Cells(ws.Rows.Count, colIndex).End(xlUp).Row
        For i = 2 To lastRow
            siteValue = Trim(ws.Cells(i, colIndex).Value)
            If Len(siteValue) > 0 Then
                If Not siteDict.exists(siteValue) Then
                    ws.Cells(i, colIndex).Interior.Color = RGB(255, 199, 206) ' Light red
                    'ws.Cells(i, colIndex).Interior.Color = RGB(255, 255, 255) ' None
                End If
            End If
        Next i
NextSheet:
    Next sheetName

    MsgBox "Highlighting complete."
End Sub

Private Sub CheckOldDictionary()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim siteNamesRange As Range
    Dim cell As Range
    Dim searchTerms As Variant
    Dim i As Long, j As Long
    Dim matchFound As Boolean
    Dim colIndex As Long

    ' Define the worksheet
    Set ws = ThisWorkbook.sheets("Site")

    ' Define the search terms
    searchTerms = Array("Aqeb k104c well", "As-sammar old", "Check gate 24 (CG)", "Cg 25", "Cg 26", "Cg 27", "Cg 29", "Cg 30", "Cg31", "Dabouq Reservoir", "H", "Mudawwara Border PS & RSV", _
            "Mudawwara Balad PS & RSV", "MIZFFER", "Miyahuna Main SCADA Center", "Pump Station 11 and check gate 3", "RAHA PUMP STATION ", "Rashdieh ps", "T3 Abu Siah Booster", "T3 Abu Sayah Booster", _
            "T3 Mazra A well11", "T3 Booster 11", "T3 10 Well", "T3 17 Well", "To 21 and Check Gate 8", "TO 33", "Wadi Arab Pumping Station 1 New-Old")

    ' Find the last row in the "Site Name" column
    colIndex = GetColumnIndex(ws, "Site Name")
    lastRow = ws.Cells(ws.Rows.Count, colIndex).End(xlUp).Row


    ' Loop through each cell in the Site Name column
    For i = 2 To lastRow
        matchFound = False
        For j = LBound(searchTerms) To UBound(searchTerms)
            If StrComp(ws.Cells(i, colIndex).Value, searchTerms(j), vbTextCompare) = 0 Then
                matchFound = True
                Exit For
            End If
        Next j

        ' Highlight the cell if a match is found
        If matchFound Then
            ws.Cells(i, colIndex).Interior.Color = RGB(255, 255, 0) ' Yellow highlight
        End If
    Next i
    MsgBox "Search complete. Matches highlighted.", vbInformation
End Sub













