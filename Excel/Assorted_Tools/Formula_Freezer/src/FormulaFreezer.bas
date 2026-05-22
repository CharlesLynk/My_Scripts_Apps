Attribute VB_Name = "FormulaFreezer"
' VBA Module (Reusable Formula Freeze/Refresh)
Option Explicit

' Hidden worksheet used to store one formula template per target reference.
' This lets the refresh macro restore formulas after a target range has been frozen to values.
Private Const FORMULA_STATE_SHEET_NAME As String = "__FrozenFormulaState"

' Button-friendly wrapper that discovers targets from the hidden state sheet.
Public Sub RefreshDefaultFrozenColumn()
   RefreshFrozenFormulaTargetsFromState
End Sub

' Discovers stored targets from __FrozenFormulaState!A:A and applies the adjacent formula in column B.
Private Sub RefreshFrozenFormulaTargetsFromState()
   Dim ws As Worksheet
   Dim lastRow As Long
   Dim i As Long
   Dim targetRef As String
   Dim formulaText As String
   Dim failed As String
   Dim processedCount As Long
   Dim totalCount As Long

   Set ws = GetFormulaStateSheet()
   lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

   If lastRow = 1 And Len(Trim$(CStr(ws.Cells(1, 1).Value2))) = 0 Then
      MsgBox "No stored frozen formula targets were found on " & FORMULA_STATE_SHEET_NAME & ".", vbExclamation
      Exit Sub
   End If

   For i = 1 To lastRow
      targetRef = Trim$(CStr(ws.Cells(i, 1).Value2))
      If Len(targetRef) > 0 Then totalCount = totalCount + 1
   Next i

   If totalCount = 0 Then
      MsgBox "No stored frozen formula targets were found on " & FORMULA_STATE_SHEET_NAME & ".", vbExclamation
      Exit Sub
   End If

   For i = 1 To lastRow
      targetRef = Trim$(CStr(ws.Cells(i, 1).Value2))
      If Len(targetRef) > 0 Then
         formulaText = ReadStoredFormulaFromCell(ws.Cells(i, 2))
         processedCount = processedCount + 1
         SetProgressStatus "Refreshing frozen formulas from state: target " & CStr(processedCount) & " of " & CStr(totalCount) & "..."

         If Not RefreshFrozenFormulaTargetCore(targetRef, False, formulaText) Then
            If Len(failed) > 0 Then failed = failed & vbCrLf
            failed = failed & targetRef
         End If
      End If
   Next i

   ClearProgressStatus

   If Len(failed) = 0 Then
      MsgBox "Refreshed " & CStr(processedCount) & " frozen formula target(s) from " & FORMULA_STATE_SHEET_NAME & ".", vbInformation
   Else
      MsgBox "Some state-sheet targets failed to refresh:" & vbCrLf & failed, vbExclamation
   End If
End Sub

' Batch-friendly wrapper for multiple targets.
' Delimiters supported in targetReferences:
' 1. Pipe (|)
' 2. Semicolon (;)
' 3. New line
Private Sub RefreshFrozenFormulaTargets(ByVal targetReferences As String)
   Dim normalized As String
   Dim parts() As String
   Dim i As Long
   Dim targetRef As String
   Dim failed As String
   Dim processedCount As Long
   Dim totalCount As Long

   normalized = Replace(targetReferences, vbCrLf, "|")
   normalized = Replace(normalized, vbLf, "|")
   normalized = Replace(normalized, ";", "|")

   parts = Split(normalized, "|")

   For i = LBound(parts) To UBound(parts)
      If Len(Trim$(parts(i))) > 0 Then totalCount = totalCount + 1
   Next i

   For i = LBound(parts) To UBound(parts)
      targetRef = Trim$(parts(i))
      If Len(targetRef) > 0 Then
         processedCount = processedCount + 1
         SetProgressStatus "Refreshing frozen formulas: target " & CStr(processedCount) & " of " & CStr(totalCount) & "..."
         If Not RefreshFrozenFormulaTargetCore(targetRef, False) Then
            If Len(failed) > 0 Then failed = failed & vbCrLf
            failed = failed & targetRef
         End If
      End If
   Next i

   If processedCount = 0 Then
      ClearProgressStatus
      MsgBox "No target references were provided.", vbExclamation
      Exit Sub
   End If

   ClearProgressStatus

   If Len(failed) = 0 Then
      MsgBox "Refreshed " & CStr(processedCount) & " frozen formula target(s).", vbInformation
   Else
      MsgBox "Some targets failed to refresh:" & vbCrLf & failed, vbExclamation
   End If
End Sub

' Generic entry point.
' Use this for any calculated target that should:
' 1. Restore formulas if the column is frozen or partially frozen.
' 2. Recalculate the workbook.
' 3. Freeze the target column back to values.
'
' Supported target formats:
' 1. TableName[Column Name]
' 2. SheetName!A2:A500
' 3. Workbook-level named ranges
'
' Important: this assumes the target range uses one uniform formula pattern.
' Multiple different formulas within the same target range are not supported.
Private Sub RefreshFrozenFormulaTarget(ByVal targetReference As String)
   Call RefreshFrozenFormulaTargetCore(targetReference, True)
End Sub

' Single-target refresh implementation with optional user prompts.
Private Function RefreshFrozenFormulaTargetCore(ByVal targetReference As String, ByVal showMessages As Boolean, Optional ByVal storedFormulaOverride As String = "") As Boolean
   Dim targetRng As Range
   Dim formulaText As String
   Dim storeKey As String

   storeKey = BuildFormulaStoreKey(targetReference)

   Set targetRng = ResolveTargetRange(targetReference)
   If targetRng Is Nothing Then
      If IsStructuredTargetReferenceEmpty(targetReference) Then
         RefreshFrozenFormulaTargetCore = True
         Exit Function
      End If
      If showMessages Then MsgBox "Target range not found or empty: " & targetReference, vbExclamation
      Exit Function
   End If

   If Len(storedFormulaOverride) > 0 Then
      formulaText = NormalizeFormulaTemplateText(storedFormulaOverride)
   Else
      formulaText = ReadFormulaTemplate(targetRng, storeKey)
   End If
   If Len(formulaText) = 0 Then
      If IsRangeBlank(targetRng) Then
         RefreshFrozenFormulaTargetCore = True
         Exit Function
      End If
      If showMessages Then MsgBox "No stored formula template found. Put formula back in first target cell, then run again.", vbExclamation
      Exit Function
   End If

   Application.ScreenUpdating = False
   Application.EnableEvents = False
   On Error GoTo SafeExit

   WriteFormulaTemplate storeKey, formulaText
   SetProgressStatus "Restoring formulas for " & targetReference & "..."
   RestoreColumnFormulasIfNeeded targetRng, formulaText

   ReleaseStatusBarForCalculation
   Application.CalculateFull
   WaitForCalcDone
   SetProgressStatus "Freezing calculated values for " & targetReference & "..."
   targetRng.Value2 = targetRng.Value2

SafeExit:
   Application.EnableEvents = True
   Application.ScreenUpdating = True

   If Err.Number <> 0 Then
      If showMessages Then MsgBox "Refresh failed: " & Err.Description, vbExclamation
      Exit Function
   End If

   RefreshFrozenFormulaTargetCore = True
End Function

' Convenience wrapper if you prefer passing table and column separately in VBA.
Private Sub RefreshFrozenFormulaColumn(ByVal tableName As String, ByVal targetColumnName As String)
   RefreshFrozenFormulaTarget BuildStructuredTargetReference(tableName, targetColumnName)
End Sub

' Finds a table anywhere in the workbook by name.
Private Function FindWorkbookTable(ByVal tableName As String) As ListObject
   Dim ws As Worksheet
   Dim lo As ListObject

   For Each ws In ThisWorkbook.Worksheets
      For Each lo In ws.ListObjects
         If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
            Set FindWorkbookTable = lo
            Exit Function
         End If
      Next lo
   Next ws
End Function

' Resolves either a structured table-column reference or a standard range/name reference.
Private Function ResolveTargetRange(ByVal targetReference As String) As Range
   Dim lo As ListObject
   Dim tableName As String
   Dim columnName As String
   Dim bangPos As Long
   Dim wsName As String
   Dim localAddress As String

   If TryParseStructuredTargetReference(targetReference, tableName, columnName) Then
      Set lo = FindWorkbookTable(tableName)
      If Not lo Is Nothing Then
         On Error Resume Next
         Set ResolveTargetRange = lo.ListColumns(columnName).DataBodyRange
         On Error GoTo 0
      End If
      Exit Function
   End If

   bangPos = InStr(1, targetReference, "!", vbTextCompare)
   If bangPos > 0 Then
      wsName = Left$(targetReference, bangPos - 1)
      localAddress = Mid$(targetReference, bangPos + 1)
      wsName = Replace(wsName, "'", "")

      On Error Resume Next
      Set ResolveTargetRange = ThisWorkbook.Worksheets(wsName).Range(localAddress)
      On Error GoTo 0
      If Not ResolveTargetRange Is Nothing Then Exit Function
   End If

   On Error Resume Next
   Set ResolveTargetRange = ThisWorkbook.Names(targetReference).RefersToRange
   On Error GoTo 0
   If Not ResolveTargetRange Is Nothing Then Exit Function

   On Error Resume Next
   Set ResolveTargetRange = Application.Range(targetReference)
   On Error GoTo 0
End Function

' Builds a structured reference string from separate table/column inputs.
Private Function BuildStructuredTargetReference(ByVal tableName As String, ByVal columnName As String) As String
   BuildStructuredTargetReference = tableName & "[" & columnName & "]"
End Function

' Parses references of the form TableName[Column Name].
Private Function TryParseStructuredTargetReference(ByVal targetReference As String, ByRef tableName As String, ByRef columnName As String) As Boolean
   Dim openPos As Long
   Dim closePos As Long

   openPos = InStr(1, targetReference, "[", vbTextCompare)
   closePos = InStrRev(targetReference, "]", -1, vbTextCompare)

   If openPos <= 1 Or closePos <= openPos Then Exit Function

   tableName = Trim$(Left$(targetReference, openPos - 1))
   columnName = Trim$(Mid$(targetReference, openPos + 1, closePos - openPos - 1))

   TryParseStructuredTargetReference = Len(tableName) > 0 And Len(columnName) > 0
End Function

' Uses the live formula when available, otherwise falls back to the stored template.
Private Function ReadFormulaTemplate(ByVal targetRng As Range, ByVal storeKey As String) As String
   Dim formulaCell As Range

   Set formulaCell = FindFirstFormulaCell(targetRng)

   If Not formulaCell Is Nothing Then
      On Error Resume Next
      ReadFormulaTemplate = formulaCell.Formula2
      If Len(ReadFormulaTemplate) = 0 Then ReadFormulaTemplate = formulaCell.Formula
      On Error GoTo 0
   Else
      ReadFormulaTemplate = ReadStoredFormulaTemplate(storeKey)
   End If
End Function

' Returns the first formula cell in a range, if one exists.
Private Function FindFirstFormulaCell(ByVal rng As Range) As Range
   Dim formulaCells As Range

   On Error Resume Next
   Set formulaCells = rng.SpecialCells(xlCellTypeFormulas)
   On Error GoTo 0

   If Not formulaCells Is Nothing Then Set FindFirstFormulaCell = formulaCells.Cells(1, 1)
End Function

' Restores formulas across the entire target range if any cells were frozen to values.
Private Sub RestoreColumnFormulasIfNeeded(ByVal targetRng As Range, ByVal formulaText As String)
   Dim formulaToApply As String

   If IsEntireRangeFormulaBacked(targetRng) Then Exit Sub

   formulaToApply = NormalizeFormulaTemplateText(formulaText)
   If Len(formulaToApply) = 0 Then
      Err.Raise vbObjectError + 501, "RestoreColumnFormulasIfNeeded", "Stored formula template was empty after normalization."
   End If

   On Error Resume Next
   targetRng.Formula2 = formulaToApply
   If Err.Number <> 0 Then
      Err.Clear
      targetRng.Formula = formulaToApply
   End If
   On Error GoTo 0

    If Not targetRng.Cells(1, 1).HasFormula Then
      Err.Raise vbObjectError + 502, "RestoreColumnFormulasIfNeeded", _
                "Stored template could not be applied as a formula: " & Left$(formulaToApply, 120)
   End If
End Sub

' Detects mixed state after an interrupted run so the next refresh can self-heal.
Private Function IsEntireRangeFormulaBacked(ByVal rng As Range) As Boolean
   Dim c As Range

   For Each c In rng.Cells
      If c.HasFormula Then
         ' Formula-backed cell.
      ElseIf IsError(c.Value2) Then
         IsEntireRangeFormulaBacked = False
         Exit Function
      ElseIf IsEmpty(c.Value2) Then
         ' Treat blank cells as acceptable so sparse ranges do not force a rewrite.
      Else
         IsEntireRangeFormulaBacked = False
         Exit Function
      End If
   Next c

   IsEntireRangeFormulaBacked = True
End Function

' True when the requested structured target exists but has no data rows.
Private Function IsStructuredTargetReferenceEmpty(ByVal targetReference As String) As Boolean
   Dim lo As ListObject
   Dim tableName As String
   Dim columnName As String
   Dim targetCol As ListColumn

   If Not TryParseStructuredTargetReference(targetReference, tableName, columnName) Then Exit Function

   Set lo = FindWorkbookTable(tableName)
   If lo Is Nothing Then Exit Function

   On Error Resume Next
   Set targetCol = lo.ListColumns(columnName)
   On Error GoTo 0

   If targetCol Is Nothing Then Exit Function

   IsStructuredTargetReferenceEmpty = targetCol.DataBodyRange Is Nothing
End Function

' True when no cell in the range contains a value or formula result.
Private Function IsRangeBlank(ByVal rng As Range) As Boolean
   On Error Resume Next
   IsRangeBlank = (WorksheetFunction.CountA(rng) = 0)
   On Error GoTo 0
End Function

' Waits until Excel reports that calculation has finished before freezing values.
Private Sub WaitForCalcDone()
   Do While Application.CalculationState <> xlDone
      DoEvents
   Loop
End Sub

' Stores one reusable formula template per target reference.
Private Sub WriteFormulaTemplate(ByVal storeKey As String, ByVal formulaText As String)
   Dim ws As Worksheet
   Dim keyCell As Range
   Dim formulaCell As Range

   Set ws = GetFormulaStateSheet()
   Set keyCell = FindStoredFormulaKeyCell(ws, storeKey)

   If keyCell Is Nothing Then
      Set keyCell = ws.Cells(ws.Rows.Count, 1).End(xlUp).Offset(1, 0)
      If Len(ws.Range("A1").Value2) = 0 Then Set keyCell = ws.Range("A1")
      keyCell.Value2 = storeKey
   End If

   Set formulaCell = keyCell.Offset(0, 1)
   formulaCell.NumberFormat = "@"
   If Len(formulaText) > 0 Then
      If Left$(formulaText, 1) = "'" Then
         formulaCell.Value2 = formulaText
      Else
         formulaCell.Value2 = "'" & formulaText
      End If
   Else
      formulaCell.Value2 = formulaText
   End If
End Sub

' Reads the stored template for a specific target reference.
Private Function ReadStoredFormulaTemplate(ByVal storeKey As String) As String
   Dim ws As Worksheet
   Dim keyCell As Range
   Dim formulaCell As Range

   Set ws = GetFormulaStateSheet()
   Set keyCell = FindStoredFormulaKeyCell(ws, storeKey)

   If Not keyCell Is Nothing Then
      Set formulaCell = keyCell.Offset(0, 1)
      ReadStoredFormulaTemplate = ReadStoredFormulaFromCell(formulaCell)
   End If
End Function

' Reads one stored formula cell and normalizes text payloads like '=... and ''=....
Private Function ReadStoredFormulaFromCell(ByVal formulaCell As Range) As String
   Dim rawValue As Variant
   Dim formulaText As String

   If formulaCell Is Nothing Then Exit Function

   If formulaCell.HasFormula Then
      On Error Resume Next
      formulaText = formulaCell.Formula2
      If Len(formulaText) = 0 Then formulaText = formulaCell.Formula
      On Error GoTo 0
   Else
      rawValue = formulaCell.Value2
      If Not IsError(rawValue) Then formulaText = CStr(rawValue)
   End If

   ReadStoredFormulaFromCell = NormalizeFormulaTemplateText(formulaText)
End Function

' Accepts formula templates stored as text in either "=..." or "'=..." format.
Private Function NormalizeFormulaTemplateText(ByVal formulaText As String) As String
   Dim i As Long

   i = 1
   Do While i <= Len(formulaText) And Mid$(formulaText, i, 1) = "'"
      i = i + 1
   Loop

   If i > 1 Then
      If i <= Len(formulaText) Then
         If Mid$(formulaText, i, 1) = "=" Then
            NormalizeFormulaTemplateText = Mid$(formulaText, i)
            Exit Function
         End If
      End If
   End If

   NormalizeFormulaTemplateText = formulaText
End Function

' Creates the lookup key used in the hidden state sheet.
Private Function BuildFormulaStoreKey(ByVal targetReference As String) As String
   BuildFormulaStoreKey = UCase$(Trim$(targetReference))
End Function

' Finds an existing stored formula record on the hidden state sheet.
Private Function FindStoredFormulaKeyCell(ByVal ws As Worksheet, ByVal storeKey As String) As Range
   Dim lastRow As Long
   Dim keys As Range

   lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
   If lastRow < 1 Then Exit Function
   If lastRow = 1 And Len(ws.Range("A1").Value2) = 0 Then Exit Function

   Set keys = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 1))
   Set FindStoredFormulaKeyCell = keys.Find(What:=storeKey, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
End Function

' Creates or returns the hidden storage sheet used by this utility module.
Private Function GetFormulaStateSheet() As Worksheet
   Dim ws As Worksheet

   On Error Resume Next
   Set ws = ThisWorkbook.Worksheets(FORMULA_STATE_SHEET_NAME)
   On Error GoTo 0

   If ws Is Nothing Then
      Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
      ws.Name = FORMULA_STATE_SHEET_NAME
      ws.Visible = xlSheetVeryHidden
   End If

   Set GetFormulaStateSheet = ws
End Function

' Shows lightweight progress feedback without needing a custom form.
Private Sub SetProgressStatus(ByVal messageText As String)
   On Error Resume Next
   Application.Cursor = xlWait
   Application.StatusBar = messageText
   DoEvents
   On Error GoTo 0
End Sub

' Releases the status bar so Excel can show its own calculation progress UI.
Private Sub ReleaseStatusBarForCalculation()
   On Error Resume Next
   Application.StatusBar = False
   DoEvents
   On Error GoTo 0
End Sub

' Restores the default Excel UI after work completes.
Private Sub ClearProgressStatus()
   On Error Resume Next
   Application.StatusBar = False
   Application.Cursor = xlDefault
   DoEvents
   On Error GoTo 0
End Sub