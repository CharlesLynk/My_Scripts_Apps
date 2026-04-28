VBA Module (Reusable Formula Freeze/Refresh)
Option Explicit

' Default wrapper target for a button-assigned macro in this workbook.
' You can use either a structured reference like TableName[Column Name]
' or a normal worksheet/name reference like Sheet1!D2:D500 or MyNamedRange.
Private Const DEFAULT_TARGET_REFERENCE As String = "MASTERV2_MASK_DB[ADDRESS]"

' Hidden worksheet used to store one formula template per target reference.
' This lets the refresh macro restore formulas after a target range has been frozen to values.
Private Const FORMULA_STATE_SHEET_NAME As String = "__FrozenFormulaState"

' Button-friendly wrapper for this workbook.
Public Sub RefreshDefaultFrozenColumn()
   RefreshFrozenFormulaTarget DEFAULT_TARGET_REFERENCE
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
Public Sub RefreshFrozenFormulaTarget(ByVal targetReference As String)
   Dim targetRng As Range
   Dim formulaText As String
   Dim storeKey As String

   storeKey = BuildFormulaStoreKey(targetReference)

   Set targetRng = ResolveTargetRange(targetReference)
   If targetRng Is Nothing Then
      MsgBox "Target range not found or empty: " & targetReference, vbExclamation
      Exit Sub
   End If

   formulaText = ReadFormulaTemplate(targetRng, storeKey)
   If Len(formulaText) = 0 Then
      MsgBox "No stored formula template found. Put formula back in first target cell, then run again.", vbExclamation
      Exit Sub
   End If

   WriteFormulaTemplate storeKey, formulaText
   RestoreColumnFormulasIfNeeded targetRng, formulaText

   Application.ScreenUpdating = False
   Application.EnableEvents = False
   On Error GoTo SafeExit

   Application.CalculateFull
   WaitForCalcDone
   targetRng.Value2 = targetRng.Value2

SafeExit:
   Application.EnableEvents = True
   Application.ScreenUpdating = True

   If Err.Number <> 0 Then
      MsgBox "Refresh failed: " & Err.Description, vbExclamation
   End If
End Sub

' Convenience wrapper if you prefer passing table and column separately in VBA.
Public Sub RefreshFrozenFormulaColumn(ByVal tableName As String, ByVal targetColumnName As String)
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
   If targetRng.Cells(1, 1).HasFormula Then
      On Error Resume Next
      ReadFormulaTemplate = targetRng.Cells(1, 1).Formula2
      If Len(ReadFormulaTemplate) = 0 Then ReadFormulaTemplate = targetRng.Cells(1, 1).Formula
      On Error GoTo 0
   Else
      ReadFormulaTemplate = ReadStoredFormulaTemplate(storeKey)
   End If
End Function

' Restores formulas across the entire target range if any cells were frozen to values.
Private Sub RestoreColumnFormulasIfNeeded(ByVal targetRng As Range, ByVal formulaText As String)
   If IsEntireRangeFormulaBacked(targetRng) Then Exit Sub

   On Error Resume Next
   targetRng.Formula2 = formulaText
   If Err.Number <> 0 Then
      Err.Clear
      targetRng.Formula = formulaText
   End If
   On Error GoTo 0
End Sub

' Detects mixed state after an interrupted run so the next refresh can self-heal.
Private Function IsEntireRangeFormulaBacked(ByVal rng As Range) As Boolean
   Dim c As Range

   For Each c In rng.Cells
      If Not c.HasFormula Then
         IsEntireRangeFormulaBacked = False
         Exit Function
      End If
   Next c

   IsEntireRangeFormulaBacked = True
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

   Set ws = GetFormulaStateSheet()
   Set keyCell = FindStoredFormulaKeyCell(ws, storeKey)

   If keyCell Is Nothing Then
      Set keyCell = ws.Cells(ws.Rows.Count, 1).End(xlUp).Offset(1, 0)
      If Len(ws.Range("A1").Value2) = 0 Then Set keyCell = ws.Range("A1")
      keyCell.Value2 = storeKey
   End If

   keyCell.Offset(0, 1).Value2 = formulaText
End Sub

' Reads the stored template for a specific target reference.
Private Function ReadStoredFormulaTemplate(ByVal storeKey As String) As String
   Dim ws As Worksheet
   Dim keyCell As Range

   Set ws = GetFormulaStateSheet()
   Set keyCell = FindStoredFormulaKeyCell(ws, storeKey)

   If Not keyCell Is Nothing Then
      ReadStoredFormulaTemplate = CStr(keyCell.Offset(0, 1).Value2)
   End If
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