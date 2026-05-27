Attribute VB_Name = "FormulaFreezer"
' VBA Module (Reusable Formula Freeze/Refresh)
Option Explicit

' Hidden worksheet used to store one formula template per target reference.
' This lets the refresh macro restore formulas after a target range has been frozen to values.
Private Const FORMULA_STATE_SHEET_NAME As String = "__FrozenFormulaState"

' __FrozenFormulaState layout:
' A = target reference, B = formula template, C = group tags (optional).
Private Const STATE_COL_TARGET_REFERENCE As Long = 1
Private Const STATE_COL_FORMULA_TEMPLATE As Long = 2
Private Const STATE_COL_GROUP_TAGS As Long = 3
Private Const CANCEL_VK_ESCAPE As Long = 27
Private Const ERR_USER_CANCELLED As Long = vbObjectError + 540

#If VBA7 Then
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#Else
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#End If

' Cached table lookup map (table name -> ListObject) for faster structured-reference resolution.
Private mTableCache As Object

' Button-friendly entry point that refreshes all stored targets.
Public Sub RefreshAllFrozenFormulas()
   RunFormulaRefresh "", "", True
End Sub

' Button-friendly entry point that prompts for groups and shows available group tags.
' Leaving the prompt blank runs all targets.
Public Sub RefreshFrozenFormulasByGroupPrompt()
   Dim response As Variant
   Dim promptText As String
   Dim availableGroupTags As Variant
   Dim groupPreview As String
   Dim resolvedGroupTags As String
   Dim validationMessage As String

   availableGroupTags = GetSortedAvailableGroupTags()
   groupPreview = BuildNumberedGroupTagPreview(availableGroupTags)

   promptText = "Enter one or more group numbers delimited by | ; , or new lines." & vbCrLf & _
                "You can also enter group tags directly." & vbCrLf & _
                "Leave blank to run all targets."

   If Len(groupPreview) > 0 Then
      promptText = promptText & vbCrLf & vbCrLf & "Available groups:" & vbCrLf & groupPreview
   Else
      promptText = promptText & vbCrLf & vbCrLf & "No group tags found in state-sheet column C."
   End If

   response = Application.InputBox( _
      Prompt:=promptText, _
      Title:="Refresh Frozen Formulas By Group", _
      Type:=2)

   If VarType(response) = vbBoolean Then
      If response = False Then Exit Sub
   End If

   If Not TryResolveGroupSelectionInput(CStr(response), availableGroupTags, resolvedGroupTags, validationMessage) Then
      MsgBox validationMessage, vbExclamation
      Exit Sub
   End If

   RunFormulaRefresh "", resolvedGroupTags, True
End Sub

' Discovers stored targets from __FrozenFormulaState and refreshes selected targets.
Private Sub RunFormulaRefresh(ByVal targetReferences As String, ByVal groupTags As String, ByVal showMessages As Boolean)
   Dim ws As Worksheet
   Dim lastRow As Long
   Dim i As Long
   Dim rowIdx As Long
   Dim rowGroupTags As String
   Dim targetRef As String
   Dim formulaText As String
   Dim failed As String
   Dim processedCount As Long
   Dim totalCount As Long
   Dim selectedTargetSet As Object
   Dim selectedGroupSet As Object
   Dim previousCalcMode As XlCalculation
   Dim previousEnableCancelKey As XlEnableCancelKey
   Dim previousScreenUpdating As Boolean
   Dim previousEnableEvents As Boolean
   Dim stateData As Variant
   Dim selectedRows() As Long
   Dim selectedRowCount As Long
   Dim preparedRanges() As Range
   Dim preparedCount As Long
   Dim preparedRange As Range
   Dim prepareError As String

   Set ws = GetOrCreateStateSheet()
   lastRow = ws.Cells(ws.Rows.Count, STATE_COL_TARGET_REFERENCE).End(xlUp).Row

   Set selectedTargetSet = BuildTargetFilterSet(targetReferences)
   Set selectedGroupSet = BuildGroupFilterSet(groupTags)

   If Not selectedTargetSet Is Nothing Then
      If selectedTargetSet.Count = 0 Then
         If showMessages Then MsgBox "No target references were provided.", vbExclamation
         Exit Sub
      End If
   End If

   If Not selectedGroupSet Is Nothing Then
      If selectedGroupSet.Count = 0 Then
         If showMessages Then MsgBox "No group tags were provided.", vbExclamation
         Exit Sub
      End If
   End If

   If lastRow = 1 And Len(Trim$(CStr(ws.Cells(1, STATE_COL_TARGET_REFERENCE).Value2))) = 0 Then
      If showMessages Then MsgBox "No stored formula targets were found on " & FORMULA_STATE_SHEET_NAME & ".", vbExclamation
      Exit Sub
   End If

   stateData = GetStateSheetData(ws, lastRow)
   If Not IsArray(stateData) Then
      If showMessages Then MsgBox "No stored formula targets were found on " & FORMULA_STATE_SHEET_NAME & ".", vbExclamation
      Exit Sub
   End If

   ReDim selectedRows(1 To UBound(stateData, 1))

   For i = 1 To UBound(stateData, 1)
      targetRef = Trim$(CStr(stateData(i, STATE_COL_TARGET_REFERENCE)))
      rowGroupTags = Trim$(CStr(stateData(i, STATE_COL_GROUP_TAGS)))
      If IsRowSelectedForRefresh(targetRef, rowGroupTags, selectedTargetSet, selectedGroupSet) Then
         selectedRowCount = selectedRowCount + 1
         selectedRows(selectedRowCount) = i
      End If
   Next i

   totalCount = selectedRowCount

   If totalCount = 0 Then
      If showMessages Then MsgBox "No matching targets were found in " & FORMULA_STATE_SHEET_NAME & ".", vbExclamation
      Exit Sub
   End If

   ReDim preparedRanges(1 To totalCount)

   previousCalcMode = Application.Calculation
   previousEnableCancelKey = Application.EnableCancelKey
   previousScreenUpdating = Application.ScreenUpdating
   previousEnableEvents = Application.EnableEvents

   Application.Calculation = xlCalculationManual
   Application.EnableCancelKey = xlDisabled
   Application.ScreenUpdating = False
   Application.EnableEvents = False
   ResetWorkbookTableCache
   InitializeWorkbookTableCache
   On Error GoTo SafeExit

   ' Stage 1: restore formulas across all selected targets.
   For i = 1 To selectedRowCount
      ThrowIfCancelRequested "preparing targets"

      rowIdx = selectedRows(i)
      targetRef = Trim$(CStr(stateData(rowIdx, STATE_COL_TARGET_REFERENCE)))
      formulaText = NormalizeFormulaTemplateText(CStr(stateData(rowIdx, STATE_COL_FORMULA_TEMPLATE)))

      processedCount = processedCount + 1
      UpdateProgressStatus "Preparing formula targets: " & CStr(processedCount) & " of " & CStr(totalCount) & "..."

      Set preparedRange = Nothing
      prepareError = ""

      If TryPrepareTargetForBatch(targetRef, formulaText, preparedRange, prepareError) Then
         If Not preparedRange Is Nothing Then
            preparedCount = preparedCount + 1
            Set preparedRanges(preparedCount) = preparedRange
         End If
      Else
         If Len(failed) > 0 Then failed = failed & vbCrLf
         If Len(prepareError) > 0 Then
            failed = failed & targetRef & " - " & prepareError
         Else
            failed = failed & targetRef
         End If
      End If
   Next i

   ' Stage 2: calculate once and then freeze all prepared targets.
   If preparedCount > 0 Then
      ThrowIfCancelRequested "before calculation"

      ReleaseStatusBarForCalc
      Application.Calculate
      WaitForCalculationComplete

      For i = 1 To preparedCount
         ThrowIfCancelRequested "freezing values"

         UpdateProgressStatus "Freezing calculated values: " & CStr(i) & " of " & CStr(preparedCount) & "..."
         preparedRanges(i).Value2 = preparedRanges(i).Value2
      Next i
   End If

SafeExit:
   Application.Calculation = previousCalcMode
   Application.EnableCancelKey = previousEnableCancelKey
   Application.EnableEvents = previousEnableEvents
   Application.ScreenUpdating = previousScreenUpdating
   ResetWorkbookTableCache
   ClearProgressStatusUI

   If Err.Number <> 0 Then
      If Len(failed) > 0 Then failed = failed & vbCrLf
      If Err.Number = ERR_USER_CANCELLED Then
         failed = failed & "(cancelled) " & Err.Description
      Else
         failed = failed & "(dispatcher error) " & Err.Description
      End If
   End If

   If showMessages Then
      If Len(failed) = 0 Then
         MsgBox "Refreshed " & CStr(processedCount) & " formula target(s) from " & FORMULA_STATE_SHEET_NAME & ".", vbInformation
      Else
         MsgBox "Some state-sheet targets failed to refresh:" & vbCrLf & failed, vbExclamation
      End If
   End If
End Sub

' Reads A:C rows from the state sheet in one call for faster dispatch operations.
Private Function GetStateSheetData(ByVal ws As Worksheet, ByVal lastRow As Long) As Variant
   If lastRow < 1 Then Exit Function
   GetStateSheetData = ws.Range(ws.Cells(1, STATE_COL_TARGET_REFERENCE), ws.Cells(lastRow, STATE_COL_GROUP_TAGS)).Value2
End Function

' Resolves, validates, and restores formulas for one target prior to batch calculation.
' Returns True when the target should be considered successful (including no-op rows).
Private Function TryPrepareTargetForBatch(ByVal targetReference As String, ByVal storedFormulaOverride As String, ByRef targetRng As Range, ByRef failureReason As String) As Boolean
   Dim formulaText As String
   Dim storeKey As String

   storeKey = BuildStoreKey(targetReference)

   Set targetRng = ResolveReferenceRange(targetReference)
   If targetRng Is Nothing Then
      If IsStructuredReferenceEmpty(targetReference) Then
         TryPrepareTargetForBatch = True
         Exit Function
      End If

      failureReason = "Target range not found or empty"
      Exit Function
   End If

   If Len(storedFormulaOverride) > 0 Then
      formulaText = NormalizeFormulaTemplateText(storedFormulaOverride)
   Else
      formulaText = ReadFormulaTemplateForTarget(targetRng, storeKey)
   End If

   If Len(formulaText) = 0 Then
      If IsRangeBlank(targetRng) Then
         Set targetRng = Nothing
         TryPrepareTargetForBatch = True
         Exit Function
      End If

      failureReason = "No stored formula template found"
      Exit Function
   End If

   On Error GoTo HandleFailure
   WriteStoredFormulaTemplate storeKey, formulaText
   RestoreTargetFormulasIfNeeded targetRng, formulaText
   TryPrepareTargetForBatch = True
   Exit Function

HandleFailure:
   failureReason = Err.Description
   Set targetRng = Nothing
End Function

' True when the row should run for this dispatch.
' A row is selected when:
' 1. Target filter is empty or contains the target reference.
' 2. Group filter is empty or intersects the row's group-tag list.
Private Function IsRowSelectedForRefresh(ByVal targetReference As String, ByVal rowGroupTags As String, ByVal selectedTargetSet As Object, ByVal selectedGroupSet As Object) As Boolean
   If Len(targetReference) = 0 Then Exit Function
   If Not MatchesTargetFilter(targetReference, selectedTargetSet) Then Exit Function
   If Not MatchesGroupFilter(rowGroupTags, selectedGroupSet) Then Exit Function
   IsRowSelectedForRefresh = True
End Function

' Builds a case-insensitive lookup set from a delimited target list.
' Returns Nothing when no explicit subset was provided (meaning "select all").
Private Function BuildTargetFilterSet(ByVal targetReferences As String) As Object
   Set BuildTargetFilterSet = BuildFilterSetFromDelimitedList(targetReferences, True)
End Function

' Builds a case-insensitive lookup set from a delimited group-tag list.
' Returns Nothing when no explicit subset was provided (meaning "all groups").
Private Function BuildGroupFilterSet(ByVal groupTags As String) As Object
   Set BuildGroupFilterSet = BuildFilterSetFromDelimitedList(groupTags, False)
End Function

' Returns known group tags from the state sheet as a sorted array.
Private Function GetSortedAvailableGroupTags() As Variant
   Dim tagSet As Object
   Dim groupTags As Variant

   Set tagSet = BuildAvailableGroupTagSet()
   If tagSet Is Nothing Then Exit Function

   groupTags = tagSet.Keys
   SortStringArray groupTags
   GetSortedAvailableGroupTags = groupTags
End Function

' Builds a numbered preview list of known group tags.
Private Function BuildNumberedGroupTagPreview(ByVal groupTags As Variant, Optional ByVal maxItems As Long = 0) As String
   Dim i As Long
   Dim shownCount As Long
   Dim groupCount As Long
   Dim displayCount As Long
   Dim outText As String

   If Not IsArray(groupTags) Then Exit Function

   groupCount = UBound(groupTags) - LBound(groupTags) + 1
   If groupCount <= 0 Then Exit Function

   If maxItems <= 0 Or maxItems > groupCount Then
      displayCount = groupCount
   Else
      displayCount = maxItems
   End If

   For i = LBound(groupTags) To LBound(groupTags) + displayCount - 1
      shownCount = shownCount + 1
      If Len(outText) > 0 Then outText = outText & vbCrLf
      outText = outText & CStr(shownCount) & ") " & CStr(groupTags(i))
   Next i

   If displayCount < groupCount Then
      outText = outText & vbCrLf & "+" & CStr(groupCount - displayCount) & " more groups not shown"
   End If

   BuildNumberedGroupTagPreview = outText
End Function

' Resolves prompt input into a group-tag filter string.
' Numeric tokens map to group list indices (1-based), while non-numeric tokens are treated as group tags.
Private Function TryResolveGroupSelectionInput(ByVal userInput As String, ByVal availableGroupTags As Variant, ByRef resolvedGroupTags As String, ByRef validationMessage As String) As Boolean
   Dim normalized As String
   Dim parts() As String
   Dim i As Long
   Dim token As String
   Dim selectedGroups As Object
   Dim selectedIndex As Long
   Dim mappedTag As String

   normalized = NormalizeDelimitedList(userInput)
   If Len(normalized) = 0 Then
      resolvedGroupTags = ""
      TryResolveGroupSelectionInput = True
      Exit Function
   End If

   Set selectedGroups = CreateTextCompareSet()
   parts = Split(normalized, "|")

   For i = LBound(parts) To UBound(parts)
      token = Trim$(parts(i))
      If Len(token) = 0 Then GoTo NextToken

      If TryParsePositiveIndex(token, selectedIndex) Then
         If Not IsArray(availableGroupTags) Then
            validationMessage = "No numbered groups are available right now. Enter group names directly or leave blank to run all."
            Exit Function
         End If

         If selectedIndex < 1 Or selectedIndex > (UBound(availableGroupTags) - LBound(availableGroupTags) + 1) Then
            validationMessage = "Group number " & token & " is out of range."
            Exit Function
         End If

         mappedTag = CStr(availableGroupTags(LBound(availableGroupTags) + selectedIndex - 1))
         selectedGroups(mappedTag) = True
      Else
         selectedGroups(token) = True
      End If

NextToken:
   Next i

   resolvedGroupTags = JoinVariantArray(selectedGroups.Keys, "|")
   TryResolveGroupSelectionInput = True
End Function

' Parses a strictly positive integer token.
Private Function TryParsePositiveIndex(ByVal token As String, ByRef indexValue As Long) As Boolean
   If Len(token) = 0 Then Exit Function
   If token Like "*[!0-9]*" Then Exit Function

   indexValue = CLng(token)
   TryParsePositiveIndex = indexValue > 0
End Function

' Reads column C group tags from the state sheet and returns a distinct token set.
Private Function BuildAvailableGroupTagSet() As Object
   Dim ws As Worksheet
   Dim lastRow As Long
   Dim i As Long
   Dim tagSet As Object
   Dim stateData As Variant

   Set ws = GetOrCreateStateSheet()
   lastRow = ws.Cells(ws.Rows.Count, STATE_COL_TARGET_REFERENCE).End(xlUp).Row

   Set tagSet = CreateTextCompareSet()
   stateData = GetStateSheetData(ws, lastRow)
   If Not IsArray(stateData) Then Exit Function

   For i = 1 To UBound(stateData, 1)
      AddDelimitedTokensToSet CStr(stateData(i, STATE_COL_GROUP_TAGS)), False, tagSet
   Next i

   If tagSet.Count = 0 Then Exit Function
   Set BuildAvailableGroupTagSet = tagSet
End Function

' Shared parser for delimited filter inputs.
Private Function BuildFilterSetFromDelimitedList(ByVal rawList As String, ByVal useTargetKey As Boolean) As Object
   Dim outSet As Object

   Set outSet = CreateTextCompareSet()
   AddDelimitedTokensToSet rawList, useTargetKey, outSet
   If outSet.Count = 0 Then Exit Function

   Set BuildFilterSetFromDelimitedList = outSet
End Function

' Adds normalized delimited tokens to an existing dictionary set.
Private Sub AddDelimitedTokensToSet(ByVal rawList As String, ByVal useTargetKey As Boolean, ByRef tokenSet As Object)
   Dim normalized As String
   Dim parts() As String
   Dim i As Long
   Dim token As String
   Dim normalizedToken As String

   normalized = NormalizeDelimitedList(rawList)
   If Len(normalized) = 0 Then Exit Sub

   parts = Split(normalized, "|")
   For i = LBound(parts) To UBound(parts)
      token = Trim$(parts(i))
      If Len(token) > 0 Then
         If useTargetKey Then
            normalizedToken = BuildStoreKey(token)
         Else
            normalizedToken = NormalizeGroupTagToken(token)
         End If

         If Len(normalizedToken) > 0 Then tokenSet(normalizedToken) = True
      End If
   Next i
End Sub

' Normalizes delimiter variants to a pipe-delimited token list.
Private Function NormalizeDelimitedList(ByVal rawList As String) As String
   Dim normalized As String

   normalized = rawList
   normalized = Replace(normalized, vbCrLf, "|")
   normalized = Replace(normalized, vbLf, "|")
   normalized = Replace(normalized, ";", "|")
   normalized = Replace(normalized, ",", "|")
   NormalizeDelimitedList = Trim$(normalized)
End Function

' Creates a case-insensitive dictionary for set-like membership checks.
Private Function CreateTextCompareSet() As Object
   Dim outSet As Object

   Set outSet = CreateObject("Scripting.Dictionary")
   outSet.CompareMode = vbTextCompare
   Set CreateTextCompareSet = outSet
End Function

' Joins a variant array into one string using the requested delimiter.
Private Function JoinVariantArray(ByVal values As Variant, ByVal delimiter As String) As String
   Dim i As Long
   Dim outText As String

   If Not IsArray(values) Then Exit Function

   For i = LBound(values) To UBound(values)
      If Len(outText) > 0 Then outText = outText & delimiter
      outText = outText & CStr(values(i))
   Next i

   JoinVariantArray = outText
End Function

' True when the target belongs to the requested subset, or when no subset was provided.
Private Function MatchesTargetFilter(ByVal targetReference As String, ByVal selectedSet As Object) As Boolean
   If selectedSet Is Nothing Then
      MatchesTargetFilter = True
   Else
      MatchesTargetFilter = selectedSet.Exists(BuildStoreKey(targetReference))
   End If
End Function

' True when rowGroupTags intersects selectedGroupSet.
' If selectedGroupSet is Nothing, all rows pass.
Private Function MatchesGroupFilter(ByVal rowGroupTags As String, ByVal selectedGroupSet As Object) As Boolean
   Dim rowGroupSet As Object
   Dim groupKey As Variant

   If selectedGroupSet Is Nothing Then
      MatchesGroupFilter = True
      Exit Function
   End If

   Set rowGroupSet = BuildFilterSetFromDelimitedList(rowGroupTags, False)
   If rowGroupSet Is Nothing Then Exit Function

   For Each groupKey In rowGroupSet.Keys
      If selectedGroupSet.Exists(CStr(groupKey)) Then
         MatchesGroupFilter = True
         Exit Function
      End If
   Next groupKey
End Function

' Sorts a variant array of strings in place (case-insensitive ascending).
Private Sub SortStringArray(ByRef values As Variant)
   Dim i As Long
   Dim j As Long
   Dim tmp As String

   If Not IsArray(values) Then Exit Sub
   If UBound(values) <= LBound(values) Then Exit Sub

   For i = LBound(values) To UBound(values) - 1
      For j = i + 1 To UBound(values)
         If StrComp(CStr(values(i)), CStr(values(j)), vbTextCompare) > 0 Then
            tmp = CStr(values(i))
            values(i) = values(j)
            values(j) = tmp
         End If
      Next j
   Next i
End Sub

' Canonical representation of a group tag.
Private Function NormalizeGroupTagToken(ByVal groupTag As String) As String
   NormalizeGroupTagToken = UCase$(Trim$(groupTag))
End Function

' Single-target refresh implementation with optional user prompts.
Private Function RefreshSingleTargetAndFreezeCore(ByVal targetReference As String, ByVal showMessages As Boolean, Optional ByVal storedFormulaOverride As String = "") As Boolean
   Dim targetRng As Range
   Dim formulaText As String
   Dim storeKey As String
   Dim previousEnableCancelKey As XlEnableCancelKey
   Dim previousEnableEvents As Boolean
   Dim previousScreenUpdating As Boolean

   storeKey = BuildStoreKey(targetReference)

   ResetWorkbookTableCache
   InitializeWorkbookTableCache

   Set targetRng = ResolveReferenceRange(targetReference)
   If targetRng Is Nothing Then
      If IsStructuredReferenceEmpty(targetReference) Then
         RefreshSingleTargetAndFreezeCore = True
         Exit Function
      End If
      If showMessages Then MsgBox "Target range not found or empty: " & targetReference, vbExclamation
      Exit Function
   End If

   If Len(storedFormulaOverride) > 0 Then
      formulaText = NormalizeFormulaTemplateText(storedFormulaOverride)
   Else
      formulaText = ReadFormulaTemplateForTarget(targetRng, storeKey)
   End If
   If Len(formulaText) = 0 Then
      If IsRangeBlank(targetRng) Then
         RefreshSingleTargetAndFreezeCore = True
         Exit Function
      End If
      If showMessages Then MsgBox "No stored formula template found. Put formula back in first target cell, then run again.", vbExclamation
      Exit Function
   End If

   previousEnableCancelKey = Application.EnableCancelKey
   previousEnableEvents = Application.EnableEvents
   previousScreenUpdating = Application.ScreenUpdating

   Application.EnableCancelKey = xlDisabled
   Application.ScreenUpdating = False
   Application.EnableEvents = False
   On Error GoTo SafeExit

   ThrowIfCancelRequested "single-target refresh"

   WriteStoredFormulaTemplate storeKey, formulaText
   UpdateProgressStatus "Restoring formulas for " & targetReference & "..."
   RestoreTargetFormulasIfNeeded targetRng, formulaText

   ReleaseStatusBarForCalc
   Application.Calculate
   WaitForCalculationComplete
   UpdateProgressStatus "Freezing calculated values for " & targetReference & "..."
   targetRng.Value2 = targetRng.Value2

SafeExit:
   Application.EnableCancelKey = previousEnableCancelKey
   Application.EnableEvents = previousEnableEvents
   Application.ScreenUpdating = previousScreenUpdating

   If Err.Number <> 0 Then
      If showMessages Then MsgBox "Refresh failed: " & Err.Description, vbExclamation
      Exit Function
   End If

   RefreshSingleTargetAndFreezeCore = True
End Function

' Finds a table anywhere in the workbook by name.
Private Function FindWorkbookTableByName(ByVal tableName As String) As ListObject
   Dim cacheKey As String
   Dim ws As Worksheet
   Dim lo As ListObject

   cacheKey = Trim$(tableName)
   If Len(cacheKey) = 0 Then Exit Function

   InitializeWorkbookTableCache
   If Not mTableCache Is Nothing Then
      If mTableCache.Exists(cacheKey) Then
         Set FindWorkbookTableByName = mTableCache(cacheKey)
         Exit Function
      End If
   End If

   ' Fallback path (should be rare) in case cache priming missed a table.
   For Each ws In ThisWorkbook.Worksheets
      For Each lo In ws.ListObjects
         If StrComp(lo.Name, cacheKey, vbTextCompare) = 0 Then
            Set FindWorkbookTableByName = lo
            Exit Function
         End If
      Next lo
   Next ws
End Function

' Builds a fresh cache of workbook ListObjects for structured-reference lookups.
Private Sub InitializeWorkbookTableCache()
   Dim ws As Worksheet
   Dim lo As ListObject

   If Not mTableCache Is Nothing Then Exit Sub

   Set mTableCache = CreateTextCompareSet()
   For Each ws In ThisWorkbook.Worksheets
      For Each lo In ws.ListObjects
         If Not mTableCache.Exists(lo.Name) Then mTableCache.Add lo.Name, lo
      Next lo
   Next ws
End Sub

' Clears the table lookup cache so each run starts from a current workbook snapshot.
Private Sub ResetWorkbookTableCache()
   Set mTableCache = Nothing
End Sub

' Resolves either a structured table-column reference or a standard range/name reference.
Private Function ResolveReferenceRange(ByVal targetReference As String) As Range
   Dim lo As ListObject
   Dim tableName As String
   Dim columnName As String
   Dim bangPos As Long
   Dim wsName As String
   Dim localAddress As String

   If TryParseStructuredReference(targetReference, tableName, columnName) Then
      Set lo = FindWorkbookTableByName(tableName)
      If Not lo Is Nothing Then
         On Error Resume Next
         Set ResolveReferenceRange = lo.ListColumns(columnName).DataBodyRange
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
      Set ResolveReferenceRange = ThisWorkbook.Worksheets(wsName).Range(localAddress)
      On Error GoTo 0
      If Not ResolveReferenceRange Is Nothing Then Exit Function
   End If

   On Error Resume Next
   Set ResolveReferenceRange = ThisWorkbook.Names(targetReference).RefersToRange
   On Error GoTo 0
   If Not ResolveReferenceRange Is Nothing Then Exit Function

   On Error Resume Next
   Set ResolveReferenceRange = Application.Range(targetReference)
   On Error GoTo 0
End Function

' Parses references of the form TableName[Column Name].
Private Function TryParseStructuredReference(ByVal targetReference As String, ByRef tableName As String, ByRef columnName As String) As Boolean
   Dim openPos As Long
   Dim closePos As Long

   openPos = InStr(1, targetReference, "[", vbTextCompare)
   closePos = InStrRev(targetReference, "]", -1, vbTextCompare)

   If openPos <= 1 Or closePos <= openPos Then Exit Function

   tableName = Trim$(Left$(targetReference, openPos - 1))
   columnName = Trim$(Mid$(targetReference, openPos + 1, closePos - openPos - 1))

   TryParseStructuredReference = Len(tableName) > 0 And Len(columnName) > 0
End Function

' Uses the live formula when available, otherwise falls back to the stored template.
Private Function ReadFormulaTemplateForTarget(ByVal targetRng As Range, ByVal storeKey As String) As String
   Dim formulaCell As Range

   Set formulaCell = FindFirstFormulaCellInRange(targetRng)

   If Not formulaCell Is Nothing Then
      On Error Resume Next
      ReadFormulaTemplateForTarget = formulaCell.Formula2
      If Len(ReadFormulaTemplateForTarget) = 0 Then ReadFormulaTemplateForTarget = formulaCell.Formula
      On Error GoTo 0
   Else
      ReadFormulaTemplateForTarget = ReadStoredFormulaTemplate(storeKey)
   End If
End Function

' Returns the first formula cell in a range, if one exists.
Private Function FindFirstFormulaCellInRange(ByVal rng As Range) As Range
   Dim formulaCells As Range

   On Error Resume Next
   Set formulaCells = rng.SpecialCells(xlCellTypeFormulas)
   On Error GoTo 0

   If Not formulaCells Is Nothing Then Set FindFirstFormulaCellInRange = formulaCells.Cells(1, 1)
End Function

' Restores formulas across the entire target range if any cells were frozen to values.
Private Sub RestoreTargetFormulasIfNeeded(ByVal targetRng As Range, ByVal formulaText As String)
   Dim formulaToApply As String

   If IsRangeFormulaBacked(targetRng) Then Exit Sub

   formulaToApply = NormalizeFormulaTemplateText(formulaText)
   If Len(formulaToApply) = 0 Then
      Err.Raise vbObjectError + 501, "RestoreTargetFormulasIfNeeded", "Stored formula template was empty after normalization."
   End If

   On Error Resume Next
   targetRng.Formula2 = formulaToApply
   If Err.Number <> 0 Then
      Err.Clear
      targetRng.Formula = formulaToApply
   End If
   On Error GoTo 0

    If Not targetRng.Cells(1, 1).HasFormula Then
      Err.Raise vbObjectError + 502, "RestoreTargetFormulasIfNeeded", _
                "Stored template could not be applied as a formula: " & Left$(formulaToApply, 120)
   End If
End Sub

' Detects mixed state after an interrupted run so the next refresh can self-heal.
Private Function IsRangeFormulaBacked(ByVal rng As Range) As Boolean
   Dim constants As Range

   On Error Resume Next
   Set constants = rng.SpecialCells(xlCellTypeConstants)
   On Error GoTo 0

   IsRangeFormulaBacked = constants Is Nothing
End Function

' True when the requested structured target exists but has no data rows.
Private Function IsStructuredReferenceEmpty(ByVal targetReference As String) As Boolean
   Dim lo As ListObject
   Dim tableName As String
   Dim columnName As String
   Dim targetCol As ListColumn

   If Not TryParseStructuredReference(targetReference, tableName, columnName) Then Exit Function

   Set lo = FindWorkbookTableByName(tableName)
   If lo Is Nothing Then Exit Function

   On Error Resume Next
   Set targetCol = lo.ListColumns(columnName)
   On Error GoTo 0

   If targetCol Is Nothing Then Exit Function

   IsStructuredReferenceEmpty = targetCol.DataBodyRange Is Nothing
End Function

' True when no cell in the range contains a value or formula result.
Private Function IsRangeBlank(ByVal rng As Range) As Boolean
   On Error Resume Next
   IsRangeBlank = (WorksheetFunction.CountA(rng) = 0)
   On Error GoTo 0
End Function

' Waits until Excel reports that calculation has finished before freezing values.
Private Sub WaitForCalculationComplete()
   Do While Application.CalculationState <> xlDone
      ThrowIfCancelRequested "waiting for calculation"
      DoEvents
   Loop
End Sub

' True when the user is requesting cancellation via Esc.
Private Function IsCancelRequested() As Boolean
   IsCancelRequested = (GetAsyncKeyState(CANCEL_VK_ESCAPE) <> 0)
End Function

' Raises a controlled cancellation error so callers can unwind cleanly.
Private Sub ThrowIfCancelRequested(Optional ByVal operationName As String = "")
   If Not IsCancelRequested() Then Exit Sub

   If Len(operationName) > 0 Then
      Err.Raise ERR_USER_CANCELLED, "FormulaFreezer", "Operation cancelled by user during " & operationName & "."
   Else
      Err.Raise ERR_USER_CANCELLED, "FormulaFreezer", "Operation cancelled by user."
   End If
End Sub

' Stores one reusable formula template per target reference.
Private Sub WriteStoredFormulaTemplate(ByVal storeKey As String, ByVal formulaText As String)
   Dim ws As Worksheet
   Dim keyCell As Range
   Dim formulaCell As Range

   Set ws = GetOrCreateStateSheet()
   Set keyCell = FindStoreKeyCell(ws, storeKey)

   If keyCell Is Nothing Then
      Set keyCell = ws.Cells(ws.Rows.Count, STATE_COL_TARGET_REFERENCE).End(xlUp).Offset(1, 0)
      If Len(ws.Cells(1, STATE_COL_TARGET_REFERENCE).Value2) = 0 Then Set keyCell = ws.Cells(1, STATE_COL_TARGET_REFERENCE)
      keyCell.Value2 = storeKey
   End If

   Set formulaCell = keyCell.Offset(0, STATE_COL_FORMULA_TEMPLATE - STATE_COL_TARGET_REFERENCE)
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

   Set ws = GetOrCreateStateSheet()
   Set keyCell = FindStoreKeyCell(ws, storeKey)

   If Not keyCell Is Nothing Then
      Set formulaCell = keyCell.Offset(0, STATE_COL_FORMULA_TEMPLATE - STATE_COL_TARGET_REFERENCE)
      ReadStoredFormulaTemplate = ReadTemplateTextFromCell(formulaCell)
   End If
End Function

' Reads one stored formula cell and normalizes text payloads like '=... and ''=....
Private Function ReadTemplateTextFromCell(ByVal formulaCell As Range) As String
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

   ReadTemplateTextFromCell = NormalizeFormulaTemplateText(formulaText)
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
Private Function BuildStoreKey(ByVal targetReference As String) As String
   BuildStoreKey = UCase$(Trim$(targetReference))
End Function

' Finds an existing stored formula record on the hidden state sheet.
Private Function FindStoreKeyCell(ByVal ws As Worksheet, ByVal storeKey As String) As Range
   Dim lastRow As Long
   Dim keys As Range

   lastRow = ws.Cells(ws.Rows.Count, STATE_COL_TARGET_REFERENCE).End(xlUp).Row
   If lastRow < 1 Then Exit Function
   If lastRow = 1 And Len(ws.Cells(1, STATE_COL_TARGET_REFERENCE).Value2) = 0 Then Exit Function

   Set keys = ws.Range(ws.Cells(1, STATE_COL_TARGET_REFERENCE), ws.Cells(lastRow, STATE_COL_TARGET_REFERENCE))
   Set FindStoreKeyCell = keys.Find(What:=storeKey, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
End Function

' Creates or returns the hidden storage sheet used by this utility module.
Private Function GetOrCreateStateSheet() As Worksheet
   Dim ws As Worksheet

   On Error Resume Next
   Set ws = ThisWorkbook.Worksheets(FORMULA_STATE_SHEET_NAME)
   On Error GoTo 0

   If ws Is Nothing Then
      Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
      ws.Name = FORMULA_STATE_SHEET_NAME
      ws.Visible = xlSheetVeryHidden
   End If

   Set GetOrCreateStateSheet = ws
End Function

' Shows lightweight progress feedback without needing a custom form.
Private Sub UpdateProgressStatus(ByVal messageText As String)
   On Error Resume Next
   Application.Cursor = xlWait
   Application.StatusBar = messageText
   DoEvents
   On Error GoTo 0
End Sub

' Releases the status bar so Excel can show its own calculation progress UI.
Private Sub ReleaseStatusBarForCalc()
   On Error Resume Next
   Application.StatusBar = False
   DoEvents
   On Error GoTo 0
End Sub

' Restores the default Excel UI after work completes.
Private Sub ClearProgressStatusUI()
   On Error Resume Next
   Application.StatusBar = False
   Application.Cursor = xlDefault
   DoEvents
   On Error GoTo 0
End Sub