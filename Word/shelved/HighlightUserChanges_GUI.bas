Attribute VB_Name = "HighlightUserChanges"
Option Explicit

' Import this module and frmHighlightOptions.frm into the same Word VBA project.
Public Enum HighlightRunMode
    RunHighlightInsertionsDeletions = 1
    RunAcceptFormatting = 2
    RunAcceptInsertionsDeletions = 3
    RunClearScriptHighlights = 4
End Enum

Public Sub HighlightUserChanges()
    Dim authors As Collection
    Dim selectedAuthors As Collection
    Dim missingColors As Collection
    Dim matchingRevisions As Collection
    Dim mode As HighlightRunMode
    Dim restrictAuthors As Boolean
    Dim useBeforeDate As Boolean
    Dim useAfterDate As Boolean
    Dim beforeDate As Date
    Dim afterDate As Date
    Dim includeBeforeDate As Boolean
    Dim includeAfterDate As Boolean
    Dim configuringColors As Boolean
    Dim revisionCount As Long

    If ActiveDocument.Revisions.Count = 0 Then
        MsgBox "This document has no tracked revisions.", vbInformation
        Exit Sub
    End If

    Set authors = GetDocumentAuthors(ActiveDocument)
    Load frmHighlightOptions
    frmHighlightOptions.InitializeAuthors authors

    Do
        frmHighlightOptions.Show vbModal

        If frmHighlightOptions.WasCancelled Then
            Unload frmHighlightOptions
            Exit Sub
        End If

        If frmHighlightOptions.ReturnedToOptions Then
            frmHighlightOptions.ClearColorRequirements
            configuringColors = False
        ElseIf Not configuringColors Then
            mode = frmHighlightOptions.RunMode
            restrictAuthors = frmHighlightOptions.AuthorFilterEnabled
            Set selectedAuthors = frmHighlightOptions.SelectedAuthors
            useBeforeDate = frmHighlightOptions.UseBeforeDate
            beforeDate = frmHighlightOptions.BeforeDate
            includeBeforeDate = frmHighlightOptions.IncludeBeforeDate
            useAfterDate = frmHighlightOptions.UseAfterDate
            afterDate = frmHighlightOptions.AfterDate
            includeAfterDate = frmHighlightOptions.IncludeAfterDate

            Set missingColors = GetMissingHighlightColors( _
                ActiveDocument, mode, restrictAuthors, selectedAuthors, _
                useBeforeDate, beforeDate, includeBeforeDate, _
                useAfterDate, afterDate, includeAfterDate, frmHighlightOptions)

            If missingColors.Count > 0 Then
                frmHighlightOptions.SetColorRequirements missingColors
                configuringColors = True
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop

    If frmHighlightOptions.WasCancelled Then
        Unload frmHighlightOptions
        Exit Sub
    End If

    Set matchingRevisions = GetMatchingRevisions( _
        ActiveDocument, mode, restrictAuthors, selectedAuthors, _
        useBeforeDate, beforeDate, includeBeforeDate, _
        useAfterDate, afterDate, includeAfterDate)

    If matchingRevisions.Count = 0 Then
        Unload frmHighlightOptions
        If mode = RunClearScriptHighlights Then
            MsgBox "No matching Chase Lynk color-formatting bursts were found.", vbInformation
        Else
            MsgBox "No revisions match the selected function and filters.", vbInformation
        End If
        Exit Sub
    End If

    If mode = RunClearScriptHighlights Then
        If MsgBox("Found " & matchingRevisions.Count & _
                  " Chase Lynk font-color revisions in bursts of at least three, " & _
                  "with no more than one second between revisions." & vbCrLf & vbCrLf & _
                  "The macro will reject these formatting revisions, restoring the font colors " & _
                  "that were in place before each change. Other rapid Chase color changes " & _
                  "matching this rule will also be rejected. Continue?", _
                  vbExclamation + vbYesNo + vbDefaultButton2, _
                  "Confirm highlight cleanup") <> vbYes Then
            Unload frmHighlightOptions
            Exit Sub
        End If
    ElseIf mode <> RunHighlightInsertionsDeletions Then
        If MsgBox("Accept " & matchingRevisions.Count & _
                  " matching revisions? This action cannot be undone.", _
                  vbQuestion + vbYesNo + vbDefaultButton2, _
                  "Confirm revision acceptance") <> vbYes Then
            Unload frmHighlightOptions
            Exit Sub
        End If
    End If

    Select Case mode
        Case RunHighlightInsertionsDeletions
            HighlightMatchingRevisions matchingRevisions, frmHighlightOptions
            revisionCount = matchingRevisions.Count
            MsgBox "Highlighted " & revisionCount & " matching revisions.", vbInformation

        Case RunAcceptFormatting, RunAcceptInsertionsDeletions
            AcceptMatchingRevisions matchingRevisions
            revisionCount = matchingRevisions.Count
            MsgBox "Accepted " & revisionCount & " matching revisions.", vbInformation

        Case RunClearScriptHighlights
            ClearScriptHighlightRevisions matchingRevisions
            revisionCount = matchingRevisions.Count
            MsgBox "Rejected " & revisionCount & " script-generated highlight revisions.", vbInformation
    End Select

    Unload frmHighlightOptions
End Sub

Public Function GetHardcodedAuthorColor(ByVal author As String) As Long
    Select Case author
        Case "Chase Lynk"
            GetHardcodedAuthorColor = wdColorRed
        Case "Hossein Targhi"
            GetHardcodedAuthorColor = wdColorBlue
        Case "Tolulope Adekanye"
            GetHardcodedAuthorColor = wdColorViolet
        Case Else
            GetHardcodedAuthorColor = -1
    End Select
End Function

Private Function GetDocumentAuthors(ByVal doc As Document) As Collection
    Dim authors As Collection
    Dim sortedAuthors As Collection
    Dim rev As Revision
    Dim names() As String
    Dim author As String
    Dim i As Long
    Dim j As Long
    Dim temp As String

    Set authors = New Collection

    For Each rev In doc.Revisions
        author = rev.Author
        If Not CollectionContainsAuthor(authors, author) Then authors.Add author
    Next rev

    If authors.Count = 0 Then
        Set GetDocumentAuthors = authors
        Exit Function
    End If

    ReDim names(1 To authors.Count)
    For i = 1 To authors.Count
        names(i) = CStr(authors(i))
    Next i

    For i = 1 To UBound(names) - 1
        For j = i + 1 To UBound(names)
            If StrComp(names(i), names(j), vbTextCompare) > 0 Then
                temp = names(i)
                names(i) = names(j)
                names(j) = temp
            End If
        Next j
    Next i

    Set sortedAuthors = New Collection
    For i = 1 To UBound(names)
        sortedAuthors.Add names(i)
    Next i
    Set GetDocumentAuthors = sortedAuthors
End Function

Private Function CollectionContainsAuthor( _
    ByVal authors As Collection, ByVal author As String) As Boolean
    Dim candidate As Variant

    For Each candidate In authors
        If StrComp(CStr(candidate), author, vbBinaryCompare) = 0 Then
            CollectionContainsAuthor = True
            Exit Function
        End If
    Next candidate
End Function

Private Function IsRevisionForMode( _
    ByVal rev As Revision, ByVal mode As HighlightRunMode) As Boolean
    Select Case mode
        Case RunHighlightInsertionsDeletions
            IsRevisionForMode = _
                (rev.Type = wdRevisionInsert Or _
                 rev.Type = wdRevisionMovedTo Or _
                 rev.Type = wdRevisionDelete)

        Case RunAcceptFormatting
            IsRevisionForMode = IsFormattingRevision(rev.Type)

        Case RunAcceptInsertionsDeletions
            IsRevisionForMode = _
                (rev.Type = wdRevisionInsert Or rev.Type = wdRevisionDelete)
    End Select
End Function

Private Function IsFormattingRevision(ByVal revisionType As WdRevisionType) As Boolean
    Select Case revisionType
        Case wdRevisionProperty, wdRevisionParagraphNumber, wdRevisionStyle, _
             wdRevisionParagraphProperty, wdRevisionTableProperty, _
             wdRevisionSectionProperty, wdRevisionStyleDefinition
            IsFormattingRevision = True
    End Select
End Function

Private Function RevisionMatchesFilters( _
    ByVal rev As Revision, _
    ByVal restrictAuthors As Boolean, _
    ByVal selectedAuthors As Collection, _
    ByVal useBeforeDate As Boolean, _
    ByVal beforeDate As Date, _
    ByVal includeBeforeDate As Boolean, _
    ByVal useAfterDate As Boolean, _
    ByVal afterDate As Date, _
    ByVal includeAfterDate As Boolean) As Boolean
    Dim selectedAuthor As Variant
    Dim revisionDate As Date
    Dim authorMatched As Boolean

    If restrictAuthors Then
        For Each selectedAuthor In selectedAuthors
            If StrComp(CStr(selectedAuthor), rev.Author, vbBinaryCompare) = 0 Then
                authorMatched = True
                Exit For
            End If
        Next selectedAuthor

        If Not authorMatched Then Exit Function
    End If

    revisionDate = DateValue(rev.Date)

    If useBeforeDate Then
        If includeBeforeDate Then
            If revisionDate > beforeDate Then Exit Function
        Else
            If revisionDate >= beforeDate Then Exit Function
        End If
    End If

    If useAfterDate Then
        If includeAfterDate Then
            If revisionDate < afterDate Then Exit Function
        Else
            If revisionDate <= afterDate Then Exit Function
        End If
    End If

    RevisionMatchesFilters = True
End Function

Private Function GetMatchingRevisions( _
    ByVal doc As Document, _
    ByVal mode As HighlightRunMode, _
    ByVal restrictAuthors As Boolean, _
    ByVal selectedAuthors As Collection, _
    ByVal useBeforeDate As Boolean, _
    ByVal beforeDate As Date, _
    ByVal includeBeforeDate As Boolean, _
    ByVal useAfterDate As Boolean, _
    ByVal afterDate As Date, _
    ByVal includeAfterDate As Boolean) As Collection
    Dim matches As Collection
    Dim rev As Revision

    If mode = RunClearScriptHighlights Then
        Set GetMatchingRevisions = GetScriptHighlightBursts( _
            doc, useBeforeDate, beforeDate, includeBeforeDate, _
            useAfterDate, afterDate, includeAfterDate)
        Exit Function
    End If

    Set matches = New Collection

    For Each rev In doc.Revisions
        If IsRevisionForMode(rev, mode) Then
            If RevisionMatchesFilters( _
                rev, restrictAuthors, selectedAuthors, _
                useBeforeDate, beforeDate, includeBeforeDate, _
                useAfterDate, afterDate, includeAfterDate) Then
                matches.Add rev
            End If
        End If
    Next rev

    Set GetMatchingRevisions = matches
End Function

Private Function GetScriptHighlightBursts( _
    ByVal doc As Document, _
    ByVal useBeforeDate As Boolean, _
    ByVal beforeDate As Date, _
    ByVal includeBeforeDate As Boolean, _
    ByVal useAfterDate As Boolean, _
    ByVal afterDate As Date, _
    ByVal includeAfterDate As Boolean) As Collection
    Dim candidates As Collection
    Dim matches As Collection
    Dim emptyAuthors As Collection
    Dim orderedRevisions() As Revision
    Dim rev As Revision
    Dim tempRevision As Revision
    Dim runStart As Long
    Dim gapSeconds As Long
    Dim i As Long
    Dim j As Long

    Set candidates = New Collection
    Set matches = New Collection
    Set emptyAuthors = New Collection

    For Each rev In doc.Revisions
        If IsChaseFontColorRevision(rev) Then
            If RevisionMatchesFilters( _
                rev, False, emptyAuthors, _
                useBeforeDate, beforeDate, includeBeforeDate, _
                useAfterDate, afterDate, includeAfterDate) Then
                candidates.Add rev
            End If
        End If
    Next rev

    If candidates.Count < 3 Then
        Set GetScriptHighlightBursts = matches
        Exit Function
    End If

    ReDim orderedRevisions(1 To candidates.Count)
    For i = 1 To candidates.Count
        Set orderedRevisions(i) = candidates(i)
    Next i

    For i = 1 To UBound(orderedRevisions) - 1
        For j = i + 1 To UBound(orderedRevisions)
            If orderedRevisions(i).Date > orderedRevisions(j).Date Then
                Set tempRevision = orderedRevisions(i)
                Set orderedRevisions(i) = orderedRevisions(j)
                Set orderedRevisions(j) = tempRevision
            End If
        Next j
    Next i

    runStart = 1
    For i = 2 To UBound(orderedRevisions) + 1
        If i <= UBound(orderedRevisions) Then
            gapSeconds = DateDiff("s", orderedRevisions(i - 1).Date, orderedRevisions(i).Date)
        Else
            gapSeconds = 2
        End If

        If gapSeconds > 1 Or gapSeconds < 0 Then
            If i - runStart >= 3 Then
                For j = runStart To i - 1
                    matches.Add orderedRevisions(j)
                Next j
            End If
            runStart = i
        End If
    Next i

    Set GetScriptHighlightBursts = matches
End Function

Private Function IsChaseFontColorRevision(ByVal rev As Revision) As Boolean
    Dim description As String

    If StrComp(rev.Author, "Chase Lynk", vbBinaryCompare) <> 0 Then Exit Function
    If rev.Type <> wdRevisionProperty Then Exit Function

    description = rev.FormatDescription
    If InStr(1, description, "Font", vbTextCompare) = 0 Then Exit Function
    If InStr(1, description, "Color", vbTextCompare) = 0 Then Exit Function

    IsChaseFontColorRevision = True
End Function

Private Function GetMissingHighlightColors( _
    ByVal doc As Document, _
    ByVal mode As HighlightRunMode, _
    ByVal restrictAuthors As Boolean, _
    ByVal selectedAuthors As Collection, _
    ByVal useBeforeDate As Boolean, _
    ByVal beforeDate As Date, _
    ByVal includeBeforeDate As Boolean, _
    ByVal useAfterDate As Boolean, _
    ByVal afterDate As Date, _
    ByVal includeAfterDate As Boolean, _
    ByVal optionsForm As frmHighlightOptions) As Collection
    Dim missing As Collection
    Dim rev As Revision

    Set missing = New Collection

    If mode = RunHighlightInsertionsDeletions Then
        For Each rev In doc.Revisions
            If IsRevisionForMode(rev, mode) Then
                If RevisionMatchesFilters( _
                    rev, restrictAuthors, selectedAuthors, _
                    useBeforeDate, beforeDate, includeBeforeDate, _
                    useAfterDate, afterDate, includeAfterDate) Then
                    If GetHardcodedAuthorColor(rev.Author) = -1 Then
                        If Not optionsForm.HasCustomColor(rev.Author) Then
                            If Not CollectionContainsAuthor(missing, rev.Author) Then
                                missing.Add rev.Author
                            End If
                        End If
                    End If
                End If
            End If
        Next rev
    End If

    Set GetMissingHighlightColors = missing
End Function

Private Sub HighlightMatchingRevisions( _
    ByVal matchingRevisions As Collection, _
    ByVal optionsForm As frmHighlightOptions)
    Dim deletionRanges() As Range
    Dim deletionColors() As Long
    Dim rev As Revision
    Dim rng As Range
    Dim i As Long
    Dim j As Long
    Dim deletionCount As Long
    Dim tempRange As Range
    Dim tempColor As Long
    Dim wasTrackingRevisions As Boolean
    Dim savedErrorNumber As Long
    Dim savedErrorSource As String
    Dim savedErrorDescription As String

    ReDim deletionRanges(1 To matchingRevisions.Count)
    ReDim deletionColors(1 To matchingRevisions.Count)

    wasTrackingRevisions = ActiveDocument.TrackRevisions
    On Error GoTo RestoreTrackingState
    ActiveDocument.TrackRevisions = False

    For i = 1 To matchingRevisions.Count
        Set rev = matchingRevisions(i)

        Select Case rev.Type
            Case wdRevisionInsert, wdRevisionMovedTo
                rev.Range.Font.Color = GetAuthorColor(rev.Author, optionsForm)

            Case wdRevisionDelete
                deletionCount = deletionCount + 1
                Set deletionRanges(deletionCount) = rev.Range.Duplicate
                deletionRanges(deletionCount).Collapse wdCollapseStart
                deletionColors(deletionCount) = GetAuthorColor(rev.Author, optionsForm)
        End Select
    Next i

    For i = 1 To deletionCount - 1
        For j = i + 1 To deletionCount
            If deletionRanges(i).Start < deletionRanges(j).Start Then
                Set tempRange = deletionRanges(i)
                Set deletionRanges(i) = deletionRanges(j)
                Set deletionRanges(j) = tempRange

                tempColor = deletionColors(i)
                deletionColors(i) = deletionColors(j)
                deletionColors(j) = tempColor
            End If
        Next j
    Next i

    If deletionCount > 0 Then
        InsertDeletionMarkers deletionRanges, deletionColors, deletionCount
    End If

    ActiveDocument.TrackRevisions = wasTrackingRevisions
    On Error GoTo 0
    Exit Sub

RestoreTrackingState:
    savedErrorNumber = Err.Number
    savedErrorSource = Err.Source
    savedErrorDescription = Err.Description
    ActiveDocument.TrackRevisions = wasTrackingRevisions
    Err.Raise savedErrorNumber, savedErrorSource, savedErrorDescription
End Sub

Private Function GetAuthorColor( _
    ByVal author As String, _
    ByVal optionsForm As frmHighlightOptions) As Long
    Dim hardcodedColor As Long

    hardcodedColor = GetHardcodedAuthorColor(author)
    If hardcodedColor <> -1 Then
        GetAuthorColor = hardcodedColor
    ElseIf optionsForm.HasCustomColor(author) Then
        GetAuthorColor = optionsForm.CustomColor(author)
    Else
        Err.Raise vbObjectError + 2100, "HighlightUserChanges", _
                  "No highlight color was assigned to " & author & "."
    End If
End Function

Private Sub InsertDeletionMarkers( _
    ByRef deletionRanges() As Range, _
    ByRef deletionColors() As Long, _
    ByVal deletionCount As Long)
    Dim wasTrackingRevisions As Boolean
    Dim savedErrorNumber As Long
    Dim savedErrorSource As String
    Dim savedErrorDescription As String
    Dim rng As Range
    Dim i As Long

    wasTrackingRevisions = ActiveDocument.TrackRevisions
    On Error GoTo RestoreTrackingState

    ActiveDocument.TrackRevisions = False
    For i = 1 To deletionCount
        Set rng = deletionRanges(i)
        rng.InsertBefore "[DEL] "
        rng.SetRange rng.Start, rng.Start + 6
        rng.Font.Color = deletionColors(i)
        rng.Font.Bold = True
    Next i

    ActiveDocument.TrackRevisions = wasTrackingRevisions
    On Error GoTo 0
    Exit Sub

RestoreTrackingState:
    savedErrorNumber = Err.Number
    savedErrorSource = Err.Source
    savedErrorDescription = Err.Description
    ActiveDocument.TrackRevisions = wasTrackingRevisions
    Err.Raise savedErrorNumber, savedErrorSource, savedErrorDescription
End Sub

Private Sub AcceptMatchingRevisions(ByVal matchingRevisions As Collection)
    Dim orderedRevisions() As Revision
    Dim rev As Revision
    Dim tempRevision As Revision
    Dim i As Long
    Dim j As Long

    ReDim orderedRevisions(1 To matchingRevisions.Count)
    For i = 1 To matchingRevisions.Count
        Set orderedRevisions(i) = matchingRevisions(i)
    Next i

    For i = 1 To UBound(orderedRevisions) - 1
        For j = i + 1 To UBound(orderedRevisions)
            If orderedRevisions(i).Range.Start < orderedRevisions(j).Range.Start Then
                Set tempRevision = orderedRevisions(i)
                Set orderedRevisions(i) = orderedRevisions(j)
                Set orderedRevisions(j) = tempRevision
            End If
        Next j
    Next i

    For i = 1 To UBound(orderedRevisions)
        Set rev = orderedRevisions(i)
        rev.Accept
    Next i
End Sub

Private Sub ClearScriptHighlightRevisions(ByVal matchingRevisions As Collection)
    Dim orderedRevisions() As Revision
    Dim tempRevision As Revision
    Dim wasTrackingRevisions As Boolean
    Dim expectedRemaining As Long
    Dim totalBefore As Long
    Dim totalAfter As Long
    Dim savedErrorNumber As Long
    Dim savedErrorSource As String
    Dim savedErrorDescription As String
    Dim i As Long
    Dim j As Long

    ReDim orderedRevisions(1 To matchingRevisions.Count)

    For i = 1 To matchingRevisions.Count
        Set orderedRevisions(i) = matchingRevisions(i)
    Next i

    ' Reject from the end of the document backward so earlier positions stay valid.
    For i = 1 To UBound(orderedRevisions) - 1
        For j = i + 1 To UBound(orderedRevisions)
            If orderedRevisions(i).Range.Start < orderedRevisions(j).Range.Start Then
                Set tempRevision = orderedRevisions(i)
                Set orderedRevisions(i) = orderedRevisions(j)
                Set orderedRevisions(j) = tempRevision
            End If
        Next j
    Next i

    totalBefore = CountChaseFontColorRevisions()

    wasTrackingRevisions = ActiveDocument.TrackRevisions
    On Error GoTo RestoreTrackingState
    ActiveDocument.TrackRevisions = False

    For i = 1 To UBound(orderedRevisions)
        orderedRevisions(i).Reject
    Next i

    ActiveDocument.TrackRevisions = wasTrackingRevisions
    On Error GoTo 0

    ' Microsoft's docs say formatting changes cannot be rejected, so verify the count dropped.
    totalAfter = CountChaseFontColorRevisions()
    expectedRemaining = totalBefore - UBound(orderedRevisions)
    If totalAfter > expectedRemaining Then
        MsgBox CStr(totalAfter - expectedRemaining) & " Chase font-color revision(s) were not removed by rejection. Check them manually.", vbExclamation
    End If
    Exit Sub

RestoreTrackingState:
    savedErrorNumber = Err.Number
    savedErrorSource = Err.Source
    savedErrorDescription = Err.Description
    ActiveDocument.TrackRevisions = wasTrackingRevisions
    Err.Raise savedErrorNumber, savedErrorSource, savedErrorDescription
End Sub

Private Function CountChaseFontColorRevisions() As Long
    Dim rev As Revision

    For Each rev In ActiveDocument.Revisions
        If IsChaseFontColorRevision(rev) Then
            CountChaseFontColorRevisions = CountChaseFontColorRevisions + 1
        End If
    Next rev
End Function
