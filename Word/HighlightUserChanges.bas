Attribute VB_Name = "HighlightUserChanges"
Option Explicit

' Sub-based version (no UserForm). Run each public sub from Alt+F8.

Private Enum RevisionRunMode
    RunHighlight = 1
    RunAcceptFormatting = 2
    RunAcceptInsDel = 3
    RunClearBursts = 4
End Enum

Private Type RevisionFilter
    RestrictAuthors As Boolean
    Authors As Collection
    UseBefore As Boolean
    BeforeDate As Date
    IncludeBefore As Boolean
    UseAfter As Boolean
    AfterDate As Date
    IncludeAfter As Boolean
End Type

Private Const APP_TITLE As String = "Highlight User Changes"
Private Const DELETE_MARKER As String = "[DEL] "
Private Const BURST_MAX_GAP_SECONDS As Long = 1

Public Sub HighlightInsertionsDeletions()
    RunRevisionFunction RunHighlight
End Sub

Public Sub AcceptFormattingRevisions()
    RunRevisionFunction RunAcceptFormatting
End Sub

Public Sub AcceptInsertionsDeletions()
    RunRevisionFunction RunAcceptInsDel
End Sub

Public Sub CleanUpChanges()
    RunCleanUp
End Sub

' Declared as Functions, not Subs, so Word's Macros list only shows the four public entry points.
Private Function RunRevisionFunction(ByVal mode As RevisionRunMode)
    Dim revFilter As RevisionFilter
    Dim allAuthors As Collection
    Dim matches As Collection
    Dim revs() As Revision
    Dim customColors As Collection

    If ActiveDocument.Revisions.Count = 0 Then
        MsgBox "This document has no tracked revisions.", vbInformation, APP_TITLE
        Exit Function
    End If

    Set allAuthors = GetDocumentAuthors()

    If Not PromptForFilter(mode, revFilter, allAuthors) Then
        MsgBox "Run cancelled. Nothing was changed.", vbInformation, APP_TITLE
        Exit Function
    End If

    Set matches = FindMatchingRevisions(mode, revFilter)
    If matches.Count = 0 Then
        MsgBox "No revisions match the selected function and filters.", vbInformation, APP_TITLE
        Exit Function
    End If

    If Not ConfirmRun(mode, matches.Count) Then
        MsgBox "Run cancelled. Nothing was changed.", vbInformation, APP_TITLE
        Exit Function
    End If

    ' Sorted by position so ApplyRevisionFunction can work from the end of the document backward.
    LoadSortedRevisions matches, False, revs

    Set customColors = New Collection
    If mode = RunHighlight Then
        If Not PromptForMissingColors(revs, customColors) Then
            MsgBox "Run cancelled. Nothing was changed.", vbInformation, APP_TITLE
            Exit Function
        End If
    End If

    ApplyRevisionFunction mode, revs, customColors

    Select Case mode
        Case RunHighlight
            MsgBox "Highlighted " & matches.Count & " revisions.", vbInformation, APP_TITLE

        Case RunAcceptFormatting, RunAcceptInsDel
            MsgBox "Accepted " & matches.Count & " revisions.", vbInformation, APP_TITLE
    End Select
End Function

Private Function ApplyRevisionFunction(ByVal mode As RevisionRunMode, _
                                  ByRef revs() As Revision, _
                                  ByVal customColors As Collection)
    Dim wasTracking As Boolean
    Dim i As Long
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    wasTracking = ActiveDocument.TrackRevisions
    ' Highlights are tracked so the cleanup step can find and reject them later.
    ActiveDocument.TrackRevisions = (mode = RunHighlight)
    On Error GoTo RestoreTracking

    For i = UBound(revs) To LBound(revs) Step -1
        Select Case mode
            Case RunHighlight
                HighlightOne revs(i), customColors
            Case RunAcceptFormatting, RunAcceptInsDel
                revs(i).Accept
            Case RunClearBursts
                revs(i).Reject
        End Select
    Next i

    ActiveDocument.TrackRevisions = wasTracking
    Exit Function

RestoreTracking:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    ActiveDocument.TrackRevisions = wasTracking
    Err.Raise errNumber, errSource, errDescription
End Function

Private Function FindMatchingRevisions(ByVal mode As RevisionRunMode, _
                                       ByRef revFilter As RevisionFilter) As Collection
    Dim candidates As Collection
    Dim rev As Revision

    Set candidates = New Collection
    For Each rev In ActiveDocument.Revisions
        If IsRevisionForMode(rev, mode) Then
            If RevisionMatchesFilter(rev, revFilter) Then candidates.Add rev
        End If
    Next rev

    Set FindMatchingRevisions = candidates
End Function

Private Function PromptForFilter(ByVal mode As RevisionRunMode, _
                                 ByRef revFilter As RevisionFilter, _
                                 ByVal allAuthors As Collection) As Boolean
    Dim chosen As Collection

    If mode <> RunClearBursts Then
        If MsgBox("Limit this run to specific authors?", vbQuestion + vbYesNo, APP_TITLE) = vbYes Then
            If Not PromptForAuthors(allAuthors, chosen) Then Exit Function
            Set revFilter.Authors = chosen
            revFilter.RestrictAuthors = True
        End If
    End If

    If MsgBox("Limit this run to revisions BEFORE a date?", vbQuestion + vbYesNo, APP_TITLE) = vbYes Then
        If Not PromptForDate("before", revFilter.BeforeDate, revFilter.IncludeBefore) Then Exit Function
        revFilter.UseBefore = True
    End If

    If MsgBox("Limit this run to revisions AFTER a date?", vbQuestion + vbYesNo, APP_TITLE) = vbYes Then
        If Not PromptForDate("after", revFilter.AfterDate, revFilter.IncludeAfter) Then Exit Function
        revFilter.UseAfter = True
    End If

    PromptForFilter = True
End Function

Private Function PromptForAuthors(ByVal allAuthors As Collection, _
                                  ByRef chosen As Collection) As Boolean
    Dim listText As String
    Dim answer As String
    Dim picked() As Boolean
    Dim includeListed As Boolean
    Dim i As Long

    For i = 1 To allAuthors.Count
        listText = listText & i & ". " & allAuthors(i) & vbCrLf
    Next i

    Do
        answer = InputBox("Authors in this document:" & vbCrLf & listText & vbCrLf & _
                          "Enter numbers or ranges, separated by commas (for example: 1, 3-5, 8):", _
                          APP_TITLE & ": select authors")
        If Trim$(answer) = "" Then Exit Function
        If ParseNumberList(answer, allAuthors.Count, picked) Then Exit Do
        MsgBox "Enter numbers or ranges from 1 to " & allAuthors.Count & ", separated by commas (for example: 1, 3-5, 8).", _
               vbExclamation, APP_TITLE
    Loop

    includeListed = (MsgBox("Yes: run only on the numbered authors." & vbCrLf & _
                            "No: run on every author EXCEPT the numbered ones.", _
                            vbQuestion + vbYesNo, APP_TITLE) = vbYes)

    Set chosen = New Collection
    For i = 1 To allAuthors.Count
        If picked(i) = includeListed Then chosen.Add allAuthors(i)
    Next i

    PromptForAuthors = True
End Function

Private Function PromptForDate(ByVal label As String, _
                               ByRef pickedDate As Date, _
                               ByRef inclusive As Boolean) As Boolean
    Dim answer As String

    Do
        answer = InputBox("Enter the " & label & " date (for example 5/1/2025):", APP_TITLE)
        If Trim$(answer) = "" Then Exit Function
        If IsDate(answer) Then Exit Do
        MsgBox "That is not a recognized date.", vbExclamation, APP_TITLE
    Loop

    pickedDate = DateValue(CDate(answer))
    inclusive = (MsgBox("Include revisions made ON " & Format$(pickedDate, "yyyy-mm-dd") & "?" & vbCrLf & _
                        "Yes: inclusive (the date itself counts)" & vbCrLf & _
                        "No: exclusive", vbQuestion + vbYesNo, APP_TITLE) = vbYes)

    PromptForDate = True
End Function

' Accepts single numbers and inclusive ranges, for example "1, 3-5, 8".
Private Function ParseNumberList(ByVal text As String, _
                                 ByVal maxValue As Long, _
                                 ByRef picked() As Boolean) As Boolean
    Dim tokens() As String
    Dim bounds() As String
    Dim lowValue As Long
    Dim highValue As Long
    Dim n As Long
    Dim i As Long

    ReDim picked(1 To maxValue)
    tokens = Split(text, ",")

    For i = LBound(tokens) To UBound(tokens)
        bounds = Split(tokens(i), "-")
        If UBound(bounds) > 1 Then Exit Function
        If Not IsDigitsOnly(bounds(0)) Then Exit Function
        lowValue = CLng(Trim$(bounds(0)))
        highValue = lowValue
        If UBound(bounds) = 1 Then
            If Not IsDigitsOnly(bounds(1)) Then Exit Function
            highValue = CLng(Trim$(bounds(1)))
        End If
        If lowValue < 1 Or highValue > maxValue Or lowValue > highValue Then Exit Function
        For n = lowValue To highValue
            picked(n) = True
        Next n
    Next i

    ParseNumberList = True
End Function

Private Function IsDigitsOnly(ByVal s As String) As Boolean
    s = Trim$(s)
    IsDigitsOnly = (Len(s) > 0 And Len(s) <= 6 And Not s Like "*[!0-9]*")
End Function

Private Function ConfirmRun(ByVal mode As RevisionRunMode, ByVal count As Long) As Boolean
    Select Case mode
        Case RunHighlight
            ConfirmRun = True
        Case Else
            ConfirmRun = (MsgBox("Accept " & count & " matching revisions? This cannot be undone.", _
                                 vbQuestion + vbYesNo + vbDefaultButton2, APP_TITLE) = vbYes)
    End Select
End Function

Private Function GetDocumentAuthors() As Collection
    Dim found As Collection
    Dim sorted As Collection
    Dim names() As String
    Dim rev As Revision
    Dim temp As String
    Dim i As Long
    Dim j As Long

    Set found = New Collection
    For Each rev In ActiveDocument.Revisions
        If Not CollectionHasAuthor(found, rev.Author) Then found.Add rev.Author
    Next rev

    Set sorted = New Collection
    If found.Count = 0 Then
        Set GetDocumentAuthors = sorted
        Exit Function
    End If

    ReDim names(1 To found.Count)
    For i = 1 To found.Count
        names(i) = CStr(found(i))
    Next i

    For i = 2 To UBound(names)
        temp = names(i)
        j = i - 1
        Do While j >= 1
            If StrComp(names(j), temp, vbTextCompare) <= 0 Then Exit Do
            names(j + 1) = names(j)
            j = j - 1
        Loop
        names(j + 1) = temp
    Next i

    For i = 1 To UBound(names)
        sorted.Add names(i)
    Next i

    Set GetDocumentAuthors = sorted
End Function

Private Function CollectionHasAuthor(ByVal authors As Collection, ByVal author As String) As Boolean
    Dim candidate As Variant

    For Each candidate In authors
        If StrComp(CStr(candidate), author, vbBinaryCompare) = 0 Then
            CollectionHasAuthor = True
            Exit Function
        End If
    Next candidate
End Function

Private Function IsRevisionForMode(ByVal rev As Revision, ByVal mode As RevisionRunMode) As Boolean
    Select Case mode
        Case RunHighlight
            IsRevisionForMode = ((rev.Type = wdRevisionInsert) And Not IsDeleteMarkerText(rev)) Or _
                                (rev.Type = wdRevisionMovedTo) Or _
                                (rev.Type = wdRevisionDelete)
        Case RunAcceptFormatting
            IsRevisionForMode = IsFormattingRevision(rev.Type)
        Case RunAcceptInsDel
            IsRevisionForMode = (rev.Type = wdRevisionInsert) Or (rev.Type = wdRevisionDelete)
        Case RunClearBursts
            IsRevisionForMode = IsChaseFontColorRevision(rev)
    End Select
End Function

' Markers are matched by text so highlighting never treats its own [DEL] tags as user changes.
Private Function IsDeleteMarkerText(ByVal rev As Revision) As Boolean
    IsDeleteMarkerText = (Left$(rev.Range.Text, Len(DELETE_MARKER)) = DELETE_MARKER)
End Function

Private Function IsFormattingRevision(ByVal revisionType As WdRevisionType) As Boolean
    Select Case revisionType
        Case wdRevisionProperty, wdRevisionParagraphNumber, wdRevisionStyle, _
             wdRevisionParagraphProperty, wdRevisionTableProperty, _
             wdRevisionSectionProperty, wdRevisionStyleDefinition
            IsFormattingRevision = True
    End Select
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

Private Function CountChaseFontColorRevisions() As Long
    Dim rev As Revision

    For Each rev In ActiveDocument.Revisions
        If IsChaseFontColorRevision(rev) Then
            CountChaseFontColorRevisions = CountChaseFontColorRevisions + 1
        End If
    Next rev
End Function

Private Function RevisionMatchesFilter(ByVal rev As Revision, _
                                       ByRef revFilter As RevisionFilter) As Boolean
    Dim revDate As Date

    If revFilter.RestrictAuthors Then
        If Not CollectionHasAuthor(revFilter.Authors, rev.Author) Then Exit Function
    End If

    revDate = DateValue(rev.Date)

    If revFilter.UseBefore Then
        If revFilter.IncludeBefore Then
            If revDate > revFilter.BeforeDate Then Exit Function
        Else
            If revDate >= revFilter.BeforeDate Then Exit Function
        End If
    End If

    If revFilter.UseAfter Then
        If revFilter.IncludeAfter Then
            If revDate < revFilter.AfterDate Then Exit Function
        Else
            If revDate <= revFilter.AfterDate Then Exit Function
        End If
    End If

    RevisionMatchesFilter = True
End Function

' Loads revisions into an array sorted ascending by Date or by Range.Start.
Private Function LoadSortedRevisions(ByVal items As Collection, ByVal byDate As Boolean, _
                                ByRef revs() As Revision)
    Dim keys() As Double
    Dim tempRev As Revision
    Dim tempKey As Double
    Dim i As Long
    Dim j As Long

    ReDim revs(1 To items.Count)
    ReDim keys(1 To items.Count)

    For i = 1 To items.Count
        Set revs(i) = items(i)
        If byDate Then
            keys(i) = CDbl(revs(i).Date)
        Else
            keys(i) = CDbl(revs(i).Range.Start)
        End If
    Next i

    For i = 2 To UBound(revs)
        Set tempRev = revs(i)
        tempKey = keys(i)
        j = i - 1
        Do While j >= 1
            If keys(j) <= tempKey Then Exit Do
            Set revs(j + 1) = revs(j)
            keys(j + 1) = keys(j)
            j = j - 1
        Loop
        Set revs(j + 1) = tempRev
        keys(j + 1) = tempKey
    Next i
End Function

Private Function GetHardcodedAuthorColor(ByVal author As String) As Long
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

' Palette entries are chosen so none equals a hardcoded author color.
Private Function GetPaletteNames() As Variant
    GetPaletteNames = Array("Dark green", "Orange", "Teal", "Brown", "Dark red", _
                            "Dark blue", "Olive", "Dark cyan", "Dark magenta")
End Function

Private Function GetPaletteColors() As Variant
    GetPaletteColors = Array(RGB(0, 100, 0), RGB(255, 140, 0), RGB(0, 128, 128), _
                             RGB(139, 69, 19), RGB(139, 0, 0), RGB(0, 0, 139), _
                             RGB(128, 128, 0), RGB(0, 139, 139), RGB(139, 0, 139))
End Function

Private Function FindCustomColor(ByVal customColors As Collection, ByVal author As String) As Long
    On Error Resume Next
    FindCustomColor = customColors("[" & author & "]")
    If Err.Number <> 0 Then FindCustomColor = -1
    On Error GoTo 0
End Function

Private Function GetAuthorColor(ByVal author As String, ByVal customColors As Collection) As Long
    Dim hardcodedColor As Long

    hardcodedColor = GetHardcodedAuthorColor(author)
    If hardcodedColor <> -1 Then
        GetAuthorColor = hardcodedColor
    ElseIf FindCustomColor(customColors, author) <> -1 Then
        GetAuthorColor = FindCustomColor(customColors, author)
    Else
        Err.Raise vbObjectError + 2100, APP_TITLE, "No highlight color was assigned to " & author & "."
    End If
End Function

Private Function PromptForMissingColors(ByRef revs() As Revision, _
                                        ByVal customColors As Collection) As Boolean
    Dim i As Long
    Dim author As String
    Dim color As Long

    For i = LBound(revs) To UBound(revs)
        author = revs(i).Author
        If GetHardcodedAuthorColor(author) = -1 Then
            If FindCustomColor(customColors, author) = -1 Then
                If Not PromptForAuthorColor(author, color) Then Exit Function
                customColors.Add color, "[" & author & "]"
            End If
        End If
    Next i

    PromptForMissingColors = True
End Function

Private Function PromptForAuthorColor(ByVal author As String, ByRef color As Long) As Boolean
    Dim names As Variant
    Dim colors As Variant
    Dim listText As String
    Dim answer As String
    Dim i As Long
    Dim choice As Long

    names = GetPaletteNames()
    colors = GetPaletteColors()

    For i = LBound(names) To UBound(names)
        listText = listText & (i - LBound(names) + 1) & ". " & names(i) & vbCrLf
    Next i

    Do
        answer = InputBox("Choose a highlight color for " & author & ". Enter a number from the list:" & _
                          vbCrLf & vbCrLf & listText & vbCrLf & _
                          "(Several authors may share one color.)", APP_TITLE & ": assign color")
        If Trim$(answer) = "" Then Exit Function
        If Len(Trim$(answer)) > 0 And Not Trim$(answer) Like "*[!0-9]*" Then
            choice = CLng(Trim$(answer))
            If choice >= 1 And choice <= UBound(names) - LBound(names) + 1 Then
                color = colors(LBound(colors) + choice - 1)
                PromptForAuthorColor = True
                Exit Function
            End If
        End If
        MsgBox "Enter a number from 1 to " & (UBound(names) - LBound(names) + 1) & ".", _
               vbExclamation, APP_TITLE
    Loop
End Function

Private Function HighlightOne(ByVal rev As Revision, ByVal customColors As Collection)
    Dim color As Long
    Dim rng As Range
    Dim markerStart As Long

    color = GetAuthorColor(rev.Author, customColors)

    If rev.Type = wdRevisionDelete Then
        markerStart = rev.Range.Start
        If Not HasMarkerBefore(markerStart) Then
            Set rng = rev.Range.Duplicate
            rng.Collapse wdCollapseStart
            rng.InsertBefore DELETE_MARKER
            Set rng = ActiveDocument.Range(markerStart, markerStart + Len(DELETE_MARKER))
            rng.Font.Color = color
            rng.Font.Bold = True
        End If
    Else
        rev.Range.Font.Color = color
    End If
End Function

' Stops a second run from stacking another [DEL] tag in front of a deletion that already has one.
Private Function HasMarkerBefore(ByVal pos As Long) As Boolean
    If pos < Len(DELETE_MARKER) Then Exit Function
    HasMarkerBefore = (ActiveDocument.Range(pos - Len(DELETE_MARKER), pos).Text = DELETE_MARKER)
End Function

' One run for both cleanups: rejects Chase font-color bursts, then removes [DEL] markers.
' Markers are matched by text, so they are found whether or not they are tracked.
Private Function RunCleanUp()
    Dim revFilter As RevisionFilter
    Dim chaseMatches As Collection
    Dim chaseRevs() As Revision
    Dim customColors As Collection
    Dim markers As Collection
    Dim markerCount As Long
    Dim untimed As Long
    Dim bursts As Long
    Dim chaseBefore As Long
    Dim chaseAfter As Long
    Dim expectedAfter As Long
    Dim remaining As Long
    Dim chaseLine As String
    Dim markerLine As String
    Dim message As String
    Dim icon As Long

    If Not PromptForFilter(RunClearBursts, revFilter, New Collection) Then
        MsgBox "Run cancelled. Nothing was changed.", vbInformation, APP_TITLE
        Exit Function
    End If

    Set chaseMatches = FindMatchingRevisions(RunClearBursts, revFilter)
    Set markers = FindDeleteMarkers()
    If chaseMatches.Count = 0 And markers.Count = 0 Then
        MsgBox "Nothing to clean up. No Chase font-color bursts or [DEL] markers were found.", _
               vbInformation, APP_TITLE
        Exit Function
    End If

    markerCount = markers.Count
    bursts = CountMarkerBursts(markers, untimed)

    If Not ConfirmCleanUp(chaseMatches.Count, markerCount, bursts, untimed) Then
        MsgBox "Run cancelled. Nothing was changed.", vbInformation, APP_TITLE
        Exit Function
    End If

    chaseBefore = CountChaseFontColorRevisions()
    If chaseMatches.Count > 0 Then
        LoadSortedRevisions chaseMatches, False, chaseRevs
        Set customColors = New Collection
        ApplyRevisionFunction RunClearBursts, chaseRevs, customColors
    End If

    ' Rejecting formatting does not move text, but re-find the markers so their ranges are current.
    Set markers = FindDeleteMarkers()
    If markers.Count > 0 Then DeleteMarkers markers

    icon = vbInformation

    If chaseMatches.Count > 0 Then
        ' Microsoft's docs say formatting changes cannot be rejected, so confirm the count dropped.
        chaseAfter = CountChaseFontColorRevisions()
        expectedAfter = chaseBefore - chaseMatches.Count
        If chaseAfter > expectedAfter Then
            chaseLine = CStr(chaseAfter - expectedAfter) & " Chase font-color revision(s) were not removed " & _
                        "by rejection. Check them manually."
            icon = vbExclamation
        Else
            chaseLine = "Rejected " & chaseMatches.Count & " Chase font-color revisions."
        End If
    End If

    If markerCount > 0 Then
        remaining = FindDeleteMarkers().Count
        If remaining > 0 Then
            markerLine = remaining & " [DEL] marker(s) remain. Check them manually."
            icon = vbExclamation
        Else
            markerLine = "Removed " & markerCount & " [DEL] markers."
        End If
    End If

    message = chaseLine
    If chaseLine <> "" And markerLine <> "" Then message = message & vbCrLf
    message = message & markerLine
    MsgBox message, icon, APP_TITLE
End Function

Private Function ConfirmCleanUp(ByVal chaseCount As Long, ByVal markerCount As Long, _
                                ByVal bursts As Long, ByVal untimed As Long) As Boolean
    Dim message As String

    If chaseCount > 0 Then
        message = "Chase Lynk font-color revisions to reject: " & chaseCount & ". " & _
                  "Font colors will return to what they were before each change." & vbCrLf & vbCrLf
    End If

    If markerCount > 0 Then
        message = message & "[DEL] markers to remove: " & markerCount & " (" & bursts & " burst(s) of timed edits, " & _
                  untimed & " with no edit timestamp). Removal is permanent and not tracked." & vbCrLf & vbCrLf
    End If

    ConfirmCleanUp = (MsgBox(message & "Continue?", vbExclamation + vbYesNo + vbDefaultButton2, APP_TITLE) = vbYes)
End Function

Private Function FindDeleteMarkers() As Collection
    Dim found As Collection
    Dim rng As Range

    Set found = New Collection
    Set rng = ActiveDocument.Content

    With rng.Find
        .ClearFormatting
        .Text = DELETE_MARKER
        .MatchCase = True
        .MatchWildcards = False
        .Forward = True
        .Wrap = wdFindStop
    End With

    Do While rng.Find.Execute
        found.Add rng.Duplicate
        rng.Collapse wdCollapseEnd
    Loop

    Set FindDeleteMarkers = found
End Function

' Counts bursts among markers that carry a revision date. Markers without one are counted in untimed.
Private Function CountMarkerBursts(ByVal markers As Collection, ByRef untimed As Long) As Long
    Dim dates() As Double
    Dim rng As Range
    Dim timedCount As Long
    Dim i As Long
    Dim j As Long
    Dim temp As Double

    ReDim dates(1 To markers.Count)
    For i = 1 To markers.Count
        Set rng = markers(i)
        If rng.Revisions.Count > 0 Then
            timedCount = timedCount + 1
            dates(timedCount) = CDbl(rng.Revisions(1).Date)
        Else
            untimed = untimed + 1
        End If
    Next i

    If timedCount = 0 Then Exit Function

    For i = 2 To timedCount
        temp = dates(i)
        j = i - 1
        Do While j >= 1
            If dates(j) <= temp Then Exit Do
            dates(j + 1) = dates(j)
            j = j - 1
        Loop
        dates(j + 1) = temp
    Next i

    CountMarkerBursts = 1
    For i = 2 To timedCount
        If Round((dates(i) - dates(i - 1)) * 86400#) > BURST_MAX_GAP_SECONDS Then
            CountMarkerBursts = CountMarkerBursts + 1
        End If
    Next i
End Function

' Deletes from the end backward. Tracking is off so the removal is permanent, and it is restored on error.
Private Function DeleteMarkers(ByVal markers As Collection)
    Dim wasTracking As Boolean
    Dim i As Long
    Dim rng As Range
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    wasTracking = ActiveDocument.TrackRevisions
    ActiveDocument.TrackRevisions = False
    On Error GoTo RestoreTracking

    For i = markers.Count To 1 Step -1
        Set rng = markers(i)
        rng.Delete
    Next i

    ActiveDocument.TrackRevisions = wasTracking
    Exit Function

RestoreTracking:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    ActiveDocument.TrackRevisions = wasTracking
    Err.Raise errNumber, errSource, errDescription
End Function
