Sub HighlightUserChanges()

    Dim rev As Revision
    Dim AuthorColor As WdColor

    Dim DelPos() As Long
    Dim DelColor() As Long
    Dim DelCount As Long

    Dim rng As Range
    Dim i As Long

    ' PASS 1
    ' Process insertions and collect deletion locations

    For Each rev In ActiveDocument.Revisions

        AuthorColor = -1

        Select Case rev.Author

            Case "Chase Lynk"
                AuthorColor = wdColorRed

            Case "Hossein Targhi"
                AuthorColor = wdColorBlue

            Case "Tolulope Adekanye"
                AuthorColor = wdColorViolet

        End Select

        If AuthorColor <> -1 Then

            Select Case rev.Type

                ' Color insertions
                Case wdRevisionInsert

                    rev.Range.Font.Color = AuthorColor

                ' Treat moved text like inserted text
                Case wdRevisionMovedTo

                    rev.Range.Font.Color = AuthorColor

                ' Store deletion locations for later
                Case wdRevisionDelete

                    DelCount = DelCount + 1

                    ReDim Preserve DelPos(1 To DelCount)
                    ReDim Preserve DelColor(1 To DelCount)

                    DelPos(DelCount) = rev.Range.Start
                    DelColor(DelCount) = AuthorColor

            End Select

        End If

    Next rev

    ' PASS 2
    ' Insert deletion markers from end of document to beginning

    For i = DelCount To 1 Step -1

        Set rng = ActiveDocument.Range(DelPos(i), DelPos(i))

        rng.InsertBefore "[DEL] "

        rng.SetRange rng.Start, rng.Start + 6

        rng.Font.Color = DelColor(i)
        rng.Font.Bold = True

    Next i

End Sub
`

' Suggested colors
'    wdColorRed
'    wdColorBlue
'    wdColorDarkGreen
'    wdColorViolet
'    wdColorOrange
'    wdColorTeal
'    wdColorBrown
'    wdColorDarkRed


