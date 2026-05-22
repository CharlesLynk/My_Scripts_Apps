Attribute VB_Name = "StringWrapVBA"
'
'==========================================
' Module: StringWrapVBA
' Purpose: Wrap long delimiter-separated strings into VBA continuation lines.
' Scope: String formatting helper for code-generation style text blocks.
'==========================================

'================== PUBLIC FUNCTIONS ==================
Public Function StringWrapVBA(txt As String, pos As Long) As String
    Dim delim As String: delim = ""","""
    Dim nearest As Long: nearest = -1
    Dim i As Long
    Dim minDist As Long: minDist = Len(txt) + 1
    Dim splitPos As Long
    Dim part1 As String, part2 As String
    Dim remainingDelims As Long

    ' Find nearest delimiter
    For i = 1 To Len(txt) - 2
        If Mid(txt, i, 3) = delim Then
            If Abs(i - pos) < minDist Then
                minDist = Abs(i - pos)
                nearest = i
            End If
        End If
    Next i

    If nearest = -1 Then
        StringWrapVBA = TrimTrailingComma(txt)
        Exit Function
    End If

    splitPos = nearest + 1
    part1 = Left(txt, splitPos) & " _"
    part2 = Mid(txt, splitPos + 1)

    ' Count remaining delimiters in part2
    remainingDelims = 0
    For i = 1 To Len(part2) - 2
        If Mid(part2, i, 3) = delim Then
            remainingDelims = remainingDelims + 1
        End If
    Next i

    If Len(part2) <= pos Or remainingDelims = 0 Then
        StringWrapVBA = part1 & IIf(part2 <> "", vbLf & TrimTrailingComma(part2), "")
    Else
        StringWrapVBA = part1 & vbLf & StringWrapVBA(part2, pos)
    End If
End Function

'================== PRIVATE HELPERS ==================
' Remove a trailing comma when present.
Private Function TrimTrailingComma(s As String) As String
    If Right(s, 1) = "," Then
        TrimTrailingComma = Left(s, Len(s) - 1)
    Else
        TrimTrailingComma = s
    End If
End Function


