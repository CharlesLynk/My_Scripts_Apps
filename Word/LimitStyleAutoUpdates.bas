Attribute VB_Name = "LimitStyleAutoUpdates"
Sub LimitStyleAutoUpdates()
    'Charles Kenyon
    ' https://answers.microsoft.com/en-us/msoffice/forum/all/how-do-i-keep-from-automatically-updating-style-in/55ff38b8-b83f-4bd8-9f73-0acb5fd241af
    'Changes all the TOC styles and the Table of Authorities/Figures styles to Automatically Update, others not
    ' This should be the default setting
    '
    Dim aStyle As Style
    On Error Resume Next
    For Each aStyle In ActiveDocument.Styles
        Select Case aStyle.NameLocal
            Case "Table of Authorities", "Table of Figures", "TOC 1", "TOC 2", "TOC 3", "TOC 4", "TOC 5", "TOC 6", "TOC 7", "TOC 8", "TOC 9"
                Let aStyle.AutomaticallyUpdate = True
            Case Else
                Let aStyle.AutomaticallyUpdate = False
        End Select
    Next aStyle
    Set aStyle = Nothing
    On Error GoTo -1
End Sub



