Attribute VB_Name = "MasterWorksMapping"
Sub MasterWorksKeepOption()
    Application.Run "KeepOption"
End Sub

Sub MasterWorksKeepAllOptions()
    Application.Run "KeepAllOptions"
End Sub

Sub MasterWorksNextOption()
    Application.Run "NextOption"
End Sub

Sub MasterWorksPreviousOption()
    Application.Run "PreviousOption"
End Sub

Sub MasterWorksDeleteOption()
    Application.Run "DeleteOption"
End Sub

Sub MasterWorksDeleteAllOptionsInParagraph()
    Application.Run "DeleteAllOptionsInParagraph"
End Sub

Sub MasterWorksDeleteParagraph()
    Application.Run "DeleteParagraph"
End Sub

Sub MasterWorksDeleteSubParas()
    Application.Run "DeleteSubParas"
End Sub

Sub MasterWorksMoveDownOneParagraph()
    Application.Run "MoveDownOneParagraph"
End Sub


Sub MasterWorksMoveUpOneParagraph()
    Application.Run "MoveUpOneParagraph"
End Sub


Sub MasterWorkssEditorsNotes()
    Application.Run "sEditorsNotes"
End Sub


Sub MasterWorkssOfficeMasterNotes()
    Application.Run "sOfficeMasterNotes"
End Sub


Sub MasterWorkssProjectNotes()
    Application.Run "sProjectNotes"
End Sub

Sub MasterWorksAddSpace()
    Application.Run "AddSpace"
End Sub

Sub MasterWorksRemoveSpace()
    Application.Run "RemoveSpace"
End Sub


Sub MasterWorksToggleAllNotes()
    Application.Run "sProjectNotes"
    Application.Run "sOfficeMasterNotes"
    Application.Run "sEditorsNotes"
    ActiveWindow.ScrollIntoView Selection.Range, True
End Sub

Sub ToggleRevisions()
    With ActiveWindow.View.RevisionsFilter
        If .Markup = wdRevisionsMarkupAll Then
            .Markup = wdRevisionsMarkupNone
        Else
            .Markup = wdRevisionsMarkupAll
        End If
        .View = wdRevisionsViewFinal
    End With
    ActiveWindow.ScrollIntoView Selection.Range, True
End Sub



