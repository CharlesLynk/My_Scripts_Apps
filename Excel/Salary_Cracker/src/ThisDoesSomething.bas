Attribute VB_Name = "ThisDoesSomething"
'==========================================
' Module: ThisDoesSomething
' Purpose: Advance a formula reference and refresh table filters.
' Scope: Salary_Cracker workbook automation helper.
'==========================================

Private Sub ThisDoesSomething() 'ByVal Target As Range)
Attribute ThisDoesSomething.VB_ProcData.VB_Invoke_Func = "k\n14"
    ' Cell containing the formula to advance.
    Dim FormulaCell As Range
    Set FormulaCell = ThisWorkbook.Sheets("GuessAndCheck1").Range("B4")

    ' Event trigger check kept for future worksheet-change wiring.
    'If Not Intersect(Target, FormulaCell) Is Nothing Then
        ' Read current formula.
        Dim CurrentFormula As String
        CurrentFormula = FormulaCell.Formula

        ' Extract row number from the existing B-reference.
        Dim CurrentRow As Long
        CurrentRow = Mid(CurrentFormula, InStrRev(CurrentFormula, "B") + 1)

        ' Advance to the next row.
        CurrentRow = CurrentRow + 1

        ' Build and apply updated formula.
        Dim NewFormula As String
        NewFormula = "=Semiprimes!B" & CurrentRow

        FormulaCell.Formula = NewFormula
    'End If

    ' Reapply table filter after formula update.
    With ActiveWorkbook.Worksheets("GuessAndCheck1").ListObjects("Table4")
         .AutoFilter.ApplyFilter
    End With
End Sub

