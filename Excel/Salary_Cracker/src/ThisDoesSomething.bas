Attribute VB_Name = "ThisDoesSomething"
Private Sub ThisDoesSomething() 'ByVal Target As Range)
Attribute ThisDoesSomething.VB_ProcData.VB_Invoke_Func = "k\n14"
    ' Define the cell that contains the formula you want to change
    Dim FormulaCell As Range
    Set FormulaCell = ThisWorkbook.Sheets("GuessAndCheck1").Range("B4")

    ' Check if the changed cell is the one that triggers the formula change
    'If Not Intersect(Target, FormulaCell) Is Nothing Then
        ' Get the current formula from the cell
        Dim CurrentFormula As String
        CurrentFormula = FormulaCell.Formula

        ' Extract the current row number from the formula
        Dim CurrentRow As Long
        CurrentRow = Mid(CurrentFormula, InStrRev(CurrentFormula, "B") + 1)

        ' Increment the row number
        CurrentRow = CurrentRow + 1

        ' Build the new formula
        Dim NewFormula As String
        NewFormula = "=Semiprimes!B" & CurrentRow

        ' Update the cell with the new formula
        FormulaCell.Formula = NewFormula
    'End If

    ' Apply the filter
    With ActiveWorkbook.Worksheets("GuessAndCheck1").ListObjects("Table4")
         .AutoFilter.ApplyFilter
    End With
End Sub

