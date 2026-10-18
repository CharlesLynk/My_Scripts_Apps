VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHighlightOptions
   Caption         =   "Highlight / Accept Revisions"
   ClientHeight    =   6900
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7800
   StartUpPosition =   1
End
Attribute VB_Name = "frmHighlightOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mHighlightMode As MSForms.OptionButton
Private WithEvents mAcceptFormattingMode As MSForms.OptionButton
Private WithEvents mAcceptInsertionsMode As MSForms.OptionButton
Private WithEvents mClearMode As MSForms.OptionButton
Private WithEvents mAuthorFilter As MSForms.CheckBox
Private WithEvents mBeforeFilter As MSForms.CheckBox
Private WithEvents mAfterFilter As MSForms.CheckBox
Private WithEvents mAuthorList As MSForms.ListBox
Private WithEvents mBeforeDateBox As MSForms.TextBox
Private WithEvents mAfterDateBox As MSForms.TextBox
Private WithEvents mIncludeBeforeDate As MSForms.CheckBox
Private WithEvents mIncludeAfterDate As MSForms.CheckBox
Private WithEvents mColorAuthorCombo As MSForms.ComboBox
Private WithEvents mColorPalette As MSForms.ComboBox
Private WithEvents mSetColorButton As MSForms.CommandButton
Private WithEvents mRunButton As MSForms.CommandButton
Private WithEvents mCancelButton As MSForms.CommandButton
Private WithEvents mBackButton As MSForms.CommandButton

Private mInstructions As MSForms.Label
Private mColorStatus As MSForms.Label
Private mAuthors As Collection
Private mAuthorColors As Object
Private mColorRequirements As Collection
Private mColorSetupMode As Boolean

Public RunMode As HighlightRunMode
Public AuthorFilterEnabled As Boolean
Public SelectedAuthors As Collection
Public UseBeforeDate As Boolean
Public BeforeDate As Date
Public IncludeBeforeDate As Boolean
Public UseAfterDate As Boolean
Public AfterDate As Date
Public IncludeAfterDate As Boolean
Public WasCancelled As Boolean
Public ReturnedToOptions As Boolean

Private Sub UserForm_Initialize()
    Me.Caption = "Highlight / Accept Revisions"
    Me.Width = 570
    Me.Height = 454

    Set mAuthorColors = CreateObject("Scripting.Dictionary")
    mAuthorColors.CompareMode = vbBinaryCompare
    Set mAuthors = New Collection
    Set mColorRequirements = New Collection

    AddLabel "lblMode", "Choose one function:", 12, 8, 250, 18
    Set mHighlightMode = AddOption("optHighlight", "Highlight insertions / deletions", 20, 28, 270, 18)
    Set mAcceptFormattingMode = AddOption("optAcceptFormatting", "Accept formatting revisions", 20, 50, 270, 18)
    Set mAcceptInsertionsMode = AddOption("optAcceptInsertions", "Accept insertions / deletions", 20, 72, 270, 18)
    Set mClearMode = AddOption("optClearHighlights", "Clear script-generated color bursts", 20, 94, 350, 18)
    mHighlightMode.GroupName = "RevisionFunction"
    mAcceptFormattingMode.GroupName = "RevisionFunction"
    mAcceptInsertionsMode.GroupName = "RevisionFunction"
    mClearMode.GroupName = "RevisionFunction"

    Set mAuthorFilter = AddCheckBox("chkAuthorFilter", "Restrict to selected authors", 12, 122, 260, 18)
    AddLabel "lblAuthors", "Authors (multi-select when restricted):", 22, 144, 340, 18
    Set mAuthorList = Me.Controls.Add("Forms.ListBox.1", "lstAuthors", True)
    mAuthorList.Left = 22
    mAuthorList.Top = 164
    mAuthorList.Width = 330
    mAuthorList.Height = 68
    mAuthorList.MultiSelect = fmMultiSelectMulti
    mAuthorList.Enabled = False

    AddLabel "lblColor", "Highlight colors (hardcoded author colors are reserved):", 12, 237, 500, 18
    Set mColorAuthorCombo = Me.Controls.Add("Forms.ComboBox.1", "cmbColorAuthor", True)
    mColorAuthorCombo.Left = 22
    mColorAuthorCombo.Top = 258
    mColorAuthorCombo.Width = 175
    mColorAuthorCombo.Style = fmStyleDropDownList
    Set mColorPalette = Me.Controls.Add("Forms.ComboBox.1", "cmbColorPalette", True)
    mColorPalette.Left = 207
    mColorPalette.Top = 258
    mColorPalette.Width = 130
    mColorPalette.Style = fmStyleDropDownList
    mColorPalette.ColumnCount = 2
    mColorPalette.ColumnWidths = "105 pt;0 pt"
    Set mSetColorButton = AddButton("cmdSetColor", "Assign color", 347, 256, 98, 24)
    Set mColorStatus = AddLabel("lblColorStatus", "Select an author to see its color.", 22, 283, 430, 18)

    AddLabel "lblDateHelp", "Dates use your Windows date format; comparisons use calendar dates.", 12, 304, 520, 18
    Set mAfterFilter = AddCheckBox("chkAfterFilter", "After:", 12, 326, 85, 20)
    Set mAfterDateBox = AddTextBox("txtAfterDate", 100, 324, 105, 22)
    Set mIncludeAfterDate = AddCheckBox("chkIncludeAfter", "Include this date", 218, 326, 150, 20)
    mIncludeAfterDate.Value = True
    Set mBeforeFilter = AddCheckBox("chkBeforeFilter", "Before:", 12, 352, 85, 20)
    Set mBeforeDateBox = AddTextBox("txtBeforeDate", 100, 350, 105, 22)
    Set mIncludeBeforeDate = AddCheckBox("chkIncludeBefore", "Include this date", 218, 352, 150, 20)
    mIncludeBeforeDate.Value = True

    Set mInstructions = AddLabel("lblInstructions", "", 12, 378, 535, 18)
    Set mRunButton = AddButton("cmdRun", "Run", 270, 404, 86, 26)
    Set mBackButton = AddButton("cmdBack", "Back to options", 362, 404, 92, 26)
    Set mCancelButton = AddButton("cmdCancel", "Cancel", 460, 404, 86, 26)
    mBackButton.Visible = False

    AddColorOption "Dark green", RGB(0, 100, 0)
    AddColorOption "Orange", RGB(255, 140, 0)
    AddColorOption "Teal", RGB(0, 128, 128)
    AddColorOption "Brown", RGB(139, 69, 19)
    AddColorOption "Dark red", RGB(139, 0, 0)
    AddColorOption "Dark blue", RGB(0, 0, 139)
    AddColorOption "Olive", RGB(128, 128, 0)
    AddColorOption "Dark cyan", RGB(0, 100, 100)
    AddColorOption "Dark magenta", RGB(139, 0, 139)
    mColorPalette.ListIndex = 0

    mHighlightMode.Value = True
    UpdateEnabledState
End Sub

Public Sub InitializeAuthors(ByVal authors As Collection)
    Dim author As Variant

    Set mAuthors = New Collection
    mAuthorList.Clear
    mColorAuthorCombo.Clear

    For Each author In authors
        mAuthors.Add CStr(author)
        mAuthorList.AddItem DisplayAuthor(CStr(author))
        mColorAuthorCombo.AddItem DisplayAuthor(CStr(author))
    Next author

    If mColorAuthorCombo.ListCount > 0 Then mColorAuthorCombo.ListIndex = 0
    UpdateColorStatus
End Sub

Public Sub SetColorRequirements(ByVal authors As Collection)
    Set mColorRequirements = authors
    mColorSetupMode = True
    mInstructions.Caption = "Assign a color to each listed author, or go back to change the filters."
    mBackButton.Visible = True
    UpdateEnabledState
End Sub

Public Sub ClearColorRequirements()
    Set mColorRequirements = New Collection
    mColorSetupMode = False
    mInstructions.Caption = "Choose one function and any optional filters."
    mBackButton.Visible = False
    UpdateEnabledState
End Sub

Public Function HasCustomColor(ByVal author As String) As Boolean
    HasCustomColor = mAuthorColors.Exists(author)
End Function

Public Function CustomColor(ByVal author As String) As Long
    If Not mAuthorColors.Exists(author) Then
        Err.Raise vbObjectError + 2101, "frmHighlightOptions.CustomColor", _
                  "No custom color is assigned to " & author & "."
    End If
    CustomColor = CLng(mAuthorColors.Item(author))
End Function

Private Function AddLabel( _
    ByVal controlName As String, ByVal caption As String, _
    ByVal left As Single, ByVal top As Single, _
    ByVal width As Single, ByVal height As Single) As MSForms.Label
    Set AddLabel = Me.Controls.Add("Forms.Label.1", controlName, True)
    With AddLabel
        .Caption = caption
        .Left = left
        .Top = top
        .Width = width
        .Height = height
    End With
End Function

Private Function AddOption( _
    ByVal controlName As String, ByVal caption As String, _
    ByVal left As Single, ByVal top As Single, _
    ByVal width As Single, ByVal height As Single) As MSForms.OptionButton
    Set AddOption = Me.Controls.Add("Forms.OptionButton.1", controlName, True)
    With AddOption
        .Caption = caption
        .Left = left
        .Top = top
        .Width = width
        .Height = height
    End With
End Function

Private Function AddCheckBox( _
    ByVal controlName As String, ByVal caption As String, _
    ByVal left As Single, ByVal top As Single, _
    ByVal width As Single, ByVal height As Single) As MSForms.CheckBox
    Set AddCheckBox = Me.Controls.Add("Forms.CheckBox.1", controlName, True)
    With AddCheckBox
        .Caption = caption
        .Left = left
        .Top = top
        .Width = width
        .Height = height
    End With
End Function

Private Function AddTextBox( _
    ByVal controlName As String, ByVal left As Single, _
    ByVal top As Single, ByVal width As Single, _
    ByVal height As Single) As MSForms.TextBox
    Set AddTextBox = Me.Controls.Add("Forms.TextBox.1", controlName, True)
    With AddTextBox
        .Left = left
        .Top = top
        .Width = width
        .Height = height
    End With
End Function

Private Function AddButton( _
    ByVal controlName As String, ByVal caption As String, _
    ByVal left As Single, ByVal top As Single, _
    ByVal width As Single, ByVal height As Single) As MSForms.CommandButton
    Set AddButton = Me.Controls.Add("Forms.CommandButton.1", controlName, True)
    With AddButton
        .Caption = caption
        .Left = left
        .Top = top
        .Width = width
        .Height = height
    End With
End Function

Private Sub AddColorOption(ByVal colorName As String, ByVal colorValue As Long)
    Dim row As Long

    mColorPalette.AddItem colorName
    row = mColorPalette.ListCount - 1
    mColorPalette.List(row, 1) = CStr(colorValue)
End Sub

Private Function DisplayAuthor(ByVal author As String) As String
    If Len(author) = 0 Then
        DisplayAuthor = "(Unknown author)"
    Else
        DisplayAuthor = author
    End If
End Function

Private Sub UpdateEnabledState()
    Dim highlightModeEnabled As Boolean
    Dim clearModeEnabled As Boolean
    Dim filtersEnabled As Boolean

    highlightModeEnabled = mHighlightMode.Value
    clearModeEnabled = mClearMode.Value
    filtersEnabled = Not mColorSetupMode

    mHighlightMode.Enabled = filtersEnabled
    mAcceptFormattingMode.Enabled = filtersEnabled
    mAcceptInsertionsMode.Enabled = filtersEnabled
    mClearMode.Enabled = filtersEnabled
    mAuthorFilter.Enabled = filtersEnabled And Not clearModeEnabled
    mAuthorList.Enabled = filtersEnabled And Not clearModeEnabled And mAuthorFilter.Value
    mBeforeFilter.Enabled = filtersEnabled
    mAfterFilter.Enabled = filtersEnabled
    mBeforeDateBox.Enabled = filtersEnabled And mBeforeFilter.Value
    mAfterDateBox.Enabled = filtersEnabled And mAfterFilter.Value
    mIncludeBeforeDate.Enabled = filtersEnabled And mBeforeFilter.Value
    mIncludeAfterDate.Enabled = filtersEnabled And mAfterFilter.Value
    mColorAuthorCombo.Enabled = highlightModeEnabled And Not mColorSetupMode
    mColorPalette.Enabled = highlightModeEnabled And Not mColorSetupMode
    mSetColorButton.Enabled = highlightModeEnabled And Not mColorSetupMode
    mRunButton.Caption = IIf(mColorSetupMode, "Continue", "Run")
    mBackButton.Visible = mColorSetupMode

    If mColorSetupMode Then
        mInstructions.Caption = "Assign a color to each listed author, or go back to change the filters."
    ElseIf clearModeEnabled Then
        mInstructions.Caption = "Find Chase Lynk color bursts (3+ revisions, gaps <= 1 second)."
    ElseIf highlightModeEnabled Then
        mInstructions.Caption = "Highlight insertions/deletions; unassigned authors will be prompted for a color."
    Else
        mInstructions.Caption = "Only the selected revision type will be accepted."
    End If
End Sub

Private Sub UpdateColorStatus()
    Dim authorIndex As Long
    Dim author As String
    Dim hardcodedColor As Long

    authorIndex = mColorAuthorCombo.ListIndex
    If authorIndex < 0 Then
        mColorStatus.Caption = "No revision authors are available."
        Exit Sub
    End If

    author = CStr(mAuthors(authorIndex + 1))
    hardcodedColor = GetHardcodedAuthorColor(author)
    If hardcodedColor <> -1 Then
        mColorStatus.Caption = "Hardcoded color: " & HardcodedColorName(author) & " (reserved)."
    ElseIf mAuthorColors.Exists(author) Then
        mColorStatus.Caption = "Assigned color: " & PaletteColorName(CLng(mAuthorColors.Item(author))) & "."
    Else
        mColorStatus.Caption = "No color assigned yet."
    End If
End Sub

Private Function HardcodedColorName(ByVal author As String) As String
    Select Case author
        Case "Chase Lynk"
            HardcodedColorName = "Red"
        Case "Hossein Targhi"
            HardcodedColorName = "Blue"
        Case "Tolulope Adekanye"
            HardcodedColorName = "Violet"
    End Select
End Function

Private Function PaletteColorName(ByVal colorValue As Long) As String
    Dim i As Long

    For i = 0 To mColorPalette.ListCount - 1
        If CLng(mColorPalette.List(i, 1)) = colorValue Then
            PaletteColorName = CStr(mColorPalette.List(i, 0))
            Exit Function
        End If
    Next i
    PaletteColorName = "Custom"
End Function

Private Function FindAuthorIndex(ByVal author As String) As Long
    Dim i As Long

    For i = 1 To mAuthors.Count
        If StrComp(CStr(mAuthors(i)), author, vbBinaryCompare) = 0 Then
            FindAuthorIndex = i - 1
            Exit Function
        End If
    Next i
    FindAuthorIndex = -1
End Function

Private Sub mHighlightMode_Click()
    UpdateEnabledState
End Sub

Private Sub mAcceptFormattingMode_Click()
    UpdateEnabledState
End Sub

Private Sub mAcceptInsertionsMode_Click()
    UpdateEnabledState
End Sub

Private Sub mClearMode_Click()
    UpdateEnabledState
End Sub

Private Sub mAuthorFilter_Click()
    UpdateEnabledState
End Sub

Private Sub mBeforeFilter_Click()
    UpdateEnabledState
End Sub

Private Sub mAfterFilter_Click()
    UpdateEnabledState
End Sub

Private Sub mAuthorList_Click()
    UpdateColorStatus
End Sub

Private Sub mColorAuthorCombo_Change()
    UpdateColorStatus
End Sub

Private Sub mSetColorButton_Click()
    Dim authorIndex As Long
    Dim author As String
    Dim colorValue As Long

    authorIndex = mColorAuthorCombo.ListIndex
    If authorIndex < 0 Then
        MsgBox "Select an author before assigning a color.", vbExclamation
        Exit Sub
    End If

    author = CStr(mAuthors(authorIndex + 1))
    If GetHardcodedAuthorColor(author) <> -1 Then
        MsgBox "This author already has a hardcoded color, which is reserved.", vbInformation
        Exit Sub
    End If

    If mColorPalette.ListIndex < 0 Then
        MsgBox "Select a color from the palette.", vbExclamation
        Exit Sub
    End If

    colorValue = CLng(mColorPalette.List(mColorPalette.ListIndex, 1))
    If mAuthorColors.Exists(author) Then
        mAuthorColors.Item(author) = colorValue
    Else
        mAuthorColors.Add author, colorValue
    End If

    UpdateColorStatus
End Sub

Private Sub mRunButton_Click()
    Dim author As Variant
    Dim authorIndex As Long

    If mColorSetupMode Then
        For Each author In mColorRequirements
            If Not mAuthorColors.Exists(CStr(author)) Then
                MsgBox "Assign a color to " & DisplayAuthor(CStr(author)) & _
                       " before continuing.", vbExclamation
                authorIndex = FindAuthorIndex(CStr(author))
                If authorIndex >= 0 Then mColorAuthorCombo.ListIndex = authorIndex
                Exit Sub
            End If
        Next author
        WasCancelled = False
        ReturnedToOptions = False
        Me.Hide
        Exit Sub
    End If

    If mAuthorFilter.Value And Not mClearMode.Value Then
        Set SelectedAuthors = New Collection
        For authorIndex = 0 To mAuthorList.ListCount - 1
            If mAuthorList.Selected(authorIndex) Then
                SelectedAuthors.Add CStr(mAuthors(authorIndex + 1))
            End If
        Next authorIndex

        If SelectedAuthors.Count = 0 Then
            MsgBox "Select at least one author or clear the author restriction.", vbExclamation
            Exit Sub
        End If
    Else
        Set SelectedAuthors = New Collection
    End If

    If mBeforeFilter.Value Then
        If Not IsDate(mBeforeDateBox.Text) Then
            MsgBox "Enter a valid date for the Before filter.", vbExclamation
            mBeforeDateBox.SetFocus
            Exit Sub
        End If
        BeforeDate = DateValue(CDate(mBeforeDateBox.Text))
    End If

    If mAfterFilter.Value Then
        If Not IsDate(mAfterDateBox.Text) Then
            MsgBox "Enter a valid date for the After filter.", vbExclamation
            mAfterDateBox.SetFocus
            Exit Sub
        End If
        AfterDate = DateValue(CDate(mAfterDateBox.Text))
    End If

    If mHighlightMode.Value Then
        RunMode = RunHighlightInsertionsDeletions
    ElseIf mAcceptFormattingMode.Value Then
        RunMode = RunAcceptFormatting
    ElseIf mAcceptInsertionsMode.Value Then
        RunMode = RunAcceptInsertionsDeletions
    Else
        RunMode = RunClearScriptHighlights
    End If

    AuthorFilterEnabled = mAuthorFilter.Value And Not mClearMode.Value
    UseBeforeDate = mBeforeFilter.Value
    IncludeBeforeDate = mIncludeBeforeDate.Value
    UseAfterDate = mAfterFilter.Value
    IncludeAfterDate = mIncludeAfterDate.Value
    WasCancelled = False
    ReturnedToOptions = False
    Me.Hide
End Sub

Private Sub mBackButton_Click()
    ReturnedToOptions = True
    WasCancelled = False
    Me.Hide
End Sub

Private Sub mCancelButton_Click()
    WasCancelled = True
    ReturnedToOptions = False
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        WasCancelled = True
        ReturnedToOptions = False
        Me.Hide
    End If
End Sub
