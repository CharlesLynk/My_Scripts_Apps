Attribute VB_Name = "func1_Binary"
' ===========================
' Author:  Chase Lynk
' ===========================

Option Explicit

'==========================
' === Module-Level Globals
'==========================

' Primary worksheet hosting the layout (configured in InitGlobals)
Private gWS As Worksheet

' --- Configuration: cell addresses (sheet-driven inputs/outputs)
Private Const SIGNED_CELL     As String = "J10"   ' e.g., "no", "two's comp.", "ones' comp.", "magnitude"
Private Const WORD_CELL       As String = "J13"   ' bits per word (4/8/16/32/64)
Private Const ENDIAN_CELL     As String = "J16"   ' "little" or "big"
Private Const LAST_CONV_CELL  As String = "B19"   ' stores last conversion direction

' --- Layout rows/columns (set once; can be changed and re-initialized)
Private gBitLabelRow  As Long  ' row with bit index labels
Private gBinRow       As Long  ' row with individual bit cells
Private gDecRow       As Long  ' row with merged decimal cells
Private gBinStartCol  As Long  ' first bit column
Private gBinEndCol    As Long  ' last bit column

' --- Key ranges
Private gBinRng   As Range   ' range containing bits for all words
Private gDecRng   As Range   ' range containing decimals (merged by word)
Private gInputRng As Range   ' user input/output word strings (vertical list)

' --- Settings (sheet-driven) + derived values
Private gSignedFmt  As String ' "no", "two's comp.", "ones' comp.", "magnitude"
Private gWordSize   As Long   ' bits per word
Private gNumWords   As Long   ' number of words across the bit row
Private gBitLen     As Long   ' total bit cells = gBinEndCol - gBinStartCol + 1
Private gIsLE       As Boolean ' True if little-endian
Private gIsSigned   As Boolean ' True if signed mode selected
Private gLastConv   As String  ' "binary to decimal" | "decimal to binary"

' --- State
Private gInitialized As Boolean
Private gSubBusy     As Boolean
Private gOldValue    As Variant ' Optional: set via Worksheet_Change if you want revert behavior

'==========================
' === Initialization
'==========================

' Initializes worksheet, layout, and static ranges.
' Call once (auto-called by other public subs when needed).
Public Sub InitGlobals(Optional ByVal force As Boolean = False)
    If gInitialized And Not force Then Exit Sub

    Set gWS = ThisWorkbook.sheets("Binary")

    ' Layout configuration (change if your sheet differs)
    gBitLabelRow = 5
    gBinRow = 6
    gDecRow = 7
    gBinStartCol = 2
    gBinEndCol = 65

    ' Wire ranges
    Set gBinRng = gWS.Range(gWS.Cells(gBinRow, gBinStartCol), gWS.Cells(gBinRow, gBinEndCol))
    Set gDecRng = gWS.Range(gWS.Cells(gDecRow, gBinStartCol), gWS.Cells(gDecRow, gBinEndCol))
    Set gInputRng = gWS.Range("B22:B37")

    ' Read sheet settings and compute derived values
    RefreshSettings

    gInitialized = True
End Sub

' Re-reads current settings from the sheet and recomputes derived values.
Private Sub RefreshSettings()
    gBitLen = gBinEndCol - gBinStartCol + 1
    gSignedFmt = LCase$(CStr(gWS.Range(SIGNED_CELL).value))      ' "no", "two's comp.", etc.
    gWordSize = CLng(gWS.Range(WORD_CELL).value)
    gIsLE = (LCase$(CStr(gWS.Range(ENDIAN_CELL).value)) = "little")
    gIsSigned = (gSignedFmt <> "no")
    gNumWords = IIf(gWordSize > 0, gBitLen \ gWordSize, 0)
    gLastConv = LCase$(CStr(gWS.Range(LAST_CONV_CELL).value))
End Sub

'==========================
' === Layout / Clear
'==========================

' Clears bit/decimal rows and input cells to zeros.
Public Sub ClearAll()
    InitGlobals
    If gSubBusy Then Exit Sub
    gSubBusy = True

    Dim c As Range
    Application.ScreenUpdating = False

    gBinRng.value = 0
    gDecRng.value = 0

    For Each c In gInputRng
        c.value = String$(gWordSize, "0")
    Next c

CleanExit:
    Application.ScreenUpdating = True
    gSubBusy = False
End Sub

' Clears only the binary bit row.
Private Sub ClearBinary()
    InitGlobals
    gBinRng.value = 0
End Sub

' Merges/labels decimal cells by word and centers values.
' Also refreshes bit position labels.
Public Sub ResizeDecimalCells()
InitGlobals:     RefreshSettings
    If gSubBusy Then Exit Sub
    gSubBusy = True

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo CleanExit

    If CheckSize Then GoTo CleanExit

    LabelBitPositions

    gDecRng.UnMerge
    gDecRng.ClearContents

    Dim i As Long, c1 As Long, c2 As Long
    For i = 0 To gNumWords - 1
        c1 = gBinStartCol + i * gWordSize
        c2 = c1 + gWordSize - 1
        With gWS.Range(gWS.Cells(gDecRow, c1), gWS.Cells(gDecRow, c2))
            .Merge
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .value = 0
            With .Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
        End With
    Next i

CleanExit:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    gSubBusy = False
End Sub

' Writes bit index labels above the bit row, honoring endianness.
Private Sub LabelBitPositions()
InitGlobals:     RefreshSettings

    Dim i As Long, j As Long, startCol As Long, bitIdx As Long
    For i = 0 To gNumWords - 1
        startCol = gBinStartCol + i * gWordSize
        For j = 0 To gWordSize - 1
            bitIdx = IIf(gIsLE, j, gWordSize - 1 - j)
            gWS.Cells(gBitLabelRow, startCol + j).value = bitIdx
        Next j
    Next i
End Sub

'==========================
' === Conversions
'==========================

' Converts all words from the bit row to merged decimal cells.
Public Sub ConvertBinaryToDecimal()
InitGlobals:     RefreshSettings
    If gSubBusy Then Exit Sub
    gSubBusy = True
    On Error GoTo CleanExit

    If CheckSize Then GoTo CleanExit

    Dim i As Long, col As Long, binStr As String, decVal As Double
    col = gBinStartCol

    For i = 1 To gNumWords
        binStr = ReadWordBits(col, gWordSize)
        decVal = BinaryToDecimal(binStr, gIsSigned)
        gWS.Cells(gDecRow, col).value = decVal
        col = col + gWordSize
    Next i

    gWS.Range(LAST_CONV_CELL).value = "Binary To Decimal"

CleanExit:
    gSubBusy = False
End Sub

' Converts merged decimal cells to their bit representations in the bit row.
Public Sub ConvertDecimalToBinary()
InitGlobals:     RefreshSettings
    If gSubBusy Then Exit Sub
    gSubBusy = True
    On Error GoTo CleanExit

    If CheckSize Then GoTo CleanExit

    Dim i As Long, col As Long, decVal As Double, binStr As String
    Dim minVal As Double, maxVal As Double
    col = gBinStartCol

    ' Valid ranges per signed mode (note: Double precision may limit very large word sizes)
    If gIsSigned Then
        minVal = -2 ^ (gWordSize - 1)
        maxVal = 2 ^ (gWordSize - 1) - 1
    Else
        minVal = 0
        maxVal = 2 ^ gWordSize - 1
    End If

    For i = 1 To gNumWords
        decVal = CDbl(gWS.Cells(gDecRow, col).value)

        If decVal < minVal Or decVal > maxVal Then
            MsgBox "Decimal value out of range in word " & i & "." & vbCrLf & _
                   "Valid range: " & minVal & " to " & maxVal, vbCritical
            ' If you capture gOldValue via Worksheet_Change, it will revert:
            If Not IsEmpty(gOldValue) Then gWS.Cells(gDecRow, col).value = gOldValue
            GoTo CleanExit
        End If

        If Not gIsSigned And decVal < 0 Then decVal = Abs(decVal)
        binStr = DecimalToBinary(decVal, gWordSize, gIsSigned)
        WriteWordBits col, gWordSize, binStr
        col = col + gWordSize
    Next i

    gWS.Range(LAST_CONV_CELL).value = "Decimal To Binary"

CleanExit:
    gSubBusy = False
End Sub

' Runs whichever conversion was recorded last.
Public Sub RunLastConversion()
InitGlobals:     RefreshSettings
    Select Case gLastConv
        Case "binary to decimal": ConvertBinaryToDecimal
        Case "decimal to binary": ConvertDecimalToBinary
    End Select
End Sub

'==========================
' === Helpers (binary <-> decimal)
'==========================

' Reads a word's bits from the bit row into a string, honoring endianness.
Private Function ReadWordBits(ByVal startCol As Long, ByVal bits As Long) As String
    Dim j As Long, c2 As Long, s As String
    c2 = startCol + bits - 1
    s = ""
    If gIsLE Then
        For j = c2 To startCol Step -1
            s = s & CStr(gWS.Cells(gBinRow, j).value)
        Next j
    Else
        For j = startCol To c2
            s = s & CStr(gWS.Cells(gBinRow, j).value)
        Next j
    End If
    ReadWordBits = s
End Function

' Writes a binary string into the bit row for a word, honoring endianness.
Private Sub WriteWordBits(ByVal startCol As Long, ByVal bits As Long, ByVal binStr As String)
    Dim j As Long, c2 As Long
    c2 = startCol + bits - 1
    For j = 0 To bits - 1
        If gIsLE Then
            gWS.Cells(gBinRow, c2 - j).value = Mid$(binStr, j + 1, 1)
        Else
            gWS.Cells(gBinRow, startCol + j).value = Mid$(binStr, j + 1, 1)
        End If
    Next j
End Sub

' Decimal -> Binary string (MSB -> LSB), considering signed format.
Private Function DecimalToBinary(ByVal decVal As Double, ByVal bits As Long, ByVal isSigned As Boolean) As String
    Dim s As String, i As Long, tmp As Double

    ' Build magnitude representation first (absolute value if signed negative)
    tmp = IIf(isSigned And decVal < 0, Abs(decVal), decVal)
    s = String$(bits, "0")
    For i = bits - 1 To 0 Step -1
        If tmp >= 2 ^ i Then
            Mid$(s, bits - i, 1) = "1"
            tmp = tmp - 2 ^ i
        End If
    Next i

    ' Apply the chosen signed encoding
    If isSigned Then
        Select Case LCase$(gSignedFmt)
            Case "two's comp.", "two's comp", "two's complement"
                If decVal < 0 Then
                    tmp = 2 ^ bits + decVal ' decVal is negative here
                    s = String$(bits, "0")
                    For i = bits - 1 To 0 Step -1
                        If tmp >= 2 ^ i Then
                            Mid$(s, bits - i, 1) = "1"
                            tmp = tmp - 2 ^ i
                        End If
                    Next i
                End If

            Case "ones' comp.", "ones", "one's complement"
                If decVal < 0 Then s = FlipBits(s)

            Case "magnitude"
                ' MSB is the sign; remainder is magnitude
                Mid$(s, 1, 1) = IIf(decVal < 0, "1", "0")
        End Select
    End If

    DecimalToBinary = s
End Function

' Binary string -> Decimal (assumes MSB first), considering signed format.
Private Function BinaryToDecimal(ByVal binStr As String, ByVal isSigned As Boolean) As Double
    Dim dec As Double, i As Long, l As Long, bitVal As Long

    binStr = Trim$(binStr)
    If Not IsValidBinaryString(binStr) Then
        BinaryToDecimal = 0
        Exit Function
    End If

    l = Len(binStr)
    dec = 0

    ' Accumulate value from MSB to LSB
    For i = 1 To l
        bitVal = Asc(Mid$(binStr, i, 1)) - 48  ' "0"/"1" -> 0/1
        dec = dec + bitVal * 2 ^ (l - i)
    Next i

    If isSigned Then
        Select Case LCase$(gSignedFmt)
            Case "two's comp.", "two's comp", "two's complement"
                If Left$(binStr, 1) = "1" Then dec = dec - 2 ^ l

            Case "ones' comp.", "ones", "one's complement"
                If Left$(binStr, 1) = "1" Then dec = dec - (2 ^ l - 1)

            Case "magnitude"
                If Left$(binStr, 1) = "1" Then
                    dec = -BinaryToDecimal(Mid$(binStr, 2), False)
                End If
        End Select
    End If

    BinaryToDecimal = dec
End Function

' Bitwise invert a "0"/"1" string.
Private Function FlipBits(ByVal binStr As String) As String
    Dim i As Long, outStr As String
    outStr = binStr
    For i = 1 To Len(outStr)
        Mid$(outStr, i, 1) = IIf(Mid$(outStr, i, 1) = "0", "1", "0")
    Next i
    FlipBits = outStr
End Function

'==========================
' === Validation / Parsing
'==========================

' Ensures configuration makes sense (bit length divisible by word size, etc.)
Private Function CheckSize() As Boolean
    RefreshSettings

    If gWordSize <= 0 Then
        MsgBox "Invalid word size.", vbExclamation
        CheckSize = True
        Exit Function
    End If

    If gBitLen Mod gWordSize <> 0 Then
        MsgBox "Invalid bit length (" & gBitLen & ") for word size " & gWordSize & ".", vbExclamation
        CheckSize = True
    ElseIf (gBinEndCol - gBinStartCol + 1) Mod gNumWords <> 0 Then
        MsgBox "Column count mismatch.", vbExclamation
        CheckSize = True
    Else
        CheckSize = False
    End If
End Function

' Returns True only if the string is non-empty and contains only 0/1.
Private Function IsValidBinaryString(ByVal binStr As String) As Boolean
    Dim i As Long, ch As String
    If Len(binStr) = 0 Then
        IsValidBinaryString = False
        Exit Function
    End If
    For i = 1 To Len(binStr)
        ch = Mid$(binStr, i, 1)
        If ch <> "0" And ch <> "1" Then
            IsValidBinaryString = False
            Exit Function
        End If
    Next i
    IsValidBinaryString = True
End Function

' Trims and removes spaces and tabs from a candidate binary string.
Private Function SanitizeBinaryString(ByVal s As String) As String
    s = Trim$(s)
    s = Replace$(s, " ", "")
    s = Replace$(s, vbTab, "")
    SanitizeBinaryString = s
End Function

' Parses binary strings from gInputRng, selects the smallest fitting word size,
' pads as needed (respecting endianness), and writes bits into gBinRow.
' Then resizes decimal cells and runs the last conversion for convenience.
Public Sub ParseBinaryStrings()
InitGlobals:     RefreshSettings
    If gSubBusy Then Exit Sub
    gSubBusy = True
    On Error GoTo CleanExit

    Dim c As Range, binStr As String, maxLen As Long
    Dim wordSizes As Variant: wordSizes = Array(4, 8, 16, 32, 64)
    Dim chosen As Long, i As Long, j As Long
    Dim bitIndex As Long: bitIndex = 0

    ClearBinary

    ' First pass: longest valid, non-empty string
    maxLen = 0
    For Each c In gInputRng
        binStr = SanitizeBinaryString(CStr(c.value))
        If binStr <> "" And IsValidBinaryString(binStr) Then
            If Len(binStr) > maxLen Then maxLen = Len(binStr)
        End If
    Next c

    ' Choose smallest word size that fits
    chosen = -1
    For i = LBound(wordSizes) To UBound(wordSizes)
        If maxLen <= CLng(wordSizes(i)) Then
            chosen = CLng(wordSizes(i))
            Exit For
        End If
    Next i

    If chosen = -1 Then
        MsgBox "No suitable word size found for the provided binary strings.", vbExclamation
        GoTo CleanExit
    End If

    ' Second pass: sanitize, validate, pad, and lay out into the bit row
    For Each c In gInputRng
        binStr = SanitizeBinaryString(CStr(c.value))
        If binStr = "" Then Exit For
        If Not IsValidBinaryString(binStr) Then
            MsgBox "Invalid binary string in " & c.Address(0, 0), vbExclamation
            GoTo CleanExit
        End If

        ' Pad (LE: right-pad; BE: left-pad)
        If gIsLE Then
            binStr = Right$(String$(chosen, "0") & binStr, chosen)
            c.value = binStr
        Else
            binStr = Left$(binStr & String$(chosen, "0"), chosen)
            c.value = binStr
        End If

        ' Ensure capacity
        If bitIndex + chosen > gBitLen Then Exit For

        ' Insert bits contiguously
        For j = 1 To chosen
            gWS.Cells(gBinRow, gBinStartCol + bitIndex + j - 1).value = Mid$(binStr, j, 1)
        Next j
        bitIndex = bitIndex + chosen
    Next c

    ' Update globals if the chosen size differs
    gWordSize = chosen
    gWS.Range(WORD_CELL).value = gWordSize
    gNumWords = gBitLen \ gWordSize

    ResizeDecimalCells
    RunLastConversion

CleanExit:
    gSubBusy = False
End Sub

' Dumps each word in the bit row back to gInputRng as contiguous strings.
Public Sub DumpBinaryWordsToCells()
InitGlobals:     RefreshSettings
    If gSubBusy Then Exit Sub
    gSubBusy = True
    On Error GoTo CleanExit

    Dim outputRng As Range, i As Long, j As Long
    Dim c1 As Long, c2 As Long, binStr As String, maxCells As Long

    Set outputRng = gInputRng
    maxCells = outputRng.Rows.Count

    c1 = gBinStartCol
    For i = 1 To maxCells
        If i <= gNumWords Then
            c2 = c1 + gWordSize - 1
            binStr = ""
            For j = c1 To c2
                binStr = binStr & CStr(gWS.Cells(gBinRow, j).value)
            Next j
            outputRng.Cells(i, 1).value = binStr
            c1 = c2 + 1
        Else
            outputRng.Cells(i, 1).value = String$(gWordSize, "0")
        End If
    Next i

CleanExit:
    gSubBusy = False
End Sub

