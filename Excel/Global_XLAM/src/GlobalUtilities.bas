Attribute VB_Name = "GlobalUtilities"
Option Explicit
'
'==========================================
' Module: GlobalUtilities
' Purpose: Shared helper routines used by XLAM and workbook projects.
' Scope: General worksheet helpers, normalization, and workbook text cleanup.
' Dependencies: None
' Notes:
' - XLAM is expected in %APPDATA%\Microsoft\AddIns.
' - Managed from File -> Options -> Add-Ins.
' - Available as VBAProject(Global).ReloadManager.
'
' CLEANUP ARCHITECTURE NOTES
'---------------------------
' 1. READ OPERATIONS (CollectUniquePLCs, BuildSummary, PLC_RebuildPLCList):
'    - Use RAW cell values (no cleaning applied)
'    - Matching is done via NormalizeKey() which applies full cleaning pipeline
'    - This preserves user's original formatting while enabling fuzzy matching
'
' 2. WRITE OPERATIONS (BuildSummary output):
'    - Write RAW values from source cells
'    - Users see their original data (with bad characters) in output
'    - To clean data, user must manually run dedicated cleanup functions
'
' 3. CLEANUP OPERATIONS (Manual only):
'    - StandardizeAllWorkbookText() - cleans ALL sheets (called by user)
'    - CleanBlankRowsAcrossSheets() - removes blank rows
'    - StandardizePLCNames() - normalizes PLC naming
'    - These are NEVER called automatically from read/filter operations
'
' 4. NORMALIZATION FUNCTIONS (Used for matching, not display):
'    - NormalizeKey() - strips all non-alphanumeric for matching
'    - StandardizeSymbols() - converts lookalikes to standard ASCII
'    - CleanWhitespace() - removes Alt+0160, zero-width chars
'    - CleanText() - removes non-printable characters
'    - These functions are ONLY used for comparisons, not data modification
'
' 5. PERFORMANCE OPTIMIZATIONS:
'    - Early exits for empty/ASCII-only strings
'    - Regex cache singleton pattern
'    - Batch array reads for multi-row operations
'    - Error handling to prevent crashes on bad data
'
' 6. GENERAL-PURPOSE WORKSHEET HELPERS:
'    - FindHeaderFlexible() - flexible header search with normalization
'    - HeaderMapFromRow() - build column index dictionary from header row
'    - LookupHeader() - normalized column lookup
'    - DataStartRow() - calculate first data row after header
'    - RangesMatch() - compare two range objects
'    These are reusable across any Excel project
'
'==========================================

' Module-level RegExp object cache for performance
Private m_RegExpCache As Object

'================== APPLICATION STATE MANAGEMENT ==================
Public Type CalcState
    ScreenUpdating As Boolean
    EnableEvents As Boolean
    DisplayAlerts As Boolean
    Calculation As XlCalculation
End Type

Public Function PushAppState() As CalcState
    Dim result As CalcState
    
    ' Force Excel into a clean state before changing settings
    On Error Resume Next
    
    ' Exit edit mode if a cell is being edited
    Application.Interactive = True
    SendKeys "{ESCAPE}", True
    DoEvents
    
    On Error GoTo 0
    
    With Application
        ' Capture current state
        On Error Resume Next
        result.ScreenUpdating = .ScreenUpdating
        result.EnableEvents = .EnableEvents
        result.DisplayAlerts = .DisplayAlerts
        result.Calculation = .Calculation
        On Error GoTo 0
        
        ' Apply optimized state with error handling for each property
        On Error Resume Next
        .ScreenUpdating = False
        If Err.Number <> 0 Then
            Err.Clear
            result.ScreenUpdating = True  ' Mark that we couldn't change it
        End If
        
        .EnableEvents = False
        If Err.Number <> 0 Then
            Err.Clear
            result.EnableEvents = True
        End If
        
        .DisplayAlerts = False
        If Err.Number <> 0 Then
            Err.Clear
            result.DisplayAlerts = True
        End If
        
        .Calculation = xlCalculationManual
        If Err.Number <> 0 Then
            Err.Clear
            result.Calculation = xlCalculationAutomatic
        End If
        On Error GoTo 0
    End With
    
    PushAppState = result
End Function

Public Sub PopAppState(ByRef state As CalcState)  ' Must be ByRef for UDTs
    ' Work with local copy for safety in case of interruption
    Dim localState As CalcState
    localState = state
    
    On Error Resume Next  ' Add error handling
    With Application
        .Calculation = localState.Calculation  ' Restore calculation first
        .ScreenUpdating = localState.ScreenUpdating
        .EnableEvents = localState.EnableEvents
        .DisplayAlerts = localState.DisplayAlerts
    End With
    On Error GoTo 0
End Sub

'================== WORKSHEET HELPERS ==================
Public Function LastUsedRow(ws As Worksheet) As Long
    Dim f As Range
    On Error Resume Next
    Set f = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
                          LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0
    LastUsedRow = IIf(f Is Nothing, 0, f.Row)
End Function

Public Function LastUsedCol(ws As Worksheet) As Long
    Dim f As Range
    On Error Resume Next
    Set f = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
                          LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    On Error GoTo 0
    LastUsedCol = IIf(f Is Nothing, 0, f.Column)
End Function

Private Function LastNonEmptyRowInRange(ws As Worksheet, ByVal startRow As Long, _
                                       ByVal startCol As Long, ByVal endCol As Long) As Long
    Dim r As Long, c As Long
    Dim lastRow As Long: lastRow = LastUsedRow(ws)
    For r = lastRow To startRow Step -1
        For c = startCol To endCol
            If Not IsBlankAfterNormalize(ws.Cells(r, c).Value2) Then
                LastNonEmptyRowInRange = r
                Exit Function
            End If
        Next c
    Next r
    LastNonEmptyRowInRange = startRow - 1
End Function

Public Function SheetByName(ByVal nm As String) As Worksheet
    On Error Resume Next
    Set SheetByName = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
End Function

'================== TEXT NORMALIZATION & CLEANING ==================

' Remove single-line whitespace errors (Alt+0160, etc.) but preserve intentional line breaks
' This function cleans invisible/problematic whitespace while keeping user formatting
Public Function CleanWhitespace(ByVal s As String) As String
    Dim t As String: t = s
    
    ' Quick exit for empty strings
    If LenB(t) = 0 Then
        CleanWhitespace = t
        Exit Function
    End If
    
    ' CRITICAL: Alt+0160 (Non-breaking space) - often invisible, causes matching issues
    t = Replace$(t, ChrW$(160), " ")     ' NBSP (Alt+0160) -> regular space
    
    ' Other problematic whitespace characters (NOT line breaks)
    t = Replace$(t, vbTab, " ")          ' Tab -> space
    t = Replace$(t, ChrW$(8239), " ")    ' Narrow NBSP -> space
    t = Replace$(t, ChrW$(8287), " ")    ' Medium mathematical space -> space
    t = Replace$(t, ChrW$(12288), " ")   ' Ideographic space -> space
    
    ' Zero-width characters (completely invisible, cause mismatch issues)
    t = Replace$(t, ChrW$(8203), "")     ' Zero-width space (ZWSP)
    t = Replace$(t, ChrW$(8204), "")     ' Zero-width non-joiner (ZWNJ)
    t = Replace$(t, ChrW$(8205), "")     ' Zero-width joiner (ZWJ)
    t = Replace$(t, ChrW$(65279), "")    ' BOM / Zero-width no-break space
    t = Replace$(t, ChrW$(8288), "")     ' Word joiner
    t = Replace$(t, ChrW$(8290), "")     ' Word joiner (variant)
    t = Replace$(t, ChrW$(8291), "")     ' Function application (invisible)
    t = Replace$(t, ChrW$(8292), "")     ' Invisible times
    t = Replace$(t, ChrW$(8298), "")     ' Invisible separator
    t = Replace$(t, ChrW$(8299), "")     ' Invisible plus
    
    ' PRESERVE intentional line breaks (vbCr, vbLf, vbCrLf)
    ' Users may use Alt+Enter (Chr(10)) to format multi-line cell content
    ' We do NOT replace these - they are intentional formatting
    
    ' Collapse multiple consecutive spaces (but keep single spaces)
    Do While InStr(t, "  ") > 0
        t = Replace$(t, "  ", " ")
    Loop
    
    CleanWhitespace = Trim$(t)
End Function

' Apply WorksheetFunction.Clean if available
Public Function CleanText(ByVal s As String) As String
    Dim t As String: t = s
    
    ' Quick exit for empty strings
    If LenB(t) = 0 Then
        CleanText = t
        Exit Function
    End If
    
    On Error Resume Next
    t = WorksheetFunction.Clean(t)
    If Err.Number <> 0 Then
        ' WorksheetFunction.Clean failed - return original (rare edge case)
        Err.Clear
        t = s
    End If
    On Error GoTo 0
    CleanText = t
End Function

' Standardize dashes, underscores, and other symbols commonly misused in tags/filenames
' Converts all dash variants to standard hyphen-minus (ASCII 45)
' Converts all underscore variants to standard underscore (ASCII 95)
Public Function StandardizeSymbols(ByVal s As String) As String
    Dim t As String: t = s
    
    ' Quick exit for empty strings
    If LenB(t) = 0 Then
        StandardizeSymbols = t
        Exit Function
    End If
    
    ' Quick check: if string is pure ASCII (no high-code characters), skip processing
    ' This optimization helps when most data is already clean
    Dim i As Long, charCode As Long, hasHighChar As Boolean
    For i = 1 To Len(t)
        charCode = AscW(Mid$(t, i, 1))
        If charCode > 127 Then
            hasHighChar = True
            Exit For
        End If
    Next i
    
    ' If no high characters found, return unchanged (already standard ASCII)
    If Not hasHighChar Then
        StandardizeSymbols = t
        Exit Function
    End If
    
    ' ===== DASHES: Standardize ALL to hyphen-minus (ASCII 45, keyboard "-") =====
    t = Replace$(t, ChrW$(8211), "-")    ' En dash
    t = Replace$(t, ChrW$(8212), "-")    ' Em dash
    t = Replace$(t, ChrW$(8210), "-")    ' Figure dash
    t = Replace$(t, ChrW$(8213), "-")    ' Horizontal bar
    t = Replace$(t, ChrW$(8722), "-")    ' Minus sign (-) - looks like hyphen but different!
    t = Replace$(t, ChrW$(8209), "-")    ' Non-breaking hyphen
    t = Replace$(t, ChrW$(8208), "-")    ' Hyphen (punctuation)
    t = Replace$(t, ChrW$(8722), "-")    ' Mathematical minus
    t = Replace$(t, ChrW$(65293), "-")   ' Fullwidth hyphen-minus
    t = Replace$(t, ChrW$(11834), "-")   ' Heavy minus sign
    t = Replace$(t, ChrW$(11835), "-")   ' Heavy hyphen
    t = Replace$(t, ChrW$(8259), "-")    ' Hyphen bullet
    
    ' ===== UNDERSCORES: Standardize ALL to low line (ASCII 95, keyboard "_") =====
    t = Replace$(t, ChrW$(65075), "_")   ' Fullwidth low line
    t = Replace$(t, ChrW$(8255), "_")    ' Undertie (looks like underscore)
    t = Replace$(t, ChrW$(8275), "_")    ' Double low line
    t = Replace$(t, ChrW$(65343), "_")   ' Halfwidth low line
    
    ' ===== PERIODS/DOTS: Standardize to ASCII period (46, keyboard ".") =====
    t = Replace$(t, ChrW$(8228), ".")    ' One dot leader
    t = Replace$(t, ChrW$(65294), ".")   ' Fullwidth full stop
    t = Replace$(t, ChrW$(12290), ".")   ' Ideographic full stop
    
    ' ===== SLASHES: Standardize to ASCII forward/backslash =====
    t = Replace$(t, ChrW$(8260), "/")    ' Fraction slash -> forward slash
    t = Replace$(t, ChrW$(8725), "/")    ' Division slash -> forward slash
    t = Replace$(t, ChrW$(65295), "/")   ' Fullwidth solidus -> forward slash
    t = Replace$(t, ChrW$(65340), "\")   ' Fullwidth reverse solidus -> backslash
    
    ' ===== COLONS: Standardize to ASCII colon (58, keyboard ":") =====
    t = Replace$(t, ChrW$(65306), ":")   ' Fullwidth colon
    t = Replace$(t, ChrW$(8758), ":")    ' Two-dot punctuation
    
    ' ===== PARENTHESES: Standardize to ASCII parens =====
    t = Replace$(t, ChrW$(65288), "(")   ' Fullwidth left paren
    t = Replace$(t, ChrW$(65289), ")")   ' Fullwidth right paren
    
    ' ===== BRACKETS: Standardize to ASCII brackets =====
    t = Replace$(t, ChrW$(65339), "[")   ' Fullwidth left square bracket
    t = Replace$(t, ChrW$(65341), "]")   ' Fullwidth right square bracket
    
    ' ===== APOSTROPHES/QUOTES: Standardize to ASCII straight quotes =====
    t = Replace$(t, ChrW$(8216), "'")    ' Left single quote
    t = Replace$(t, ChrW$(8217), "'")    ' Right single quote
    t = Replace$(t, ChrW$(8220), """")   ' Left double quote
    t = Replace$(t, ChrW$(8221), """")   ' Right double quote
    t = Replace$(t, ChrW$(8219), "'")    ' Single high-reversed-9 quote
    t = Replace$(t, ChrW$(8223), """")   ' Double high-reversed-9 quote
    t = Replace$(t, ChrW$(65287), "'")   ' Fullwidth apostrophe
    t = Replace$(t, ChrW$(65282), """")  ' Fullwidth quotation mark
    
    ' ===== NUMBERS: Standardize lookalike characters (common in OCR/paste errors) =====
    t = Replace$(t, ChrW$(65296), "0")   ' Fullwidth digit 0
    t = Replace$(t, ChrW$(65297), "1")   ' Fullwidth digit 1
    t = Replace$(t, ChrW$(65298), "2")   ' Fullwidth digit 2
    t = Replace$(t, ChrW$(65299), "3")   ' Fullwidth digit 3
    t = Replace$(t, ChrW$(65300), "4")   ' Fullwidth digit 4
    t = Replace$(t, ChrW$(65301), "5")   ' Fullwidth digit 5
    t = Replace$(t, ChrW$(65302), "6")   ' Fullwidth digit 6
    t = Replace$(t, ChrW$(65303), "7")   ' Fullwidth digit 7
    t = Replace$(t, ChrW$(65304), "8")   ' Fullwidth digit 8
    t = Replace$(t, ChrW$(65305), "9")   ' Fullwidth digit 9
    
    ' ===== LETTERS: Standardize lookalike letters (common confusions) =====
    ' Cyrillic/Greek letters that look like Latin but have different codes
    t = Replace$(t, ChrW$(1040), "A")    ' Cyrillic A
    t = Replace$(t, ChrW$(1042), "B")    ' Cyrillic Be (looks like B)
    t = Replace$(t, ChrW$(1045), "E")    ' Cyrillic Ie (looks like E)
    t = Replace$(t, ChrW$(1047), "Z")    ' Cyrillic Ze (looks like 3/Z)
    t = Replace$(t, ChrW$(1050), "K")    ' Cyrillic Ka
    t = Replace$(t, ChrW$(1052), "M")    ' Cyrillic Em
    t = Replace$(t, ChrW$(1053), "H")    ' Cyrillic En (looks like H)
    t = Replace$(t, ChrW$(1054), "O")    ' Cyrillic O
    t = Replace$(t, ChrW$(1056), "P")    ' Cyrillic Pe (looks like P)
    t = Replace$(t, ChrW$(1057), "C")    ' Cyrillic Es (looks like C)
    t = Replace$(t, ChrW$(1058), "T")    ' Cyrillic Te
    t = Replace$(t, ChrW$(1061), "X")    ' Cyrillic Kha (looks like X)
    t = Replace$(t, ChrW$(1059), "Y")    ' Cyrillic U (looks like Y)
    
    ' Lowercase Cyrillic
    t = Replace$(t, ChrW$(1072), "a")    ' Cyrillic a
    t = Replace$(t, ChrW$(1077), "e")    ' Cyrillic ie (looks like e)
    t = Replace$(t, ChrW$(1086), "o")    ' Cyrillic o
    t = Replace$(t, ChrW$(1088), "p")    ' Cyrillic pe (looks like p)
    t = Replace$(t, ChrW$(1089), "c")    ' Cyrillic es (looks like c)
    t = Replace$(t, ChrW$(1091), "y")    ' Cyrillic u (looks like y)
    t = Replace$(t, ChrW$(1093), "x")    ' Cyrillic kha (looks like x)
    
    StandardizeSymbols = t
End Function

' Check if a value is blank after normalization (handles hidden chars)
Public Function IsBlankAfterNormalize(ByVal v As Variant) As Boolean
    ' Safe blank check that handles errors, empty cells, and hidden characters
    On Error Resume Next
    Dim t As String
    
    ' Handle error cells and null values
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        IsBlankAfterNormalize = True
        Exit Function
    End If
    
    t = CStr(v)
    If Err.Number <> 0 Then
        IsBlankAfterNormalize = True
        Err.Clear
        Exit Function
    End If
    
    t = CleanText(t)
    t = CleanWhitespace(t)
    IsBlankAfterNormalize = (LenB(Trim$(t)) = 0)
    
    On Error GoTo 0
End Function

' Robust key normalizer for matching (uppercase, remove all non-alphanumeric)
' Public so worksheet formulas can call it
Public Function NormalizeKey(ByVal v As Variant) As String
    On Error Resume Next
    Dim t As String
    
    ' Handle error cells, null, empty
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        NormalizeKey = ""
        Exit Function
    End If
    
    t = CStr(v)
    If Err.Number <> 0 Then
        NormalizeKey = ""
        Err.Clear
        Exit Function
    End If
    
    ' Apply comprehensive cleaning pipeline
    t = CleanText(t)              ' Remove non-printable characters
    t = StandardizeSymbols(t)     ' Standardize all dashes, underscores, lookalikes
    t = CleanWhitespace(t)        ' Clean whitespace (preserves line breaks)
    
    t = UCase$(t)
    
    ' Remove all non-alphanumeric characters for matching
    t = RxReplace(t, "[^A-Z0-9]+", "", True)
    
    NormalizeKey = t
    On Error GoTo 0
End Function

' HTML entity decoder (for pasted config values)
Private Function SanitizeHTMLEntities(ByVal s As String) As String
    Dim t As String: t = CStr(s)
    ' Double-escaped
    t = Replace$(t, "&amp;amp;", "&amp;")
    t = Replace$(t, "&amp;lt;", "<")
    t = Replace$(t, "&amp;gt;", ">")
    t = Replace$(t, "&amp;quot;", """")
    t = Replace$(t, "&amp;#39;", "'")
    ' Single-escaped
    t = Replace$(t, "&amp;", "&")
    t = Replace$(t, "&lt;", "<")
    t = Replace$(t, "&gt;", ">")
    t = Replace$(t, "&quot;", """")
    t = Replace$(t, "&#39;", "'")
    SanitizeHTMLEntities = t
End Function

'================== REGEX HELPER ==================
' Late-bound RegExp (no reference needed) - uses cached object for performance
Public Function RxReplace(ByVal s As String, ByVal pattern As String, _
                          ByVal replacement As String, ByVal globalReplace As Boolean) As String
    ' Safety check: protect against excessively large strings
    If Len(s) > 32767 Then
        ' Excel cell limit is 32,767 chars - anything larger is suspicious
        RxReplace = s
        Exit Function
    End If
    
    ' Create RegExp object once and reuse it (singleton pattern)
    If m_RegExpCache Is Nothing Then
        Set m_RegExpCache = CreateObject("VBScript.RegExp")
    End If
    
    On Error Resume Next
    With m_RegExpCache
        .Global = globalReplace
        .Multiline = True
        .IgnoreCase = True
        .pattern = pattern
        RxReplace = .Replace(s, replacement)
    End With
    
    ' If regex failed, return original string
    If Err.Number <> 0 Then
        Err.Clear
        RxReplace = s
    End If
    On Error GoTo 0
End Function

' Clear regex cache (useful for shared workbooks or memory management)
Private Sub ClearRegexCache()
    Set m_RegExpCache = Nothing
End Sub

'================== SHARED WORKBOOK DETECTION ==================

' Detect if workbook is in traditional shared mode
Public Function IsWorkbookShared(Optional ByVal wb As Workbook) As Boolean
    On Error Resume Next
    If wb Is Nothing Then Set wb = ActiveWorkbook
    IsWorkbookShared = wb.MultiUserEditing
    If Err.Number <> 0 Then
        Err.Clear
        IsWorkbookShared = False
    End If
    On Error GoTo 0
End Function

' Detect number of active users (Excel 2016+ co-authoring)
Public Function GetActiveUserCount(Optional ByVal wb As Workbook) As Long
    On Error Resume Next
    Dim userCount As Long
    Dim wbObj As Object  ' Late-bound to avoid compile error on local files
    
    If wb Is Nothing Then
        Set wbObj = ActiveWorkbook
    Else
        Set wbObj = wb
    End If
    
    ' Try to access co-authoring users collection (Office 365/2016+)
    ' Use late binding to avoid compile-time error when Users property doesn't exist
    If Not wbObj.IsInplace Then
        ' Workbook is not embedded, check for co-authors
        ' This will fail gracefully on local files or older Excel versions
        userCount = wbObj.Users.count
    Else
        userCount = 1
    End If
    
    ' Fallback if Users collection doesn't exist (older Excel, local files, etc.)
    If Err.Number <> 0 Then
        Err.Clear
        userCount = 1  ' Assume solo user (local file or no co-authoring)
    End If
    
    ' Minimum count is 1 (current user)
    If userCount < 1 Then userCount = 1
    
    GetActiveUserCount = userCount
    On Error GoTo 0
End Function

' Check if it's safe to run destructive operations
Public Function IsSafeToModifyWorkbook() As Boolean
    IsSafeToModifyWorkbook = (GetActiveUserCount() <= 1) And Not IsWorkbookShared()
End Function

'================== ARRAY HELPERS ==================
Public Function ArrayCount(ByRef arr() As String) As Long
    On Error Resume Next
    ArrayCount = UBound(arr) - LBound(arr) + 1
    If Err.Number <> 0 Then ArrayCount = 0
    On Error GoTo 0
End Function

Public Sub QuickSortStrings(a() As String, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, p As String, t As String
    If lo >= hi Then Exit Sub
    
    i = lo: j = hi: p = a((lo + hi) \ 2)
    Do While i <= j
        Do While StrComp(a(i), p, vbTextCompare) < 0: i = i + 1: Loop
        Do While StrComp(a(j), p, vbTextCompare) > 0: j = j - 1: Loop
        If i <= j Then
            If i <> j Then
                t = a(i): a(i) = a(j): a(j) = t
            End If
            i = i + 1: j = j - 1
        End If
    Loop
    
    If lo < j Then QuickSortStrings a, lo, j
    If i < hi Then QuickSortStrings a, i, hi
End Sub

' Check if value exists in array (for small arrays)
Private Function ExistsInArray(arr() As Long, ByVal countSoFar As Long, ByVal v As Long) As Boolean
    Dim i As Long
    For i = 1 To countSoFar
        If arr(i) = v Then
            ExistsInArray = True
            Exit Function
        End If
    Next i
    ExistsInArray = False
End Function

'================== MATH HELPERS ==================
Public Function Min2(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then Min2 = a Else Min2 = b
End Function

Private Function Max2(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then Max2 = a Else Max2 = b
End Function

'================== COLUMN CONVERSION ==================
' Accepts "A".."ZZZ" or a number. Returns 1-based column index.
Public Function ColRefToNum(ByVal v As Variant) As Long
    If IsNumeric(v) Then
        ColRefToNum = CLng(v)
    Else
        Dim s As String: s = UCase$(Trim$(CStr(v)))
        On Error Resume Next
        ColRefToNum = Range(s & "1").Column
        If Err.Number <> 0 Then ColRefToNum = 0
        On Error GoTo 0
    End If
End Function

'================== HEADER & DATA ROW HELPERS ==================

' Flexible header finder that tolerates line breaks, spaces, punctuation
' Scans first maxRows x maxCols for a header matching headerText (after normalization)
Public Function FindHeaderFlexible(ws As Worksheet, ByVal headerText As String, _
                                   ByRef headerRow As Long, ByRef headerCol As Long, _
                                   Optional ByVal maxRows As Long = 10, _
                                   Optional ByVal maxCols As Long = 20) As Boolean
    Dim r As Long, c As Long, maxCol As Long
    Dim target As String, cellText As String, key As String
    
    target = NormalizeKey(headerText)
    maxCol = Min2(maxCols, ws.Columns.count)
    
    For r = 1 To maxRows
        For c = 1 To maxCol
            On Error Resume Next
            cellText = CStr(ws.Cells(r, c).Value2)
            If Err.Number <> 0 Then
                Err.Clear
                cellText = ""
            End If
            On Error GoTo 0
            
            key = NormalizeKey(cellText)
            
            If key = target Then
                headerRow = r
                headerCol = c
                FindHeaderFlexible = True
                Exit Function
            End If
        Next c
    Next r
    FindHeaderFlexible = False
End Function

' Build a dictionary mapping normalized header names to column indices
Public Function HeaderMapFromRow(ws As Worksheet, ByVal hdrRow As Long) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Dim lastCol As Long: lastCol = LastUsedCol(ws)
    Dim c As Long, key As String, val As String
    
    For c = 1 To lastCol
        On Error Resume Next
        val = Trim$(CStr(ws.Cells(hdrRow, c).Value2))
        If Err.Number <> 0 Then
            Err.Clear
            val = ""
        End If
        On Error GoTo 0
        
        If Len(val) > 0 Then
            key = NormalizeKey(val)
            If Not d.Exists(key) Then d(key) = c
        End If
    Next c
    Set HeaderMapFromRow = d
End Function

' Lookup column index from header map (case-insensitive, normalized)
Public Function LookupHeader(map As Object, ByVal hdr As String) As Long
    Dim key As String: key = NormalizeKey(hdr)
    If map.Exists(key) Then LookupHeader = CLng(map(key)) Else LookupHeader = 0
End Function

' Determine first data row (directly after header row)
Public Function DataStartRow(ws As Worksheet, ByVal hdrRow As Long) As Long
    DataStartRow = hdrRow + 1
End Function

' Check if two ranges are identical (by address)
Public Function RangesMatch(rng1 As Range, rng2 As Range) As Boolean
    If rng1 Is Nothing Or rng2 Is Nothing Then
        RangesMatch = False
        Exit Function
    End If
    
    RangesMatch = (rng1.Address = rng2.Address)
End Function

'================== WORKBOOK-WIDE STANDARDIZATION ==================

' Standardize all text in workbook: clean whitespace + standardize symbols
' Targets all visible sheets (excludes hidden/very hidden sheets)
' Optional: specify sheet names to target specific sheets only
' IMPORTANT: This modifies cell contents - should only be called manually by user
Public Sub StandardizeAllWorkbookText(Optional ByVal targetSheets As Variant)
    Dim state As CalcState: state = PushAppState()
    On Error GoTo CleanFail
    
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "No active workbook found.", vbExclamation
        GoTo CleanExit
    End If
    
    ' Safety Gate 1: Check for other active users
    Dim userCount As Long: userCount = GetActiveUserCount(wb)
    Dim isShared As Boolean: isShared = IsWorkbookShared(wb)
    Dim msg As String
    Dim userConfirm As VbMsgBoxResult
    
    If userCount > 1 Or isShared Then
        ' Multiple users detected - show critical warning
        msg = "[!] WARNING: Multiple Users Detected!" & vbCrLf & vbCrLf
        msg = msg & "Active users: " & userCount & vbCrLf
        If isShared Then msg = msg & "Shared workbook mode: ENABLED" & vbCrLf
        msg = msg & vbCrLf & "This operation will modify cell contents across the workbook." & vbCrLf
        msg = msg & "Running this now may cause conflicts or data loss." & vbCrLf & vbCrLf
        msg = msg & "RECOMMENDED: Ask other users to close the file first." & vbCrLf & vbCrLf
        msg = msg & "Do you want to continue anyway?"
        
        userConfirm = MsgBox(msg, vbYesNo + vbCritical, "Multiple Users Detected")
    Else
        ' Safety Gate 2: Solo user - standard confirmation
        msg = "This will modify cell contents across the workbook." & vbCrLf & _
              "Do you want to continue?"
        
        userConfirm = MsgBox(msg, vbYesNo + vbQuestion, "Confirm Data Modification")
    End If
    
    If userConfirm <> vbYes Then
        GoTo CleanExit
    End If
    
    Dim ws As Worksheet
    Dim processedSheets As Long, processedCells As Long, changedCells As Long
    Dim totalSheets As Long
    Dim sheetList As Object
    Dim startTime As Double: startTime = Timer
    
    ' Build list of sheets to process
    Set sheetList = CreateObject("Scripting.Dictionary")
    sheetList.CompareMode = vbTextCompare
    
    If IsMissing(targetSheets) Then
        ' Process all visible sheets
        For Each ws In wb.Worksheets
            If ws.Visible = xlSheetVisible Then
                sheetList.Add ws.Name, ws
            End If
        Next ws
    Else
        ' Process specified sheets only
        Dim sheetName As Variant
        If IsArray(targetSheets) Then
            For Each sheetName In targetSheets
                Set ws = SheetByName(CStr(sheetName))
                If Not ws Is Nothing Then sheetList.Add ws.Name, ws
            Next sheetName
        Else
            Set ws = SheetByName(CStr(targetSheets))
            If Not ws Is Nothing Then sheetList.Add ws.Name, ws
        End If
    End If
    
    totalSheets = sheetList.count
    
    If totalSheets = 0 Then
        MsgBox "No sheets to process.", vbInformation
        GoTo CleanExit
    End If
    
    ' Process each sheet (with error handling per sheet)
    Dim wsKey As Variant
    For Each wsKey In sheetList.Keys
        Set ws = sheetList(wsKey)
        
        Dim sheetChanges As Long
        On Error Resume Next
        sheetChanges = StandardizeSheetText(ws)
        If Err.Number <> 0 Then
            ' Log error but continue with other sheets
            Debug.Print "Warning: Failed to process sheet '" & ws.Name & "': " & Err.Description
            Err.Clear
        Else
            changedCells = changedCells + sheetChanges
            processedSheets = processedSheets + 1
        End If
        On Error GoTo CleanFail
    Next wsKey
    
    Dim elapsed As Double: elapsed = Timer - startTime
    
    ' Report results
    msg = "Text Standardization Complete" & vbCrLf & vbCrLf
    msg = msg & "Sheets processed: " & processedSheets & vbCrLf
    msg = msg & "Cells changed: " & changedCells & vbCrLf
    msg = msg & "Time elapsed: " & Format$(elapsed, "0.0") & " seconds" & vbCrLf & vbCrLf
    
    If changedCells > 0 Then
        msg = msg & "Changes applied:" & vbCrLf
        msg = msg & "  - Removed Alt+0160 (non-breaking spaces)" & vbCrLf
        msg = msg & "  - Standardized all dashes to keyboard '-'" & vbCrLf
        msg = msg & "  - Standardized all underscores to '_'" & vbCrLf
        msg = msg & "  - Cleaned invisible/zero-width characters" & vbCrLf
        msg = msg & "  - Normalized lookalike letters/numbers"
    Else
        msg = msg & "No changes needed - data is already clean!"
    End If
    
    MsgBox msg, vbInformation, "Standardization Results"
    
CleanExit:
    PopAppState state
    Exit Sub
    
CleanFail:
    MsgBox "Error in StandardizeAllWorkbookText: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

' Standardize text in a single worksheet (returns count of changed cells)
Private Function StandardizeSheetText(ws As Worksheet) As Long
    Dim lastRow As Long, lastCol As Long
    Dim usedRange As Range
    Dim cell As Range
    Dim originalValue As Variant, cleanedValue As String
    Dim changedCount As Long
    
    changedCount = 0
    
    lastRow = LastUsedRow(ws)
    lastCol = LastUsedCol(ws)
    
    If lastRow = 0 Or lastCol = 0 Then
        StandardizeSheetText = 0
        Exit Function
    End If
    
    ' Get all constant cells (includes text AND numbers with text content)
    On Error Resume Next
    Set usedRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).SpecialCells(xlCellTypeConstants)
    On Error GoTo 0
    
    ' If no constants found, exit early
    If usedRange Is Nothing Then
        StandardizeSheetText = 0
        Exit Function
    End If
    
    ' Process all constant cells that contain string data
    For Each cell In usedRange
        ' Skip formulas and empty cells
        If cell.HasFormula Or IsEmpty(cell.Value2) Then GoTo NextCell
        
        originalValue = cell.Value2
        
        ' Only process string-compatible types (skip pure numbers, dates, booleans, errors)
        Select Case VarType(originalValue)
            Case vbString
                ' Standard text cell - process it
            Case vbDouble, vbDate, vbBoolean, vbError, vbEmpty, vbNull
                ' Skip these types entirely
                GoTo NextCell
            Case Else
                ' Variant subtypes, currency, decimal - can contain text, process them
        End Select
        
        ' Convert once and reuse
        cleanedValue = CStr(originalValue)
        
        ' Apply full cleaning pipeline
        cleanedValue = CleanText(cleanedValue)
        cleanedValue = StandardizeSymbols(cleanedValue)
        cleanedValue = CleanWhitespace(cleanedValue)
        
        ' Only update if changed (preserve formatting, notes, etc.)
        If cleanedValue <> CStr(originalValue) Then
            cell.Value2 = cleanedValue
            changedCount = changedCount + 1
        End If
        
NextCell:
    Next cell
    
    StandardizeSheetText = changedCount
End Function

' Comprehensive diagnostic: check workbook for problematic characters
' Returns a report showing which sheets have issues
' This is READ-ONLY and safe to run with multiple users
' maxSheets: Administrative safety limit (default 50, modify in code if needed)
Private Function DiagnoseWorkbookCharacters(Optional ByVal maxSheets As Long = 50) As String
    Dim ws As Worksheet
    Dim report As String
    Dim sheetCount As Long
    Dim issueCount As Long
    Dim totalIssues As Long
    
    report = "CHARACTER DIAGNOSIS REPORT" & vbCrLf
    report = report & String$(50, "=") & vbCrLf & vbCrLf
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible <> xlSheetVisible Then GoTo NextDiagSheet
        
        sheetCount = sheetCount + 1
        If sheetCount > maxSheets Then Exit For
        
        ' Comprehensive scan (all rows, not sampled)
        issueCount = DiagnoseSheetCharacters(ws)
        
        If issueCount > 0 Then
            report = report & ws.Name & ": " & issueCount & " cells with issues" & vbCrLf
            totalIssues = totalIssues + issueCount
        End If
        
NextDiagSheet:
    Next ws
    
    report = report & vbCrLf & "TOTAL: " & totalIssues & " cells across " & sheetCount & " sheets"
    
    If totalIssues = 0 Then
        report = report & vbCrLf & vbCrLf & "[OK] No issues detected!"
    Else
        report = report & vbCrLf & vbCrLf
        report = report & "Run StandardizeAllWorkbookText() to fix all issues."
    End If
    
    DiagnoseWorkbookCharacters = report
End Function

' Quick character issue survey for health check (configurable scan depth)
' Returns total count of cells with character issues across all sheets
' scanDepth: Row sampling percentage (0.25 = 25%, 0.5 = 50%, 1.0 = 100%)
Public Function QuickScanCharacterIssues(Optional ByVal scanDepth As Double = 1) As Long
    Dim ws As Worksheet
    Dim totalIssues As Long
    Dim lastRow As Long, lastCol As Long
    Dim r As Long, c As Long
    Dim originalValue As Variant, cleanedValue As String
    Dim scanLimit As Long
    Dim sheetCount As Long
    Const MAX_SHEETS As Long = 50  ' Administrative safety limit
    
    totalIssues = 0
    sheetCount = 0
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible <> xlSheetVisible Then GoTo NextQuickScan
        
        sheetCount = sheetCount + 1
        If sheetCount > MAX_SHEETS Then Exit For
        
        lastRow = LastUsedRow(ws)
        lastCol = LastUsedCol(ws)
        
        If lastRow = 0 Or lastCol = 0 Then GoTo NextQuickScan
        
        ' Calculate sample limit based on scan depth
        scanLimit = CLng(lastRow * scanDepth)
        If scanLimit < 1 Then scanLimit = 1
        If scanLimit > lastRow Then scanLimit = lastRow
        
        For r = 1 To scanLimit
            For c = 1 To lastCol
                On Error Resume Next
                If Not ws.Cells(r, c).HasFormula Then
                    originalValue = ws.Cells(r, c).Value2
                    
                    If VarType(originalValue) = vbString Then
                        cleanedValue = StandardizeSymbols(CleanWhitespace(CStr(originalValue)))
                        
                        If cleanedValue <> CStr(originalValue) Then
                            totalIssues = totalIssues + 1
                        End If
                    End If
                End If
                On Error GoTo 0
            Next c
        Next r
        
NextQuickScan:
    Next ws
    
    QuickScanCharacterIssues = totalIssues
End Function

' Diagnose a single sheet for character issues (returns issue count)
' Comprehensive scan - checks ALL rows (not sampled)
Private Function DiagnoseSheetCharacters(ws As Worksheet) As Long
    Dim lastRow As Long, lastCol As Long
    Dim r As Long, c As Long
    Dim cell As Range
    Dim originalValue As Variant, cleanedValue As String
    Dim issueCount As Long
    
    On Error Resume Next
    
    issueCount = 0
    lastRow = LastUsedRow(ws)
    lastCol = LastUsedCol(ws)
    
    If lastRow = 0 Or lastCol = 0 Then
        DiagnoseSheetCharacters = 0
        Exit Function
    End If
    
    For r = 1 To lastRow
        For c = 1 To lastCol
            Set cell = ws.Cells(r, c)
            
            ' Skip formulas and error cells
            If Not cell.HasFormula And Not IsError(cell.Value2) Then
                If VarType(cell.Value2) = vbString Then
                    originalValue = cell.Value2
                    cleanedValue = StandardizeSymbols(CleanWhitespace(CStr(originalValue)))
                    
                    If cleanedValue <> CStr(originalValue) Then
                        issueCount = issueCount + 1
                    End If
                End If
            End If
        Next c
    Next r
    
    On Error GoTo 0
    DiagnoseSheetCharacters = issueCount
End Function

' Show detailed character codes for a specific cell (debugging helper)
' Usage: Call ShowCharacterCodes("Sheet1", "A1") or ShowCharacterCodes(ActiveSheet.Name, Selection.Address)
Private Sub ShowCharacterCodes(ByVal sheetName As String, ByVal cellAddress As String)
    Dim ws As Worksheet
    Dim cell As Range
    Dim s As String
    Dim i As Long
    Dim charCode As Long
    Dim charName As String
    Dim output As String
    
    Set ws = SheetByName(sheetName)
    If ws Is Nothing Then
        MsgBox "Sheet '" & sheetName & "' not found.", vbExclamation
        Exit Sub
    End If
    
    On Error Resume Next
    Set cell = ws.Range(cellAddress)
    On Error GoTo 0
    
    If cell Is Nothing Then
        MsgBox "Invalid cell address: " & cellAddress, vbExclamation
        Exit Sub
    End If
    
    s = CStr(cell.Value2)
    
    If Len(s) = 0 Then
        MsgBox "Cell is empty.", vbInformation
        Exit Sub
    End If
    
    output = "CHARACTER ANALYSIS" & vbCrLf
    output = output & "Sheet: " & sheetName & vbCrLf
    output = output & "Cell: " & cellAddress & vbCrLf
    output = output & "Length: " & Len(s) & " characters" & vbCrLf
    output = output & String$(60, "=") & vbCrLf & vbCrLf
    
    For i = 1 To Len(s)
        charCode = AscW(Mid$(s, i, 1))
        charName = GetCharacterName(charCode)
        
        output = output & "Position " & i & ": "
        
        ' Show the character (or description if invisible)
        If charCode < 32 Or charCode = 127 Or charCode = 160 Or _
           (charCode >= 8203 And charCode <= 8299) Or charCode = 65279 Then
            output = output & "[" & charName & "]"
        Else
            output = output & "'" & Mid$(s, i, 1) & "'"
        End If
        
        output = output & " = Code " & charCode
        
        If charName <> "" Then
            output = output & " (" & charName & ")"
        End If
        
        ' Flag problematic characters
        If charCode = 160 Then
            output = output & " ?? NON-BREAKING SPACE (Alt+0160)"
        ElseIf charCode >= 8203 And charCode <= 8299 Then
            output = output & " ?? INVISIBLE CHARACTER"
        ElseIf charCode = 65279 Then
            output = output & " ?? BOM/ZERO-WIDTH"
        ElseIf charCode >= 8208 And charCode <= 8213 Then
            output = output & " ?? NON-STANDARD DASH"
        ElseIf charCode = 8722 Then
            output = output & " ?? MINUS SIGN (not hyphen)"
        ElseIf (charCode >= 1040 And charCode <= 1103) Then
            output = output & " ?? CYRILLIC LETTER"
        ElseIf (charCode >= 65280 And charCode <= 65519) Then
            output = output & " ?? FULLWIDTH CHARACTER"
        End If
        
        output = output & vbCrLf
    Next i
    
    ' Output to Immediate Window and show summary dialog
    Debug.Print output
    
    Dim summary As String
    summary = "Character analysis complete!" & vbCrLf & vbCrLf
    summary = summary & "Cell: " & sheetName & "!" & cellAddress & vbCrLf
    summary = summary & "Length: " & Len(s) & " characters" & vbCrLf & vbCrLf
    summary = summary & "Detailed output sent to Immediate Window (Ctrl+G)."
    
    MsgBox summary, vbInformation, "Character Analysis"
End Sub

' Helper function to get friendly names for character codes
Private Function GetCharacterName(ByVal code As Long) As String
    Select Case code
        Case 9: GetCharacterName = "TAB"
        Case 10: GetCharacterName = "LINE FEED (Alt+Enter)"
        Case 13: GetCharacterName = "CARRIAGE RETURN"
        Case 32: GetCharacterName = "SPACE"
        Case 160: GetCharacterName = "NON-BREAKING SPACE"
        Case 8203: GetCharacterName = "ZERO-WIDTH SPACE"
        Case 8204: GetCharacterName = "ZERO-WIDTH NON-JOINER"
        Case 8205: GetCharacterName = "ZERO-WIDTH JOINER"
        Case 8208: GetCharacterName = "HYPHEN"
        Case 8209: GetCharacterName = "NON-BREAKING HYPHEN"
        Case 8210: GetCharacterName = "FIGURE DASH"
        Case 8211: GetCharacterName = "EN DASH"
        Case 8212: GetCharacterName = "EM DASH"
        Case 8213: GetCharacterName = "HORIZONTAL BAR"
        Case 8216: GetCharacterName = "LEFT SINGLE QUOTE"
        Case 8217: GetCharacterName = "RIGHT SINGLE QUOTE"
        Case 8220: GetCharacterName = "LEFT DOUBLE QUOTE"
        Case 8221: GetCharacterName = "RIGHT DOUBLE QUOTE"
        Case 8722: GetCharacterName = "MINUS SIGN"
        Case 65279: GetCharacterName = "BOM/ZERO-WIDTH NO-BREAK"
        Case 8288: GetCharacterName = "WORD JOINER"
        Case Else: GetCharacterName = ""
    End Select
End Function

' Quick cell inspector - call from Immediate Window or button
' Usage: Call InspectActiveCell()
Private Sub InspectActiveCell()
    On Error Resume Next
    If TypeName(Selection) = "Range" Then
        ShowCharacterCodes ActiveSheet.Name, Selection.Address
    Else
        MsgBox "Please select a cell first.", vbInformation
    End If
    On Error GoTo 0
End Sub



