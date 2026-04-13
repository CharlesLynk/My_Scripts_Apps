Option Explicit

'==========================================
' Module: ReloadManager
' Purpose: Sync workbook VBA standard modules from repo BAS sources.
' Scope: Workbook type tagging, signature checks, and module re-import.
' Notes:
' - XLAM is expected in %APPDATA%\Microsoft\AddIns.
' - Managed from File -> Options -> Add-Ins.
' - Available as VBAProject(Global).ReloadManager.
'==========================================

'================== CONFIGURATION ==================
Private Const REPO_ROOT As String = "D:\Utilities\My_Scripts_Apps\Excel"

' Document properties used in target workbooks
Private Const PROP_TYPE_ID As String = "WB_TYPE_ID"
Private Const PROP_SRC_SIG As String = "WB_SRC_SIGNATURE"
Private Const PROP_LAST_SYNC As String = "WB_LAST_SYNC_UTC"

'================== PUBLIC ENTRY POINTS ==================

' Sync the active workbook against the central repo for its type id.
Public Sub SyncActiveWorkbook()
    Dim wb As Workbook
    Set wb = ActiveWorkbook

    If wb Is Nothing Then Exit Sub
    If wb.Name Like "*.xlam" Then Exit Sub ' don't sync the add-in itself

    Dim typeId As String
    typeId = GetDocProp(wb, PROP_TYPE_ID)
    If Len(typeId) = 0 Then
        MsgBox "This workbook is missing the custom property " & PROP_TYPE_ID & "." & vbCrLf & _
               "Run SetActiveWorkbookTypeId first.", vbExclamation
        Exit Sub
    End If

    Dim srcPath As String
    srcPath = REPO_ROOT & Application.PathSeparator & typeId & _
              Application.PathSeparator & "src" & Application.PathSeparator

    If Len(Dir(srcPath, vbDirectory)) = 0 Then
        MsgBox "Source folder not found:" & vbCrLf & srcPath, vbExclamation
        Exit Sub
    End If

    ' Compute signature of current repo source
    Dim repoSig As String
    repoSig = ComputeFolderSignature(srcPath, "*.bas")

    If Len(repoSig) = 0 Then
        MsgBox "No .bas files found in:" & vbCrLf & srcPath, vbExclamation
        Exit Sub
    End If

    Dim wbSig As String
    wbSig = GetDocProp(wb, PROP_SRC_SIG)

    If StrComp(repoSig, wbSig, vbBinaryCompare) = 0 Then
        ' Up-to-date
        Exit Sub
    End If

    ' Perform sync: remove + import standard modules
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ImportBasFolder wb, srcPath

    SetDocProp wb, PROP_SRC_SIG, repoSig
    SetDocProp wb, PROP_LAST_SYNC, Format$(NowUTC(), "yyyy-mm-dd\THH:NN:SS\Z")

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub


' Helper: set the workbook type id once (stores as a custom doc property)
Public Sub SetActiveWorkbookTypeId()
    Dim wb As Workbook
    Set wb = ActiveWorkbook

    Dim typeId As String
    typeId = InputBox("Enter WB_TYPE_ID (e.g., ABProgramming):", "Set Workbook Type ID")
    If Len(typeId) = 0 Then Exit Sub

    SetDocProp wb, PROP_TYPE_ID, typeId
    MsgBox "WB_TYPE_ID set to: " & typeId, vbInformation
End Sub


'================== CORE IMPORT LOGIC ==================

Private Sub ImportBasFolder(ByVal wb As Workbook, ByVal folderPath As String)
    Dim f As String
    f = Dir(folderPath & "*.bas")

    Do While Len(f) > 0
        Dim compName As String
        compName = Left$(f, Len(f) - 4)

        ' Remove existing standard module of same name, then import file
        RemoveStdModuleIfExists wb, compName
        wb.VBProject.VBComponents.Import folderPath & f

        f = Dir()
    Loop
End Sub

Private Sub RemoveStdModuleIfExists(ByVal wb As Workbook, ByVal compName As String)
    Dim vbComp As Object

    On Error Resume Next
    Set vbComp = wb.VBProject.VBComponents(compName)
    On Error GoTo 0

    If Not vbComp Is Nothing Then
        ' vbext_ct_StdModule = 1. Document modules are type 100 and should not be removed. [3](https://bettersolutions.com/vba/visual-basic-editor/extensibility-object-model.htm)
        If vbComp.Type = 1 Then
            wb.VBProject.VBComponents.Remove vbComp
        End If
    End If
End Sub


'================== SIGNATURE / VERSION CHECK ==================
' Creates a stable signature based on file name + size + modified time (UTC-ish).
' This avoids re-importing when nothing changed.

Private Function ComputeFolderSignature(ByVal folderPath As String, ByVal pattern As String) As String
    Dim items() As String
    Dim count As Long
    count = 0

    Dim f As String
    f = Dir(folderPath & pattern)
    Do While Len(f) > 0
        count = count + 1
        ReDim Preserve items(1 To count)

        Dim fullPath As String
        fullPath = folderPath & f

        Dim sizeBytes As Double
        sizeBytes = FileLen(fullPath)

        Dim modDate As Date
        modDate = FileDateTime(fullPath)

        ' Normalize record
        items(count) = LCase$(f) & "|" & CStr(sizeBytes) & "|" & Format$(modDate, "yyyy-mm-dd hh:nn:ss")
        f = Dir()
    Loop

    If count = 0 Then
        ComputeFolderSignature = ""
        Exit Function
    End If

    QuickSort items, LBound(items), UBound(items)

    Dim i As Long, sig As String
    sig = ""
    For i = LBound(items) To UBound(items)
        sig = sig & items(i) & vbLf
    Next

    ' Simple folding hash (not cryptographic) to keep property short
    ComputeFolderSignature = CStr(SimpleHash32(sig)) & ":" & CStr(count)
End Function

Private Function SimpleHash32(ByVal s As String) As Long
    Dim i As Long
    Dim h As Long
    Dim hd As Double
    h = -2128831035

    For i = 1 To Len(s)
        h = h Xor AscW(Mid$(s, i, 1))
        hd = CDbl(h) * 16777619#
        hd = hd - 4294967296# * Fix(hd / 4294967296#) ' wrap to unsigned 32-bit
        If hd > 2147483647# Then
            hd = hd - 4294967296#
        ElseIf hd < -2147483648# Then
            hd = hd + 4294967296#
        End If
        h = CLng(hd)
    Next

    If h > 2147483647# Then
        h = h - 4294967296#
    End If

    SimpleHash32 = CLng(h)
End Function

Private Sub QuickSort(arr() As String, ByVal first As Long, ByVal last As Long)
    Dim i As Long, j As Long
    Dim pivot As String, temp As String
    i = first
    j = last
    pivot = arr((first + last) \ 2)

    Do While i <= j
        Do While arr(i) < pivot
            i = i + 1
        Loop
        Do While arr(j) > pivot
            j = j - 1
        Loop
        If i <= j Then
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
            i = i + 1
            j = j - 1
        End If
    Loop

    If first < j Then QuickSort arr, first, j
    If i < last Then QuickSort arr, i, last
End Sub


'================== DOCUMENT PROPERTY HELPERS ==================

Private Function GetDocProp(ByVal wb As Workbook, ByVal propName As String) As String
    On Error Resume Next
    GetDocProp = CStr(wb.CustomDocumentProperties(propName).Value)
    On Error GoTo 0
End Function

Private Sub SetDocProp(ByVal wb As Workbook, ByVal propName As String, ByVal propValue As String)
    On Error Resume Next
    wb.CustomDocumentProperties(propName).Value = propValue
    If Err.Number <> 0 Then
        Err.Clear
        wb.CustomDocumentProperties.Add Name:=propName, LinkToContent:=False, _
                                       Type:=msoPropertyTypeString, Value:=propValue
    End If
    On Error GoTo 0
End Sub

Private Function NowUTC() As Date
    ' VBA Date is local; approximate UTC using system bias
    NowUTC = DateAdd("n", -CLng(GetTimeZoneBiasMinutes()), Now)
End Function

Private Function GetTimeZoneBiasMinutes() As Long
    ' Returns minutes to add to local time to get UTC (bias).
    ' Uses WMI via COM; if unavailable, returns 0.
    On Error GoTo Fail
    Dim wmi As Object, col As Object, itm As Object
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Set col = wmi.ExecQuery("Select Bias from Win32_TimeZone")
    For Each itm In col
        GetTimeZoneBiasMinutes = CLng(itm.Bias)
        Exit Function
    Next
Fail:
    GetTimeZoneBiasMinutes = 0
End Function