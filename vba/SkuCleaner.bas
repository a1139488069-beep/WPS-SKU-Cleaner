Attribute VB_Name = "SkuCleaner"
Option Explicit

' WPS VBA / Microsoft Excel. No external reference required.
Public Const SKU_HEADER_ROW As Long = 0
Public Const SKU_SCAN_ROWS As Long = 20
Public Const SKU_HEADER_CONTAINS As Boolean = True
Public SKU_AutoEnabled As Boolean
Private SKU_Busy As Boolean
Private SKU_LastReport As String

Private Type SkuStats
    Visited As Long
    Columns As Long
    Converted As Long
    Numeric As Long
    TextFormatted As Long
    Blank As Long
    Invalid As Long
    TooLong As Long
    Formulas As Long
    Merged As Long
    ProtectedSheets As Long
    Failed As Long
End Type

Private Function CleanSkuText(ByVal text As String) As String
    Dim i As Long, code As Long, ch As String, result As String
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        code = AscW(ch)
        If code < 0 Then code = code + 65536
        Select Case code
            Case 9 To 13, 32, 160, 5760, 8192 To 8207, 8232 To 8238, 8239, 8287, 8288, 12288, 65279
                ' Remove whitespace and common invisible characters.
            Case 65296 To 65305
                result = result & Chr$(code - 65296 + 48)
            Case Else
                result = result & ch
        End Select
    Next i
    Do While Len(result) > 0
        code = AscW(Left$(result, 1))
        If code = 39 Or code = 8216 Or code = 8217 Then
            result = Mid$(result, 2)
        Else
            Exit Do
        End If
    Loop
    CleanSkuText = result
End Function

Private Function IsSkuHeader(ByVal cell As Object) As Boolean
    Dim text As String
    If cell.HasFormula Or cell.MergeCells Then Exit Function
    If IsError(cell.Value2) Or VarType(cell.Value2) <> vbString Then Exit Function
    text = LCase$(CleanSkuText(CStr(cell.Value2)))
    If SKU_HEADER_CONTAINS Then
        IsSkuHeader = (InStr(1, text, "sku", vbBinaryCompare) > 0)
    Else
        IsSkuHeader = (text = "sku")
    End If
End Function

Private Function FindHeaderRow(ByVal sh As Object) As Long
    Dim used As Object, r As Long, c As Long, firstRow As Long, lastRow As Long, lastCol As Long
    Set used = sh.UsedRange
    lastCol = used.Column + used.Columns.Count - 1
    If lastCol > 256 Then lastCol = 256
    If SKU_HEADER_ROW > 0 Then
        firstRow = SKU_HEADER_ROW
        lastRow = firstRow
    Else
        firstRow = used.Row
        lastRow = used.Row + used.Rows.Count - 1
        If lastRow > SKU_SCAN_ROWS Then lastRow = SKU_SCAN_ROWS
    End If
    For r = firstRow To lastRow
        For c = used.Column To lastCol
            If IsSkuHeader(sh.Cells(r, c)) Then
                FindHeaderRow = r
                Exit Function
            End If
        Next c
    Next r
End Function

Private Sub ProcessCell(ByVal cell As Object, ByRef stats As SkuStats)
    Dim value As Variant, text As String, i As Long, code As Long, number As Double
    On Error GoTo Failed
    stats.Visited = stats.Visited + 1
    If cell.MergeCells Then stats.Merged = stats.Merged + 1: Exit Sub
    If cell.HasFormula Then stats.Formulas = stats.Formulas + 1: Exit Sub
    value = cell.Value2
    If IsError(value) Then GoTo Invalid
    If IsEmpty(value) Then GoTo Blank
    If VarType(value) = vbString Then
        If Len(value) = 0 Then GoTo Blank
        text = CleanSkuText(CStr(value))
        If Len(text) = 0 Then
            cell.ClearContents
            GoTo Blank
        End If
        For i = 1 To Len(text)
            code = AscW(Mid$(text, i, 1))
            If code < 48 Or code > 57 Then GoTo TextSku
        Next i
        Do While Len(text) > 1 And Left$(text, 1) = "0"
            text = Mid$(text, 2)
        Loop
        If Len(text) > 15 Then GoTo TooLong
        number = CDbl(text)
        If cell.NumberFormat <> "0" Then cell.NumberFormat = "0"
        cell.Value2 = number
        stats.Converted = stats.Converted + 1
        Exit Sub
    End If
    If VarType(value) = vbBoolean Or Not IsNumeric(value) Then GoTo Invalid
    number = CDbl(value)
    If number < 0 Or Fix(number) <> number Then GoTo Invalid
    If number > 999999999999999# Then GoTo TooLong
    If cell.NumberFormat <> "0" Then cell.NumberFormat = "0"
    stats.Numeric = stats.Numeric + 1
    Exit Sub
TextSku:
    ' Write as text first to avoid accidental formula/number/date interpretation.
    If CStr(value) <> text Then
        cell.NumberFormat = "@"
        cell.Value2 = text
    End If
    If cell.NumberFormat <> "0" Then cell.NumberFormat = "0"
    stats.TextFormatted = stats.TextFormatted + 1
    Exit Sub
Blank:
    stats.Blank = stats.Blank + 1
    Exit Sub
Invalid:
    stats.Invalid = stats.Invalid + 1
    Exit Sub
TooLong:
    stats.TooLong = stats.TooLong + 1
    Exit Sub
Failed:
    stats.Failed = stats.Failed + 1
End Sub

Private Sub ProcessSheet(ByVal sh As Object, ByVal target As Object, ByRef stats As SkuStats)
    Dim used As Object, headerRow As Long, lastRow As Long, c As Long, lastCol As Long
    Dim data As Object, scope As Object, cell As Object, headerHit As Object
    If sh.ProtectContents Then
        stats.ProtectedSheets = stats.ProtectedSheets + 1
        Exit Sub
    End If
    headerRow = FindHeaderRow(sh)
    If headerRow = 0 Then Exit Sub
    Set used = sh.UsedRange
    lastRow = used.Row + used.Rows.Count - 1
    If lastRow <= headerRow Then Exit Sub
    If target Is Nothing And lastRow - headerRow > 2000 Then Err.Raise vbObjectError + 501, , "VBA manual mode is limited to 2000 rows. Select a smaller region or use the JS add-in."
    lastCol = used.Column + used.Columns.Count - 1
    If lastCol > 256 Then lastCol = 256
    For c = used.Column To lastCol
        If IsSkuHeader(sh.Cells(headerRow, c)) Then
            Set data = sh.Range(sh.Cells(headerRow + 1, c), sh.Cells(lastRow, c))
            Set scope = Nothing
            If target Is Nothing Then
                Set scope = data
            Else
                Set scope = Application.Intersect(data, target)
            End If
            If Not scope Is Nothing Then
                If scope.Cells.Count > 2000 - stats.Visited Then Err.Raise vbObjectError + 502, , "VBA manual mode is limited to 2000 SKU cells per run. Select a smaller region."
                stats.Columns = stats.Columns + 1
                For Each cell In scope.Cells
                    ProcessCell cell, stats
                Next cell
            End If
        End If
    Next c
End Sub

Private Function Report(ByRef stats As SkuStats) As String
    Report = "SKU cleanup complete" & vbCrLf & "Matched columns: " & stats.Columns & _
        vbCrLf & "Text to number: " & stats.Converted & vbCrLf & "Numeric formatted: " & stats.Numeric & _
        vbCrLf & "Text identifiers formatted: " & stats.TextFormatted & vbCrLf & "Blank: " & stats.Blank & vbCrLf & "Unsupported values skipped: " & stats.Invalid & _
        vbCrLf & "Long IDs preserved: " & stats.TooLong & vbCrLf & "Formulas skipped: " & stats.Formulas & _
        vbCrLf & "Merged cells skipped: " & stats.Merged & vbCrLf & "Protected sheets skipped: " & stats.ProtectedSheets & _
        vbCrLf & "Write failures: " & stats.Failed
End Function

Private Sub Execute(ByVal book As Object, ByVal sh As Object, ByVal target As Object, ByVal showReport As Boolean)
    Dim stats As SkuStats, ws As Object, oldEvents As Boolean, oldScreen As Boolean
    Dim captured As Boolean, errorNumber As Long, errorMessage As String
    If SKU_Busy Then Exit Sub
    SKU_Busy = True
    On Error GoTo Failed
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    captured = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    If sh Is Nothing Then
        For Each ws In book.Worksheets
            ProcessSheet ws, Nothing, stats
        Next ws
    Else
        ProcessSheet sh, target, stats
    End If
    SKU_LastReport = Report(stats)
    GoTo Cleanup
Failed:
    errorNumber = Err.Number
    errorMessage = Err.Description
Cleanup:
    On Error Resume Next
    If captured Then
        Application.ScreenUpdating = oldScreen
        Application.EnableEvents = oldEvents
    End If
    SKU_Busy = False
    On Error GoTo 0
    If errorNumber <> 0 Then
        SKU_AutoEnabled = False
        MsgBox "SKU cleanup failed; auto mode paused: " & errorMessage, vbExclamation
    ElseIf showReport Then
        MsgBox SKU_LastReport, vbInformation
    End If
End Sub

Public Sub SKU_CleanCurrentSheet()
    If Application.ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(Application.ActiveSheet) <> "Worksheet" Then Exit Sub
    Execute Application.ActiveWorkbook, Application.ActiveSheet, Nothing, True
End Sub
Public Sub SKU_CleanWorkbook()
    If Application.ActiveWorkbook Is Nothing Then Exit Sub
    Execute Application.ActiveWorkbook, Nothing, Nothing, True
End Sub
Public Sub SKU_EnableAuto()
    SKU_AutoEnabled = False
    MsgBox "Use the JS add-in for automatic processing. This VBA fallback supports small manual runs only.", vbInformation
End Sub
Public Sub SKU_DisableAuto()
    SKU_AutoEnabled = False
    MsgBox "SKU auto mode disabled.", vbInformation
End Sub
Public Sub SKU_Status()
    MsgBox "Auto mode: " & IIf(SKU_AutoEnabled, "ON", "OFF") & vbCrLf & SKU_LastReport, vbInformation
End Sub
Public Sub SKU_HandleChange(ByVal sh As Object, ByVal target As Object)
    If Not SKU_AutoEnabled Or SKU_Busy Then Exit Sub
    Execute ThisWorkbook, sh, target, False
End Sub
Public Sub SKU_OnOpen()
    SKU_AutoEnabled = False
End Sub
Public Sub SKU_CleanSelection()
    If Application.ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(Application.Selection) <> "Range" Then Exit Sub
    Execute Application.ActiveWorkbook, Application.ActiveSheet, Application.Selection, True
End Sub
