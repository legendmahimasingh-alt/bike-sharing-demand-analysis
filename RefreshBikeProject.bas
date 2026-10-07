Attribute VB_Name = "BikeProject"
Option Explicit

Public Sub RefreshBikeProject()
    Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet
    Dim master As Worksheet, analysis As Worksheet, dash As Worksheet
    Dim records As Object, atempMap As Object
    Dim lastRow As Long, r As Long, c As Long, key As Variant
    Dim rec As Variant, keys As Variant, outRow As Long, n As Long
    Dim meanCnt As Double, sdCnt As Double, cntRange As Range
    Dim i As Long, h As Long, w As Long, k As Long
    Dim headers As Variant, formulaRow As Long

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo CleanFail

    Set ws1 = ThisWorkbook.Worksheets("Raw_1")
    Set ws2 = ThisWorkbook.Worksheets("Raw_2")
    Set ws3 = ThisWorkbook.Worksheets("Raw_3")
    Set master = ThisWorkbook.Worksheets("Master")
    Set analysis = ThisWorkbook.Worksheets("Analysis")
    Set dash = ThisWorkbook.Worksheets("Dashboard")
    Set records = CreateObject("Scripting.Dictionary")
    Set atempMap = CreateObject("Scripting.Dictionary")

    ' Dataset 2: index atemp by instant (instant is column B; atemp is column C).
    lastRow = ws2.Cells(ws2.Rows.Count, 2).End(xlUp).Row
    For r = 2 To lastRow
        If Len(ws2.Cells(r, 2).Value) > 0 Then
            atempMap(CStr(CLng(ws2.Cells(r, 2).Value))) = ws2.Cells(r, 3).Value
        End If
    Next r

    ' Dataset 1: skip blank keys and keep one row per instant.
    lastRow = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If Len(ws1.Cells(r, 1).Value) > 0 Then
            key = CStr(CLng(ws1.Cells(r, 1).Value))
            If Not records.Exists(key) Then
                rec = ws1.Cells(r, 1).Resize(1, 16).Value
                If Len(rec(1, 11)) = 0 And atempMap.Exists(key) Then rec(1, 11) = atempMap(key)
                records.Add key, rec
            End If
        End If
    Next r

    ' Dataset 3: append new instants only.
    lastRow = ws3.Cells(ws3.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If Len(ws3.Cells(r, 1).Value) > 0 Then
            key = CStr(CLng(ws3.Cells(r, 1).Value))
            If Not records.Exists(key) Then records.Add key, ws3.Cells(r, 1).Resize(1, 16).Value
        End If
    Next r

    headers = Array("instant", "dteday", "season", "yr", "mnth", "hr", "holiday", "weekday", "weathersit", "temp", "atemp", "hum", "windspeed", "casual", "registered", "cnt", "day_type", "demand_band", "anomaly_flag")
    master.Cells.ClearContents
    For c = LBound(headers) To UBound(headers)
        master.Cells(1, c + 1).Value = headers(c)
    Next c
    keys = records.Keys
    ' The supplied files use consecutive instant keys; sort to keep the master in time order.
    SortKeys keys, LBound(keys), UBound(keys)
    outRow = 2
    For i = LBound(keys) To UBound(keys)
        rec = records(keys(i))
        master.Cells(outRow, 1).Resize(1, 16).Value = rec
        If CLng(rec(1, 8)) = 0 Or CLng(rec(1, 8)) = 6 Then
            master.Cells(outRow, 17).Value = "Weekend"
        Else
            master.Cells(outRow, 17).Value = "Weekday"
        End If
        outRow = outRow + 1
    Next i

    n = outRow - 2
    If n = 0 Then Err.Raise vbObjectError + 100, , "No valid records found."
    Set cntRange = master.Range("P2:P" & outRow - 1)
    meanCnt = Application.WorksheetFunction.Average(cntRange)
    sdCnt = Application.WorksheetFunction.StDev_S(cntRange)
    For r = 2 To outRow - 1
        If master.Cells(r, 16).Value >= meanCnt + sdCnt Then
            master.Cells(r, 18).Value = "High"
        Else
            master.Cells(r, 18).Value = "Typical"
        End If
        If master.Cells(r, 16).Value > meanCnt + 2 * sdCnt Then
            master.Cells(r, 19).Value = "Check"
        Else
            master.Cells(r, 19).Value = "-"
        End If
    Next r

    master.Range("A1:S" & outRow - 1).AutoFilter
    master.Rows(1).Font.Bold = True
    master.Columns("B").NumberFormat = "yyyy-mm-dd"
    master.Columns("A:S").AutoFit

    ' Refresh hour, weather, weekday/weekend, correlation, and baseline forecast tables.
    For h = 0 To 23
        analysis.Cells(h + 2, 1).Value = h
        analysis.Cells(h + 2, 2).Formula = "=AVERAGEIF(Master!$F$2:$F$" & outRow - 1 & ",A" & h + 2 & ",Master!$P$2:$P$" & outRow - 1 & ")"
        analysis.Cells(h + 2, 3).Formula = "=COUNTIF(Master!$F$2:$F$" & outRow - 1 & ",A" & h + 2 & ")"
        analysis.Cells(h + 2, 4).Formula = "=AVERAGEIF(Master!$F$2:$F$" & outRow - 1 & ",A" & h + 2 & ",Master!$O$2:$O$" & outRow - 1 & ")"
        analysis.Cells(h + 2, 17).Value = h
        analysis.Cells(h + 2, 18).Formula = "=B" & h + 2
        analysis.Cells(h + 2, 19).Value = "Historical mean by hour"
    Next h
    For w = 1 To 4
        analysis.Cells(w + 1, 6).Value = w
        analysis.Cells(w + 1, 7).Formula = "=AVERAGEIF(Master!$I$2:$I$" & outRow - 1 & ",F" & w + 1 & ",Master!$P$2:$P$" & outRow - 1 & ")"
        analysis.Cells(w + 1, 8).Formula = "=COUNTIF(Master!$I$2:$I$" & outRow - 1 & ",F" & w + 1 & ")"
    Next w
    analysis.Range("J2:J3").Value = Application.Transpose(Array("Weekday", "Weekend"))
    For w = 0 To 1
        analysis.Cells(w + 2, 11).Formula = "=AVERAGEIF(Master!$Q$2:$Q$" & outRow - 1 & ",J" & w + 2 & ",Master!$P$2:$P$" & outRow - 1 & ")"
        analysis.Cells(w + 2, 12).Formula = "=COUNTIF(Master!$Q$2:$Q$" & outRow - 1 & ",J" & w + 2 & ")"
    Next w
    analysis.Range("O2").Formula = "=CORREL(Master!$J$2:$J$" & outRow - 1 & ",Master!$P$2:$P$" & outRow - 1 & ")"
    analysis.Range("O3").Formula = "=CORREL(Master!$K$2:$K$" & outRow - 1 & ",Master!$P$2:$P$" & outRow - 1 & ")"
    analysis.Range("O4").Formula = "=CORREL(Master!$L$2:$L$" & outRow - 1 & ",Master!$P$2:$P$" & outRow - 1 & ")"
    analysis.Range("O5").Formula = "=CORREL(Master!$M$2:$M$" & outRow - 1 & ",Master!$P$2:$P$" & outRow - 1 & ")"
    dash.Range("B5").Formula = "=SUM(Master!P2:P" & outRow - 1 & ")"
    dash.Range("B6").Formula = "=AVERAGE(Master!P2:P" & outRow - 1 & ")"
    dash.Range("B7").Formula = "=SUM(Master!O2:O" & outRow - 1 & ")/SUM(Master!P2:P" & outRow - 1 & ")"
    Application.CalculateFull
    MsgBox n & " records refreshed. Review the Dashboard and anomaly flags.", vbInformation

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub
CleanFail:
    MsgBox "Refresh stopped: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

Private Sub SortKeys(ByRef arr As Variant, ByVal first As Long, ByVal last As Long)
    Dim i As Long, j As Long, pivot As Double, tmp As Variant
    i = first: j = last: pivot = CDbl(arr((first + last) \ 2))
    Do While i <= j
        Do While CDbl(arr(i)) < pivot: i = i + 1: Loop
        Do While CDbl(arr(j)) > pivot: j = j - 1: Loop
        If i <= j Then
            tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop
    If first < j Then SortKeys arr, first, j
    If i < last Then SortKeys arr, i, last
End Sub



