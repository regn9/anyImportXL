Attribute VB_Name = "modWriter"
Option Explicit

Public Sub ApplyOutputValues(ByVal outputs As Collection, ByVal targets As Collection)
    On Error GoTo CleanFail

    Dim calcState As XlCalculation
    Dim evtState As Boolean, suState As Boolean
    calcState = Application.Calculation
    evtState = Application.EnableEvents
    suState = Application.ScreenUpdating

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim i As Long
    For i = 1 To outputs.Count
        Dim od As COutputDef
        Set od = outputs(i)

        Dim tm As CTargetMap
        Set tm = FindTargetMap(targets, od.OutputName)
        If Len(tm.TargetSheet) > 0 And Len(tm.TargetAddress) > 0 Then
            Dim ws As Worksheet
            Set ws = ActiveWorkbook.Worksheets(tm.TargetSheet)
            ws.Range(tm.TargetAddress).Value2 = od.LastValue
        End If
    Next i

CleanExit:
    Application.ScreenUpdating = suState
    Application.EnableEvents = evtState
    Application.Calculation = calcState
    Exit Sub

CleanFail:
    LogMessage "ERROR", "ApplyOutputValues failed", Err.Description
    Resume CleanExit
End Sub

Public Function FindTargetMap(ByVal targets As Collection, ByVal outputName As String) As CTargetMap
    Dim i As Long
    For i = 1 To targets.Count
        Dim tm As CTargetMap
        Set tm = targets(i)
        If UCase$(tm.OutputName) = UCase$(outputName) Then
            Set FindTargetMap = tm
            Exit Function
        End If
    Next i
    Set FindTargetMap = NewTargetMap()
End Function

Public Sub LogMessage(ByVal level As String, ByVal msg As String, Optional ByVal details As String = "")
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = Nothing
    Set ws = ActiveWorkbook.Worksheets("_ImportLog")
    If ws Is Nothing Then
        Set ws = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
        ws.Name = "_ImportLog"
        ws.Range("A1:D1").Value = Array("Timestamp", "Level", "Message", "Details")
    End If

    Dim r As Long
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value2 = Now
    ws.Cells(r, 2).Value2 = level
    ws.Cells(r, 3).Value2 = msg
    ws.Cells(r, 4).Value2 = details
End Sub

