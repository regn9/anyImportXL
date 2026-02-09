Attribute VB_Name = "modConfig"
Option Explicit

Private Const CONFIG_SHEET As String = "zConfig"

Public Sub EnsureConfigSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = CONFIG_SHEET
        ws.Visible = xlSheetVeryHidden
        ws.Range("A1:H1").Value = Array("ProfileKey", "Kind", "Name", "Value1", "Value2", "Value3", "UpdatedAt", "Reserved")
    End If
End Sub

Public Sub SaveProfile(ByVal profileKey As String, ByVal vars As Collection, ByVal outputs As Collection, ByVal targets As Collection)
    EnsureConfigSheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET)

    Dim r As Long
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = r To 2 Step -1
        If CStr(ws.Cells(i, 1).Value2) = profileKey Then ws.Rows(i).Delete
    Next i

    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    For i = 1 To vars.Count
        Dim vb As TVarBinding
        vb = vars(i)
        WriteCfgRow ws, r, profileKey, "VAR", vb.VarName, vb.AccountCode, vb.Label, vb.Metric
        r = r + 1
    Next i

    For i = 1 To outputs.Count
        Dim od As TOutputDef
        od = outputs(i)
        WriteCfgRow ws, r, profileKey, "OUT", od.OutputName, od.FormulaText, "", ""
        r = r + 1
    Next i

    For i = 1 To targets.Count
        Dim tm As TTargetMap
        tm = targets(i)
        WriteCfgRow ws, r, profileKey, "MAP", tm.OutputName, tm.TargetSheet, tm.TargetAddress, ""
        r = r + 1
    Next i
End Sub

Public Sub LoadProfile(ByVal profileKey As String, ByRef vars As Collection, ByRef outputs As Collection, ByRef targets As Collection)
    EnsureConfigSheet
    Set vars = New Collection
    Set outputs = New Collection
    Set targets = New Collection

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = 2 To lastRow
        If CStr(ws.Cells(i, 1).Value2) = profileKey Then
            Dim kind As String
            kind = CStr(ws.Cells(i, 2).Value2)
            Select Case kind
                Case "VAR"
                    Dim vb As TVarBinding
                    vb.VarName = CStr(ws.Cells(i, 3).Value2)
                    vb.AccountCode = CStr(ws.Cells(i, 4).Value2)
                    vb.Label = CStr(ws.Cells(i, 5).Value2)
                    vb.Metric = CStr(ws.Cells(i, 6).Value2)
                    vars.Add vb
                Case "OUT"
                    Dim od As TOutputDef
                    od.OutputName = CStr(ws.Cells(i, 3).Value2)
                    od.FormulaText = CStr(ws.Cells(i, 4).Value2)
                    outputs.Add od
                Case "MAP"
                    Dim tm As TTargetMap
                    tm.OutputName = CStr(ws.Cells(i, 3).Value2)
                    tm.TargetSheet = CStr(ws.Cells(i, 4).Value2)
                    tm.TargetAddress = CStr(ws.Cells(i, 5).Value2)
                    targets.Add tm
            End Select
        End If
    Next i
End Sub

Private Sub WriteCfgRow(ByVal ws As Worksheet, ByVal r As Long, ByVal profileKey As String, ByVal kind As String, _
                        ByVal name As String, ByVal v1 As String, ByVal v2 As String, ByVal v3 As String)
    ws.Cells(r, 1).Value2 = profileKey
    ws.Cells(r, 2).Value2 = kind
    ws.Cells(r, 3).Value2 = name
    ws.Cells(r, 4).Value2 = v1
    ws.Cells(r, 5).Value2 = v2
    ws.Cells(r, 6).Value2 = v3
    ws.Cells(r, 7).Value2 = Now
End Sub

Public Function CurrentProfileKey() As String
    If ActiveWorkbook Is Nothing Then
        CurrentProfileKey = ""
    Else
        CurrentProfileKey = ActiveWorkbook.FullName
    End If
End Function
