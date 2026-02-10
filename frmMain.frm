VERSION 5.00
Begin VB.UserForm frmMain 
   Caption         =   "anyImportXL"
   ClientHeight    =   9000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12000
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnImport 
      Caption         =   "Import report..."
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1450
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   150
      Width           =   2400
   End
   Begin VB.ListBox lstRows 
      Height          =   2700
      Left            =   120
      TabIndex        =   2
      Top             =   540
      Width           =   5400
   End
   Begin VB.ComboBox cmbVar 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   840
   End
   Begin VB.ComboBox cmbMetric 
      Height          =   315
      Left            =   1080
      TabIndex        =   4
      Top             =   3360
      Width           =   1200
   End
   Begin VB.CommandButton btnBindVar 
      Caption         =   "Assign var"
      Height          =   315
      Left            =   2400
      TabIndex        =   5
      Top             =   3360
      Width           =   1080
   End
   Begin VB.ListBox lstVars 
      Height          =   1800
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   5400
   End
   Begin VB.ListBox lstOutputs 
      Height          =   2400
      Left            =   5640
      TabIndex        =   7
      Top             =   540
      Width           =   6180
   End
   Begin VB.TextBox txtFormula 
      Height          =   315
      Left            =   6600
      TabIndex        =   8
      Top             =   3120
      Width           =   1860
   End
   Begin VB.ComboBox cmbTargetSheet 
      Height          =   315
      Left            =   8580
      TabIndex        =   9
      Top             =   3120
      Width           =   1440
   End
   Begin VB.TextBox txtTargetCell 
      Height          =   315
      Left            =   10080
      TabIndex        =   10
      Top             =   3120
      Width           =   840
   End
   Begin VB.CommandButton btnPickCell 
      Caption         =   "Pick cell..."
      Height          =   315
      Left            =   10980
      TabIndex        =   11
      Top             =   3120
      Width           =   900
   End
   Begin VB.CommandButton btnSaveOutput 
      Caption         =   "Save O#"
      Height          =   315
      Left            =   5640
      TabIndex        =   12
      Top             =   3120
      Width           =   900
   End
   Begin VB.CommandButton btnPreview 
      Caption         =   "Preview"
      Height          =   360
      Left            =   5640
      TabIndex        =   13
      Top             =   3600
      Width           =   1080
   End
   Begin VB.CommandButton btnApply 
      Caption         =   "Apply"
      Height          =   360
      Left            =   6780
      TabIndex        =   14
      Top             =   3600
      Width           =   1080
   End
   Begin VB.CommandButton btnSaveProfile 
      Caption         =   "Save profile"
      Height          =   360
      Left            =   7860
      TabIndex        =   15
      Top             =   3600
      Width           =   1260
   End
   Begin VB.CommandButton btnLoadProfile 
      Caption         =   "Load profile"
      Height          =   360
      Left            =   9180
      TabIndex        =   16
      Top             =   3600
      Width           =   1260
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mRows As Collection
Private mVars As Collection
Private mOutputs As Collection
Private mTargets As Collection

Private Sub UserForm_Initialize()
    Set mRows = New Collection
    Set mVars = New Collection
    Set mOutputs = New Collection
    Set mTargets = New Collection

    Dim i As Long
    For i = 65 To 90
        cmbVar.AddItem Chr$(i)
    Next i
    cmbMetric.AddItem "Current"
    cmbMetric.AddItem "Prev"
    cmbMetric.AddItem "Change"

    lstRows.ColumnCount = 6
    lstRows.ColumnWidths = "55 pt;180 pt;70 pt;70 pt;70 pt;0 pt"
    lstVars.ColumnCount = 4
    lstVars.ColumnWidths = "35 pt;65 pt;60 pt;180 pt"
    lstOutputs.ColumnCount = 5
    lstOutputs.ColumnWidths = "35 pt;140 pt;90 pt;120 pt;80 pt"

    For i = 1 To 10
        Dim od As TOutputDef
        od.OutputName = "O" & CStr(i)
        od.FormulaText = ""
        mOutputs.Add od

        Dim tm As TTargetMap
        tm.OutputName = od.OutputName
        mTargets.Add tm
    Next i

    RefreshSheets
    RefreshOutputList

    On Error Resume Next
    LoadProfile CurrentProfileKey, mVars, mOutputs, mTargets
    If mOutputs.Count = 0 Then
        For i = 1 To 10
            od.OutputName = "O" & CStr(i)
            mOutputs.Add od
        Next i
    End If
    On Error GoTo 0

    RefreshVarList
    RefreshOutputList
End Sub

Private Sub btnImport_Click()
    On Error GoTo Fail
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Select HTML disguised XLS report"
    fd.Filters.Clear
    fd.Filters.Add "Excel/HTML", "*.xls;*.html;*.htm"
    If fd.Show <> -1 Then Exit Sub

    Set mRows = ParseHtmlReportFile(fd.SelectedItems(1))
    RefreshRowList txtSearch.Text
    LogMessage "INFO", "Report imported", fd.SelectedItems(1)
    Exit Sub
Fail:
    LogMessage "ERROR", "Import failed", Err.Description
    MsgBox "Import failed: " & Err.Description, vbExclamation
End Sub

Private Sub txtSearch_Change()
    RefreshRowList txtSearch.Text
End Sub

Private Sub RefreshRowList(ByVal filterText As String)
    lstRows.Clear
    Dim i As Long
    For i = 1 To mRows.Count
        Dim rr As TReportRow
        rr = mRows(i)
        If MatchFilter(rr, filterText) Then
            lstRows.AddItem rr.AccountCode
            lstRows.List(lstRows.ListCount - 1, 1) = rr.Label
            lstRows.List(lstRows.ListCount - 1, 2) = FormatNumber(rr.ValCurrent, 2)
            lstRows.List(lstRows.ListCount - 1, 3) = FormatNumber(rr.ValPrev, 2)
            lstRows.List(lstRows.ListCount - 1, 4) = FormatNumber(rr.ValChange, 2)
            lstRows.List(lstRows.ListCount - 1, 5) = CStr(i)
        End If
    Next i
End Sub

Private Function MatchFilter(ByVal rr As TReportRow, ByVal filterText As String) As Boolean
    If Len(Trim$(filterText)) = 0 Then
        MatchFilter = True
    Else
        MatchFilter = (InStr(1, rr.AccountCode, filterText, vbTextCompare) > 0 Or InStr(1, rr.Label, filterText, vbTextCompare) > 0)
    End If
End Function

Private Sub btnBindVar_Click()
    On Error GoTo Fail
    If lstRows.ListIndex < 0 Then Err.Raise vbObjectError + 801, , "Select a source row first"
    If Len(cmbVar.Value) = 0 Then Err.Raise vbObjectError + 802, , "Select variable A..Z"
    If Len(cmbMetric.Value) = 0 Then Err.Raise vbObjectError + 803, , "Select metric"

    Dim sourceIndex As Long
    sourceIndex = CLng(lstRows.List(lstRows.ListIndex, 5))

    Dim rr As TReportRow
    rr = mRows(sourceIndex)

    Dim vb As TVarBinding
    vb.VarName = UCase$(cmbVar.Value)
    vb.AccountCode = rr.AccountCode
    vb.Label = rr.Label
    vb.Metric = NormalizeMetric(cmbMetric.Value)
    vb.Value = MetricValue(rr, vb.Metric)

    UpsertVarBinding vb
    RefreshVarList
    Exit Sub
Fail:
    MsgBox Err.Description, vbExclamation
End Sub

Private Function MetricValue(ByVal rr As TReportRow, ByVal metric As String) As Double
    Select Case UCase$(metric)
        Case "CURRENT": MetricValue = rr.ValCurrent
        Case "PREV": MetricValue = rr.ValPrev
        Case "CHANGE": MetricValue = rr.ValChange
        Case Else: MetricValue = rr.ValCurrent
    End Select
End Function

Private Sub UpsertVarBinding(ByVal entry As TVarBinding)
    Dim i As Long
    For i = 1 To mVars.Count
        Dim v As TVarBinding
        v = mVars(i)
        If UCase$(v.VarName) = UCase$(entry.VarName) Then
            mVars.Remove i
            Exit For
        End If
    Next i
    mVars.Add entry
End Sub

Private Sub RefreshVarList()
    lstVars.Clear
    Dim i As Long
    For i = 1 To mVars.Count
        Dim v As TVarBinding
        v = mVars(i)
        lstVars.AddItem v.VarName
        lstVars.List(lstVars.ListCount - 1, 1) = v.AccountCode
        lstVars.List(lstVars.ListCount - 1, 2) = v.Metric
        lstVars.List(lstVars.ListCount - 1, 3) = FormatNumber(v.Value, 2)
    Next i
End Sub

Private Sub RefreshSheets()
    cmbTargetSheet.Clear
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        cmbTargetSheet.AddItem ws.Name
    Next ws
End Sub

Private Sub RefreshOutputList()
    lstOutputs.Clear
    Dim i As Long
    For i = 1 To mOutputs.Count
        Dim od As TOutputDef
        od = mOutputs(i)
        Dim tm As TTargetMap
        tm = FindTargetByOutput(od.OutputName)

        lstOutputs.AddItem od.OutputName
        lstOutputs.List(lstOutputs.ListCount - 1, 1) = od.FormulaText
        lstOutputs.List(lstOutputs.ListCount - 1, 2) = IIf(od.LastValue = 0, "", CStr(od.LastValue))
        lstOutputs.List(lstOutputs.ListCount - 1, 3) = tm.TargetSheet
        lstOutputs.List(lstOutputs.ListCount - 1, 4) = tm.TargetAddress
    Next i
End Sub

Private Sub lstOutputs_Click()
    If lstOutputs.ListIndex < 0 Then Exit Sub
    txtFormula.Value = lstOutputs.List(lstOutputs.ListIndex, 1)
    cmbTargetSheet.Value = lstOutputs.List(lstOutputs.ListIndex, 3)
    txtTargetCell.Value = lstOutputs.List(lstOutputs.ListIndex, 4)
End Sub

Private Sub btnSaveOutput_Click()
    On Error GoTo Fail
    If lstOutputs.ListIndex < 0 Then Err.Raise vbObjectError + 820, , "Select output row"

    Dim outName As String
    outName = lstOutputs.List(lstOutputs.ListIndex, 0)

    Dim msg As String
    If Len(Trim$(txtFormula.Value)) > 0 Then
        If Not ValidateFormula(txtFormula.Value, msg) Then
            Err.Raise vbObjectError + 821, , "Invalid formula: " & msg
        End If
    End If

    Dim i As Long
    For i = 1 To mOutputs.Count
        Dim od As TOutputDef
        od = mOutputs(i)
        If od.OutputName = outName Then
            od.FormulaText = Trim$(txtFormula.Value)
            mOutputs.Remove i
            mOutputs.Add od, , i
            Exit For
        End If
    Next i

    Dim tm As TTargetMap
    tm.OutputName = outName
    tm.TargetSheet = Trim$(cmbTargetSheet.Value)
    tm.TargetAddress = Trim$(txtTargetCell.Value)
    UpsertTarget tm

    RefreshOutputList
    Exit Sub
Fail:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub UpsertTarget(ByVal tm As TTargetMap)
    Dim i As Long
    For i = 1 To mTargets.Count
        Dim x As TTargetMap
        x = mTargets(i)
        If UCase$(x.OutputName) = UCase$(tm.OutputName) Then
            mTargets.Remove i
            Exit For
        End If
    Next i
    mTargets.Add tm
End Sub

Private Function FindTargetByOutput(ByVal outName As String) As TTargetMap
    Dim i As Long
    For i = 1 To mTargets.Count
        Dim tm As TTargetMap
        tm = mTargets(i)
        If UCase$(tm.OutputName) = UCase$(outName) Then
            FindTargetByOutput = tm
            Exit Function
        End If
    Next i
    Dim fallback As TTargetMap
    fallback.OutputName = outName
    FindTargetByOutput = fallback
End Function

Private Function BuildVarMap() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To mVars.Count
        Dim vb As TVarBinding
        vb = mVars(i)
        vb.Value = ResolveBindingValue(vb)
        d(UCase$(vb.VarName)) = vb.Value
    Next i
    Set BuildVarMap = d
End Function

Private Function ResolveBindingValue(ByVal vb As TVarBinding) As Double
    Dim i As Long
    For i = 1 To mRows.Count
        Dim rr As TReportRow
        rr = mRows(i)
        If rr.AccountCode = vb.AccountCode And rr.Label = vb.Label Then
            ResolveBindingValue = MetricValue(rr, vb.Metric)
            Exit Function
        End If
    Next i
    ResolveBindingValue = vb.Value
End Function

Private Sub btnPreview_Click()
    On Error GoTo Fail
    Dim varMap As Object
    Set varMap = BuildVarMap()

    Dim i As Long
    For i = 1 To mOutputs.Count
        Dim od As TOutputDef
        od = mOutputs(i)
        If Len(Trim$(od.FormulaText)) > 0 Then
            od.LastValue = EvaluateFormula(od.FormulaText, varMap)
        Else
            od.LastValue = 0
        End If
        mOutputs.Remove i
        mOutputs.Add od, , i

        Dim tm As TTargetMap
        tm = FindTargetByOutput(od.OutputName)
        If Len(tm.TargetSheet) = 0 Or Len(tm.TargetAddress) = 0 Then
            LogMessage "WARN", "Missing target", od.OutputName
        ElseIf Not IsValidTarget(tm.TargetSheet, tm.TargetAddress) Then
            LogMessage "WARN", "Invalid target", tm.OutputName & " -> " & tm.TargetSheet & "!" & tm.TargetAddress
        End If
    Next i

    RefreshVarList
    RefreshOutputList
    MsgBox "Preview done. Check output list and _ImportLog.", vbInformation
    Exit Sub
Fail:
    LogMessage "ERROR", "Preview failed", Err.Description
    MsgBox "Preview failed: " & Err.Description, vbExclamation
End Sub

Private Function IsValidTarget(ByVal sheetName As String, ByVal addressText As String) As Boolean
    On Error GoTo Bad
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets(sheetName)
    Dim r As Range
    Set r = ws.Range(addressText)
    IsValidTarget = True
    Exit Function
Bad:
    IsValidTarget = False
End Function

Private Sub btnApply_Click()
    On Error GoTo Fail
    btnPreview_Click
    ApplyOutputValues mOutputs, mTargets
    LogMessage "INFO", "Apply complete", CStr(mOutputs.Count) & " outputs"
    MsgBox "Applied successfully", vbInformation
    Exit Sub
Fail:
    LogMessage "ERROR", "Apply failed", Err.Description
    MsgBox "Apply failed: " & Err.Description, vbExclamation
End Sub

Private Sub btnPickCell_Click()
    On Error GoTo Fail
    If lstOutputs.ListIndex < 0 Then Err.Raise vbObjectError + 850, , "Select output row"
    Me.Hide
    Dim rng As Range
    Set rng = Application.InputBox("Select target cell", "Pick cell", Type:=8)
    Me.Show

    If Not rng Is Nothing Then
        cmbTargetSheet.Value = rng.Worksheet.Name
        txtTargetCell.Value = rng.Address(False, False)
    End If
    Exit Sub
Fail:
    Me.Show
End Sub

Private Sub btnSaveProfile_Click()
    On Error GoTo Fail
    SaveProfile CurrentProfileKey, mVars, mOutputs, mTargets
    LogMessage "INFO", "Profile saved", CurrentProfileKey
    MsgBox "Profile saved.", vbInformation
    Exit Sub
Fail:
    LogMessage "ERROR", "Save profile failed", Err.Description
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub btnLoadProfile_Click()
    On Error GoTo Fail
    LoadProfile CurrentProfileKey, mVars, mOutputs, mTargets
    If mOutputs.Count = 0 Then
        Dim i As Long
        For i = 1 To 10
            Dim od As TOutputDef
            od.OutputName = "O" & CStr(i)
            mOutputs.Add od
        Next i
    End If
    RefreshVarList
    RefreshOutputList
    LogMessage "INFO", "Profile loaded", CurrentProfileKey
    MsgBox "Profile loaded.", vbInformation
    Exit Sub
Fail:
    LogMessage "ERROR", "Load profile failed", Err.Description
    MsgBox Err.Description, vbExclamation
End Sub
