Attribute VB_Name = "modEntry"
Option Explicit

Public Sub OpenAnyImportXlUI()
    On Error GoTo Fail
    frmMain.Show
    Exit Sub
Fail:
    MsgBox "Failed to open add-in UI: " & Err.Description, vbCritical
End Sub
