Attribute VB_Name = "modModel"
Option Explicit

Public Function NewReportRow() As CReportRow
    Set NewReportRow = New CReportRow
End Function

Public Function NewVarBinding() As CVarBinding
    Set NewVarBinding = New CVarBinding
End Function

Public Function NewOutputDef() As COutputDef
    Set NewOutputDef = New COutputDef
End Function

Public Function NewTargetMap() As CTargetMap
    Set NewTargetMap = New CTargetMap
End Function

Public Function NormalizeMetric(ByVal metric As String) As String
    Select Case UCase$(Trim$(metric))
        Case "CURRENT": NormalizeMetric = "Current"
        Case "PREV": NormalizeMetric = "Prev"
        Case "CHANGE": NormalizeMetric = "Change"
        Case Else: NormalizeMetric = "Current"
    End Select
End Function
