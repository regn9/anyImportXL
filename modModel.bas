Attribute VB_Name = "modModel"
Option Explicit

Public Type TReportRow
    AccountCode As String
    Label As String
    ValCurrent As Double
    ValPrev As Double
    ValChange As Double
End Type

Public Type TVarBinding
    VarName As String
    AccountCode As String
    Label As String
    Metric As String
    Value As Double
End Type

Public Type TOutputDef
    OutputName As String
    FormulaText As String
    LastValue As Double
End Type

Public Type TTargetMap
    OutputName As String
    TargetSheet As String
    TargetAddress As String
End Type

Public Function NewReportRows() As Collection
    Set NewReportRows = New Collection
End Function

Public Function NewVarBindings() As Collection
    Set NewVarBindings = New Collection
End Function

Public Function NewOutputDefs() As Collection
    Set NewOutputDefs = New Collection
End Function

Public Function NewTargetMaps() As Collection
    Set NewTargetMaps = New Collection
End Function

Public Function NormalizeMetric(ByVal metric As String) As String
    Select Case UCase$(Trim$(metric))
        Case "CURRENT": NormalizeMetric = "Current"
        Case "PREV": NormalizeMetric = "Prev"
        Case "CHANGE": NormalizeMetric = "Change"
        Case Else: NormalizeMetric = "Current"
    End Select
End Function
