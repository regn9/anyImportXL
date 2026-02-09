Attribute VB_Name = "modFormula"
Option Explicit

Public Function ValidateFormula(ByVal expr As String, ByRef errMsg As String) As Boolean
    On Error GoTo Bad
    Dim tokens As Collection
    Set tokens = Tokenize(expr)
    Dim i As Long
    For i = 1 To tokens.Count
        Dim t As String
        t = CStr(tokens(i))
        If IsIdentifier(t) Then
            If Len(t) <> 1 Or Asc(t) < 65 Or Asc(t) > 90 Then
                errMsg = "Only variables A..Z are allowed."
                Exit Function
            End If
        End If
    Next i
    Dim dummy As Double
    dummy = EvaluateFormula(expr, CreateObject("Scripting.Dictionary"))
    ValidateFormula = True
    Exit Function
Bad:
    errMsg = Err.Description
    ValidateFormula = False
End Function

Public Function EvaluateFormula(ByVal expr As String, ByVal varMap As Object) As Double
    Dim rpn As Collection
    Set rpn = ToRpn(Tokenize(expr))
    EvaluateFormula = EvalRpn(rpn, varMap)
End Function

Private Function Tokenize(ByVal expr As String) As Collection
    Dim tokens As New Collection
    Dim i As Long
    i = 1
    Do While i <= Len(expr)
        Dim ch As String
        ch = Mid$(expr, i, 1)
        If ch = " " Or ch = vbTab Then
            i = i + 1
        ElseIf InStr(1, "+-*/()", ch, vbBinaryCompare) > 0 Then
            tokens.Add ch
            i = i + 1
        ElseIf ch Like "[0-9.]" Then
            Dim j As Long
            j = i
            Do While j <= Len(expr) And Mid$(expr, j, 1) Like "[0-9.]"
                j = j + 1
            Loop
            tokens.Add Mid$(expr, i, j - i)
            i = j
        ElseIf ch Like "[A-Za-z]" Then
            tokens.Add UCase$(ch)
            i = i + 1
        Else
            Err.Raise vbObjectError + 701, "Tokenize", "Invalid character: " & ch
        End If
    Loop
    Set Tokenize = tokens
End Function

Private Function ToRpn(ByVal tokens As Collection) As Collection
    Dim outQ As New Collection
    Dim ops As New Collection
    Dim i As Long
    For i = 1 To tokens.Count
        Dim t As String
        t = CStr(tokens(i))
        If IsNumericToken(t) Or IsIdentifier(t) Then
            outQ.Add t
        ElseIf IsOperator(t) Then
            Do While ops.Count > 0 And IsOperator(CStr(ops(ops.Count))) And Precedence(CStr(ops(ops.Count))) >= Precedence(t)
                outQ.Add ops(ops.Count)
                ops.Remove ops.Count
            Loop
            ops.Add t
        ElseIf t = "(" Then
            ops.Add t
        ElseIf t = ")" Then
            Do While ops.Count > 0 And CStr(ops(ops.Count)) <> "("
                outQ.Add ops(ops.Count)
                ops.Remove ops.Count
            Loop
            If ops.Count = 0 Then Err.Raise vbObjectError + 702, "ToRpn", "Mismatched parentheses"
            ops.Remove ops.Count
        End If
    Next i

    Do While ops.Count > 0
        If CStr(ops(ops.Count)) = "(" Then Err.Raise vbObjectError + 703, "ToRpn", "Mismatched parentheses"
        outQ.Add ops(ops.Count)
        ops.Remove ops.Count
    Loop

    Set ToRpn = outQ
End Function

Private Function EvalRpn(ByVal rpn As Collection, ByVal varMap As Object) As Double
    Dim st As New Collection
    Dim i As Long
    For i = 1 To rpn.Count
        Dim t As String
        t = CStr(rpn(i))
        If IsNumericToken(t) Then
            st.Add CDbl(t)
        ElseIf IsIdentifier(t) Then
            If varMap.Exists(t) Then
                st.Add CDbl(varMap(t))
            Else
                st.Add 0#
            End If
        ElseIf IsOperator(t) Then
            If st.Count < 2 Then Err.Raise vbObjectError + 704, "EvalRpn", "Invalid expression"
            Dim b As Double, a As Double
            b = CDbl(st(st.Count)): st.Remove st.Count
            a = CDbl(st(st.Count)): st.Remove st.Count
            Select Case t
                Case "+": st.Add a + b
                Case "-": st.Add a - b
                Case "*": st.Add a * b
                Case "/"
                    If b = 0 Then Err.Raise vbObjectError + 705, "EvalRpn", "Division by zero"
                    st.Add a / b
            End Select
        End If
    Next i

    If st.Count <> 1 Then Err.Raise vbObjectError + 706, "EvalRpn", "Invalid expression"
    EvalRpn = CDbl(st(1))
End Function

Private Function IsOperator(ByVal t As String) As Boolean
    IsOperator = (t = "+" Or t = "-" Or t = "*" Or t = "/")
End Function

Private Function IsIdentifier(ByVal t As String) As Boolean
    IsIdentifier = (Len(t) > 0 And t Like "[A-Z]*" And Not IsNumericToken(t))
End Function

Private Function IsNumericToken(ByVal t As String) As Boolean
    IsNumericToken = IsNumeric(t)
End Function

Private Function Precedence(ByVal op As String) As Long
    Select Case op
        Case "+", "-": Precedence = 1
        Case "*", "/": Precedence = 2
        Case Else: Precedence = 0
    End Select
End Function
