Attribute VB_Name = "modParserHtml"
Option Explicit

Private Const TABLE_CLASS As String = "contenttable"

Public Function ParseHtmlReportFile(ByVal filePath As String) As Collection
    On Error GoTo CleanFail

    Dim htmlText As String
    htmlText = ReadTextFileUtf8(filePath)

    Dim doc As Object
    Set doc = CreateObject("htmlfile")
    doc.Open
    doc.Write htmlText
    doc.Close

    Dim tbl As Object
    Set tbl = FindPrimaryTable(doc)
    If tbl Is Nothing Then Err.Raise vbObjectError + 601, "ParseHtmlReportFile", "No HTML table found"

    Dim bodyRows As Object
    Set bodyRows = tbl.getElementsByTagName("tr")

    Dim outRows As Collection
    Set outRows = New Collection

    Dim r As Long
    For r = 0 To bodyRows.Length - 1
        Dim tr As Object
        Set tr = bodyRows.Item(r)

        Dim tds As Object
        Set tds = tr.getElementsByTagName("td")
        If tds.Length >= 4 Then
            Dim c1 As String, c2 As String, c3 As String, c4 As String
            c1 = CleanText(tds.Item(0).innerText)
            c2 = CleanText(tds.Item(1).innerText)
            c3 = CleanText(tds.Item(2).innerText)
            c4 = CleanText(tds.Item(3).innerText)

            If IsLeafRow(c2, c3, c4) Then
                Dim rec As TReportRow
                rec.AccountCode = ExtractLeadingCode(c1)
                rec.Label = ExtractLabel(c1)
                rec.ValCurrent = ParseNumber(c2)
                rec.ValPrev = ParseNumber(c3)
                rec.ValChange = ParseNumber(c4)
                outRows.Add rec
            End If
        End If
    Next r

    Set ParseHtmlReportFile = outRows
    Exit Function

CleanFail:
    Err.Raise Err.Number, "ParseHtmlReportFile", Err.Description
End Function

Private Function FindPrimaryTable(ByVal doc As Object) As Object
    Dim tables As Object
    Set tables = doc.getElementsByTagName("table")

    Dim i As Long
    For i = 0 To tables.Length - 1
        Dim cls As String
        On Error Resume Next
        cls = LCase$(CStr(tables.Item(i).className))
        On Error GoTo 0
        If InStr(1, cls, TABLE_CLASS, vbTextCompare) > 0 Then
            Set FindPrimaryTable = tables.Item(i)
            Exit Function
        End If
    Next i

    If tables.Length > 0 Then Set FindPrimaryTable = tables.Item(0)
End Function

Private Function ReadTextFileUtf8(ByVal filePath As String) As String
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile filePath
    ReadTextFileUtf8 = stm.ReadText
    stm.Close
End Function

Private Function IsLeafRow(ByVal currentTxt As String, ByVal prevTxt As String, ByVal changeTxt As String) As Boolean
    If Len(Trim$(currentTxt)) = 0 And Len(Trim$(prevTxt)) = 0 And Len(Trim$(changeTxt)) = 0 Then
        IsLeafRow = False
    Else
        IsLeafRow = True
    End If
End Function

Private Function CleanText(ByVal txt As String) As String
    Dim s As String
    s = Replace(txt, ChrW$(160), " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = CollapseSpaces(s)
    CleanText = Trim$(s)
End Function

Private Function CollapseSpaces(ByVal s As String) As String
    Do While InStr(1, s, "  ", vbBinaryCompare) > 0
        s = Replace(s, "  ", " ")
    Loop
    CollapseSpaces = s
End Function

Private Function ExtractLeadingCode(ByVal s As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^\s*(\d{4})\b"
    re.IgnoreCase = True
    If re.Test(s) Then
        ExtractLeadingCode = re.Execute(s)(0).SubMatches(0)
    Else
        ExtractLeadingCode = ""
    End If
End Function

Private Function ExtractLabel(ByVal s As String) As String
    Dim code As String
    code = ExtractLeadingCode(s)
    If Len(code) > 0 Then
        ExtractLabel = Trim$(Mid$(s, InStr(1, s, code, vbTextCompare) + Len(code)))
    Else
        ExtractLabel = Trim$(s)
    End If
End Function

Public Function ParseNumber(ByVal s As String) As Double
    Dim t As String
    t = Trim$(s)
    If Len(t) = 0 Then
        ParseNumber = 0
        Exit Function
    End If

    t = Replace(t, " ", "")
    t = Replace(t, ".", "")
    t = Replace(t, ",", ".")
    t = Replace(t, "(", "-")
    t = Replace(t, ")", "")

    If IsNumeric(t) Then
        ParseNumber = CDbl(t)
    Else
        ParseNumber = 0
    End If
End Function
