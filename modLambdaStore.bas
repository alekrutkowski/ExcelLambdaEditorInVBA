Attribute VB_Name = "modLambdaStore"
Option Explicit

Public Function IsLambdaName(ByVal nm As Name) As Boolean
    Dim s As String

    s = UCase$(Replace(nm.RefersTo, vbCrLf, ""))
    s = Replace(s, vbLf, "")

    IsLambdaName = InStr(1, s, "=LAMBDA(", vbTextCompare) > 0 _
        Or InStr(1, s, "=LAMBDA (", vbTextCompare) > 0
End Function

Public Function CleanNameText(ByVal s As String) As String
    s = Trim$(s)
    If Left$(s, 1) = "=" Then s = Mid$(s, 2)
    CleanNameText = s
End Function

Public Function CleanFormulaText(ByVal s As String) As String
    s = NormalizeFormulaForName(s)

    If Len(s) = 0 Then
        CleanFormulaText = ""
    ElseIf Left$(s, 1) = "=" Then
        CleanFormulaText = s
    Else
        CleanFormulaText = "=" & s
    End If
End Function

Private Function NormalizeFormulaForName(ByVal s As String) As String
    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    s = Replace(s, vbTab, " ")
    s = Replace(s, vbLf, " ")
    s = Trim$(s)

    Do While InStr(1, s, "  ", vbBinaryCompare) > 0
        s = Replace(s, "  ", " ")
    Loop

    NormalizeFormulaForName = s
End Function

Public Function LambdaNameExists(ByVal wb As Workbook, ByVal lambdaName As String) As Boolean
    Dim nm As Name

    On Error Resume Next
    Set nm = wb.Names(CleanNameText(lambdaName))
    LambdaNameExists = Not nm Is Nothing
    On Error GoTo 0
End Function

Public Function GetLambdaName(ByVal wb As Workbook, ByVal lambdaName As String) As Name
    On Error Resume Next
    Set GetLambdaName = wb.Names(CleanNameText(lambdaName))
    On Error GoTo 0
End Function

Public Function GetLambdaNames(ByVal wb As Workbook) As Collection
    Dim out As New Collection
    Dim nm As Name

    For Each nm In wb.Names
        If IsLambdaName(nm) Then out.Add nm.Name
    Next nm

    Set GetLambdaNames = out
End Function

Public Sub SaveLambdaName(ByVal wb As Workbook, ByVal lambdaName As String, ByVal formulaText As String, Optional ByVal commentText As String = "")
    Dim cleanedName As String
    Dim cleanedFormula As String
    Dim nm As Name

    cleanedName = CleanNameText(lambdaName)
    cleanedFormula = CleanFormulaText(formulaText)

    If Len(cleanedName) = 0 Then Err.Raise vbObjectError + 1000, , "Function name is required."
    If Len(cleanedFormula) = 0 Then Err.Raise vbObjectError + 1001, , "Formula is required."
    If InStr(1, cleanedFormula, "=LAMBDA", vbTextCompare) <> 1 Then Err.Raise vbObjectError + 1002, , "Formula must start with =LAMBDA(...)."

    Set nm = GetLambdaName(wb, cleanedName)

    If nm Is Nothing Then
        wb.Names.Add Name:=cleanedName, RefersTo:=cleanedFormula, Visible:=True
        Set nm = wb.Names(cleanedName)
    Else
        nm.RefersTo = cleanedFormula
    End If

    On Error Resume Next
    nm.Comment = Left$(commentText, 255)
    On Error GoTo 0
End Sub

Public Sub DeleteLambdaName(ByVal wb As Workbook, ByVal lambdaName As String)
    Dim nm As Name

    Set nm = GetLambdaName(wb, lambdaName)
    If nm Is Nothing Then Err.Raise vbObjectError + 1003, , "Name not found."

    nm.Delete
End Sub

Public Function TryEvaluateFormula(ByVal formulaText As String, Optional ByVal wb As Workbook) As Variant
    Dim oldRefStyle As XlReferenceStyle
    Dim expr As String

    expr = CleanFormulaText(formulaText)
    oldRefStyle = Application.ReferenceStyle

    On Error GoTo Fail

    If Not wb Is Nothing Then wb.Activate
    Application.ReferenceStyle = xlA1
    TryEvaluateFormula = Application.Evaluate(expr)
    Application.ReferenceStyle = oldRefStyle
    Exit Function

Fail:
    Application.ReferenceStyle = oldRefStyle
    TryEvaluateFormula = CVErr(xlErrValue)
End Function

Public Function ValueToText(ByVal v As Variant) As String
    On Error GoTo Fail

    If IsError(v) Then
        ValueToText = ErrorText(v)
    ElseIf IsArray(v) Then
        ValueToText = ArrayToText(v)
    Else
        ValueToText = CStr(v)
    End If

    Exit Function

Fail:
    ValueToText = "<unable to display result>"
End Function

Private Function ErrorText(ByVal v As Variant) As String
    Select Case CLng(v)
        Case xlErrDiv0
            ErrorText = "#DIV/0!"
        Case xlErrNA
            ErrorText = "#N/A"
        Case xlErrName
            ErrorText = "#NAME?"
        Case xlErrNull
            ErrorText = "#NULL!"
        Case xlErrNum
            ErrorText = "#NUM!"
        Case xlErrRef
            ErrorText = "#REF!"
        Case xlErrValue
            ErrorText = "#VALUE!"
        Case Else
            ErrorText = "#ERROR"
    End Select
End Function

Private Function ArrayToText(ByVal v As Variant) As String
    Dim r As Long
    Dim c As Long
    Dim r1 As Long
    Dim r2 As Long
    Dim c1 As Long
    Dim c2 As Long
    Dim lines() As String
    Dim cells() As String

    r1 = LBound(v, 1)
    r2 = UBound(v, 1)
    c1 = LBound(v, 2)
    c2 = UBound(v, 2)

    ReDim lines(r1 To r2)

    For r = r1 To r2
        ReDim cells(c1 To c2)

        For c = c1 To c2
            cells(c) = ValueToText(v(r, c))
        Next c

        lines(r) = Join(cells, vbTab)
    Next r

    ArrayToText = Join(lines, vbCrLf)
End Function

Public Function FormatLambdaFormula(ByVal s As String) As String
    Dim t As String

    t = CleanFormulaText(s)
    t = Replace(t, ",", "," & vbCrLf & "    ")
    t = Replace(t, "LET(", "LET(" & vbCrLf & "    ", , , vbTextCompare)
    t = Replace(t, "LAMBDA(", "LAMBDA(" & vbCrLf & "    ", , , vbTextCompare)

    FormatLambdaFormula = t
End Function
