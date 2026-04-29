Attribute VB_Name = "modLambdaStore"
Option Explicit

Public Function IsLambdaName(ByVal nm As Name) As Boolean
    Dim s As String

    s = FormulaHeadForDetection(nm.RefersTo)

    IsLambdaName = Left$(s, 8) = "=LAMBDA("
End Function

Private Function FormulaHeadForDetection(ByVal formulaText As String) As String
    Dim s As String
    Dim i As Long
    Dim ch As String

    s = formulaText
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    s = Trim$(s)

    If Left$(s, 1) = "=" Then
        i = 2

        Do While i <= Len(s)
            ch = Mid$(s, i, 1)

            If ch <> " " Then Exit Do
            i = i + 1
        Loop

        s = "=" & Mid$(s, i)
    End If

    FormulaHeadForDetection = UCase$(s)
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

    If wb Is Nothing Then Set wb = ActiveWorkbook
    If Not wb Is Nothing Then wb.Activate

    Application.ReferenceStyle = xlA1
    TryEvaluateFormula = EvaluateWithFormula2Spill(expr, wb)
    Application.ReferenceStyle = oldRefStyle
    Exit Function

Fail:
    Application.ReferenceStyle = oldRefStyle
    TryEvaluateFormula = CVErr(xlErrValue)
End Function

Private Function EvaluateWithFormula2Spill(ByVal expr As String, ByVal wb As Workbook) As Variant
    Dim ws As Worksheet
    Dim cell As Range
    Dim spillRange As Range

    On Error GoTo Fallback

    Set ws = GetScratchSheet(wb)
    Set cell = ws.Range("A1")

    ws.Cells.Clear
    cell.Formula2 = expr
    cell.Calculate

    On Error Resume Next
    Set spillRange = cell.SpillingToRange
    On Error GoTo Fallback

    If Not spillRange Is Nothing Then
        EvaluateWithFormula2Spill = spillRange.Value
    Else
        EvaluateWithFormula2Spill = cell.Value
    End If

    ws.Cells.Clear
    Exit Function

Fallback:
    On Error Resume Next
    If Not ws Is Nothing Then ws.Cells.Clear
    EvaluateWithFormula2Spill = Application.Evaluate(expr)
    On Error GoTo 0
End Function

Private Function GetScratchSheet(ByVal wb As Workbook) As Worksheet
    Const scratchName As String = "__LambdaEditorScratch"
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Worksheets(scratchName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = scratchName
    End If

    On Error Resume Next
    ws.Visible = xlSheetVeryHidden
    On Error GoTo 0

    Set GetScratchSheet = ws
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
    Dim dims As Long

    dims = ArrayDimensions(v)

    If dims = 1 Then
        ArrayToText = Array1DToText(v)
    ElseIf dims = 2 Then
        ArrayToText = Array2DToText(v)
    Else
        ArrayToText = "<array with " & CStr(dims) & " dimensions>"
    End If
End Function

Private Function ArrayDimensions(ByVal v As Variant) As Long
    Dim n As Long
    Dim tmp As Long

    On Error GoTo Done

    For n = 1 To 60
        tmp = LBound(v, n)
    Next n

Done:
    ArrayDimensions = n - 1
End Function

Private Function Array1DToText(ByVal v As Variant) As String
    Dim i As Long
    Dim i1 As Long
    Dim i2 As Long
    Dim cells() As String

    i1 = LBound(v, 1)
    i2 = UBound(v, 1)

    ReDim cells(i1 To i2)

    For i = i1 To i2
        cells(i) = ValueToText(v(i))
    Next i

    Array1DToText = Join(cells, vbTab)
End Function

Private Function Array2DToText(ByVal v As Variant) As String
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

    Array2DToText = Join(lines, vbCrLf)
End Function

Public Function FormatLambdaFormula(ByVal s As String) As String
    Dim t As String

    t = CleanFormulaText(s)
    t = Replace(t, ",", "," & vbCrLf & "    ")
    t = Replace(t, "LET(", "LET(" & vbCrLf & "    ", , , vbTextCompare)
    t = Replace(t, "LAMBDA(", "LAMBDA(" & vbCrLf & "    ", , , vbTextCompare)

    FormatLambdaFormula = t
End Function


Public Function MinifyLambdaDefinition(ByVal formulaText As String, Optional ByVal shortenNames As Boolean = True) As String
    Dim s As String
    Dim binders As Object
    Dim allIds As Object
    Dim mapping As Object

    s = CleanFormulaText(formulaText)
    s = StripWhitespaceOutsideStrings(s)

    If shortenNames Then
        Set binders = CreateObject("Scripting.Dictionary")
        Set allIds = CreateObject("Scripting.Dictionary")
        CollectIdentifierTokens s, allIds
        CollectLambdaAndLetBinders s, binders
        Set mapping = BuildShortNameMap(binders, allIds)
        s = RenameFormulaIdentifiers(s, mapping)
    End If

    MinifyLambdaDefinition = s
End Function

Private Function StripWhitespaceOutsideStrings(ByVal s As String) As String
    Dim i As Long
    Dim ch As String
    Dim out As String
    Dim inString As Boolean

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)

        If ch = Chr$(34) Then
            out = out & ch

            If inString And i < Len(s) And Mid$(s, i + 1, 1) = Chr$(34) Then
                i = i + 1
                out = out & Chr$(34)
            Else
                inString = Not inString
            End If
        ElseIf inString Then
            out = out & ch
        ElseIf Not IsFormulaWhitespace(ch) Then
            out = out & ch
        End If
    Next i

    StripWhitespaceOutsideStrings = out
End Function

Private Function IsFormulaWhitespace(ByVal ch As String) As Boolean
    IsFormulaWhitespace = ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf
End Function

Private Sub CollectIdentifierTokens(ByVal s As String, ByVal ids As Object)
    Dim i As Long
    Dim token As String
    Dim ch As String
    Dim inString As Boolean

    i = 1
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)

        If ch = Chr$(34) Then
            inString = Not inString
            i = i + 1
        ElseIf Not inString And IsNameStartChar(ch) Then
            token = ReadIdentifierToken(s, i)
            If Not ids.Exists(UCase$(token)) Then ids.Add UCase$(token), token
            i = i + Len(token)
        Else
            i = i + 1
        End If
    Loop
End Sub

Private Sub CollectLambdaAndLetBinders(ByVal s As String, ByVal binders As Object)
    Dim i As Long
    Dim nameText As String
    Dim openPos As Long
    Dim closePos As Long
    Dim args As Collection
    Dim j As Long
    Dim argText As String
    Dim inString As Boolean
    Dim ch As String

    i = 1
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)

        If ch = Chr$(34) Then
            inString = Not inString
            i = i + 1
        ElseIf Not inString And IsNameStartChar(ch) Then
            nameText = ReadIdentifierToken(s, i)
            openPos = i + Len(nameText)

            If openPos <= Len(s) And Mid$(s, openPos, 1) = "(" Then
                If StrComp(nameText, "LAMBDA", vbTextCompare) = 0 Or StrComp(nameText, "LET", vbTextCompare) = 0 Then
                    closePos = FindMatchingParen(s, openPos)

                    If closePos > openPos Then
                        Set args = SplitTopLevelArgs(Mid$(s, openPos + 1, closePos - openPos - 1))

                        If StrComp(nameText, "LAMBDA", vbTextCompare) = 0 Then
                            For j = 1 To args.Count - 1
                                argText = CStr(args(j))
                                If IsValidLambdaName(argText) Then AddBinder binders, argText
                            Next j
                        ElseIf StrComp(nameText, "LET", vbTextCompare) = 0 Then
                            For j = 1 To args.Count - 1 Step 2
                                argText = CStr(args(j))
                                If IsValidLambdaName(argText) Then AddBinder binders, argText
                            Next j
                        End If

                        CollectLambdaAndLetBinders Mid$(s, openPos + 1, closePos - openPos - 1), binders
                        i = closePos + 1
                    Else
                        i = openPos + 1
                    End If
                Else
                    i = openPos + 1
                End If
            Else
                i = i + Len(nameText)
            End If
        Else
            i = i + 1
        End If
    Loop
End Sub

Private Sub AddBinder(ByVal binders As Object, ByVal nameText As String)
    Dim key As String

    key = UCase$(nameText)
    If Not binders.Exists(key) Then binders.Add key, nameText
End Sub

Private Function BuildShortNameMap(ByVal binders As Object, ByVal allIds As Object) As Object
    Dim map As Object
    Dim used As Object
    Dim k As Variant
    Dim oldName As String
    Dim newName As String
    Dim n As Long

    Set map = CreateObject("Scripting.Dictionary")
    Set used = CreateObject("Scripting.Dictionary")

    For Each k In allIds.Keys
        used(k) = True
    Next k

    n = 1
    For Each k In binders.Keys
        oldName = CStr(binders(k))
        newName = NextSafeShortName(n, used)

        If Len(newName) < Len(oldName) Then
            map(UCase$(oldName)) = newName
            used(UCase$(newName)) = True
        End If
    Next k

    Set BuildShortNameMap = map
End Function

Private Function NextSafeShortName(ByRef n As Long, ByVal used As Object) As String
    Dim candidate As String

    Do
        candidate = "_" & Base26Name(n)
        n = n + 1
    Loop While used.Exists(UCase$(candidate)) Or LooksLikeExcelReference(candidate)

    NextSafeShortName = candidate
End Function

Private Function Base26Name(ByVal n As Long) As String
    Dim x As Long
    Dim remVal As Long
    Dim out As String

    x = n
    Do
        remVal = (x - 1) Mod 26
        out = Chr$(97 + remVal) & out
        x = (x - 1) \ 26
    Loop While x > 0

    Base26Name = out
End Function

Private Function RenameFormulaIdentifiers(ByVal s As String, ByVal mapping As Object) As String
    Dim i As Long
    Dim ch As String
    Dim token As String
    Dim out As String
    Dim inString As Boolean
    Dim key As String

    i = 1
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)

        If ch = Chr$(34) Then
            out = out & ch

            If inString And i < Len(s) And Mid$(s, i + 1, 1) = Chr$(34) Then
                i = i + 1
                out = out & Chr$(34)
            Else
                inString = Not inString
            End If

            i = i + 1
        ElseIf Not inString And IsNameStartChar(ch) Then
            token = ReadIdentifierToken(s, i)
            key = UCase$(token)

            If mapping.Exists(key) Then
                out = out & CStr(mapping(key))
            Else
                out = out & token
            End If

            i = i + Len(token)
        Else
            out = out & ch
            i = i + 1
        End If
    Loop

    RenameFormulaIdentifiers = out
End Function

Private Function SplitTopLevelArgs(ByVal s As String) As Collection
    Dim args As New Collection
    Dim i As Long
    Dim startPos As Long
    Dim depth As Long
    Dim ch As String
    Dim inString As Boolean

    startPos = 1
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)

        If ch = Chr$(34) Then
            If inString And i < Len(s) And Mid$(s, i + 1, 1) = Chr$(34) Then
                i = i + 1
            Else
                inString = Not inString
            End If
        ElseIf Not inString Then
            If ch = "(" Then
                depth = depth + 1
            ElseIf ch = ")" Then
                If depth > 0 Then depth = depth - 1
            ElseIf ch = "," And depth = 0 Then
                args.Add Mid$(s, startPos, i - startPos)
                startPos = i + 1
            End If
        End If
    Next i

    args.Add Mid$(s, startPos)
    Set SplitTopLevelArgs = args
End Function

Private Function FindMatchingParen(ByVal s As String, ByVal openPos As Long) As Long
    Dim i As Long
    Dim depth As Long
    Dim ch As String
    Dim inString As Boolean

    For i = openPos To Len(s)
        ch = Mid$(s, i, 1)

        If ch = Chr$(34) Then
            If inString And i < Len(s) And Mid$(s, i + 1, 1) = Chr$(34) Then
                i = i + 1
            Else
                inString = Not inString
            End If
        ElseIf Not inString Then
            If ch = "(" Then
                depth = depth + 1
            ElseIf ch = ")" Then
                depth = depth - 1
                If depth = 0 Then
                    FindMatchingParen = i
                    Exit Function
                End If
            End If
        End If
    Next i

    FindMatchingParen = 0
End Function

Private Function ReadIdentifierToken(ByVal s As String, ByVal startPos As Long) As String
    Dim i As Long
    Dim ch As String

    i = startPos
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        If IsNameBodyChar(ch) Then
            i = i + 1
        Else
            Exit Do
        End If
    Loop

    ReadIdentifierToken = Mid$(s, startPos, i - startPos)
End Function

Private Function IsNameStartChar(ByVal ch As String) As Boolean
    If Len(ch) <> 1 Then Exit Function
    IsNameStartChar = (ch Like "[A-Za-z_\\]")
End Function

Private Function IsNameBodyChar(ByVal ch As String) As Boolean
    If Len(ch) <> 1 Then Exit Function
    IsNameBodyChar = (ch Like "[A-Za-z0-9_.\\]")
End Function

Private Function IsValidLambdaName(ByVal s As String) As Boolean
    s = Trim$(s)

    If Len(s) = 0 Then Exit Function
    If Not IsNameStartChar(Left$(s, 1)) Then Exit Function
    If LooksLikeExcelReference(s) Then Exit Function
    If InStr(1, s, ".", vbBinaryCompare) > 0 Then Exit Function
    If InStr(1, s, "!", vbBinaryCompare) > 0 Then Exit Function
    If InStr(1, s, "[", vbBinaryCompare) > 0 Then Exit Function
    If InStr(1, s, "]", vbBinaryCompare) > 0 Then Exit Function

    IsValidLambdaName = True
End Function

Private Function LooksLikeExcelReference(ByVal s As String) As Boolean
    Dim t As String
    Dim i As Long
    Dim letters As String
    Dim digits As String

    t = UCase$(Replace(s, "$", ""))
    If Len(t) = 0 Then Exit Function

    If t Like "R[1-9]*C[1-9]*" Then
        LooksLikeExcelReference = True
        Exit Function
    End If

    For i = 1 To Len(t)
        If Mid$(t, i, 1) Like "[A-Z]" Then
            letters = letters & Mid$(t, i, 1)
        Else
            Exit For
        End If
    Next i

    If Len(letters) > 0 And Len(letters) <= 3 Then
        digits = Mid$(t, Len(letters) + 1)
        If Len(digits) > 0 And digits Like "[0-9]*" Then
            LooksLikeExcelReference = True
        End If
    End If
End Function

Public Sub OpenFormulaBoostVisualization(ByVal formulaText As String)
    Dim url As String
    Dim f As String

    f = CleanFormulaText(formulaText)
    url = "https://www.formulaboost.com/parse?f=" & UrlEncodeFormulaBoost(f)

    On Error GoTo Fail
    ActiveWorkbook.FollowHyperlink Address:=url, NewWindow:=True
    Exit Sub

Fail:
    MsgBox "Could not open visualization link." & vbCrLf & vbCrLf & url & vbCrLf & vbCrLf & Err.Description, vbExclamation, "Visualize Formula"
End Sub

Public Function FormulaBoostVisualizationUrl(ByVal formulaText As String) As String
    FormulaBoostVisualizationUrl = "https://www.formulaboost.com/parse?f=" & UrlEncodeFormulaBoost(CleanFormulaText(formulaText))
End Function

Private Function UrlEncodeFormulaBoost(ByVal s As String) As String
    Dim i As Long
    Dim ch As String
    Dim code As Long
    Dim out As String

    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    s = Replace(s, vbLf, "")
    s = Replace(s, vbTab, "")

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)

        If IsFormulaBoostSafeChar(ch) Then
            out = out & ch
        ElseIf code >= 0 And code <= 255 Then
            out = out & "%" & Right$("0" & Hex$(code), 2)
        Else
            out = out & "%3F"
        End If
    Next i

    UrlEncodeFormulaBoost = out
End Function

Private Function IsFormulaBoostSafeChar(ByVal ch As String) As Boolean
    Dim code As Long

    code = AscW(ch)

    Select Case code
        Case 48 To 57, 65 To 90, 97 To 122
            IsFormulaBoostSafeChar = True
        Case 40, 41, 44, 45, 46, 61, 95, 126
            ' Preserve parentheses, commas and equals so URLs look like:
            ' ?f==lambda(x,y,let(z,2,x%2By-z))
            IsFormulaBoostSafeChar = True
        Case Else
            IsFormulaBoostSafeChar = False
    End Select
End Function

