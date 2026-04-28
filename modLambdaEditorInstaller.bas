Attribute VB_Name = "modLambdaEditorInstaller"
Option Explicit

Private mInstallStep As String

Public Sub InstallLambdaEditor()
    Dim vbProj As Object
    Dim formClassName As String

    On Error GoTo InstallError

    mInstallStep = "Accessing this workbook's VBA project"
    Set vbProj = ThisWorkbook.VBProject

    If MsgBox("Install or replace the LAMBDA editor components in this workbook?", vbQuestion + vbYesNo, "Install LAMBDA Editor") <> vbYes Then Exit Sub

    Application.ScreenUpdating = False

    mInstallStep = "Removing old modLambdaEditor"
    RemoveComponentIfExists vbProj, "modLambdaEditor"

    mInstallStep = "Removing old modLambdaStore"
    RemoveComponentIfExists vbProj, "modLambdaStore"

    mInstallStep = "Removing old frmLambdaEditor if it exists"
    RemoveComponentIfExists vbProj, "frmLambdaEditor"

    mInstallStep = "Adding modLambdaStore"
    AddStandardModule vbProj, "modLambdaStore", Code_modLambdaStore()

    mInstallStep = "Building LAMBDA editor UserForm"
    formClassName = BuildLambdaEditorForm(vbProj)

    mInstallStep = "Adding modLambdaEditor for form " & formClassName
    AddStandardModule vbProj, "modLambdaEditor", Code_modLambdaEditorForForm(formClassName)

    Application.ScreenUpdating = True

    MsgBox "Installed LAMBDA Function Editor." & vbCrLf & vbCrLf & _
           "Created form class: " & formClassName & vbCrLf & _
           "Run ShowLambdaEditor to open it.", vbInformation, "Install LAMBDA Editor"
    Exit Sub

InstallError:
    Application.ScreenUpdating = True
    MsgBox "Install failed at step:" & vbCrLf & mInstallStep & vbCrLf & vbCrLf & _
           "Error " & CStr(Err.Number) & ": " & Err.Description & vbCrLf & vbCrLf & _
           "A partly created UserForm may remain. You can delete it and run InstallLambdaEditor again.", _
           vbExclamation, "Install LAMBDA Editor"
End Sub

Private Sub AddStandardModule(ByVal vbProj As Object, ByVal moduleName As String, ByVal codeText As String)
    Dim comp As Object

    Set comp = vbProj.VBComponents.Add(1)
    comp.Name = moduleName
    comp.CodeModule.AddFromString StripAttributeLines(codeText)
End Sub

Private Function BuildLambdaEditorForm(ByVal vbProj As Object) As String
    Dim comp As Object
    Dim frm As Object

    Dim txtFormulaCtl As Object
    Dim txtResultCtl As Object

    mInstallStep = "Creating UserForm component"
    Set comp = vbProj.VBComponents.Add(3)

    ' Do not rename the component here. Some VBE environments raise error 75 on comp.Name = ...
    ' We use the VBE-assigned name, usually UserForm1 or UserForm2, and generate ShowLambdaEditor for it.
    BuildLambdaEditorForm = comp.Name

    mInstallStep = "Accessing UserForm designer"
    Set frm = comp.Designer

    mInstallStep = "Setting UserForm size and caption"
    SafeSet frm, "Caption", "LAMBDA Function Editor"
    SafeSet frm, "Width", 900
    SafeSet frm, "Height", 600

    mInstallStep = "Adding labels and function list"
    AddLabel frm, "lblFunctions", "Functions", 12, 8, 120, 18
    AddListBox frm, "lstNames", 12, 28, 190, 450

    mInstallStep = "Adding name field and top buttons"
    AddLabel frm, "lblName", "Name", 220, 8, 80, 18
    AddTextBox frm, "txtName", 220, 28, 250, 22, False, False

    AddButton frm, "cmdNew", "New", 488, 26, 58, 24
    AddButton frm, "cmdSave", "Save", 552, 26, 58, 24
    AddButton frm, "cmdDelete", "Delete", 616, 26, 62, 24
    AddButton frm, "cmdRefresh", "Refresh", 684, 26, 70, 24
    AddButton frm, "cmdClose", "Close", 760, 26, 62, 24

    mInstallStep = "Adding comment field"
    AddLabel frm, "lblComment", "Comment", 220, 58, 100, 18
    AddTextBox frm, "txtComment", 220, 78, 602, 46, True, True

    mInstallStep = "Adding formula editor"
    AddLabel frm, "lblFormula", "Formula", 220, 132, 100, 18
    Set txtFormulaCtl = AddTextBox(frm, "txtFormula", 220, 152, 602, 245, True, False)

    mInstallStep = "Adding test controls"
    AddLabel frm, "lblTest", "Test formula", 220, 406, 100, 18
    AddTextBox frm, "txtTestFormula", 220, 426, 395, 22, False, False
    AddButton frm, "cmdValidate", "Validate", 624, 424, 80, 24
    AddButton frm, "cmdTest", "Test", 710, 424, 64, 24

    mInstallStep = "Adding result controls"
    AddLabel frm, "lblResult", "Result", 220, 458, 100, 18
    Set txtResultCtl = AddTextBox(frm, "txtResult", 220, 478, 602, 48, True, True)

    AddLabel frm, "lblStatus", "", 12, 520, 810, 18

    mInstallStep = "Configuring formula editor"
    If Not txtFormulaCtl Is Nothing Then
        SafeSet txtFormulaCtl.Font, "Name", "Consolas"
        SafeSet txtFormulaCtl.Font, "Size", 10
        SafeSet txtFormulaCtl, "EnterKeyBehavior", True
        SafeSet txtFormulaCtl, "ScrollBars", 3
        SafeSet txtFormulaCtl, "WordWrap", False
    End If

    mInstallStep = "Configuring result field"
    If Not txtResultCtl Is Nothing Then
        SafeSet txtResultCtl, "Locked", True
        SafeSet txtResultCtl, "ScrollBars", 2
    End If

    mInstallStep = "Adding UserForm code"
    comp.CodeModule.AddFromString Code_frmLambdaEditor()
End Function

Private Function AddControl(ByVal frm As Object, ByVal progId As String, ByVal controlName As String) As Object
    mInstallStep = "Adding control " & controlName & " using " & progId
    Set AddControl = frm.Controls.Add(progId, controlName, True)
End Function

Private Sub PlaceControl(ByVal c As Object, ByVal leftPos As Single, ByVal topPos As Single, ByVal widthVal As Single, ByVal heightVal As Single)
    SafeSet c, "Left", leftPos
    SafeSet c, "Top", topPos
    SafeSet c, "Width", widthVal
    SafeSet c, "Height", heightVal
End Sub

Private Sub AddLabel(ByVal frm As Object, ByVal controlName As String, ByVal captionText As String, ByVal leftPos As Single, ByVal topPos As Single, ByVal widthVal As Single, ByVal heightVal As Single)
    Dim c As Object

    Set c = AddControl(frm, "Forms.Label.1", controlName)
    SafeSet c, "Caption", captionText
    PlaceControl c, leftPos, topPos, widthVal, heightVal
End Sub

Private Function AddTextBox(ByVal frm As Object, ByVal controlName As String, ByVal leftPos As Single, ByVal topPos As Single, ByVal widthVal As Single, ByVal heightVal As Single, ByVal multilineVal As Boolean, ByVal wrapVal As Boolean) As Object
    Dim c As Object

    Set c = AddControl(frm, "Forms.TextBox.1", controlName)
    PlaceControl c, leftPos, topPos, widthVal, heightVal
    SafeSet c, "MultiLine", multilineVal
    SafeSet c, "WordWrap", wrapVal

    If multilineVal Then
        SafeSet c, "EnterKeyBehavior", True
        SafeSet c, "ScrollBars", 2
    End If

    Set AddTextBox = c
End Function

Private Sub AddListBox(ByVal frm As Object, ByVal controlName As String, ByVal leftPos As Single, ByVal topPos As Single, ByVal widthVal As Single, ByVal heightVal As Single)
    Dim c As Object

    Set c = AddControl(frm, "Forms.ListBox.1", controlName)
    PlaceControl c, leftPos, topPos, widthVal, heightVal
End Sub

Private Sub AddButton(ByVal frm As Object, ByVal controlName As String, ByVal captionText As String, ByVal leftPos As Single, ByVal topPos As Single, ByVal widthVal As Single, ByVal heightVal As Single)
    Dim c As Object

    Set c = AddControl(frm, "Forms.CommandButton.1", controlName)
    SafeSet c, "Caption", captionText
    PlaceControl c, leftPos, topPos, widthVal, heightVal
End Sub

Private Sub SafeSet(ByVal obj As Object, ByVal propertyName As String, ByVal value As Variant)
    On Error Resume Next
    CallByName obj, propertyName, VbLet, value
    On Error GoTo 0
End Sub

Private Sub RemoveComponentIfExists(ByVal vbProj As Object, ByVal componentName As String)
    Dim comp As Object

    On Error Resume Next
    Set comp = vbProj.VBComponents(componentName)
    On Error GoTo 0

    If Not comp Is Nothing Then vbProj.VBComponents.Remove comp
End Sub

Private Function StripAttributeLines(ByVal codeText As String) As String
    Dim lines As Variant
    Dim kept() As String
    Dim i As Long
    Dim n As Long

    lines = Split(codeText, vbCrLf)
    ReDim kept(0 To UBound(lines))

    For i = LBound(lines) To UBound(lines)
        If Left$(Trim$(CStr(lines(i))), 10) <> "Attribute " Then
            kept(n) = CStr(lines(i))
            n = n + 1
        End If
    Next i

    If n = 0 Then
        StripAttributeLines = vbNullString
    Else
        ReDim Preserve kept(0 To n - 1)
        StripAttributeLines = Join(kept, vbCrLf)
    End If
End Function


Private Function Code_modLambdaEditorForForm(ByVal formClassName As String) As String
    Dim s As String

    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub ShowLambdaEditor()" & vbCrLf
    s = s & "    Dim f As Object" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo Fail" & vbCrLf
    s = s & "    Set f = VBA.UserForms.Add(" & Chr$(34) & formClassName & Chr$(34) & ")" & vbCrLf
    s = s & "    f.Show vbModeless" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Fail:" & vbCrLf
    s = s & "    MsgBox " & Chr$(34) & "Could not open the LAMBDA editor form: " & Chr$(34) & " & Err.Description, vbExclamation, " & Chr$(34) & "LAMBDA Editor" & Chr$(34) & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub InstallLambdaEditorShortcut()" & vbCrLf
    s = s & "    Application.OnKey " & Chr$(34) & "^+l" & Chr$(34) & ", " & Chr$(34) & "ShowLambdaEditor" & Chr$(34) & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub RemoveLambdaEditorShortcut()" & vbCrLf
    s = s & "    Application.OnKey " & Chr$(34) & "^+l" & Chr$(34) & vbCrLf
    s = s & "End Sub" & vbCrLf

    Code_modLambdaEditorForForm = s
End Function


Private Function Code_modLambdaStore() As String
    Dim s As String
    s = s & "Attribute VB_Name = ""modLambdaStore""" & vbCrLf
    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function IsLambdaName(ByVal nm As Name) As Boolean" & vbCrLf
    s = s & "    Dim s As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    s = UCase$(Replace(nm.RefersTo, vbCrLf, """"))" & vbCrLf
    s = s & "    s = Replace(s, vbLf, """")" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    IsLambdaName = InStr(1, s, ""=LAMBDA("", vbTextCompare) > 0 _" & vbCrLf
    s = s & "        Or InStr(1, s, ""=LAMBDA ("", vbTextCompare) > 0" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function CleanNameText(ByVal s As String) As String" & vbCrLf
    s = s & "    s = Trim$(s)" & vbCrLf
    s = s & "    If Left$(s, 1) = ""="" Then s = Mid$(s, 2)" & vbCrLf
    s = s & "    CleanNameText = s" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function CleanFormulaText(ByVal s As String) As String" & vbCrLf
    s = s & "    s = NormalizeFormulaForName(s)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Len(s) = 0 Then" & vbCrLf
    s = s & "        CleanFormulaText = """"" & vbCrLf
    s = s & "    ElseIf Left$(s, 1) = ""="" Then" & vbCrLf
    s = s & "        CleanFormulaText = s" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        CleanFormulaText = ""="" & s" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function NormalizeFormulaForName(ByVal s As String) As String" & vbCrLf
    s = s & "    s = Replace(s, vbCrLf, vbLf)" & vbCrLf
    s = s & "    s = Replace(s, vbCr, vbLf)" & vbCrLf
    s = s & "    s = Replace(s, vbTab, "" "")" & vbCrLf
    s = s & "    s = Replace(s, vbLf, "" "")" & vbCrLf
    s = s & "    s = Trim$(s)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Do While InStr(1, s, ""  "", vbBinaryCompare) > 0" & vbCrLf
    s = s & "        s = Replace(s, ""  "", "" "")" & vbCrLf
    s = s & "    Loop" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    NormalizeFormulaForName = s" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function LambdaNameExists(ByVal wb As Workbook, ByVal lambdaName As String) As Boolean" & vbCrLf
    s = s & "    Dim nm As Name" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Set nm = wb.Names(CleanNameText(lambdaName))" & vbCrLf
    s = s & "    LambdaNameExists = Not nm Is Nothing" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function GetLambdaName(ByVal wb As Workbook, ByVal lambdaName As String) As Name" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Set GetLambdaName = wb.Names(CleanNameText(lambdaName))" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function GetLambdaNames(ByVal wb As Workbook) As Collection" & vbCrLf
    s = s & "    Dim out As New Collection" & vbCrLf
    s = s & "    Dim nm As Name" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    For Each nm In wb.Names" & vbCrLf
    s = s & "        If IsLambdaName(nm) Then out.Add nm.Name" & vbCrLf
    s = s & "    Next nm" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set GetLambdaNames = out" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub SaveLambdaName(ByVal wb As Workbook, ByVal lambdaName As String, ByVal formulaText As String, Optional ByVal commentText As String = """")" & vbCrLf
    s = s & "    Dim cleanedName As String" & vbCrLf
    s = s & "    Dim cleanedFormula As String" & vbCrLf
    s = s & "    Dim nm As Name" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    cleanedName = CleanNameText(lambdaName)" & vbCrLf
    s = s & "    cleanedFormula = CleanFormulaText(formulaText)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Len(cleanedName) = 0 Then Err.Raise vbObjectError + 1000, , ""Function name is required.""" & vbCrLf
    s = s & "    If Len(cleanedFormula) = 0 Then Err.Raise vbObjectError + 1001, , ""Formula is required.""" & vbCrLf
    s = s & "    If InStr(1, cleanedFormula, ""=LAMBDA"", vbTextCompare) <> 1 Then Err.Raise vbObjectError + 1002, , ""Formula must start with =LAMBDA(...).""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set nm = GetLambdaName(wb, cleanedName)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If nm Is Nothing Then" & vbCrLf
    s = s & "        wb.Names.Add Name:=cleanedName, RefersTo:=cleanedFormula, Visible:=True" & vbCrLf
    s = s & "        Set nm = wb.Names(cleanedName)" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        nm.RefersTo = cleanedFormula" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    nm.Comment = Left$(commentText, 255)" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub DeleteLambdaName(ByVal wb As Workbook, ByVal lambdaName As String)" & vbCrLf
    s = s & "    Dim nm As Name" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set nm = GetLambdaName(wb, lambdaName)" & vbCrLf
    s = s & "    If nm Is Nothing Then Err.Raise vbObjectError + 1003, , ""Name not found.""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    nm.Delete" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function TryEvaluateFormula(ByVal formulaText As String, Optional ByVal wb As Workbook) As Variant" & vbCrLf
    s = s & "    Dim oldRefStyle As XlReferenceStyle" & vbCrLf
    s = s & "    Dim expr As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    expr = CleanFormulaText(formulaText)" & vbCrLf
    s = s & "    oldRefStyle = Application.ReferenceStyle" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo Fail" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Not wb Is Nothing Then wb.Activate" & vbCrLf
    s = s & "    Application.ReferenceStyle = xlA1" & vbCrLf
    s = s & "    TryEvaluateFormula = Application.Evaluate(expr)" & vbCrLf
    s = s & "    Application.ReferenceStyle = oldRefStyle" & vbCrLf
    s = s & "    Exit Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Fail:" & vbCrLf
    s = s & "    Application.ReferenceStyle = oldRefStyle" & vbCrLf
    s = s & "    TryEvaluateFormula = CVErr(xlErrValue)" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function ValueToText(ByVal v As Variant) As String" & vbCrLf
    s = s & "    On Error GoTo Fail" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If IsError(v) Then" & vbCrLf
    s = s & "        ValueToText = ErrorText(v)" & vbCrLf
    s = s & "    ElseIf IsArray(v) Then" & vbCrLf
    s = s & "        ValueToText = ArrayToText(v)" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        ValueToText = CStr(v)" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Exit Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Fail:" & vbCrLf
    s = s & "    ValueToText = ""<unable to display result>""" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function ErrorText(ByVal v As Variant) As String" & vbCrLf
    s = s & "    Select Case CLng(v)" & vbCrLf
    s = s & "        Case xlErrDiv0" & vbCrLf
    s = s & "            ErrorText = ""#DIV/0!""" & vbCrLf
    s = s & "        Case xlErrNA" & vbCrLf
    s = s & "            ErrorText = ""#N/A""" & vbCrLf
    s = s & "        Case xlErrName" & vbCrLf
    s = s & "            ErrorText = ""#NAME?""" & vbCrLf
    s = s & "        Case xlErrNull" & vbCrLf
    s = s & "            ErrorText = ""#NULL!""" & vbCrLf
    s = s & "        Case xlErrNum" & vbCrLf
    s = s & "            ErrorText = ""#NUM!""" & vbCrLf
    s = s & "        Case xlErrRef" & vbCrLf
    s = s & "            ErrorText = ""#REF!""" & vbCrLf
    s = s & "        Case xlErrValue" & vbCrLf
    s = s & "            ErrorText = ""#VALUE!""" & vbCrLf
    s = s & "        Case Else" & vbCrLf
    s = s & "            ErrorText = ""#ERROR""" & vbCrLf
    s = s & "    End Select" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function ArrayToText(ByVal v As Variant) As String" & vbCrLf
    s = s & "    Dim r As Long" & vbCrLf
    s = s & "    Dim c As Long" & vbCrLf
    s = s & "    Dim r1 As Long" & vbCrLf
    s = s & "    Dim r2 As Long" & vbCrLf
    s = s & "    Dim c1 As Long" & vbCrLf
    s = s & "    Dim c2 As Long" & vbCrLf
    s = s & "    Dim lines() As String" & vbCrLf
    s = s & "    Dim cells() As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    r1 = LBound(v, 1)" & vbCrLf
    s = s & "    r2 = UBound(v, 1)" & vbCrLf
    s = s & "    c1 = LBound(v, 2)" & vbCrLf
    s = s & "    c2 = UBound(v, 2)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ReDim lines(r1 To r2)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    For r = r1 To r2" & vbCrLf
    s = s & "        ReDim cells(c1 To c2)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        For c = c1 To c2" & vbCrLf
    s = s & "            cells(c) = ValueToText(v(r, c))" & vbCrLf
    s = s & "        Next c" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        lines(r) = Join(cells, vbTab)" & vbCrLf
    s = s & "    Next r" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ArrayToText = Join(lines, vbCrLf)" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function FormatLambdaFormula(ByVal s As String) As String" & vbCrLf
    s = s & "    Dim t As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    t = CleanFormulaText(s)" & vbCrLf
    s = s & "    t = Replace(t, "","", "","" & vbCrLf & ""    "")" & vbCrLf
    s = s & "    t = Replace(t, ""LET("", ""LET("" & vbCrLf & ""    "", , , vbTextCompare)" & vbCrLf
    s = s & "    t = Replace(t, ""LAMBDA("", ""LAMBDA("" & vbCrLf & ""    "", , , vbTextCompare)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    FormatLambdaFormula = t" & vbCrLf
    s = s & "End Function" & vbCrLf
    Code_modLambdaStore = s
End Function

Private Function Code_frmLambdaEditor() As String
    Dim s As String
    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private mLoading As Boolean" & vbCrLf
    s = s & "Private mDirty As Boolean" & vbCrLf
    s = s & "Private mCurrentName As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    Me.Caption = ""LAMBDA Function Editor""" & vbCrLf
    s = s & "    ConfigureEditor" & vbCrLf
    s = s & "    RefreshList" & vbCrLf
    s = s & "    ClearEditor" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub ConfigureEditor()" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    txtFormula.Font.Name = ""Consolas""" & vbCrLf
    s = s & "    txtFormula.Font.Size = 10" & vbCrLf
    s = s & "    txtFormula.MultiLine = True" & vbCrLf
    s = s & "    txtFormula.EnterKeyBehavior = True" & vbCrLf
    s = s & "    txtFormula.ScrollBars = fmScrollBarsBoth" & vbCrLf
    s = s & "    txtFormula.WordWrap = False" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    txtComment.MultiLine = True" & vbCrLf
    s = s & "    txtComment.ScrollBars = fmScrollBarsVertical" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    txtResult.Locked = True" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub RefreshList()" & vbCrLf
    s = s & "    Dim names As Collection" & vbCrLf
    s = s & "    Dim x As Variant" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    mLoading = True" & vbCrLf
    s = s & "    lstNames.Clear" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set names = GetLambdaNames(ActiveWorkbook)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    For Each x In names" & vbCrLf
    s = s & "        lstNames.AddItem CStr(x)" & vbCrLf
    s = s & "    Next x" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    lblStatus.Caption = CStr(lstNames.ListCount) & "" LAMBDA function(s) found.""" & vbCrLf
    s = s & "    mLoading = False" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub ClearEditor()" & vbCrLf
    s = s & "    mLoading = True" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    mCurrentName = vbNullString" & vbCrLf
    s = s & "    txtName.Text = vbNullString" & vbCrLf
    s = s & "    txtComment.Text = vbNullString" & vbCrLf
    s = s & "    txtFormula.Text = ""=LAMBDA("" & vbCrLf & ""    x,"" & vbCrLf & ""    x"" & vbCrLf & "")""" & vbCrLf
    s = s & "    txtTestFormula.Text = vbNullString" & vbCrLf
    s = s & "    txtResult.Text = vbNullString" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    mDirty = False" & vbCrLf
    s = s & "    mLoading = False" & vbCrLf
    s = s & "    txtName.SetFocus" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub LoadLambda(ByVal lambdaName As String)" & vbCrLf
    s = s & "    Dim nm As Name" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set nm = GetLambdaName(ActiveWorkbook, lambdaName)" & vbCrLf
    s = s & "    If nm Is Nothing Then Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    mLoading = True" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    mCurrentName = nm.Name" & vbCrLf
    s = s & "    txtName.Text = nm.Name" & vbCrLf
    s = s & "    txtFormula.Text = nm.RefersTo" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    txtComment.Text = nm.Comment" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    txtTestFormula.Text = ""="" & nm.Name & ""()""" & vbCrLf
    s = s & "    txtResult.Text = vbNullString" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    mDirty = False" & vbCrLf
    s = s & "    mLoading = False" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function ConfirmDiscardChanges() As Boolean" & vbCrLf
    s = s & "    If Not mDirty Then" & vbCrLf
    s = s & "        ConfirmDiscardChanges = True" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        ConfirmDiscardChanges = MsgBox(""Discard unsaved changes?"", vbQuestion + vbYesNo, ""LAMBDA Editor"") = vbYes" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub MarkDirty()" & vbCrLf
    s = s & "    If mLoading Then Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    mDirty = True" & vbCrLf
    s = s & "    lblStatus.Caption = ""Unsaved changes""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub lstNames_Click()" & vbCrLf
    s = s & "    If mLoading Then Exit Sub" & vbCrLf
    s = s & "    If lstNames.ListIndex < 0 Then Exit Sub" & vbCrLf
    s = s & "    If Not ConfirmDiscardChanges Then Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    LoadLambda CStr(lstNames.Value)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdNew_Click()" & vbCrLf
    s = s & "    If Not ConfirmDiscardChanges Then Exit Sub" & vbCrLf
    s = s & "    ClearEditor" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdRefresh_Click()" & vbCrLf
    s = s & "    If Not ConfirmDiscardChanges Then Exit Sub" & vbCrLf
    s = s & "    RefreshList" & vbCrLf
    s = s & "    ClearEditor" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdSave_Click()" & vbCrLf
    s = s & "    SaveCurrentLambda" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub SaveCurrentLambda()" & vbCrLf
    s = s & "    On Error GoTo Fail" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    SaveLambdaName ActiveWorkbook, txtName.Text, txtFormula.Text, txtComment.Text" & vbCrLf
    s = s & "    mCurrentName = CleanNameText(txtName.Text)" & vbCrLf
    s = s & "    mDirty = False" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    RefreshList" & vbCrLf
    s = s & "    SelectNameInList mCurrentName" & vbCrLf
    s = s & "    LoadLambda mCurrentName" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    lblStatus.Caption = ""Saved "" & mCurrentName" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Fail:" & vbCrLf
    s = s & "    lblStatus.Caption = Err.Description" & vbCrLf
    s = s & "    MsgBox Err.Description, vbExclamation, ""Save failed""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub SelectNameInList(ByVal lambdaName As String)" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    For i = 0 To lstNames.ListCount - 1" & vbCrLf
    s = s & "        If StrComp(CStr(lstNames.List(i)), lambdaName, vbTextCompare) = 0 Then" & vbCrLf
    s = s & "            lstNames.ListIndex = i" & vbCrLf
    s = s & "            Exit Sub" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdDelete_Click()" & vbCrLf
    s = s & "    Dim lambdaName As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    lambdaName = CleanNameText(txtName.Text)" & vbCrLf
    s = s & "    If Len(lambdaName) = 0 Then Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If MsgBox(""Delete "" & lambdaName & ""?"", vbQuestion + vbYesNo, ""Delete LAMBDA"") <> vbYes Then Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo Fail" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    DeleteLambdaName ActiveWorkbook, lambdaName" & vbCrLf
    s = s & "    RefreshList" & vbCrLf
    s = s & "    ClearEditor" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    lblStatus.Caption = ""Deleted "" & lambdaName" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Fail:" & vbCrLf
    s = s & "    lblStatus.Caption = Err.Description" & vbCrLf
    s = s & "    MsgBox Err.Description, vbExclamation, ""Delete failed""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdValidate_Click()" & vbCrLf
    s = s & "    ValidateCurrentLambda" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub ValidateCurrentLambda()" & vbCrLf
    s = s & "    Dim v As Variant" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    v = TryEvaluateFormula(txtFormula.Text, ActiveWorkbook)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If IsError(v) Then" & vbCrLf
    s = s & "        txtResult.Text = ValueToText(v)" & vbCrLf
    s = s & "        lblStatus.Caption = ""Formula did not evaluate cleanly. This may still be valid if it expects arguments.""" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        txtResult.Text = ValueToText(v)" & vbCrLf
    s = s & "        lblStatus.Caption = ""Formula evaluates.""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdTest_Click()" & vbCrLf
    s = s & "    Dim v As Variant" & vbCrLf
    s = s & "    Dim expr As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    expr = Trim$(txtTestFormula.Text)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Len(expr) = 0 Then" & vbCrLf
    s = s & "        MsgBox ""Enter a test formula, for example =MyLambda(1,2)."", vbInformation, ""Test LAMBDA""" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If mDirty Then SaveCurrentLambda" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    v = TryEvaluateFormula(expr, ActiveWorkbook)" & vbCrLf
    s = s & "    txtResult.Text = ValueToText(v)" & vbCrLf
    s = s & "    lblStatus.Caption = ""Test completed.""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdClose_Click()" & vbCrLf
    s = s & "    If ConfirmDiscardChanges Then Unload Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub txtName_Change()" & vbCrLf
    s = s & "    MarkDirty" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub txtFormula_Change()" & vbCrLf
    s = s & "    MarkDirty" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub txtComment_Change()" & vbCrLf
    s = s & "    MarkDirty" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub txtFormula_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)" & vbCrLf
    s = s & "    If Shift = 2 And KeyCode = vbKeyS Then" & vbCrLf
    s = s & "        KeyCode = 0" & vbCrLf
    s = s & "        SaveCurrentLambda" & vbCrLf
    s = s & "    ElseIf Shift = 2 And KeyCode = vbKeyR Then" & vbCrLf
    s = s & "        KeyCode = 0" & vbCrLf
    s = s & "        txtFormula.Text = FormatLambdaFormula(txtFormula.Text)" & vbCrLf
    s = s & "        txtFormula.SelStart = Len(txtFormula.Text)" & vbCrLf
    s = s & "    ElseIf KeyCode = vbKeyTab Then" & vbCrLf
    s = s & "        KeyCode = 0" & vbCrLf
    s = s & "        InsertAtCursor txtFormula, ""    """ & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub txtTestFormula_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)" & vbCrLf
    s = s & "    If KeyCode = vbKeyReturn Then" & vbCrLf
    s = s & "        KeyCode = 0" & vbCrLf
    s = s & "        cmdTest_Click" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub InsertAtCursor(ByVal tb As MSForms.TextBox, ByVal s As String)" & vbCrLf
    s = s & "    Dim p As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    p = tb.SelStart" & vbCrLf
    s = s & "    tb.Text = Left$(tb.Text, p) & s & Mid$(tb.Text, p + tb.SelLength + 1)" & vbCrLf
    s = s & "    tb.SelStart = p + Len(s)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    Code_frmLambdaEditor = s
End Function