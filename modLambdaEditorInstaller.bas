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
    SafeSet frm, "Caption", " "
    SafeSet frm, "Width", 1020
    SafeSet frm, "Height", 730
    SafeSet frm, "InsideWidth", 880
    SafeSet frm, "InsideHeight", 585
    SafeSetCompProperty comp, "Width", 1020
    SafeSetCompProperty comp, "Height", 730

    mInstallStep = "Adding labels and function list"
    AddLabel frm, "lblAppTitle", "LAMBDA Function Editor", 12, 8, 220, 18
    AddLabel frm, "lblFilter", "Filter regex", 12, 34, 80, 18
    AddTextBox frm, "txtFilter", 12, 54, 190, 22, False, False
    AddLabel frm, "lblFunctions", "Functions", 12, 84, 120, 18
    AddListBox frm, "lstNames", 12, 104, 190, 548

    mInstallStep = "Adding name field and top buttons"
    AddLabel frm, "lblName", "Name", 220, 8, 80, 18
    AddTextBox frm, "txtName", 220, 28, 250, 22, False, False

    AddButton frm, "cmdNew", "New", 488, 26, 58, 24
    AddButton frm, "cmdSave", "Save", 552, 26, 58, 24
    AddButton frm, "cmdDelete", "Delete", 616, 26, 62, 24
    AddButton frm, "cmdRefresh", "Refresh", 684, 26, 70, 24
    AddButton frm, "cmdClose", "Close", 760, 26, 62, 24
    AddLambdaIcon frm

    mInstallStep = "Adding comment field"
    AddLabel frm, "lblComment", "Comment", 220, 58, 100, 18
    AddTextBox frm, "txtComment", 220, 78, 690, 46, True, True

    mInstallStep = "Adding formula editor"
    AddLabel frm, "lblFormula", "Formula", 220, 132, 100, 18
    Set txtFormulaCtl = AddTextBox(frm, "txtFormula", 220, 152, 690, 300, True, False)

    mInstallStep = "Adding test controls"
    AddLabel frm, "lblTest", "Test formula", 220, 462, 100, 18
    AddTextBox frm, "txtTestFormula", 220, 426, 325, 22, False, False
    AddButton frm, "cmdValidate", "Validate", 554, 424, 76, 24
    AddButton frm, "cmdMinify", "Minify", 636, 424, 64, 24
    AddButton frm, "cmdVisualize", "Visualize", 706, 424, 80, 24
    AddButton frm, "cmdTest", "Test", 792, 424, 50, 24

    mInstallStep = "Adding result controls"
    AddLabel frm, "lblResult", "Result", 220, 514, 100, 18
    Set txtResultCtl = AddTextBox(frm, "txtResult", 220, 534, 690, 86, True, True)

    AddLabel frm, "lblStatus", "", 12, 668, 960, 18

    AddButton frm, "cmdImportFile", "Import file", 220, 638, 95, 24
    AddButton frm, "cmdImportUrl", "Import URL", 322, 638, 95, 24
    AddButton frm, "cmdExport", "Export", 424, 638, 95, 24

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

    mInstallStep = "Auto-fitting UserForm size"
    AutoFitForm frm

    mInstallStep = "Configuring fixed-width fonts"
    SafeSet frm.Controls("lstNames").Font, "Name", "Consolas"
    SafeSet frm.Controls("lstNames").Font, "Size", 10
    SafeSet frm.Controls("txtFilter").Font, "Name", "Consolas"
    SafeSet frm.Controls("txtFilter").Font, "Size", 10
    SafeSet frm.Controls("txtName").Font, "Name", "Consolas"
    SafeSet frm.Controls("txtName").Font, "Size", 10
    SafeSet frm.Controls("txtComment").Font, "Name", "Consolas"
    SafeSet frm.Controls("txtComment").Font, "Size", 10
    SafeSet frm.Controls("txtFormula").Font, "Name", "Consolas"
    SafeSet frm.Controls("txtFormula").Font, "Size", 10
    SafeSet frm.Controls("txtTestFormula").Font, "Name", "Consolas"
    SafeSet frm.Controls("txtTestFormula").Font, "Size", 10
    SafeSet frm.Controls("txtResult").Font, "Name", "Consolas"
    SafeSet frm.Controls("txtResult").Font, "Size", 10
    SafeSet frm.Controls("lblAppTitle").Font, "Bold", True

    mInstallStep = "Adding UserForm code"
    comp.CodeModule.AddFromString Code_frmLambdaEditor()
End Function

Private Sub AutoFitForm(ByVal frm As Object)
    Dim c As Object
    Dim maxRight As Double
    Dim maxBottom As Double

    For Each c In frm.Controls
        If c.Left + c.Width > maxRight Then maxRight = c.Left + c.Width
        If c.Top + c.Height > maxBottom Then maxBottom = c.Top + c.Height
    Next c

    maxRight = maxRight + 24
    maxBottom = maxBottom + 36

    If maxRight < 860 Then maxRight = 860
    If maxBottom < 560 Then maxBottom = 560

    SafeSet frm, "InsideWidth", maxRight
    SafeSet frm, "InsideHeight", maxBottom

    SafeSet frm, "Width", maxRight + 40
    SafeSet frm, "Height", maxBottom + 60
End Sub

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

Private Sub AddLambdaIcon(ByVal frm As Object)
    Dim c As Object

    Set c = AddControl(frm, "Forms.Label.1", "lblLambdaIcon")
    SafeSet c, "Caption", ChrW$(&H3BB)
    SafeSet c, "BackStyle", 0
    SafeSet c, "BorderStyle", 0
    SafeSet c, "SpecialEffect", 0
    SafeSet c, "TextAlign", 2
    SafeSet c, "ControlTipText", "LAMBDA Function Editor"
    SafeSet c.Font, "Name", "Segoe UI Symbol"
    SafeSet c.Font, "Size", 34
    SafeSet c.Font, "Bold", True
    SafeSet c, "ForeColor", RGB(86, 65, 170)
    PlaceControl c, 936, 8, 50, 50
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

    If controlName = "lstNames" Then
        SafeSet c, "MultiSelect", 2
    End If
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

Private Sub SafeSetCompProperty(ByVal comp As Object, ByVal propertyName As String, ByVal value As Variant)
    On Error Resume Next
    comp.Properties(propertyName).Value = value
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
    s = s & "    s = FormulaHeadForDetection(nm.RefersTo)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    IsLambdaName = Left$(s, 8) = ""=LAMBDA(""" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function FormulaHeadForDetection(ByVal formulaText As String) As String" & vbCrLf
    s = s & "    Dim s As String" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim ch As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    s = formulaText" & vbCrLf
    s = s & "    s = Replace(s, vbCrLf, "" "")" & vbCrLf
    s = s & "    s = Replace(s, vbCr, "" "")" & vbCrLf
    s = s & "    s = Replace(s, vbLf, "" "")" & vbCrLf
    s = s & "    s = Replace(s, vbTab, "" "")" & vbCrLf
    s = s & "    s = Trim$(s)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Left$(s, 1) = ""="" Then" & vbCrLf
    s = s & "        i = 2" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        Do While i <= Len(s)" & vbCrLf
    s = s & "            ch = Mid$(s, i, 1)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "            If ch <> "" "" Then Exit Do" & vbCrLf
    s = s & "            i = i + 1" & vbCrLf
    s = s & "        Loop" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        s = ""="" & Mid$(s, i)" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    FormulaHeadForDetection = UCase$(s)" & vbCrLf
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
    s = s & "    If wb Is Nothing Then Set wb = ActiveWorkbook" & vbCrLf
    s = s & "    If Not wb Is Nothing Then wb.Activate" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Application.ReferenceStyle = xlA1" & vbCrLf
    s = s & "    TryEvaluateFormula = EvaluateWithFormula2Spill(expr, wb)" & vbCrLf
    s = s & "    Application.ReferenceStyle = oldRefStyle" & vbCrLf
    s = s & "    Exit Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Fail:" & vbCrLf
    s = s & "    Application.ReferenceStyle = oldRefStyle" & vbCrLf
    s = s & "    TryEvaluateFormula = CVErr(xlErrValue)" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function EvaluateWithFormula2Spill(ByVal expr As String, ByVal wb As Workbook) As Variant" & vbCrLf
    s = s & "    Dim ws As Worksheet" & vbCrLf
    s = s & "    Dim cell As Range" & vbCrLf
    s = s & "    Dim spillRange As Range" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo Fallback" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set ws = GetScratchSheet(wb)" & vbCrLf
    s = s & "    Set cell = ws.Range(""A1"")" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ws.Cells.Clear" & vbCrLf
    s = s & "    cell.Formula2 = expr" & vbCrLf
    s = s & "    cell.Calculate" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Set spillRange = cell.SpillingToRange" & vbCrLf
    s = s & "    On Error GoTo Fallback" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Not spillRange Is Nothing Then" & vbCrLf
    s = s & "        EvaluateWithFormula2Spill = spillRange.Value" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        EvaluateWithFormula2Spill = cell.Value" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ws.Cells.Clear" & vbCrLf
    s = s & "    Exit Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Fallback:" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    If Not ws Is Nothing Then ws.Cells.Clear" & vbCrLf
    s = s & "    EvaluateWithFormula2Spill = Application.Evaluate(expr)" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function GetScratchSheet(ByVal wb As Workbook) As Worksheet" & vbCrLf
    s = s & "    Const scratchName As String = ""__LambdaEditorScratch""" & vbCrLf
    s = s & "    Dim ws As Worksheet" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Set ws = wb.Worksheets(scratchName)" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If ws Is Nothing Then" & vbCrLf
    s = s & "        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))" & vbCrLf
    s = s & "        ws.Name = scratchName" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    ws.Visible = xlSheetVeryHidden" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set GetScratchSheet = ws" & vbCrLf
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
    s = s & "    Dim dims As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    dims = ArrayDimensions(v)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If dims = 1 Then" & vbCrLf
    s = s & "        ArrayToText = Array1DToText(v)" & vbCrLf
    s = s & "    ElseIf dims = 2 Then" & vbCrLf
    s = s & "        ArrayToText = Array2DToText(v)" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        ArrayToText = ""<array with "" & CStr(dims) & "" dimensions>""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function ArrayDimensions(ByVal v As Variant) As Long" & vbCrLf
    s = s & "    Dim n As Long" & vbCrLf
    s = s & "    Dim tmp As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo Done" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    For n = 1 To 60" & vbCrLf
    s = s & "        tmp = LBound(v, n)" & vbCrLf
    s = s & "    Next n" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Done:" & vbCrLf
    s = s & "    ArrayDimensions = n - 1" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function Array1DToText(ByVal v As Variant) As String" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim i1 As Long" & vbCrLf
    s = s & "    Dim i2 As Long" & vbCrLf
    s = s & "    Dim cells() As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    i1 = LBound(v, 1)" & vbCrLf
    s = s & "    i2 = UBound(v, 1)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ReDim cells(i1 To i2)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    For i = i1 To i2" & vbCrLf
    s = s & "        cells(i) = ValueToText(v(i))" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Array1DToText = Join(cells, ""  "")" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function Array2DToText(ByVal v As Variant) As String" & vbCrLf
    s = s & "    Dim r As Long" & vbCrLf
    s = s & "    Dim c As Long" & vbCrLf
    s = s & "    Dim r1 As Long" & vbCrLf
    s = s & "    Dim r2 As Long" & vbCrLf
    s = s & "    Dim c1 As Long" & vbCrLf
    s = s & "    Dim c2 As Long" & vbCrLf
    s = s & "    Dim lines() As String" & vbCrLf
    s = s & "    Dim cells() As String" & vbCrLf
    s = s & "    Dim widths() As Long" & vbCrLf
    s = s & "    Dim textValue As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    r1 = LBound(v, 1)" & vbCrLf
    s = s & "    r2 = UBound(v, 1)" & vbCrLf
    s = s & "    c1 = LBound(v, 2)" & vbCrLf
    s = s & "    c2 = UBound(v, 2)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ReDim widths(c1 To c2)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    For r = r1 To r2" & vbCrLf
    s = s & "        For c = c1 To c2" & vbCrLf
    s = s & "            textValue = ValueToText(v(r, c))" & vbCrLf
    s = s & "            If Len(textValue) > widths(c) Then widths(c) = Len(textValue)" & vbCrLf
    s = s & "        Next c" & vbCrLf
    s = s & "    Next r" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ReDim lines(r1 To r2)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    For r = r1 To r2" & vbCrLf
    s = s & "        ReDim cells(c1 To c2)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        For c = c1 To c2" & vbCrLf
    s = s & "            textValue = ValueToText(v(r, c))" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "            If c < c2 Then" & vbCrLf
    s = s & "                cells(c) = PadRightText(textValue, widths(c)) & ""  """ & vbCrLf
    s = s & "            Else" & vbCrLf
    s = s & "                cells(c) = textValue" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        Next c" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        lines(r) = Join(cells, vbNullString)" & vbCrLf
    s = s & "    Next r" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Array2DToText = Join(lines, vbCrLf)" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function PadRightText(ByVal s As String, ByVal width As Long) As String" & vbCrLf
    s = s & "    If Len(s) >= width Then" & vbCrLf
    s = s & "        PadRightText = s" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        PadRightText = s & Space$(width - Len(s))" & vbCrLf
    s = s & "    End If" & vbCrLf
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
    s = s & "" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function MinifyLambdaDefinition(ByVal formulaText As String, Optional ByVal shortenNames As Boolean = True) As String" & vbCrLf
    s = s & "    Dim s As String" & vbCrLf
    s = s & "    Dim binders As Object" & vbCrLf
    s = s & "    Dim allIds As Object" & vbCrLf
    s = s & "    Dim mapping As Object" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    s = CleanFormulaText(formulaText)" & vbCrLf
    s = s & "    s = StripWhitespaceOutsideStrings(s)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If shortenNames Then" & vbCrLf
    s = s & "        Set binders = CreateObject(""Scripting.Dictionary"")" & vbCrLf
    s = s & "        Set allIds = CreateObject(""Scripting.Dictionary"")" & vbCrLf
    s = s & "        CollectIdentifierTokens s, allIds" & vbCrLf
    s = s & "        CollectLambdaAndLetBinders s, binders" & vbCrLf
    s = s & "        Set mapping = BuildShortNameMap(binders, allIds)" & vbCrLf
    s = s & "        s = RenameFormulaIdentifiers(s, mapping)" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    MinifyLambdaDefinition = s" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function StripWhitespaceOutsideStrings(ByVal s As String) As String" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim ch As String" & vbCrLf
    s = s & "    Dim out As String" & vbCrLf
    s = s & "    Dim inString As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    For i = 1 To Len(s)" & vbCrLf
    s = s & "        ch = Mid$(s, i, 1)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        If ch = Chr$(34) Then" & vbCrLf
    s = s & "            out = out & ch" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "            If inString And i < Len(s) And Mid$(s, i + 1, 1) = Chr$(34) Then" & vbCrLf
    s = s & "                i = i + 1" & vbCrLf
    s = s & "                out = out & Chr$(34)" & vbCrLf
    s = s & "            Else" & vbCrLf
    s = s & "                inString = Not inString" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        ElseIf inString Then" & vbCrLf
    s = s & "            out = out & ch" & vbCrLf
    s = s & "        ElseIf Not IsFormulaWhitespace(ch) Then" & vbCrLf
    s = s & "            out = out & ch" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    StripWhitespaceOutsideStrings = out" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function IsFormulaWhitespace(ByVal ch As String) As Boolean" & vbCrLf
    s = s & "    IsFormulaWhitespace = ch = "" "" Or ch = vbTab Or ch = vbCr Or ch = vbLf" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub CollectIdentifierTokens(ByVal s As String, ByVal ids As Object)" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim token As String" & vbCrLf
    s = s & "    Dim ch As String" & vbCrLf
    s = s & "    Dim inString As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    i = 1" & vbCrLf
    s = s & "    Do While i <= Len(s)" & vbCrLf
    s = s & "        ch = Mid$(s, i, 1)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        If ch = Chr$(34) Then" & vbCrLf
    s = s & "            inString = Not inString" & vbCrLf
    s = s & "            i = i + 1" & vbCrLf
    s = s & "        ElseIf Not inString And IsNameStartChar(ch) Then" & vbCrLf
    s = s & "            token = ReadIdentifierToken(s, i)" & vbCrLf
    s = s & "            If Not ids.Exists(UCase$(token)) Then ids.Add UCase$(token), token" & vbCrLf
    s = s & "            i = i + Len(token)" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            i = i + 1" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Loop" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub CollectLambdaAndLetBinders(ByVal s As String, ByVal binders As Object)" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim nameText As String" & vbCrLf
    s = s & "    Dim openPos As Long" & vbCrLf
    s = s & "    Dim closePos As Long" & vbCrLf
    s = s & "    Dim args As Collection" & vbCrLf
    s = s & "    Dim j As Long" & vbCrLf
    s = s & "    Dim argText As String" & vbCrLf
    s = s & "    Dim inString As Boolean" & vbCrLf
    s = s & "    Dim ch As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    i = 1" & vbCrLf
    s = s & "    Do While i <= Len(s)" & vbCrLf
    s = s & "        ch = Mid$(s, i, 1)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        If ch = Chr$(34) Then" & vbCrLf
    s = s & "            inString = Not inString" & vbCrLf
    s = s & "            i = i + 1" & vbCrLf
    s = s & "        ElseIf Not inString And IsNameStartChar(ch) Then" & vbCrLf
    s = s & "            nameText = ReadIdentifierToken(s, i)" & vbCrLf
    s = s & "            openPos = i + Len(nameText)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "            If openPos <= Len(s) And Mid$(s, openPos, 1) = ""("" Then" & vbCrLf
    s = s & "                If StrComp(nameText, ""LAMBDA"", vbTextCompare) = 0 Or StrComp(nameText, ""LET"", vbTextCompare) = 0 Then" & vbCrLf
    s = s & "                    closePos = FindMatchingParen(s, openPos)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "                    If closePos > openPos Then" & vbCrLf
    s = s & "                        Set args = SplitTopLevelArgs(Mid$(s, openPos + 1, closePos - openPos - 1))" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "                        If StrComp(nameText, ""LAMBDA"", vbTextCompare) = 0 Then" & vbCrLf
    s = s & "                            For j = 1 To args.Count - 1" & vbCrLf
    s = s & "                                argText = CStr(args(j))" & vbCrLf
    s = s & "                                If IsValidLambdaName(argText) Then AddBinder binders, argText" & vbCrLf
    s = s & "                            Next j" & vbCrLf
    s = s & "                        ElseIf StrComp(nameText, ""LET"", vbTextCompare) = 0 Then" & vbCrLf
    s = s & "                            For j = 1 To args.Count - 1 Step 2" & vbCrLf
    s = s & "                                argText = CStr(args(j))" & vbCrLf
    s = s & "                                If IsValidLambdaName(argText) Then AddBinder binders, argText" & vbCrLf
    s = s & "                            Next j" & vbCrLf
    s = s & "                        End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "                        CollectLambdaAndLetBinders Mid$(s, openPos + 1, closePos - openPos - 1), binders" & vbCrLf
    s = s & "                        i = closePos + 1" & vbCrLf
    s = s & "                    Else" & vbCrLf
    s = s & "                        i = openPos + 1" & vbCrLf
    s = s & "                    End If" & vbCrLf
    s = s & "                Else" & vbCrLf
    s = s & "                    i = openPos + 1" & vbCrLf
    s = s & "                End If" & vbCrLf
    s = s & "            Else" & vbCrLf
    s = s & "                i = i + Len(nameText)" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            i = i + 1" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Loop" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub AddBinder(ByVal binders As Object, ByVal nameText As String)" & vbCrLf
    s = s & "    Dim key As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    key = UCase$(nameText)" & vbCrLf
    s = s & "    If Not binders.Exists(key) Then binders.Add key, nameText" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function BuildShortNameMap(ByVal binders As Object, ByVal allIds As Object) As Object" & vbCrLf
    s = s & "    Dim map As Object" & vbCrLf
    s = s & "    Dim used As Object" & vbCrLf
    s = s & "    Dim k As Variant" & vbCrLf
    s = s & "    Dim oldName As String" & vbCrLf
    s = s & "    Dim newName As String" & vbCrLf
    s = s & "    Dim n As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set map = CreateObject(""Scripting.Dictionary"")" & vbCrLf
    s = s & "    Set used = CreateObject(""Scripting.Dictionary"")" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    For Each k In allIds.Keys" & vbCrLf
    s = s & "        used(k) = True" & vbCrLf
    s = s & "    Next k" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    n = 1" & vbCrLf
    s = s & "    For Each k In binders.Keys" & vbCrLf
    s = s & "        oldName = CStr(binders(k))" & vbCrLf
    s = s & "        newName = NextSafeShortName(n, used)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        If Len(newName) < Len(oldName) Then" & vbCrLf
    s = s & "            map(UCase$(oldName)) = newName" & vbCrLf
    s = s & "            used(UCase$(newName)) = True" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next k" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set BuildShortNameMap = map" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function NextSafeShortName(ByRef n As Long, ByVal used As Object) As String" & vbCrLf
    s = s & "    Dim candidate As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Do" & vbCrLf
    s = s & "        candidate = ""_"" & Base26Name(n)" & vbCrLf
    s = s & "        n = n + 1" & vbCrLf
    s = s & "    Loop While used.Exists(UCase$(candidate)) Or LooksLikeExcelReference(candidate)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    NextSafeShortName = candidate" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function Base26Name(ByVal n As Long) As String" & vbCrLf
    s = s & "    Dim x As Long" & vbCrLf
    s = s & "    Dim remVal As Long" & vbCrLf
    s = s & "    Dim out As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    x = n" & vbCrLf
    s = s & "    Do" & vbCrLf
    s = s & "        remVal = (x - 1) Mod 26" & vbCrLf
    s = s & "        out = Chr$(97 + remVal) & out" & vbCrLf
    s = s & "        x = (x - 1) \ 26" & vbCrLf
    s = s & "    Loop While x > 0" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Base26Name = out" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function RenameFormulaIdentifiers(ByVal s As String, ByVal mapping As Object) As String" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim ch As String" & vbCrLf
    s = s & "    Dim token As String" & vbCrLf
    s = s & "    Dim out As String" & vbCrLf
    s = s & "    Dim inString As Boolean" & vbCrLf
    s = s & "    Dim key As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    i = 1" & vbCrLf
    s = s & "    Do While i <= Len(s)" & vbCrLf
    s = s & "        ch = Mid$(s, i, 1)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        If ch = Chr$(34) Then" & vbCrLf
    s = s & "            out = out & ch" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "            If inString And i < Len(s) And Mid$(s, i + 1, 1) = Chr$(34) Then" & vbCrLf
    s = s & "                i = i + 1" & vbCrLf
    s = s & "                out = out & Chr$(34)" & vbCrLf
    s = s & "            Else" & vbCrLf
    s = s & "                inString = Not inString" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "            i = i + 1" & vbCrLf
    s = s & "        ElseIf Not inString And IsNameStartChar(ch) Then" & vbCrLf
    s = s & "            token = ReadIdentifierToken(s, i)" & vbCrLf
    s = s & "            key = UCase$(token)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "            If mapping.Exists(key) Then" & vbCrLf
    s = s & "                out = out & CStr(mapping(key))" & vbCrLf
    s = s & "            Else" & vbCrLf
    s = s & "                out = out & token" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "            i = i + Len(token)" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            out = out & ch" & vbCrLf
    s = s & "            i = i + 1" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Loop" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    RenameFormulaIdentifiers = out" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function SplitTopLevelArgs(ByVal s As String) As Collection" & vbCrLf
    s = s & "    Dim args As New Collection" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim startPos As Long" & vbCrLf
    s = s & "    Dim depth As Long" & vbCrLf
    s = s & "    Dim ch As String" & vbCrLf
    s = s & "    Dim inString As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    startPos = 1" & vbCrLf
    s = s & "    For i = 1 To Len(s)" & vbCrLf
    s = s & "        ch = Mid$(s, i, 1)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        If ch = Chr$(34) Then" & vbCrLf
    s = s & "            If inString And i < Len(s) And Mid$(s, i + 1, 1) = Chr$(34) Then" & vbCrLf
    s = s & "                i = i + 1" & vbCrLf
    s = s & "            Else" & vbCrLf
    s = s & "                inString = Not inString" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        ElseIf Not inString Then" & vbCrLf
    s = s & "            If ch = ""("" Then" & vbCrLf
    s = s & "                depth = depth + 1" & vbCrLf
    s = s & "            ElseIf ch = "")"" Then" & vbCrLf
    s = s & "                If depth > 0 Then depth = depth - 1" & vbCrLf
    s = s & "            ElseIf ch = "","" And depth = 0 Then" & vbCrLf
    s = s & "                args.Add Mid$(s, startPos, i - startPos)" & vbCrLf
    s = s & "                startPos = i + 1" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    args.Add Mid$(s, startPos)" & vbCrLf
    s = s & "    Set SplitTopLevelArgs = args" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function FindMatchingParen(ByVal s As String, ByVal openPos As Long) As Long" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim depth As Long" & vbCrLf
    s = s & "    Dim ch As String" & vbCrLf
    s = s & "    Dim inString As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    For i = openPos To Len(s)" & vbCrLf
    s = s & "        ch = Mid$(s, i, 1)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        If ch = Chr$(34) Then" & vbCrLf
    s = s & "            If inString And i < Len(s) And Mid$(s, i + 1, 1) = Chr$(34) Then" & vbCrLf
    s = s & "                i = i + 1" & vbCrLf
    s = s & "            Else" & vbCrLf
    s = s & "                inString = Not inString" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        ElseIf Not inString Then" & vbCrLf
    s = s & "            If ch = ""("" Then" & vbCrLf
    s = s & "                depth = depth + 1" & vbCrLf
    s = s & "            ElseIf ch = "")"" Then" & vbCrLf
    s = s & "                depth = depth - 1" & vbCrLf
    s = s & "                If depth = 0 Then" & vbCrLf
    s = s & "                    FindMatchingParen = i" & vbCrLf
    s = s & "                    Exit Function" & vbCrLf
    s = s & "                End If" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    FindMatchingParen = 0" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function ReadIdentifierToken(ByVal s As String, ByVal startPos As Long) As String" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim ch As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    i = startPos" & vbCrLf
    s = s & "    Do While i <= Len(s)" & vbCrLf
    s = s & "        ch = Mid$(s, i, 1)" & vbCrLf
    s = s & "        If IsNameBodyChar(ch) Then" & vbCrLf
    s = s & "            i = i + 1" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            Exit Do" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Loop" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ReadIdentifierToken = Mid$(s, startPos, i - startPos)" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function IsNameStartChar(ByVal ch As String) As Boolean" & vbCrLf
    s = s & "    If Len(ch) <> 1 Then Exit Function" & vbCrLf
    s = s & "    IsNameStartChar = (ch Like ""[A-Za-z_\\]"")" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function IsNameBodyChar(ByVal ch As String) As Boolean" & vbCrLf
    s = s & "    If Len(ch) <> 1 Then Exit Function" & vbCrLf
    s = s & "    IsNameBodyChar = (ch Like ""[A-Za-z0-9_.\\]"")" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function IsValidLambdaName(ByVal s As String) As Boolean" & vbCrLf
    s = s & "    s = Trim$(s)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Len(s) = 0 Then Exit Function" & vbCrLf
    s = s & "    If Not IsNameStartChar(Left$(s, 1)) Then Exit Function" & vbCrLf
    s = s & "    If LooksLikeExcelReference(s) Then Exit Function" & vbCrLf
    s = s & "    If InStr(1, s, ""."", vbBinaryCompare) > 0 Then Exit Function" & vbCrLf
    s = s & "    If InStr(1, s, ""!"", vbBinaryCompare) > 0 Then Exit Function" & vbCrLf
    s = s & "    If InStr(1, s, ""["", vbBinaryCompare) > 0 Then Exit Function" & vbCrLf
    s = s & "    If InStr(1, s, ""]"", vbBinaryCompare) > 0 Then Exit Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    IsValidLambdaName = True" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function LooksLikeExcelReference(ByVal s As String) As Boolean" & vbCrLf
    s = s & "    Dim t As String" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim letters As String" & vbCrLf
    s = s & "    Dim digits As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    t = UCase$(Replace(s, ""$"", """"))" & vbCrLf
    s = s & "    If Len(t) = 0 Then Exit Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If t Like ""R[1-9]*C[1-9]*"" Then" & vbCrLf
    s = s & "        LooksLikeExcelReference = True" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    For i = 1 To Len(t)" & vbCrLf
    s = s & "        If Mid$(t, i, 1) Like ""[A-Z]"" Then" & vbCrLf
    s = s & "            letters = letters & Mid$(t, i, 1)" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            Exit For" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Len(letters) > 0 And Len(letters) <= 3 Then" & vbCrLf
    s = s & "        digits = Mid$(t, Len(letters) + 1)" & vbCrLf
    s = s & "        If Len(digits) > 0 And digits Like ""[0-9]*"" Then" & vbCrLf
    s = s & "            LooksLikeExcelReference = True" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub OpenFormulaBoostVisualization(ByVal formulaText As String)" & vbCrLf
    s = s & "    Dim url As String" & vbCrLf
    s = s & "    Dim f As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    f = CleanFormulaText(formulaText)" & vbCrLf
    s = s & "    url = ""https://www.formulaboost.com/parse?f="" & UrlEncodeFormulaBoost(f)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo Fail" & vbCrLf
    s = s & "    ActiveWorkbook.FollowHyperlink Address:=url, NewWindow:=True" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Fail:" & vbCrLf
    s = s & "    MsgBox ""Could not open visualization link."" & vbCrLf & vbCrLf & url & vbCrLf & vbCrLf & Err.Description, vbExclamation, ""Visualize Formula""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function FormulaBoostVisualizationUrl(ByVal formulaText As String) As String" & vbCrLf
    s = s & "    FormulaBoostVisualizationUrl = ""https://www.formulaboost.com/parse?f="" & UrlEncodeFormulaBoost(CleanFormulaText(formulaText))" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function UrlEncodeFormulaBoost(ByVal s As String) As String" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim ch As String" & vbCrLf
    s = s & "    Dim code As Long" & vbCrLf
    s = s & "    Dim out As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    s = Replace(s, vbCrLf, vbLf)" & vbCrLf
    s = s & "    s = Replace(s, vbCr, vbLf)" & vbCrLf
    s = s & "    s = Replace(s, vbLf, """")" & vbCrLf
    s = s & "    s = Replace(s, vbTab, """")" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    For i = 1 To Len(s)" & vbCrLf
    s = s & "        ch = Mid$(s, i, 1)" & vbCrLf
    s = s & "        code = AscW(ch)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        If IsFormulaBoostSafeChar(ch) Then" & vbCrLf
    s = s & "            out = out & ch" & vbCrLf
    s = s & "        ElseIf code >= 0 And code <= 255 Then" & vbCrLf
    s = s & "            out = out & ""%"" & Right$(""0"" & Hex$(code), 2)" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            out = out & ""%3F""" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    UrlEncodeFormulaBoost = out" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function IsFormulaBoostSafeChar(ByVal ch As String) As Boolean" & vbCrLf
    s = s & "    Dim code As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    code = AscW(ch)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Select Case code" & vbCrLf
    s = s & "        Case 48 To 57, 65 To 90, 97 To 122" & vbCrLf
    s = s & "            IsFormulaBoostSafeChar = True" & vbCrLf
    s = s & "        Case 40, 41, 44, 45, 46, 61, 95, 126" & vbCrLf
    s = s & "            ' Preserve parentheses, commas and equals so URLs look like:" & vbCrLf
    s = s & "            ' ?f==lambda(x,y,let(z,2,x%2By-z))" & vbCrLf
    s = s & "            IsFormulaBoostSafeChar = True" & vbCrLf
    s = s & "        Case Else" & vbCrLf
    s = s & "            IsFormulaBoostSafeChar = False" & vbCrLf
    s = s & "    End Select" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub ImportLambdasFromTextFile(Optional ByVal wb As Workbook)" & vbCrLf
    s = s & "    Dim filePath As Variant" & vbCrLf
    s = s & "    Dim textContent As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If wb Is Nothing Then Set wb = ActiveWorkbook" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    filePath = Application.GetOpenFilename(""Text Files (*.txt), *.txt, All Files (*.*), *.*"", , ""Import LAMBDA definitions"")" & vbCrLf
    s = s & "    If VarType(filePath) = vbBoolean Then Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    textContent = ReadTextFile(CStr(filePath))" & vbCrLf
    s = s & "    ImportLambdasFromText textContent, wb" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub ImportLambdasFromUrl(Optional ByVal wb As Workbook)" & vbCrLf
    s = s & "    Dim url As String" & vbCrLf
    s = s & "    Dim textContent As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If wb Is Nothing Then Set wb = ActiveWorkbook" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    url = InputBox(""Enter a raw text URL containing LAMBDA definitions:"", ""Import LAMBDAs from URL"")" & vbCrLf
    s = s & "    If Len(Trim$(url)) = 0 Then Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    textContent = ReadUrlText(url)" & vbCrLf
    s = s & "    ImportLambdasFromText textContent, wb" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub ExportLambdasToTextFile(Optional ByVal wb As Workbook)" & vbCrLf
    s = s & "    Dim filePath As Variant" & vbCrLf
    s = s & "    Dim textContent As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If wb Is Nothing Then Set wb = ActiveWorkbook" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    textContent = ExportLambdasToText(wb)" & vbCrLf
    s = s & "    If Len(textContent) = 0 Then" & vbCrLf
    s = s & "        MsgBox ""No LAMBDA functions found in the Name Manager."", vbInformation, ""Export LAMBDAs""" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    filePath = Application.GetSaveAsFilename(InitialFileName:=""my_excel_lambda_functions.txt"", _" & vbCrLf
    s = s & "                                             FileFilter:=""Text Files (*.txt), *.txt"", _" & vbCrLf
    s = s & "                                             Title:=""Export LAMBDA definitions"")" & vbCrLf
    s = s & "    If VarType(filePath) = vbBoolean Then Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Len(Dir$(CStr(filePath))) > 0 Then" & vbCrLf
    s = s & "        If MsgBox(""The file already exists:"" & vbCrLf & vbCrLf & _" & vbCrLf
    s = s & "                  CStr(filePath) & vbCrLf & vbCrLf & _" & vbCrLf
    s = s & "                  ""Overwrite this file?"", _" & vbCrLf
    s = s & "                  vbExclamation + vbYesNo, ""Confirm export overwrite"") <> vbYes Then Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    WriteTextFile CStr(filePath), textContent" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Shell ""notepad.exe "" & Chr$(34) & CStr(filePath) & Chr$(34), vbNormalFocus" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub ImportLambdasFromText(ByVal textContent As String, Optional ByVal wb As Workbook)" & vbCrLf
    s = s & "    Dim lines As Variant" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim line As String" & vbCrLf
    s = s & "    Dim pendingComment As String" & vbCrLf
    s = s & "    Dim commentText As String" & vbCrLf
    s = s & "    Dim lambdaName As String" & vbCrLf
    s = s & "    Dim lambdaBody As String" & vbCrLf
    s = s & "    Dim inLambda As Boolean" & vbCrLf
    s = s & "    Dim importedCount As Long" & vbCrLf
    s = s & "    Dim overwriteList As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If wb Is Nothing Then Set wb = ActiveWorkbook" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    textContent = NormalizeNewlines(textContent)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    overwriteList = ImportedOverwriteList(textContent, wb)" & vbCrLf
    s = s & "    If Len(overwriteList) > 0 Then" & vbCrLf
    s = s & "        If MsgBox(""The import will overwrite existing LAMBDA function(s):"" & vbCrLf & vbCrLf & _" & vbCrLf
    s = s & "                  overwriteList & vbCrLf & _" & vbCrLf
    s = s & "                  ""Continue and overwrite them?"", _" & vbCrLf
    s = s & "                  vbExclamation + vbYesNo, ""Confirm LAMBDA overwrite"") <> vbYes Then Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    lines = Split(textContent, vbLf)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    For i = LBound(lines) To UBound(lines)" & vbCrLf
    s = s & "        line = Trim$(CStr(lines(i)))" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        If Left$(line, 2) = ""##"" Then" & vbCrLf
    s = s & "            ' Ignore double-hash comments." & vbCrLf
    s = s & "        ElseIf Len(line) = 0 Then" & vbCrLf
    s = s & "            If inLambda Then" & vbCrLf
    s = s & "                SaveImportedLambda wb, lambdaName, lambdaBody, commentText" & vbCrLf
    s = s & "                importedCount = importedCount + 1" & vbCrLf
    s = s & "                lambdaName = vbNullString" & vbCrLf
    s = s & "                lambdaBody = vbNullString" & vbCrLf
    s = s & "                commentText = vbNullString" & vbCrLf
    s = s & "                inLambda = False" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        ElseIf Left$(line, 1) = ""#"" And Not inLambda Then" & vbCrLf
    s = s & "            If Len(pendingComment) > 0 Then pendingComment = pendingComment & vbLf" & vbCrLf
    s = s & "            pendingComment = pendingComment & Mid$(line, 2)" & vbCrLf
    s = s & "        ElseIf IsLambdaDefinitionStart(line) Then" & vbCrLf
    s = s & "            If inLambda Then" & vbCrLf
    s = s & "                SaveImportedLambda wb, lambdaName, lambdaBody, commentText" & vbCrLf
    s = s & "                importedCount = importedCount + 1" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "            ParseLambdaDefinitionStart line, lambdaName, lambdaBody" & vbCrLf
    s = s & "            commentText = pendingComment" & vbCrLf
    s = s & "            pendingComment = vbNullString" & vbCrLf
    s = s & "            inLambda = True" & vbCrLf
    s = s & "        ElseIf inLambda Then" & vbCrLf
    s = s & "            lambdaBody = lambdaBody & vbLf & line" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If inLambda Then" & vbCrLf
    s = s & "        SaveImportedLambda wb, lambdaName, lambdaBody, commentText" & vbCrLf
    s = s & "        importedCount = importedCount + 1" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    MsgBox ""Imported "" & CStr(importedCount) & "" LAMBDA definition(s)."", vbInformation, ""Import LAMBDAs""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function ExportLambdasToText(Optional ByVal wb As Workbook) As String" & vbCrLf
    s = s & "    Dim nameItem As Name" & vbCrLf
    s = s & "    Dim outputText As String" & vbCrLf
    s = s & "    Dim formulaText As String" & vbCrLf
    s = s & "    Dim commentText As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If wb Is Nothing Then Set wb = ActiveWorkbook" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    For Each nameItem In wb.Names" & vbCrLf
    s = s & "        If IsLambdaName(nameItem) Then" & vbCrLf
    s = s & "            commentText = CommentLinesForExport(nameItem.Comment)" & vbCrLf
    s = s & "            formulaText = CleanFormulaText(nameItem.RefersTo)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "            If Len(commentText) > 0 Then outputText = outputText & commentText & vbCrLf" & vbCrLf
    s = s & "            outputText = outputText & nameItem.Name & "" "" & formulaText & vbCrLf & vbCrLf" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next nameItem" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ExportLambdasToText = outputText" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function ImportedOverwriteList(ByVal textContent As String, ByVal wb As Workbook) As String" & vbCrLf
    s = s & "    Dim lines As Variant" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim line As String" & vbCrLf
    s = s & "    Dim lambdaName As String" & vbCrLf
    s = s & "    Dim lambdaBody As String" & vbCrLf
    s = s & "    Dim outputText As String" & vbCrLf
    s = s & "    Dim shownCount As Long" & vbCrLf
    s = s & "    Dim totalCount As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    lines = Split(textContent, vbLf)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    For i = LBound(lines) To UBound(lines)" & vbCrLf
    s = s & "        line = Trim$(CStr(lines(i)))" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        If IsLambdaDefinitionStart(line) Then" & vbCrLf
    s = s & "            ParseLambdaDefinitionStart line, lambdaName, lambdaBody" & vbCrLf
    s = s & "            lambdaName = CleanImportedName(lambdaName)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "            If LambdaNameExists(wb, lambdaName) Then" & vbCrLf
    s = s & "                totalCount = totalCount + 1" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "                If shownCount < 20 Then" & vbCrLf
    s = s & "                    outputText = outputText & lambdaName & vbCrLf" & vbCrLf
    s = s & "                    shownCount = shownCount + 1" & vbCrLf
    s = s & "                End If" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If totalCount > shownCount Then" & vbCrLf
    s = s & "        outputText = outputText & ""...and "" & CStr(totalCount - shownCount) & "" more."" & vbCrLf" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ImportedOverwriteList = outputText" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub SaveImportedLambda(ByVal wb As Workbook, ByVal lambdaName As String, ByVal lambdaBody As String, ByVal commentText As String)" & vbCrLf
    s = s & "    lambdaName = CleanImportedName(lambdaName)" & vbCrLf
    s = s & "    lambdaBody = CleanFormulaText(lambdaBody)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Len(lambdaName) = 0 Then Exit Sub" & vbCrLf
    s = s & "    If InStr(1, lambdaBody, ""=LAMBDA"", vbTextCompare) <> 1 Then Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    SaveLambdaName wb, lambdaName, lambdaBody, commentText" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function IsLambdaDefinitionStart(ByVal line As String) As Boolean" & vbCrLf
    s = s & "    Dim eqPos As Long" & vbCrLf
    s = s & "    Dim leftPart As String" & vbCrLf
    s = s & "    Dim rightPart As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    eqPos = InStr(1, line, ""="", vbBinaryCompare)" & vbCrLf
    s = s & "    If eqPos = 0 Then Exit Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    leftPart = Trim$(Left$(line, eqPos - 1))" & vbCrLf
    s = s & "    rightPart = Trim$(Mid$(line, eqPos + 1))" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Len(leftPart) = 0 Then Exit Function" & vbCrLf
    s = s & "    If Not IsValidImportedName(leftPart) Then Exit Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    IsLambdaDefinitionStart = Left$(UCase$(LTrim$(rightPart)), 7) = ""LAMBDA(""" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub ParseLambdaDefinitionStart(ByVal line As String, ByRef lambdaName As String, ByRef lambdaBody As String)" & vbCrLf
    s = s & "    Dim eqPos As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    eqPos = InStr(1, line, ""="", vbBinaryCompare)" & vbCrLf
    s = s & "    lambdaName = Trim$(Left$(line, eqPos - 1))" & vbCrLf
    s = s & "    lambdaBody = ""="" & Trim$(Mid$(line, eqPos + 1))" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function CleanImportedName(ByVal s As String) As String" & vbCrLf
    s = s & "    s = Trim$(s)" & vbCrLf
    s = s & "    s = Replace(s, "" "", """")" & vbCrLf
    s = s & "    s = Replace(s, vbTab, """")" & vbCrLf
    s = s & "    If Left$(s, 1) = ""="" Then s = Mid$(s, 2)" & vbCrLf
    s = s & "    CleanImportedName = s" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function IsValidImportedName(ByVal s As String) As Boolean" & vbCrLf
    s = s & "    Dim re As Object" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    s = CleanImportedName(s)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set re = CreateObject(""VBScript.RegExp"")" & vbCrLf
    s = s & "    re.Pattern = ""^[A-Za-z_\.][A-Za-z0-9_\.]*$""" & vbCrLf
    s = s & "    re.Global = False" & vbCrLf
    s = s & "    re.IgnoreCase = True" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    IsValidImportedName = re.Test(s)" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function CommentLinesForExport(ByVal commentText As String) As String" & vbCrLf
    s = s & "    Dim lines As Variant" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    commentText = NormalizeNewlines(commentText)" & vbCrLf
    s = s & "    If Len(commentText) = 0 Then Exit Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    lines = Split(commentText, vbLf)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    For i = LBound(lines) To UBound(lines)" & vbCrLf
    s = s & "        lines(i) = ""#"" & CStr(lines(i))" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    CommentLinesForExport = Join(lines, vbCrLf)" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function ReadTextFile(ByVal filePath As String) As String" & vbCrLf
    s = s & "    Dim stm As Object" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set stm = CreateObject(""ADODB.Stream"")" & vbCrLf
    s = s & "    stm.Type = 2" & vbCrLf
    s = s & "    stm.Charset = ""utf-8""" & vbCrLf
    s = s & "    stm.Open" & vbCrLf
    s = s & "    stm.LoadFromFile filePath" & vbCrLf
    s = s & "    ReadTextFile = stm.ReadText" & vbCrLf
    s = s & "    stm.Close" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub WriteTextFile(ByVal filePath As String, ByVal textContent As String)" & vbCrLf
    s = s & "    Dim stm As Object" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set stm = CreateObject(""ADODB.Stream"")" & vbCrLf
    s = s & "    stm.Type = 2" & vbCrLf
    s = s & "    stm.Charset = ""utf-8""" & vbCrLf
    s = s & "    stm.Open" & vbCrLf
    s = s & "    stm.WriteText textContent" & vbCrLf
    s = s & "    stm.SaveToFile filePath, 2" & vbCrLf
    s = s & "    stm.Close" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function ReadUrlText(ByVal url As String) As String" & vbCrLf
    s = s & "    Dim http As Object" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set http = CreateObject(""MSXML2.XMLHTTP"")" & vbCrLf
    s = s & "    http.Open ""GET"", url, False" & vbCrLf
    s = s & "    http.Send" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If http.Status < 200 Or http.Status >= 300 Then" & vbCrLf
    s = s & "        Err.Raise vbObjectError + 2000, , ""HTTP "" & CStr(http.Status) & "": "" & http.statusText" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ReadUrlText = CStr(http.responseText)" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function NormalizeNewlines(ByVal s As String) As String" & vbCrLf
    s = s & "    s = Replace(s, vbCrLf, vbLf)" & vbCrLf
    s = s & "    s = Replace(s, vbCr, vbLf)" & vbCrLf
    s = s & "    NormalizeNewlines = s" & vbCrLf
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
    s = s & "Private mHandlingListSelection As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    Me.Caption = "" """ & vbCrLf
    s = s & "    EnsureEditorSize" & vbCrLf
    s = s & "    HideDuplicateTitleLabels" & vbCrLf
    s = s & "    ConfigureEditor" & vbCrLf
    s = s & "    RefreshList" & vbCrLf
    s = s & "    ClearEditor" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub HideDuplicateTitleLabels()" & vbCrLf
    s = s & "    Dim c As Control" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    For Each c In Me.Controls" & vbCrLf
    s = s & "        If TypeName(c) = ""Label"" Then" & vbCrLf
    s = s & "            If Trim$(CStr(c.Caption)) = ""LAMBDA Function Editor"" Then" & vbCrLf
    s = s & "                If c.Name <> ""lblAppTitle"" Then" & vbCrLf
    s = s & "                    c.Caption = vbNullString" & vbCrLf
    s = s & "                    c.Visible = False" & vbCrLf
    s = s & "                    c.Height = 0" & vbCrLf
    s = s & "                End If" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next c" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Activate()" & vbCrLf
    s = s & "    EnsureEditorSize" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub EnsureEditorSize()" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Me.Width = 1020" & vbCrLf
    s = s & "    Me.Height = 730" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' InsideWidth and InsideHeight are read-only at runtime in some Excel/VBA builds." & vbCrLf
    s = s & "    ' Instead, set the outer form size and explicitly place every control." & vbCrLf
    s = s & "    LayoutControls" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub LayoutControls()" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    lblAppTitle.Left = 12" & vbCrLf
    s = s & "    lblAppTitle.Top = 8" & vbCrLf
    s = s & "    lblAppTitle.Width = 220" & vbCrLf
    s = s & "    lblAppTitle.Height = 18" & vbCrLf
    s = s & "    lblAppTitle.Font.Bold = True" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    lblFilter.Left = 12" & vbCrLf
    s = s & "    lblFilter.Top = 34" & vbCrLf
    s = s & "    lblFilter.Width = 80" & vbCrLf
    s = s & "    lblFilter.Height = 18" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    txtFilter.Left = 12" & vbCrLf
    s = s & "    txtFilter.Top = 54" & vbCrLf
    s = s & "    txtFilter.Width = 190" & vbCrLf
    s = s & "    txtFilter.Height = 22" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    lblFunctions.Left = 12" & vbCrLf
    s = s & "    lblFunctions.Top = 84" & vbCrLf
    s = s & "    lblFunctions.Width = 120" & vbCrLf
    s = s & "    lblFunctions.Height = 18" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    lstNames.Left = 12" & vbCrLf
    s = s & "    lstNames.Top = 104" & vbCrLf
    s = s & "    lstNames.Width = 190" & vbCrLf
    s = s & "    lstNames.Height = 548" & vbCrLf
    s = s & "    Me.lstNames.MultiSelect = 2" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    lblName.Left = 220" & vbCrLf
    s = s & "    lblName.Top = 8" & vbCrLf
    s = s & "    lblName.Width = 80" & vbCrLf
    s = s & "    lblName.Height = 18" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    txtName.Left = 220" & vbCrLf
    s = s & "    txtName.Top = 28" & vbCrLf
    s = s & "    txtName.Width = 250" & vbCrLf
    s = s & "    txtName.Height = 22" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    cmdNew.Left = 488" & vbCrLf
    s = s & "    cmdNew.Top = 26" & vbCrLf
    s = s & "    cmdNew.Width = 58" & vbCrLf
    s = s & "    cmdNew.Height = 24" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    cmdSave.Left = 552" & vbCrLf
    s = s & "    cmdSave.Top = 26" & vbCrLf
    s = s & "    cmdSave.Width = 58" & vbCrLf
    s = s & "    cmdSave.Height = 24" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    cmdDelete.Left = 616" & vbCrLf
    s = s & "    cmdDelete.Top = 26" & vbCrLf
    s = s & "    cmdDelete.Width = 62" & vbCrLf
    s = s & "    cmdDelete.Height = 24" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    cmdRefresh.Left = 684" & vbCrLf
    s = s & "    cmdRefresh.Top = 26" & vbCrLf
    s = s & "    cmdRefresh.Width = 70" & vbCrLf
    s = s & "    cmdRefresh.Height = 24" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    cmdClose.Left = 760" & vbCrLf
    s = s & "    cmdClose.Top = 26" & vbCrLf
    s = s & "    cmdClose.Width = 62" & vbCrLf
    s = s & "    cmdClose.Height = 24" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    lblLambdaIcon.Left = 936" & vbCrLf
    s = s & "    lblLambdaIcon.Top = 8" & vbCrLf
    s = s & "    lblLambdaIcon.Width = 50" & vbCrLf
    s = s & "    lblLambdaIcon.Height = 50" & vbCrLf
    s = s & "    lblLambdaIcon.Caption = ChrW$(&H3BB)" & vbCrLf
    s = s & "    lblLambdaIcon.BackStyle = 0" & vbCrLf
    s = s & "    lblLambdaIcon.BorderStyle = 0" & vbCrLf
    s = s & "    lblLambdaIcon.SpecialEffect = 0" & vbCrLf
    s = s & "    lblLambdaIcon.TextAlign = 2" & vbCrLf
    s = s & "    lblLambdaIcon.Font.Name = ""Segoe UI Symbol""" & vbCrLf
    s = s & "    lblLambdaIcon.Font.Size = 34" & vbCrLf
    s = s & "    lblLambdaIcon.Font.Bold = True" & vbCrLf
    s = s & "    lblLambdaIcon.ForeColor = RGB(86, 65, 170)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    lblComment.Left = 220" & vbCrLf
    s = s & "    lblComment.Top = 58" & vbCrLf
    s = s & "    lblComment.Width = 100" & vbCrLf
    s = s & "    lblComment.Height = 18" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    txtComment.Left = 220" & vbCrLf
    s = s & "    txtComment.Top = 78" & vbCrLf
    s = s & "    txtComment.Width = 690" & vbCrLf
    s = s & "    txtComment.Height = 46" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    lblFormula.Left = 220" & vbCrLf
    s = s & "    lblFormula.Top = 132" & vbCrLf
    s = s & "    lblFormula.Width = 100" & vbCrLf
    s = s & "    lblFormula.Height = 18" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    txtFormula.Left = 220" & vbCrLf
    s = s & "    txtFormula.Top = 152" & vbCrLf
    s = s & "    txtFormula.Width = 690" & vbCrLf
    s = s & "    txtFormula.Height = 300" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    lblTest.Left = 220" & vbCrLf
    s = s & "    lblTest.Top = 462" & vbCrLf
    s = s & "    lblTest.Width = 100" & vbCrLf
    s = s & "    lblTest.Height = 18" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    txtTestFormula.Left = 220" & vbCrLf
    s = s & "    txtTestFormula.Top = 482" & vbCrLf
    s = s & "    txtTestFormula.Width = 430" & vbCrLf
    s = s & "    txtTestFormula.Height = 22" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    cmdValidate.Left = 665" & vbCrLf
    s = s & "    cmdValidate.Top = 480" & vbCrLf
    s = s & "    cmdValidate.Width = 80" & vbCrLf
    s = s & "    cmdValidate.Height = 24" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    cmdTest.Left = 755" & vbCrLf
    s = s & "    cmdTest.Top = 480" & vbCrLf
    s = s & "    cmdTest.Width = 64" & vbCrLf
    s = s & "    cmdTest.Height = 24" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    cmdMinify.Left = 825" & vbCrLf
    s = s & "    cmdMinify.Top = 480" & vbCrLf
    s = s & "    cmdMinify.Width = 70" & vbCrLf
    s = s & "    cmdMinify.Height = 24" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    cmdVisualize.Left = 905" & vbCrLf
    s = s & "    cmdVisualize.Top = 480" & vbCrLf
    s = s & "    cmdVisualize.Width = 80" & vbCrLf
    s = s & "    cmdVisualize.Height = 24" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    cmdImportFile.Left = 220" & vbCrLf
    s = s & "    cmdImportFile.Top = 638" & vbCrLf
    s = s & "    cmdImportFile.Width = 95" & vbCrLf
    s = s & "    cmdImportFile.Height = 24" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    cmdImportUrl.Left = 322" & vbCrLf
    s = s & "    cmdImportUrl.Top = 638" & vbCrLf
    s = s & "    cmdImportUrl.Width = 95" & vbCrLf
    s = s & "    cmdImportUrl.Height = 24" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    cmdExport.Left = 424" & vbCrLf
    s = s & "    cmdExport.Top = 638" & vbCrLf
    s = s & "    cmdExport.Width = 95" & vbCrLf
    s = s & "    cmdExport.Height = 24" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    lblResult.Left = 220" & vbCrLf
    s = s & "    lblResult.Top = 514" & vbCrLf
    s = s & "    lblResult.Width = 100" & vbCrLf
    s = s & "    lblResult.Height = 18" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    txtResult.Left = 220" & vbCrLf
    s = s & "    txtResult.Top = 534" & vbCrLf
    s = s & "    txtResult.Width = 690" & vbCrLf
    s = s & "    txtResult.Height = 86" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    lblStatus.Left = 12" & vbCrLf
    s = s & "    lblStatus.Top = 668" & vbCrLf
    s = s & "    lblStatus.Width = 960" & vbCrLf
    s = s & "    lblStatus.Height = 18" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub ConfigureEditor()" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Me.lstNames.MultiSelect = 2" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    SetFixedWidthFont lstNames" & vbCrLf
    s = s & "    SetFixedWidthFont txtFilter" & vbCrLf
    s = s & "    SetFixedWidthFont txtName" & vbCrLf
    s = s & "    SetFixedWidthFont txtComment" & vbCrLf
    s = s & "    SetFixedWidthFont txtFormula" & vbCrLf
    s = s & "    SetFixedWidthFont txtTestFormula" & vbCrLf
    s = s & "    SetFixedWidthFont txtResult" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    txtFormula.MultiLine = True" & vbCrLf
    s = s & "    txtFormula.EnterKeyBehavior = True" & vbCrLf
    s = s & "    txtFormula.ScrollBars = fmScrollBarsBoth" & vbCrLf
    s = s & "    txtFormula.WordWrap = False" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    txtComment.MultiLine = True" & vbCrLf
    s = s & "    txtComment.ScrollBars = fmScrollBarsVertical" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    txtResult.Locked = True" & vbCrLf
    s = s & "    txtResult.WordWrap = False" & vbCrLf
    s = s & "    txtResult.ScrollBars = fmScrollBarsBoth" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub SetFixedWidthFont(ByVal ctl As Object)" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    ctl.Font.Name = ""Consolas""" & vbCrLf
    s = s & "    ctl.Font.Size = 10" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub RefreshList()" & vbCrLf
    s = s & "    Dim names As Collection" & vbCrLf
    s = s & "    Dim x As Variant" & vbCrLf
    s = s & "    Dim re As Object" & vbCrLf
    s = s & "    Dim filterText As String" & vbCrLf
    s = s & "    Dim useFilter As Boolean" & vbCrLf
    s = s & "    Dim totalCount As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo FilterError" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    mLoading = True" & vbCrLf
    s = s & "    lstNames.Clear" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    filterText = Trim$(txtFilter.Text)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Len(filterText) > 0 Then" & vbCrLf
    s = s & "        Set re = CreateObject(""VBScript.RegExp"")" & vbCrLf
    s = s & "        re.Pattern = filterText" & vbCrLf
    s = s & "        re.IgnoreCase = True" & vbCrLf
    s = s & "        re.Global = False" & vbCrLf
    s = s & "        useFilter = True" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set names = GetLambdaNames(ActiveWorkbook)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    For Each x In names" & vbCrLf
    s = s & "        totalCount = totalCount + 1" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        If Not useFilter Then" & vbCrLf
    s = s & "            lstNames.AddItem CStr(x)" & vbCrLf
    s = s & "        ElseIf re.Test(CStr(x)) Then" & vbCrLf
    s = s & "            lstNames.AddItem CStr(x)" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next x" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If useFilter Then" & vbCrLf
    s = s & "        lblStatus.Caption = CStr(lstNames.ListCount) & "" of "" & CStr(totalCount) & "" LAMBDA function(s) shown.""" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        lblStatus.Caption = CStr(lstNames.ListCount) & "" LAMBDA function(s) found.""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    mLoading = False" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "FilterError:" & vbCrLf
    s = s & "    lblStatus.Caption = ""Invalid regex filter: "" & Err.Description" & vbCrLf
    s = s & "    mLoading = False" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
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
    s = s & "Private Sub txtFilter_Change()" & vbCrLf
    s = s & "    If mLoading Then Exit Sub" & vbCrLf
    s = s & "    RefreshList" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub lstNames_Click()" & vbCrLf
    s = s & "    HandleListSelection" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub lstNames_Change()" & vbCrLf
    s = s & "    HandleListSelection" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub HandleListSelection()" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim selectedCount As Long" & vbCrLf
    s = s & "    Dim selectedIndex As Long" & vbCrLf
    s = s & "    Dim selectedName As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If mLoading Then Exit Sub" & vbCrLf
    s = s & "    If mHandlingListSelection Then Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    mHandlingListSelection = True" & vbCrLf
    s = s & "    On Error GoTo CleanUp" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    selectedIndex = -1" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    For i = 0 To lstNames.ListCount - 1" & vbCrLf
    s = s & "        If lstNames.Selected(i) Then" & vbCrLf
    s = s & "            selectedCount = selectedCount + 1" & vbCrLf
    s = s & "            If selectedIndex = -1 Then selectedIndex = i" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If selectedCount = 0 Then GoTo CleanUp" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If selectedCount > 1 Then" & vbCrLf
    s = s & "        lblStatus.Caption = CStr(selectedCount) & "" LAMBDA functions selected.""" & vbCrLf
    s = s & "        GoTo CleanUp" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    selectedName = CStr(lstNames.List(selectedIndex))" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If StrComp(selectedName, mCurrentName, vbTextCompare) = 0 And Not mDirty Then GoTo CleanUp" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Not ConfirmDiscardChanges Then GoTo CleanUp" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    LoadLambda selectedName" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "CleanUp:" & vbCrLf
    s = s & "    mHandlingListSelection = False" & vbCrLf
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
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdDelete_Click()" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim count As Long" & vbCrLf
    s = s & "    Dim lambdaName As String" & vbCrLf
    s = s & "    Dim msg As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    For i = 0 To lstNames.ListCount - 1" & vbCrLf
    s = s & "        If lstNames.Selected(i) Then count = count + 1" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If count = 0 Then" & vbCrLf
    s = s & "        lambdaName = CleanNameText(txtName.Text)" & vbCrLf
    s = s & "        If Len(lambdaName) = 0 Then Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        msg = ""Delete "" & lambdaName & ""?""" & vbCrLf
    s = s & "        If MsgBox(msg, vbQuestion + vbYesNo, ""Delete LAMBDA"") <> vbYes Then Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        On Error GoTo FailSingle" & vbCrLf
    s = s & "        DeleteLambdaName ActiveWorkbook, lambdaName" & vbCrLf
    s = s & "        RefreshList" & vbCrLf
    s = s & "        ClearEditor" & vbCrLf
    s = s & "        lblStatus.Caption = ""Deleted "" & lambdaName" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If count = 1 Then" & vbCrLf
    s = s & "        For i = 0 To lstNames.ListCount - 1" & vbCrLf
    s = s & "            If lstNames.Selected(i) Then" & vbCrLf
    s = s & "                lambdaName = CStr(lstNames.List(i))" & vbCrLf
    s = s & "                Exit For" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        Next i" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        If MsgBox(""Delete "" & lambdaName & ""?"", vbQuestion + vbYesNo, ""Delete LAMBDA"") <> vbYes Then Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        On Error GoTo FailSingle" & vbCrLf
    s = s & "        DeleteLambdaName ActiveWorkbook, lambdaName" & vbCrLf
    s = s & "        RefreshList" & vbCrLf
    s = s & "        ClearEditor" & vbCrLf
    s = s & "        lblStatus.Caption = ""Deleted "" & lambdaName" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' First confirmation for true batch deletes." & vbCrLf
    s = s & "    If MsgBox(""You are about to delete "" & CStr(count) & "" LAMBDA functions. Continue?"", _" & vbCrLf
    s = s & "              vbExclamation + vbYesNo, ""Confirm batch delete"") <> vbYes Then Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' Second stronger confirmation for true batch deletes." & vbCrLf
    s = s & "    If MsgBox(""This action cannot be undone. Delete all selected functions?"", _" & vbCrLf
    s = s & "              vbCritical + vbYesNo, ""Final confirmation"") <> vbYes Then Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo FailMulti" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    For i = lstNames.ListCount - 1 To 0 Step -1" & vbCrLf
    s = s & "        If lstNames.Selected(i) Then" & vbCrLf
    s = s & "            DeleteLambdaName ActiveWorkbook, CStr(lstNames.List(i))" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    RefreshList" & vbCrLf
    s = s & "    ClearEditor" & vbCrLf
    s = s & "    lblStatus.Caption = ""Deleted "" & CStr(count) & "" LAMBDA functions.""" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "FailSingle:" & vbCrLf
    s = s & "    lblStatus.Caption = Err.Description" & vbCrLf
    s = s & "    MsgBox Err.Description, vbExclamation, ""Delete failed""" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "FailMulti:" & vbCrLf
    s = s & "    lblStatus.Caption = Err.Description" & vbCrLf
    s = s & "    MsgBox Err.Description, vbExclamation, ""Batch delete failed""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "" & vbCrLf
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
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdMinify_Click()" & vbCrLf
    s = s & "    On Error GoTo Fail" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    txtFormula.Text = MinifyLambdaDefinition(txtFormula.Text, True)" & vbCrLf
    s = s & "    txtFormula.SelStart = Len(txtFormula.Text)" & vbCrLf
    s = s & "    mDirty = True" & vbCrLf
    s = s & "    lblStatus.Caption = ""Minified formula. Review, test, then save.""" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Fail:" & vbCrLf
    s = s & "    lblStatus.Caption = Err.Description" & vbCrLf
    s = s & "    MsgBox Err.Description, vbExclamation, ""Minify failed""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdVisualize_Click()" & vbCrLf
    s = s & "    On Error GoTo Fail" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    OpenFormulaBoostVisualization txtFormula.Text" & vbCrLf
    s = s & "    lblStatus.Caption = ""Opened FormulaBoost visualization.""" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Fail:" & vbCrLf
    s = s & "    lblStatus.Caption = Err.Description" & vbCrLf
    s = s & "    MsgBox Err.Description, vbExclamation, ""Visualize failed""" & vbCrLf
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
    s = s & "    ElseIf Shift = 2 And KeyCode = vbKeyM Then" & vbCrLf
    s = s & "        KeyCode = 0" & vbCrLf
    s = s & "        cmdMinify_Click" & vbCrLf
    s = s & "    ElseIf Shift = 3 And KeyCode = vbKeyV Then" & vbCrLf
    s = s & "        KeyCode = 0" & vbCrLf
    s = s & "        cmdVisualize_Click" & vbCrLf
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
    s = s & "" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdImportFile_Click()" & vbCrLf
    s = s & "    On Error GoTo Fail" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If mDirty Then SaveCurrentLambda" & vbCrLf
    s = s & "    ImportLambdasFromTextFile ActiveWorkbook" & vbCrLf
    s = s & "    RefreshList" & vbCrLf
    s = s & "    lblStatus.Caption = ""Import from file completed.""" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Fail:" & vbCrLf
    s = s & "    lblStatus.Caption = Err.Description" & vbCrLf
    s = s & "    MsgBox Err.Description, vbExclamation, ""Import from file failed""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdImportUrl_Click()" & vbCrLf
    s = s & "    On Error GoTo Fail" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If mDirty Then SaveCurrentLambda" & vbCrLf
    s = s & "    ImportLambdasFromUrl ActiveWorkbook" & vbCrLf
    s = s & "    RefreshList" & vbCrLf
    s = s & "    lblStatus.Caption = ""Import from URL completed.""" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Fail:" & vbCrLf
    s = s & "    lblStatus.Caption = Err.Description" & vbCrLf
    s = s & "    MsgBox Err.Description, vbExclamation, ""Import from URL failed""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdExport_Click()" & vbCrLf
    s = s & "    On Error GoTo Fail" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If mDirty Then SaveCurrentLambda" & vbCrLf
    s = s & "    ExportLambdasToTextFile ActiveWorkbook" & vbCrLf
    s = s & "    lblStatus.Caption = ""Export completed.""" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Fail:" & vbCrLf
    s = s & "    lblStatus.Caption = Err.Description" & vbCrLf
    s = s & "    MsgBox Err.Description, vbExclamation, ""Export failed""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    Code_frmLambdaEditor = s
End Function
