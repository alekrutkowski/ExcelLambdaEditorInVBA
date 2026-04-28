Attribute VB_Name = "modLambdaEditor"
Option Explicit

Public Sub ShowLambdaEditor()
    Dim f As Object

    Set f = VBA.UserForms.Add("UserForm1")
    f.Show vbModeless
End Sub

Public Sub InstallLambdaEditorShortcut()
    Application.OnKey "^+l", "ShowLambdaEditor"
End Sub

Public Sub RemoveLambdaEditorShortcut()
    Application.OnKey "^+l"
End Sub
