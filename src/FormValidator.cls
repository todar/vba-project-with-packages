VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type State
    Controls As Collection
End Type

Private this As State

Private Sub Class_Initialize()
    Set this.Controls = New Collection
End Sub

Public Function isValid() As Boolean
    Dim control As FormValidatorControl
    For Each control In this.Controls
        If Not control.isValid Then
            isValid = False
            Exit Function
        End If
    Next control
    
    isValid = True
End Function

Public Function controlIsValid(ByVal controlToValidate As MSForms.control) As Boolean
    Dim control As FormValidatorControl
    For Each control In this.Controls
        If control.control = controlToValidate Then
            controlIsValid = control.isValid
            Exit Function
        End If
    Next control
End Function

Public Sub AddTextbox(ByVal Textbox As MSForms.Textbox, ByVal RegExPattern As String, Optional ByVal helperText As String)
    Dim formControl As New FormValidatorControl
    formControl.Create Textbox, RegExPattern, helperText
    this.Controls.add formControl
End Sub

Public Property Get Controls() As Collection
    Set Controls = this.Controls
End Property
