VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormValidatorControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type State
    control As MSForms.control
    pattern As String
    helperText As String
End Type

Private this As State

Public Sub Create(ByVal control As MSForms.control, ByVal pattern As String, Optional ByVal helperText As String)
    Set this.control = control
    this.pattern = pattern
    this.helperText = helperText
End Sub

Public Function isValid() As Boolean
    isValid = RegexPatternMatches(this.pattern, this.control.value)
End Function

Public Property Get helperText() As String
    helperText = this.helperText
End Property

Public Property Get control() As MSForms.control
    Set control = this.control
End Property
