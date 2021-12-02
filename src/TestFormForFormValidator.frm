VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TestFormForFormValidator 
   Caption         =   "Form Validator Test Form"
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5775
   OleObjectBlob   =   "TestFormForFormValidator.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TestFormForFormValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private validator As FormValidator

Private Sub EmailTextbox_Change()
    EmailTextbox.BorderColor = vbRed
    EmailTextbox.BorderStyle = fmBorderStyleNone
    EmailTextbox.BorderStyle = fmBorderStyleSingle
End Sub

Private Sub RunValidationButton_Click()
    If ValidateForm = True Then
        MsgBox "Form is valid!"
    End If
End Sub

Private Sub UserForm_Initialize()
    Set validator = New FormValidator
    validator.AddTextbox EmailTextbox, REGEX_PATTERN_EMAIL, "Must be valid email"
    validator.AddTextbox PhoneNumberTextbox, REGEX_PATTERN_PHONE_NUMBER, "Must be valid Phone Number"
End Sub

Private Function ValidateForm() As Boolean
    ValidateForm = True
    
    Dim control As FormValidatorControl
    For Each control In validator.Controls
        If Not control.isValid Then
            control.control.BorderColor = vbRed
            control.control.BorderStyle = fmBorderStyleNone
            control.control.BorderStyle = fmBorderStyleSingle
            
            
            MsgBox "Control " & control.control.name & " is not valid. " & control.helperText
            ValidateForm = False
        End If
    Next control
End Function
