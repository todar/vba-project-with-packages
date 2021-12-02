VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoginForm 
   Caption         =   "Login Credentials"
   ClientHeight    =   3645
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "LoginForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type State
    isCanceled As Boolean
End Type

Private this As State

Private Sub UserForm_Initialize()
    'UsernameTextbox.value = LCase(Environ("Username"))
End Sub

Public Property Get Password() As String
    Password = PasswordTextbox.text
End Property

Public Property Get Username() As String
    Username = UsernameTextbox.text
End Property

Public Property Get IsCancelled() As Boolean
    IsCancelled = this.isCanceled
End Property

Private Sub ShowPasswordCheckbox_Click()
    PasswordTextbox.PasswordChar = IIf(ShowPasswordCheckbox.value, vbNullString, "•")
    PasswordTextbox.SetFocus
End Sub

Private Sub SubmitButton_Click()
    Me.Hide
End Sub

Private Sub CancelButton_Click()
    OnCancel
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub

Private Sub OnCancel()
    this.isCanceled = True
    Me.Hide
End Sub
