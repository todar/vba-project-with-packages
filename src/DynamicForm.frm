VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DynamicForm 
   Caption         =   "Form"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "DynamicForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DynamicForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Type State
    IsCancelled As Boolean
End Type

Private this As State

' Return current implementation of this form.
Public Property Get Self() As DynamicForm
    Set Self = Me
End Property

' Property to be used to see if the user cancelled the form.
Public Property Get IsCancelled() As Boolean
    IsCancelled = this.IsCancelled
End Property

' Listen for a cancel click by the user.
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub

' Call for when the user cancels.
Public Sub OnCancel()
    this.IsCancelled = True
    Me.Hide
End Sub
