VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AwSnapForm 
   Caption         =   "Marks Fault"
   ClientHeight    =   7125
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8730
   OleObjectBlob   =   "AwSnapForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AwSnapForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Property Let ErrorMessage(ByVal value As String)
    Me.ErrorMessageText.Caption = IIf(value <> vbNullString, "Unknown Error!", value)
End Property

Private Sub OkButton_Click()
    Me.Hide
End Sub
