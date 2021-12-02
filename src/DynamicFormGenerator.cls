VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DynamicFormGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type State
    form As DynamicForm
    padding As Long
    offsetHeight As Long
End Type

Private WithEvents emitter As EventListenerEmitter
Attribute emitter.VB_VarHelpID = -1

Private this As State

Private Sub Class_Initialize()
    Set this.form = New DynamicForm
    Set emitter = New EventListenerEmitter
    emitter.AddEventListenerAll this.form
    this.padding = 20
End Sub

Public Sub add(ByVal controlType As MSFormControls)
    ' Create the new control
    Dim control As MSForms.control
    Set control = AddFormControl(this.form, controlType)
    
    ' Set the top location of the new control.
    control.Top = this.offsetHeight + this.padding
    
    ' Update the offset height for the next control.
    this.offsetHeight = this.offsetHeight + control.height + this.padding
    
    ' Update the width of the control
    control.width = this.form.InsideWidth - (this.padding * 2)
    control.Left = this.padding
    
    emitter.AddEventListener control
End Sub

Public Sub AddTextbox(ByVal PlaceHolder As String, ByVal validationPattern As String)
    Dim inputLabel As MSForms.Label
    Set inputLabel = AddFormControl(this.form, Label)
    With inputLabel
        .Caption = PlaceHolder
        .Top = this.offsetHeight + this.padding
        
        this.offsetHeight = this.offsetHeight + .height
        .width = this.form.InsideWidth - (this.padding * 2)
        .Left = this.padding
    End With
    
    Dim inputTextbox As MSForms.Textbox
    Set inputTextbox = AddFormControl(this.form, Textbox)
    With inputTextbox
        .Top = this.offsetHeight + this.padding
        this.offsetHeight = this.offsetHeight + .height + this.padding
        
        .width = this.form.InsideWidth - (this.padding * 2)
        .Left = this.padding
    End With
    
    emitter.AddEventListener inputLabel
    emitter.AddEventListener inputTextbox
End Sub

Public Sub ShowForm()
    If (this.offsetHeight + (this.padding * 2)) > this.form.height Then
        this.form.ScrollBars = fmScrollBarsVertical
        this.form.ScrollHeight = this.offsetHeight + this.padding
    End If
    this.form.Show
End Sub

Public Property Get form() As DynamicForm
    Set form = this.form
End Property

Private Sub Emitter_CommandButtonMouseOut(commandButton As MSForms.commandButton)
    commandButton.BackColor = 8435998
End Sub

Private Sub Emitter_CommandButtonMouseOver(commandButton As MSForms.commandButton)
    commandButton.BackColor = 9029664
End Sub

Private Sub Emitter_TextboxMouseOut(Textbox As MSForms.Textbox)
    Textbox.BackColor = 8435998
End Sub

Private Sub Emitter_TextboxMouseOver(Textbox As MSForms.Textbox)
    Textbox.BackColor = 9029664
End Sub
