Attribute VB_Name = "UserformLib"
Option Explicit

' List of all the MSForms Controls.
Public Enum MSFormControls
    CheckBox
    ComboBox
    commandButton
    Frame
    Image
    Label
    ListBox
    MultiPage
    OptionButton
    ScrollBar
    SpinButton
    TabStrip
    Textbox
    ToggleButton
End Enum

' Gets the string name of each control.
Public Function GetUserformControlType(control As MSFormControls) As String
    Select Case control
      Case MSFormControls.CheckBox:       GetUserformControlType = "CheckBox"
      Case MSFormControls.ComboBox:       GetUserformControlType = "ComboBox"
      Case MSFormControls.commandButton:  GetUserformControlType = "CommandButton"
      Case MSFormControls.Frame:          GetUserformControlType = "Frame"
      Case MSFormControls.Image:          GetUserformControlType = "Image"
      Case MSFormControls.Label:          GetUserformControlType = "Label"
      Case MSFormControls.ListBox:        GetUserformControlType = "ListBox"
      Case MSFormControls.MultiPage:      GetUserformControlType = "MultiPage"
      Case MSFormControls.OptionButton:   GetUserformControlType = "OptionButton"
      Case MSFormControls.ScrollBar:      GetUserformControlType = "ScrollBar"
      Case MSFormControls.SpinButton:     GetUserformControlType = "SpinButton"
      Case MSFormControls.TabStrip:       GetUserformControlType = "TabStrip"
      Case MSFormControls.Textbox:        GetUserformControlType = "Textbox"
      Case MSFormControls.ToggleButton:   GetUserformControlType = "ToggleButton"
    End Select
End Function

' Gets the ProgID for each individual control. Used to create controls using `Object.add` method.
' @see https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/add-method-microsoft-forms
Public Function GetMSFormsProgID(control As MSFormControls) As String
    Select Case control
      Case MSFormControls.CheckBox:       GetMSFormsProgID = "Forms.CheckBox.1"
      Case MSFormControls.ComboBox:       GetMSFormsProgID = "Forms.ComboBox.1"
      Case MSFormControls.commandButton:  GetMSFormsProgID = "Forms.CommandButton.1"
      Case MSFormControls.Frame:          GetMSFormsProgID = "Forms.Frame.1"
      Case MSFormControls.Image:          GetMSFormsProgID = "Forms.Image.1"
      Case MSFormControls.Label:          GetMSFormsProgID = "Forms.Label.1"
      Case MSFormControls.ListBox:        GetMSFormsProgID = "Forms.ListBox.1"
      Case MSFormControls.MultiPage:      GetMSFormsProgID = "Forms.MultiPage.1"
      Case MSFormControls.OptionButton:   GetMSFormsProgID = "Forms.OptionButton.1"
      Case MSFormControls.ScrollBar:      GetMSFormsProgID = "Forms.ScrollBar.1"
      Case MSFormControls.SpinButton:     GetMSFormsProgID = "Forms.SpinButton.1"
      Case MSFormControls.TabStrip:       GetMSFormsProgID = "Forms.TabStrip.1"
      Case MSFormControls.Textbox:        GetMSFormsProgID = "Forms.TextBox.1"
      Case MSFormControls.ToggleButton:   GetMSFormsProgID = "Forms.ToggleButton.1"
    End Select
End Function

' Easly add control to userform or a frame.
' @returns {MSForms.control} The control that was created
Public Function AddFormControl(userformOrFrame As Object _
                         , control As MSFormControls _
                         , Optional name As String = vbNullString _
                         , Optional visable As Boolean = True _
                        ) As MSForms.control
    Set AddFormControl = userformOrFrame.Controls.add(GetMSFormsProgID(control), name, visable)
End Function

' Function to create a dynamic form. This is still being tested.
Public Function CreateForm(ByVal Caption As String, Optional width As Long = 300, Optional height As Long = 300) As DynamicFormGenerator
    Set CreateForm = New DynamicFormGenerator
    With CreateForm.form
        .Caption = Caption
        
        .width = width
        .height = height
    End With
End Function

' Selects All the Text in a textbox
Public Function TextboxWordSelect(Textbox As MSForms.Textbox) As Boolean
    Textbox.SelStart = 0
    Textbox.SelLength = Len(Textbox.text)
End Function
