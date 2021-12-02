Attribute VB_Name = "UserformLib_Tests"
Option Explicit

Sub Test_BasicDynamicUserform()
    With New DynamicForm
        AddFormControl .Self, Textbox, "Test", True
        .Show
    End With
End Sub

Sub Test_DemoAddingControlsToDynamicForm()
    With CreateForm("Test Controls")
        .AddTextbox "First Name", ""
        .AddTextbox "Last Name", ""
        .AddTextbox "Age Name", ""
    
        .add commandButton
        
        .ShowForm
    End With
End Sub
