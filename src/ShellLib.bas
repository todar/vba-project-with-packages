Attribute VB_Name = "ShellLib"
'@Folder("ShellLib")
'@IgnoreModule ProcedureNotUsed
''
' Functions to work with command prompts.
' Currently this just has windows shell, but could also implement bash, powershell, etc.
'
' @author Robert Todar <robert@roberttodar.com>
''
Option Explicit

' Helper function to run scripts from the root directory.
Public Function CommandPrompt(ByRef script As String, _
                    Optional ByRef keepCommandWindowOpen As Boolean = False, _
                    Optional ByRef cdToCurrentDirectory As Boolean, _
                    Optional ByRef WindowStyle As VbAppWinStyle = vbMinimizedFocus _
                    ) As Double
    ' cmd.exe Opens the command prompt.
    ' /S      Modifies the treatment of string after /C or /K (see below)
    ' /C      Carries out the command specified by string and then terminates
    ' /K      Carries out the command specified by string but remains
    ' cd      Change directory to the root directory.
    CommandPrompt = Shell("cmd.exe /S /" & _
                 IIf(keepCommandWindowOpen, "K", "C") & _
                 IIf(cdToCurrentDirectory, " cd " & ThisWorkbook.path & " && ", vbNullString) & _
                 script _
                 , WindowStyle _
                )
End Function


