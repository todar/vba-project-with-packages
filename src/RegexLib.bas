Attribute VB_Name = "RegexLib"
''
' This is a library of functions an constant for working with regular expressions.
'
' @ref  {Microsoft VBScript Regular Expressions 5.5}
''
Option Explicit

' This is a good place to store common patterns.
Public Const REGEX_PATTERN_EMAIL As String = "^[_a-z0-9-]+(.[a-z0-9-]+)@[a-z0-9-]+(.[a-z0-9-]+)*(.[a-z]{2,4})$"
Public Const REGEX_PATTERN_PHONE_NUMBER As String = "^\s*(?:\+?(\d{1,3}))?[-. (]*(\d{3})[-. )]*(\d{3})[-. ]*(\d{4})(?: *x(\d+))?\s*$"

' Function to see if a string matches a given pattern.
Public Function RegexPatternMatches(ByVal pattern As String, ByVal sourceString) As Boolean
    Dim re As New RegExp
    With re
        .pattern = pattern
        RegexPatternMatches = .test(sourceString)
    End With
End Function
