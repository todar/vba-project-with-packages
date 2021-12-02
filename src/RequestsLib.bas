Attribute VB_Name = "RequestsLib"
Option Explicit

''
' Sample of how to make requests. This might be able to turn into more useful functions.
'
' @ref {Microsoft XML, v6.0}
' @example:
'   ```vba
'   FetchSample("https://www.google.com").status ' -> 200
'   ```
''
Public Function FetchSample(ByVal URL As String) As MSXML2.XMLHTTP60
    Set FetchSample = New MSXML2.XMLHTTP60
    With FetchSample
        .Open "GET", URL
        .setRequestHeader "Content-Type", "text/html"
        '.setRequestHeader "Authorization", "Basic Username:Password"
        .send
    End With
End Function
