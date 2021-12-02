Attribute VB_Name = "InternetExplorerLib"
''
' A library for interacting with Internet Explorer
' @ref {Microsoft Internet Controls} InternetExplorer
''
Option Explicit

' A simple factory for creating an Internet Explorer Object
' @author: Robert Todar <robert@roberttodar.com>
' @example: Set IE = NewInternetExplorer("www.google.com")
' @ref: Microsoft Internet Controls
Public Function NewInternetExplorer(ByVal URL As String, Optional ByVal Visible As Boolean = True) As InternetExplorer
    Set NewInternetExplorer = CreateObject("InternetExplorer.Application")
    With NewInternetExplorer
        .Visible = Visible
        .Navigate URL
    End With
    
    WaitForInternetExplorer NewInternetExplorer
End Function

' Returns Instance of IE, Looks to find first matching URL
' @author: Robert Todar <robert@roberttodar.com>
' @example: Set IE = GetOpenInternetExplorer("www.google.com")
' @ref: Microsoft Internet Controls
Public Function GetOpenInternetExplorer(ByVal URL As String) As InternetExplorer
    Dim Window As Object
    For Each Window In CreateObject("Shell.Application").Windows
        If InStr(Window.LocationURL, URL) > 0 Then
            Set GetOpenInternetExplorer = Window
            Exit Function
        End If
    Next Window
End Function

' Waits for Internet Explorer to have readystate 4 and not busy
' @ref: Microsoft Internet Controls
Public Sub WaitForInternetExplorer(ByRef IE As InternetExplorer)
    While IE.ReadyState <> 4 Or IE.Busy: DoEvents: Wend
End Sub

' Simple way to execute scripts in IE.
' @ref: Microsoft Internet Controls
Public Sub InjectJavascript(ByVal IE As InternetExplorer, ByVal Code As String)
    IE.Document.parentWindow.execScript Code:=Code
End Sub
