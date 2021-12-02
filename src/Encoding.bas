Attribute VB_Name = "Encoding"
'@Folder("Encoding")
'@IgnoreModule ProcedureNotUsed

''
' Functions to help with encoding and decoding
'
' @author Robert Todar <robert@roberttodar.com>
' @ref {Microsoft Scripting Runtime} Scripting.Dictionary
''
Option Explicit

' Encode test to base64
'@Ignore UseMeaningfulName
Public Function EncodeBase64(ByVal text As String) As String
    Dim bytes() As Byte
    bytes = StrConv(text, vbFromUnicode)
    
    Dim domDocument As MSXML2.DomDocument60
    Dim xmlElement As MSXML2.IXMLDOMElement
    
    Set domDocument = New MSXML2.DomDocument60
    Set xmlElement = domDocument.createElement("b64")
    
    xmlElement.dataType = "bin.base64"
    xmlElement.nodeTypedValue = bytes
    EncodeBase64 = xmlElement.text
    
    Set xmlElement = Nothing
    Set domDocument = Nothing
End Function
