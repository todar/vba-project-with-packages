Attribute VB_Name = "Crypto"
''
' Functions for cryptography. Currently just hasing functions.
''
Option Explicit

' Hash a string to SHA256
Public Function SHA256(ByVal value As String) As String
    Dim bytes() As Byte
    bytes = StrConv(value, vbFromUnicode)
     
    Dim algo As Object
    Set algo = CreateObject("System.Security.Cryptography.SHA256Managed")
     
    Dim buffer() As Byte
    buffer = algo.ComputeHash_2(bytes)
         
    SHA256 = ToHexString(buffer)
End Function

' Convert a buffer array to Hex String.
Public Function ToHexString(ByRef buffer() As Byte) As String
    Dim result As String
    Dim index As Long
    For index = LBound(buffer) To UBound(buffer)
        result = result & Hex(buffer(index))
    Next
    ToHexString = result
End Function
