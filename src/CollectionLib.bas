Attribute VB_Name = "CollectionLib"
'@Folder("CollectionLibrary")
'@IgnoreModule ProcedureNotUsed

''
' Functions to work with collections.
'
' @author Robert Todar <robert@roberttodar.com>
''
Option Explicit

' Return a collections values as a string (better to use toString method from StringLib)
Public Function CollectionToString(ByVal target As Collection, Optional ByVal delimeter As String = ", ") As String
    Dim value As Variant
    For Each value In target
        CollectionToString = IIf(CollectionToString = vbNullString, value, CollectionToString & delimeter & value)
    Next value
End Function

' Check to see if a collection has a specified value
Public Function CollectionHasValue(ByRef target As Collection, ByVal value As Variant) As Boolean
    Dim currentValue As Variant
    For Each currentValue In target
        If currentValue = value Then
            CollectionHasValue = True
            Exit Function
        End If
    Next currentValue
End Function
