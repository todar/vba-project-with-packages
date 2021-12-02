Attribute VB_Name = "RangeLib"
'@Folder("RangeLib")
'@IgnoreModule FunctionReturnValueDiscarded, FunctionReturnValueAlwaysDiscarded, ProcedureNotUsed
Option Explicit

' Helper function that returns a collection of fields that are missing from the
' provided `requiredFields` array.
Public Function MissingFields(ByRef fields As Range, ByRef requiredFields() As String) As Collection
    Set MissingFields = New Collection
    
    Dim index As Long
    For index = LBound(requiredFields) To UBound(requiredFields)
        If fields.Find(requiredFields(index), , , xlWhole) Is Nothing Then
            MissingFields.add requiredFields(index)
        End If
    Next index
End Function

' Helper function to see if a specified range contains a value
Public Function RangeContains(ByVal target As Range, ByVal value As String) As Boolean
    Dim cell As Range
    For Each cell In target.Cells
        If cell.value Like value Then
            RangeContains = True
            Exit Function
        End If
    Next cell
End Function

''
' Finds a range based on it's value.
' This works faster than `Range.Find()` as it loops an array instead of cells.
' This also works for hidden cells where `Range.Find` does not.
'
' Note, this looks for first match, and is case sensitive by defaut, unless
' Option Match Case is used at the top of the module it is stored in.
'
' @author Robert Todar <robert@roberttodar.com>
''
Public Function FindFast(searchRange As Range, what As String) As Range
    ' Get data from range into an Array. Looping Arrays is much
    ' faster than looping cells.
    Dim data As Variant
    data = searchRange.value
    
    ' Loop every row in the array.
    Dim rowIndex As Long
    For rowIndex = LBound(data, 1) To UBound(data, 1)
        
        ' Loop every column in the array.
        Dim ColumnIndex As Long
        For ColumnIndex = LBound(data, 2) To UBound(data, 2)
        
            ' If current row/column matches the correct value then return the range.
            If data(rowIndex, ColumnIndex) Like what Then
                Set FindFast = searchRange.Cells(rowIndex, ColumnIndex)
                Exit Function
            End If
        Next ColumnIndex
    Next rowIndex
    
    ' If the range is not found then `Nothing` is returned.
    Set FindFast = Nothing
End Function

''
' Finds the next empty row on a worksheet.
'
' @author Robert Todar <robert@roberttodar.com>
''
Public Function NextAvailibleRow(ByRef ws As Worksheet) As Range
On Error GoTo catch
    Set NextAvailibleRow = ws.Cells.Find("*", , xlValues, , xlRows, xlPrevious).Offset(1).EntireRow
Exit Function
    ' If there is an error, that means the worksheet is empty.
    ' Return the first row
catch:
    Set NextAvailibleRow = ws.Rows(1)
End Function

''
' Gets the intersect between a colum and a header.
' Helpful for updating/create content to a table of data.
'
' @author Robert Todar <robert@roberttodar.com>
''
Function GetCell(row As Range, headerRow As Range, header As String) As Range
    Set GetCell = Intersect(row, FindFast(headerRow, header).EntireColumn)
End Function

