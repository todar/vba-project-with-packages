Attribute VB_Name = "RecordsetLib"
'@Folder("RecordsetLib")
'@IgnoreModule FunctionReturnValueDiscarded, FunctionReturnValueAlwaysDiscarded, ProcedureNotUsed, UseMeaningfulName

''
' @ref {Microsoft ActiveX Data Objects 6.1 Library} ADODB.Recordset
''
Option Explicit

' Easy way to get an ADODB.Recordset to an Excel.Range
' @ref Microsoft ActiveX Data Objects 6.1 Library
Public Function RecordsetToRange(ByVal rs As ADODB.Recordset, ByRef startingCell As Range) As Range
    ' Add the headers fields first
    Dim index As Long
    For index = 0 To rs.fields.count - 1
        startingCell.Offset(0, index).value = rs.fields.item(index).name
    Next index
    
    ' Add data below the headers fields
    startingCell.Offset(1, 0).CopyFromRecordset rs
    
    ' return the range that was created from the recordset
    Set RecordsetToRange = startingCell.CurrentRegion
End Function

' Returns a recordset from an Excel.Range object
' @ref Microsoft ActiveX Data Objects 6.1 Library
' @ref Microsoft XML, v6.0
Public Function RangeToRecordset(ByVal target As Range) As ADODB.Recordset
    Set RangeToRecordset = New ADODB.Recordset
    
    Dim doc As MSXML2.DomDocument60
    Set doc = New MSXML2.DomDocument60
    doc.LoadXML target.value(xlRangeValueMSPersistXML)
    RangeToRecordset.Open doc
End Function
