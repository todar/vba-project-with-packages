Attribute VB_Name = "ExcelLib"
Option Explicit
Option Compare Text
Option Private Module

'/**
' * CHECKS TO SEE IF WORKSHEET NAME EXISTS IN A WORKBOOK (DISPLAY NAME ONLY)
' *
' * @author: Robert Todar <robert.todar@albertsons.com>
' */
Public Function WorksheetExists(ByVal name As String, Optional ByRef SourceWorkbook As Workbook) As Boolean
    If SourceWorkbook Is Nothing Then
        Set SourceWorkbook = ActiveWorkbook
    End If
    
    Dim ws As Worksheet
    For Each ws In SourceWorkbook.Worksheets
        If ws.name = name Then
            WorksheetExists = True
            Exit Function
        End If
    Next ws
End Function

'/**
' * Checks to see if workbook is open.
' *
' * @author: Robert Todar <robert.todar@albertsons.com>
' */
Public Function WorkbookIsOpen(ByVal fullName As String) As Boolean
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If wb.fullName = fullName Then
            WorkbookIsOpen = True
        End If
    Next wb
End Function
