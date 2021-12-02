Attribute VB_Name = "DateTimeLib"
Option Explicit

' Stop code execution for specified milliseconds
#If VBA7 And Win64 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' Check to see if two timeframes overlap at all at a minimum of a single day.
' @example: TimeFramesOverlap("01/01/2021", "02/01/2021", "02/01/2021", "02/05/2021") -> True
Public Function TimeFramesOverlap(startDateOne As Date, endDateOne As Date, startDateTwo As Date, endDateTwo As Date) As Boolean
    TimeFramesOverlap = Not (startDateOne > endDateTwo Or endDateOne < startDateTwo)
End Function

' Test to see if a given year is a leap year
Public Function IsLeapYear(targetYear As Integer) As Boolean
    IsLeapYear = (Day(DateSerial(targetYear, 3, 0)) = 29)
End Function
