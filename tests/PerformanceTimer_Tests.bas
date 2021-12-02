Attribute VB_Name = "PerformanceTimer_Tests"
'@Folder("PerformanceTimer")
'@IgnoreModule FunctionReturnValueDiscarded, ProcedureNotUsed
Option Explicit
Option Private Module

Public Declare Sub Sleep Lib "kernel32" _
(ByVal dwMilliseconds As Long)

Public Sub TEST_PerformanceTimer()
    Dim timer As PerformanceTimer
    Set timer = New PerformanceTimer
    timer.LogResults = True
    
    Sleep 1000
    timer.Mark "a"
    Dim index As Long
    Do While index < 10000
        index = index + 1
    Loop
    
    Sleep 1000
    timer.Mark "b"
    
    timer.MeasureMarks "a", "b"
End Sub


