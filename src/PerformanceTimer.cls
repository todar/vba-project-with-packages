VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PerformanceTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("PerformanceTimer")
'@IgnoreModule ProcedureNotUsed, UnassignedVariableUsage, VariableNotAssigned

''
' Performance timer that is accurate to the millisecond.
' This is a good tool for tracking code performance or used with animations.
'
' @author Robert Todar <robert@roberttodar.com>
''
Option Explicit

''
' API's that are needed to get accurate time measurements.
''
#If VBA7 Then
    Private Declare PtrSafe Function getFrequency Lib "kernel32" Alias _
    "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
    Private Declare PtrSafe Function getTickCount Lib "kernel32" Alias _
    "QueryPerformanceCounter" (cyTickCount As Currency) As Long
#Else
    Private Declare Function getFrequency Lib "kernel32" Alias _
    "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
    Private Declare Function getTickCount Lib "kernel32" Alias _
    "QueryPerformanceCounter" (cyTickCount As Currency) As Long
#End If

''
' State is a {Type} containing all the private variables for this class.
''
Private Type State
    frequency As Currency
    startTime As Double
    Marks As Collection
    LogResults As Boolean
End Type

''
' property {state} this - The private properties stored in one place.
''
Private this As State

Private Sub Class_Initialize()
    Set this.Marks = New Collection
    startTimer
End Sub

''
' Starts the timer. This is called on initialize event.
''
Private Sub startTimer()
    getFrequency this.frequency
    this.startTime = tick / this.frequency
End Sub

''
' Needed to get accuarate time.
''
Private Property Get tick() As Currency
    Dim firstTicks As Currency
    getTickCount firstTicks
    
    Dim secondTicks As Currency
    getTickCount secondTicks
    
    If secondTicks < firstTicks Then
        tick = firstTicks
    Else
        tick = secondTicks
    End If
End Property

''
' @property {Double} Now - Returns current milliseconds.
''
Public Property Get Now() As Double
    Now = ((tick / this.frequency) - this.startTime) * 1000
End Property

''
' Setting option to log results on every mark & measure.
''
Public Property Let LogResults(ByVal value As Boolean)
    this.LogResults = value
End Property

''
' Read property Setting option to log results on every mark & measure.
''
Public Property Get LogResults() As Boolean
    LogResults = this.LogResults
End Property

''
' Logs a time and assigns it to a specific key.
''
Public Function Mark(ByVal key As String) As Double
    this.Marks.add Me.Now, key
    Mark = this.Marks.item(key)
    If this.LogResults Then Debug.Print "Mark " & key & ": " & Mark
End Function

''
' Way to compare two marks to one another.
''
Public Function MeasureMarks(ByVal firstKey As String, ByVal secondKey As String) As Double
    MeasureMarks = Abs(this.Marks.item(firstKey) - this.Marks.item(secondKey))
    If this.LogResults Then Debug.Print "Measure " & firstKey & "->" & secondKey & ": " & MeasureMarks
End Function

''
' Property to get the Marks collection. These are all the logged times.
''
Public Property Get Marks() As Collection
    Set Marks = this.Marks
End Property


