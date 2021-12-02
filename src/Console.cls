VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Console"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This is a global class that makes it easy to log to both the
' immediate window and log files.
' The log adds date and time stamps and structures the log
' into three categories and requires a source of the log.
'
' Update `LogFolders` property if you desire logging
' to different locations other then the default.
'
' This Class Module must be Imported as it is created Globally
' with: `Attribute VB_PredeclaredId = True`
'
' @author Robert Todar<robert@roberttodar.com>
' @ref {Microsoft Scripting Runtime} Scripting.FileSystemObject
Option Explicit

' This is the private state that is used within this
' Class Module, and can be accessed using `this`.
Private Type State
    fso As Scripting.FileSystemObject
End Type
Private this As State

' Each log is catergorized into these three categories for
' ease of view of what the intent of the log is.
Public Enum LogType
    ErrorLog
    Warn
    Info
End Enum

' This is a list of where the log files should
' be written to. Add/remove folders as desired.
Private Property Get LogFolders() As Collection
    Set LogFolders = New Collection
    With LogFolders
        ' By default this will create a log folder in the
        ' current directory and save there. Add/remove as
        ' desired per application.
        .add this.fso.BuildPath(ThisWorkbook.path, "logs")
    End With
End Property

Private Sub Class_Initialize()
    On Error GoTo catch
    ' Get an instance of the filesystem for creating folders
    ' and appending to the log files.
    Set this.fso = New Scripting.FileSystemObject
    
    ' Make sure all log folders are created
    Exit Sub
catch:
    Debug.Print "Console ERR! The Console expereinced an issue initializing and will have issues logging to the appropriate files."
    Debug.Print Err.description, Err.Source, Err.Number
End Sub

' Log a message to all of the appropriate log
' files/immediate window.
Public Sub Log(ByVal typeOfLog As LogType, ByVal message As String, ByVal moduleName As String, ByVal functionName As String)
    ' Must format message to remove any tabs or newlines.
    message = Replace(message, vbTab, "\t")
    message = Replace(message, vbNewLine, "\n")
    
    ' This is the final format of the message and how it will be
    ' displayed. Change this if another format is desired.
    message = format(Date, "MM-DD-YYYY") _
              & vbTab & format(Now, "hh:mm:ss") _
              & vbTab & useLogTypeName(typeOfLog) _
              & vbTab & Environ("Username") _
              & vbTab & moduleName & "." & functionName _
              & vbTab & message
    
    ' Log the message to the immediate window.
    Debug.Print message
    
    ' Log the message to the appropriate log files.
    Dim folderPath As Variant
    For Each folderPath In LogFolders
        If makeDir(folderPath) = False Then
            Debug.Print "Console ERR! Unable to create and write log to: " & folderPath
        Else
            Dim FILEPATH As String
            FILEPATH = this.fso.BuildPath(folderPath, useFileName)
            appendToFile FILEPATH, message
        End If
    Next folderPath
End Sub

' Convert the LogType Enum values into a string
Private Function useLogTypeName(ByVal typeOfLog As LogType) As String
    useLogTypeName = Switch(typeOfLog = ErrorLog, "Error" _
                          , typeOfLog = Info, "Info" _
                          , typeOfLog = Warn, "Warn")
End Function

' The filename will be by year and month for ease of indexing.
Private Property Get useFileName() As String
    useFileName = format(Date, "YYYY_MM") & ".log"
End Property

' Will append to the text file.
Private Sub appendToFile(ByVal FILEPATH As String, ByVal message As String)
    On Error GoTo catch
    this.fso.OpenTextFile(FILEPATH, ForAppending, True).WriteLine message
    Exit Sub
catch:
    Debug.Print "Console ERR! Not able to log to: " & FILEPATH
    Debug.Print Err.description, Err.Source, Err.Number
End Sub

' Recursivly make each directory in a folderpath
' Returns false with any errors.
Private Function makeDir(ByVal folderPath As String) As Boolean
    On Error GoTo catch

    Dim fso As New FileSystemObject
    Dim directories() As String
    directories = split(folderPath, "\")

    Dim index As Long
    For index = LBound(directories) To UBound(directories)
        ' The current working directory will change
        ' on each loop. This will be to recursivily create
        ' each folder directory one at a time.
        Dim cwd As String
        cwd = cwd & directories(index) & "\"

        If Not fso.FolderExists(cwd) Then
            fso.CreateFolder cwd
        End If
    Next index

    makeDir = True
    Exit Function
catch:
    'RETURNS FALSE
End Function
