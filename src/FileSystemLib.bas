Attribute VB_Name = "FileSystemLib"
''
' Functions for working with FileSytem operations
'
' @author: Robert Todar <robert@roberttodar.com>
' @Licence: MIT
' @ref {Microsoft Scripting Runtime} Scripting.FileSystemObject
''
Option Explicit
Option Compare Text
Option Private Module

' OPENS FILES WITH THEIR DEFAULT APPLICATION
Public Declare Function ShellExecute Lib "shell32.dll" _
Alias "ShellExecuteA" ( _
ByVal hwnd As Long, _
ByVal lpOperation As String, _
ByVal lpFile As String, _
ByVal lpParameters As String, _
ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long

''
' Check to see if current user has write access to a specific folder.
' This works by creating a temp file in the folder and attempting to write
' to it.
'
' @author Robert Todar <robert@roberttodar.com>
' @example HasWriteAccessToFolder("C:\Program Files") -> True || False
''
Public Function HasWriteAccessToFolder(ByVal folderPath As String) As Boolean
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    
    ' If folder doesn't exist, then user does not have access to write to it.
    If Not fso.FolderExists(folderPath) Then
        HasWriteAccessToFolder = False
        Exit Function
    End If

    ' Loop is for making sure to get a unique temp filepath so that this function
    ' doesn't overwrite something that already exists.
    Dim FILEPATH As String
    Dim count As Long
    Do
        FILEPATH = fso.BuildPath(folderPath, "TEST_WRITE_ACCESS" & count & ".tmp")
        count = count + 1
    Loop Until Not fso.FileExists(FILEPATH)
    
    'ATTEMPT TO CREATE THE TMP FILE, ERROR RETURNS FALSE
On Error GoTo catch
    fso.CreateTextFile(FILEPATH).Write ("Test Folder Access")
    Kill FILEPATH
    
    'NO ERROR, ABLE TO WRITE TO FILE; RETURN TRUE!
    HasWriteAccessToFolder = True
catch:
End Function

''
' CALL TO CREATE FILEPATH AND WRITE TO TEXT FILE
' @author: ROBERT TODAR
' @desc: CREATED FOR EASE OF USE AS THIS CODE IS REPEATED A LOT.
''
Public Sub WriteToTextFile(ByVal FILEPATH As String, ByVal value As String)
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    
    If Not fso.FileExists(FILEPATH) Then
        CreateFilePath FILEPATH
    End If
    
    Dim ts As TextStream
    Set ts = fso.OpenTextFile(FILEPATH, ForWriting, True)
    
    ts.Write value
End Sub

''
' READ ANY TEXT FILE, EMPTY STRING IF FILE DOES NOT EXIST
' @author: ROBERT TODAR
''
Public Function ReadTextFile(ByVal FILEPATH As String) As String
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    
On Error GoTo NoFile
    Dim ts As TextStream
    Set ts = fso.OpenTextFile(FILEPATH, ForReading, False)
    ReadTextFile = ts.ReadAll
    
Exit Function
NoFile:
    'FOR MY NEEDS, ERROR JUST RETURNS AND EMPTY STRING. ADJUST AS NEED FOR OTHER CODE.
End Function

''
' CALL TO CREATE FILEPATH AND APPEND TO TEXT FILE
' @author: ROBERT TODAR
''
Public Sub AppendTextFile(ByVal FILEPATH As String, ByVal value As String)
    If CreateFilePath(FILEPATH) = False Then
        Exit Sub
    End If
    
    On Error GoTo catch
    Dim ts As TextStream
    Set ts = useFile(FILEPATH, 1).OpenAsTextStream(ForAppending)
    ts.WriteLine value
    
    Exit Sub
catch:
    ' Attempt to open text file
    ' up to max attempts.
    ' This is used as multiple users could
    ' by appending a text file at the same time.
    Const MAXATTEMPTS As Long = 20
    Dim attempts As Long
    attempts = attempts + 1
    If attempts <= MAXATTEMPTS Then
        ' Wait a random amount of time
        ' before trying again.
        Sleep Rnd * 10
        
        ' Try to open connection to text file
        ' another time.
        Resume
    End If
    
    ' Only reach this if max attempts exceeded.
    Console.Error Err.description, "FileSystemFunctions.AppendTextFile"
End Sub

' Check to see if file is in use by another user
' or possiblly even another function.
Public Function FileInUse(sFileName As String) As Boolean
    On Error Resume Next
    Open sFileName For Binary Access Read Lock Read As #1
    Close #1
    FileInUse = IIf(Err.Number > 0, True, False)
    On Error GoTo 0
End Function

' Return a file object.
' Also will optionally wait for the file
' to not be in use by another user/function.
Public Function useFile(ByVal FILEPATH As String, Optional maxseconds As Double = 2) As file
    Dim fso As New Scripting.FileSystemObject
    Set useFile = fso.GetFile(FILEPATH)
    
    Dim startTime As Double
    startTime = timer
    Do While (timer - startTime) <= maxseconds
        If Not FileInUse(FILEPATH) Then
            Exit Function
        End If
    Loop
End Function

''
' CREATES FULL PATH. NORMAL CREATE FOLDER OR FILE ONLY DOES ONE LEVEL.
' @author: ROBERT TODAR
''
Public Function CreateFilePath(ByVal FullPath As String) As Boolean
    Dim paths() As String
    paths = split(FullPath, "\")
    
    Dim PathIndex As Integer
    For PathIndex = LBound(paths, 1) To UBound(paths, 1) - 1
        Dim currentPath As String
        currentPath = currentPath & paths(PathIndex) & "\"
        
        Dim fso As New Scripting.FileSystemObject
        If Not fso.FolderExists(currentPath) Then
            fso.CreateFolder currentPath
        End If
    Next PathIndex
    
    ' Create file
    If Not fso.FileExists(FullPath) Then
        fso.CreateTextFile FullPath
    End If
    
    CreateFilePath = fso.FileExists(FullPath)
End Function

''
' EASY WAY TO SEE IF FILE EXISTS
' @author: ROBERT TODAR
''
Public Function FileExists(ByVal FileSpec As String) As Boolean
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    
    FileExists = fso.FileExists(FileSpec)
End Function

''
' EASY WAY TO SEE IF FOLDER EXISTS
' @author: ROBERT TODAR
''
Public Function FolderExists(ByVal FileSpec As String) As Boolean
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    
    FolderExists = fso.FolderExists(FileSpec)
End Function

''
' EASY WAY TO CREATE A FOLDER
' @author: ROBERT TODAR
''
Public Sub CreateFolder(ByVal folderPath As String)
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
End Sub

''
' EASY WAY TO DELETE A FOLDER
' @author: ROBERT TODAR
''
Public Sub DeleteFolder(ByVal folderPath As String)
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    
    If fso.FolderExists(folderPath) Then
        fso.DeleteFolder folderPath, True
    End If
End Sub

''
' EASY WAY TO DELETE A FILE
' @author: ROBERT TODAR
''
Public Sub DeleteFile(ByVal FILEPATH As String)
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    
    If fso.FileExists(FILEPATH) Then
        fso.DeleteFile FILEPATH, True
    End If
End Sub

''
' EASY WAY TO MOVE A FILE
' @author: ROBERT TODAR
''
Public Sub MoveFile(ByVal Source As String, ByVal destination As String)
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    
    fso.MoveFile Source, destination
End Sub

Public Function getFileName(ByVal path As String) As String
    Dim fso As New Scripting.FileSystemObject
    getFileName = fso.getFileName(path)
End Function

Public Sub CopyFileOverwrite(ByVal Source As String, ByVal destination As String)
    Dim fso As New Scripting.FileSystemObject
    fso.CopyFile Source, destination, True
End Sub

''
' CHECKS TO SEE IF FILE EXISTS, THEN OPENS IT IF IT DOES
' @author: ROBERT TODAR
''
Public Function OpenAnyFile(ByVal FileToOpen As String) As Boolean
    'WILL ONLY OPEN FILE IF IT EXISTS
    If FileExists(FileToOpen) Then
        OpenAnyFile = True
        
        'API FUNCTION FOR OPENING FILES
        Call ShellExecute(0, "Open", FileToOpen & vbNullString, _
        vbNullString, vbNullString, 1)
    End If
End Function

''
' CHECKS TO SEE IF FOLDER EXISTS, THEN OPENS WINDOWS EXPLORER TO THAT PATH
' @author: ROBERT TODAR
''
Public Function OpenFileExplorer(ByVal folderPath As String) As Boolean
    If FolderExists(folderPath) Then
        OpenFileExplorer = True
        Call Shell("explorer.exe " & Chr(34) & folderPath & Chr(34), vbNormalFocus)
    End If
End Function

''
' OPEN URL IN DEFAULT BROWSER
' @author: ROBERT TODAR
''
Public Sub OpenURL(ByVal UrlToOpen As String)
    'API FUNCTION FOR OPENING FILES
    Call ShellExecute(0, "Open", UrlToOpen & vbNullString, _
    vbNullString, vbNullString, 1)
End Sub


' Easily see when a file was last updated.
Public Function FileLastUpdated(ByVal FILEPATH As String) As Date
On Error GoTo catch
    Dim fso As New Scripting.FileSystemObject
    FileLastUpdated = fso.GetFile(FILEPATH).DateLastModified
    Exit Function
catch:
    Console.Error "FileSystemLibrary.FileLastUpdated", Err.description
End Function


