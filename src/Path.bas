Attribute VB_Name = "Path"
''
' This module is created to ease working with paths. Main goals
' will to be builing paths, normalizing paths, and working from
' relative paths.
'
' @author Robert Todar <robert@roberttodar.com>
' @ref {Microsoft Scripting Runtime} Scripting.FileSystemObject
''
Option Explicit
Private Const CURRENT_WORKING_DIRECTORY = "."

' The path.join() method joins all given path segments together
' and (next statment is not developed) normalizes the resulting path.
' @status Limited Production
'   Normalize is not yet in working production
'   so it simply returns the path as of now. This shouldn't cause
'   any issues for basic usage, just don't use relative paths and
'   know excess seperators will remain.
'
' @see https://nodejs.org/api/path.html#path_path_join_paths
Public Function Join(ParamArray paths() As Variant) As String
    Dim fso As New Scripting.FileSystemObject
    Dim index As Long
    For index = LBound(paths) To UBound(paths)
        Join = fso.BuildPath(Join, Replace(paths(index), "/", "\"))
    Next
    
    Join = path.Normalize(Join)
End Function

' Resolved relative path '..' and '.' segments.
' Also will remove excess seperators
' @status Development
Public Function Normalize(ByVal path As String) As String
    ' Zero length string return the current working directory token.
    If path = vbNullString Then
        Normalize = CURRENT_WORKING_DIRECTORY
        Exit Function
    End If
    
    Normalize = path
End Function
