Attribute VB_Name = "Path_Tests"
''
' Tests for the Path Module.
'
' @author Robert Todar <robert@roberttodar.com>
' @ref {Module} Path
''
Option Explicit

' Basic tests for Joining path segments.
' Note, this will need to be adjusted once Path.Normalize is developed.
Public Function testJoin()
    ' Once Normalized is developed expected result
    ' should be: "\foo\bar\baz\asdf\"
    Debug.Assert path.Join("/foo", "bar", "baz/asdf", "quux", "..") = "\foo\bar\baz\asdf\quux\.."
    
    ' Once Normalized is developed expected result
    ' should be: "C:\dev\vba-git\src\tests\test.txt"
    Debug.Assert path.Join(Dirname, "//src/", "\\tests\\", "test.txt") = "C:\dev\vba-git\\src\\tests\\test.txt"
    
    Debug.Print "Tests Passed"
End Function
