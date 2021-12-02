Attribute VB_Name = "StringLib_Tests"
'@Folder("StringLib")
'@IgnoreModule ProcedureNotUsed
Option Explicit

''
' Tests
' @author Robert Todar <robert@robertodar.com>
''
Private Sub testsForStringFunctions()
    Debug.Print StringSimilarity("Test", "Tester")                     '->  66.6666666666667
    Debug.Print LevenshteinDistance("Test", "Tester")                  '->  2
    Debug.Print Truncate("This is a long sentence", 10)                '-> "This is..."
    Debug.Print StringBetween("Robert Paul Todar", "Robert", "Todar")  '-> "Paul"
    Debug.Print StringPadding("1001", 6, "0", True)                    '-> "100100"
    Debug.Print Inject("Hello,\nMy name is {Name} and I am {Age}!", "Robert", 31)
        '-> Hello,
        '-> My name is Robert and I am 30!
End Sub

''
' Tests for ToString Function.
''
Private Sub testToStringFunction()
    ' Test values
    Debug.Print ToString("String")
    Debug.Print ToString(31)
    Debug.Print ToString(True)
    
    ' Test Null
    Debug.Print ToString(Null)
    
    ' Test Array
    Debug.Print ToString(Array(1, 2, 3, 4))
    
    ' Test collections
    Dim col As Collection
    Set col = New Collection
    col.add "item", "key"
    col.add "item2", "key2"
    Debug.Print ToString(col)
    
    ' Test Dictionary
    Dim Dic As Scripting.Dictionary
    Set Dic = New Scripting.Dictionary
    Dic.add "Name", "Robert"
    Dic.add "Age", 31
    Debug.Print ToString(Dic)
    
    ' Test objects
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    Debug.Print ToString(fso)
End Sub



