Attribute VB_Name = "StringLib"
'@Folder("StringLib")
'@IgnoreModule UnreachableCase, ProcedureNotUsed, UseMeaningfulName

''
' This is a utility library for string functions.
'
' @author Robert Todar <robert@robertodar.com> <https://github.com/todar>
' @licence MIT
' @ref {Microsoft Scripting Runtime} Scripting.Dictionary
' @ref {Microsoft VBScript Regular Expressions 5.5} [RegExp, Match]
''
Option Explicit
Option Compare Text

''
' This returns a percentage of how similar two strings are using the levenshtein formula.
'
' @author Robert Todar <robert@robertodar.com>
' @example StringSimilarity("Test", "Tester") ->  66.6666666666667
''
Public Function StringSimilarity(ByVal firstString As String, ByVal secondString As String) As Double
    ' Levenshtein is the distance between two sequences
    Dim levenshtein As Double
    levenshtein = LevenshteinDistance(firstString, secondString)
    
    ' Convert levenshtein into a percentage(0 to 100)
    StringSimilarity = (1 - (levenshtein / Application.Max(Len(firstString), Len(secondString)))) * 100
End Function

''
' Levenshtein is the distance between two sequences of words.
'
' @author Robert Todar <robert@robertodar.com>
' @see <https://www.cuelogic.com/blog/the-levenshtein-algorithm>
' @example LevenshteinDistance("Test", "Tester") ->  2
''
Public Function LevenshteinDistance(ByVal firstString As String, ByVal secondString As String) As Double
    Dim firstLength As Long
    firstLength = Len(firstString)

    Dim secondLength As Long
    secondLength = Len(secondString)
    
    ' Prepare distance array matrix with the proper indexes
    Dim distance() As Long
    ReDim distance(firstLength, secondLength)
    
    Dim index As Long
    For index = 0 To firstLength
        distance(index, 0) = index
    Next
    
    Dim InnerIndex As Long
    For InnerIndex = 0 To secondLength
        distance(0, InnerIndex) = InnerIndex
    Next
    
    ' Outer loop is for the first string
    For index = 1 To firstLength

        ' Inner loop is for the second string
        For InnerIndex = 1 To secondLength

            ' Character matches exactly
            If Mid$(firstString, index, 1) = Mid$(secondString, InnerIndex, 1) Then
                distance(index, InnerIndex) = distance(index - 1, InnerIndex - 1)
            
            ' Character is off, offset the matrix by the appropriate number.
            Else
                Dim min1 As Long
                min1 = distance(index - 1, InnerIndex) + 1

                Dim min2 As Long
                min2 = distance(index, InnerIndex - 1) + 1

                If min2 < min1 Then
                    min1 = min2
                End If
                min2 = distance(index - 1, InnerIndex - 1) + 1
    
                If min2 < min1 Then
                    min1 = min2
                End If
                distance(index, InnerIndex) = min1

            End If
        Next
    Next
    
    ' Levenshtein is the last index of the array.
    LevenshteinDistance = distance(firstLength, secondLength)
End Function

''
' Returns a new cloned string that replaced special {keys} with its associated pair value.
' Keys can be anything since it goes off of the index, so variables must be in proper order!
' Can't have whitespace in the key.
' Also Replaces "\t" with VbTab and "\n" with VbNewLine
'
' @author Robert Todar <robert@roberttodar.com>
' @ref {Microsoft Scripting Runtime} Scripting.Dictionary
' @ref {Microsoft VBScript Regular Expressions 5.5} [RegExp, Match]
' @example Inject("Hello, {name}!\nJS Object = {name: {name}, age: {age}}\n", "Robert", 31)
''
Public Function Inject(ByVal Source As String, ParamArray values() As Variant) As String
    ' Want to get a copy and not mutate original
    Inject = Source
    
    Dim regex As RegExp
    Set regex = New RegExp ' Late Binding would be: CreateObject("vbscript.regexp")
    With regex
        .Global = True
        .MultiLine = True
        .IgnoreCase = True

        ' This section is only when user passes in variables
        If Not IsMissing(values) Then

            ' Looking for pattern like: {key}
            ' First capture group is the full pattern: {key}
            ' Second capture group is just the name:    key
            .pattern = "(?:^|[^\\])(\{([A-z�-�0-9\s]*)\})"

            ' Used to make sure there are even number of uniqueKeys and values.
            Dim keys As Scripting.Dictionary
            Set keys = New Scripting.Dictionary

            Dim keyMatch As Match
            For Each keyMatch In .Execute(Inject)

                ' Extract key name
                Dim key As Variant
                key = keyMatch.SubMatches.item(1)

                ' Only want to increment on unique keys.
                If Not keys.Exists(key) Then

                    If (keys.count) > UBound(values) Then
                        Err.Raise 9, "Inject", "Inject expects an equal amount of keys to values. Keys found: " & Join(keys.keys, ", ") & ", " & key
                    End If

                    ' Replace {key} with the pairing value.
                    Inject = Replace(Inject, keyMatch.SubMatches.item(0), values(keys.count))

                    ' Add key to make sure it isn't looped again.
                    keys.add key, vbNullString

               End If
            Next
        End If

        ' Replace extra special characters. Must allow code above to run first!
        .pattern = "(^|[^\\])\{"
        Inject = .Replace(Inject, "$1" & "{")

        .pattern = "(^|[^\\])\\t"
        Inject = .Replace(Inject, "$1" & vbTab)

        .pattern = "(^|[^\\])\\n"
        Inject = .Replace(Inject, "$1" & vbNewLine)

        .pattern = "(^|[^\\])\\"
        Inject = .Replace(Inject, "$1" & vbNullString)
    End With
End Function

''
' Create a max lenght of string and return it with extension.
'
' @author Robert Todar <robert@roberttodar.com>
' @example Truncate("This is a long sentence", 10)  -> "This is..."
''
Public Function Truncate(ByRef Source As String, ByRef maxLength As Long) As String
    If Len(Source) <= maxLength Then
        Truncate = Source
        Exit Function
    End If
    
    Const extention As String = "..."
    Source = Left$(Source, maxLength - Len(extention)) & extention
    Truncate = Source
End Function

''
' Find string between two words.
'
' @author Robert Todar <robert@roberttodar.com>
' @example StringBetween("Robert Paul Todar", "Robert", "Todar")  -> "Paul"
''
Public Function StringBetween(ByVal main As String, ByVal between1 As String, Optional ByVal between2 As String) As String
    Dim startIndex As Long
    startIndex = InStr(main, between1) + Len(between1)
    
    Dim endIndex As Long
    endIndex = IIf(between2 = vbNullString, Len(main) + 1, InStr(startIndex, main, between2))
    
    StringBetween = Trim$(Mid$(main, startIndex, endIndex - startIndex))
End Function

''
' Returns a string with the proper padding on each side.
'
' @author Robert Todar <robert@roberttodar.com>
' @example StringPadding("1001", 6, "0", True) -> "100100"
''
Public Function StringPadding(ByVal value As String, ByVal length As Long, ByVal fillValue As String, Optional ByRef afterString As Boolean = True) As String
    Dim localValue As String
    localValue = value
    
    ' Insure infinite loop doesn't occur due to an empty string.
    Dim localFillValue As String
    '@Ignore AssignmentNotUsed
    localFillValue = IIf(fillValue = vbNullString, " ", fillValue)
    
    If Len(localValue) >= length Then
        localFillValue = Left$(localValue, length)
    Else
        ' Add extra value
        Do While Len(localValue) < length
            localValue = IIf(afterString, localValue & localFillValue, localFillValue & localValue)
        Loop
    End If
    StringPadding = localValue
End Function


''
' Created this and helper functions to easily read different containers.
'
' @author Robert Todar <robert@roberttodar.com>
''
Public Function ToString(ByVal Source As Variant) As String
    Const delimiter As String = ", "
    
    Select Case True
        Case TypeName(Source) = "Dictionary"
            ToString = toStingDictionary(Source, delimiter)
        
        Case TypeName(Source) = "Collection"
            ToString = toStingCollection(Source, delimiter)
        
        Case isArrayDimension(Source, 1)
            ToString = toStingSingleDimArray(Source, delimiter)
        
        Case isArrayDimension(Source, 2)
            ToString = toStingTwoDimArray(Source, delimiter)
        
        Case IsObject(Source)
            ToString = TypeName(Source) & " {}"
            
        Case TypeName(Source) = "String"
            ToString = """" & Replace(Replace(Source, vbNewLine, "\n"), vbTab, "\t") & """"
        
        Case IsNull(Source)
            ToString = vbNullString
        
        Case Else 'IsNumeric(source), TypeName(source) = "Boolean"
            ToString = Source
            
    End Select
End Function

''
' Helper function to add lines for ToString()
''
Private Function AddLineIfNeeded(ByVal Source As Variant) As String
        If TypeName(Source) = "Dictionary" _
            Or TypeName(Source) = "Collection" _
            Or isArrayDimension(Source, 1) _
            Or isArrayDimension(Source, 2) Then
            
            AddLineIfNeeded = vbNewLine & "  "
        End If
End Function

''
' Dictionary as a string
''
Private Function toStingDictionary(ByVal Source As Scripting.Dictionary, ByVal delimiter As String) As String
    toStingDictionary = "{"
    
    Dim key As Variant
    For Each key In Source.keys
        toStingDictionary = toStingDictionary & AddLineIfNeeded(Source.item(key)) & """" & key & """" & ": " & ToString(Source.item(key)) & delimiter
    Next key
    toStingDictionary = Left$(toStingDictionary, Len(toStingDictionary) - Len(delimiter)) & IIf(InStr(toStingDictionary, vbNewLine), vbNewLine, vbNullString) & "}"
End Function

''
' Collection as a string
''
Private Function toStingCollection(ByVal Source As Collection, ByVal delimiter As String) As String
    toStingCollection = "{"
    
    Dim item As Variant
    For Each item In Source
        toStingCollection = toStingCollection & AddLineIfNeeded(item) & ToString(item) & delimiter
    Next item
    toStingCollection = Left$(toStingCollection, Len(toStingCollection) - Len(delimiter)) & IIf(InStr(toStingCollection, vbNewLine), vbNewLine, vbNullString) & "}"
End Function

''
' Single Array as a string
''
Private Function toStingSingleDimArray(ByVal Source As Variant, ByVal delimiter As String) As String
    toStingSingleDimArray = "["
    
    Dim index As Long
    For index = LBound(Source) To UBound(Source)
        toStingSingleDimArray = toStingSingleDimArray & AddLineIfNeeded(Source(index)) & ToString(Source(index)) & IIf(index < UBound(Source), delimiter, vbNullString)
    Next index
    toStingSingleDimArray = toStingSingleDimArray & IIf(InStr(toStingSingleDimArray, vbNewLine), vbNewLine, vbNullString) & "]"
End Function

''
' Two Dim Array as a string
''
Private Function toStingTwoDimArray(ByVal Source As Variant, ByVal delimiter As String) As String
    toStingTwoDimArray = "[" & vbNewLine
    Dim rowIndex As Long
    For rowIndex = LBound(Source) To UBound(Source)
        toStingTwoDimArray = toStingTwoDimArray & "  ["
        
        ' Add row elements to the string
        Dim colIndex As Long
        For colIndex = LBound(Source, 2) To UBound(Source, 2)
            toStingTwoDimArray = toStingTwoDimArray & ToString(Source(rowIndex, colIndex)) & IIf(colIndex < UBound(Source, 2), delimiter, vbNullString)
        Next colIndex
        
        toStingTwoDimArray = toStingTwoDimArray & "]" & IIf(rowIndex < UBound(Source), "," & vbNewLine, vbNullString)
    Next rowIndex
    toStingTwoDimArray = toStingTwoDimArray & vbNewLine & "]"
End Function

''
' Helper function to see if array is two dim or single
''
Public Function isArrayDimension(ByVal Source As Variant, ByVal dimension As Long) As Boolean
    If IsArray(Source) Then
        isArrayDimension = (dimension = ArrayDimensionLength(Source))
    End If
End Function

''
' Helper function to get the legnth of the dimension of an array.
''
Private Function ArrayDimensionLength(ByVal sourceArray As Variant) As Long
    ' Run loop until error. Remove one and it gives the array dimension =)
    On Error GoTo catch
    Do
        Dim iterator As Long
        iterator = iterator + 1
        
        '@Ignore VariableNotUsed
        Dim test As Long
        test = UBound(sourceArray, iterator)
    Loop
catch:
    On Error GoTo 0
    ArrayDimensionLength = iterator - 1
End Function

''
' Returns the number of elements in an Array.
''
Private Function arrayLength(ByRef Source As Variant) As Long
    arrayLength = UBound(Source) - LBound(Source) + 1
End Function


