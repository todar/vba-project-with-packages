VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ENV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type State
    fso As Scripting.FileSystemObject
    lastCached As Date
    data As Scripting.Dictionary
    Source As String
End Type
Private this As State

Private Property Get FILEPATH() As String
    FILEPATH = ThisWorkbook.path & "\.env"
End Property

Private Sub Class_Initialize()
    Set this.fso = New Scripting.FileSystemObject
    Set this.data = New Scripting.Dictionary
    
    If Not this.fso.FileExists(FILEPATH) Then
        this.fso.CreateTextFile(FILEPATH, True).WriteLine "# Auto generated .env file created in Class_Initialize."
    End If
    ParseSourceFile
End Sub

' Main function for returning a value from the config
' file. This requires
Public Property Get use(ByVal key As String, Optional fallback As Variant = vbNullString) As Variant
    ' By default it will check to see if the source
    ' data file has been modified, if so it will parse it again.
    ' Otherwise, the data is cached and will be quick to access.
    ParseSourceFile
    
    ' See that the key actually exists. return that value if so.
    If this.data.Exists(key) Then
        use = this.data(key)
        Exit Property
    End If
    
    ' If the key doesn't exists, then return the fallback value.
    use = fallback
End Property

' Main function for returning a value from the config
' file. This requires
Public Function add(ByVal key As String, ByVal value As Variant) As Boolean
    ' By default it will check to see if the source
    ' data file has been modified, if so it will parse it again.
    ' Otherwise, the data is cached and will be quick to access.
    ParseSourceFile
    
    ' Can't add a key if it already exists. Will update the value???
    If this.data.Exists(key) Then
        this.data(key) = value
        Exit Function
    End If
    
    
    
    ' If the key doesn't exists, then return the fallback value.
    use = fallback
End Function

Public Sub ParseSourceFile()
    ' Don't parse if data is already loaded and the file has not been
    ' modified.
    If this.lastCached = LastModified And Not this.data Is Nothing Then
        Exit Sub
    End If
    
    ' Get the data Object ready to be populated
    Set this.data = New Scripting.Dictionary
    
    ' Get the source file content as a string.
    this.Source = this.fso.OpenTextFile(FILEPATH, ForReading).ReadAll
    
    ' Split the source file based on each newline.
    Dim lines As Variant
    lines = split(this.Source, vbNewLine)
    
    ' Loop each line and parse accordingly
    Dim row As Long
    For row = LBound(lines) To UBound(lines)
        ' Store the current line string into
        ' a variable to make it easier to access.
        Dim line As String
        line = Trim(lines(row))
        
        ' Skip blank lines and comments as they can't be parsed into
        ' anything.
        If line = vbNullString Or Left(line, 1) = "#" Then
            GoTo Continue
        End If
        
        ' Find the index of the assignment operator,
        ' this seperates the key and the values.
        Dim assign As Long
        assign = InStr(line, "=")
        If assign <= 0 Then
            Debug.Print "Assign = is not found in the current line: " & line
            GoTo Continue
        End If
        
        ' Read the key based on the assignment operator location
        Dim key As String
        key = Trim(Left(line, assign - 1))
        
        ' Skip if the key already exists as this must have
        ' been done in error.
        If this.data.Exists(key) Then
            Debug.Print key & " already exists in " & FILEPATH
            GoTo Continue
        End If
        
        ' Get the value based on the assignment operator.
        Dim StringValue As String
        StringValue = Trim(Mid(line, assign + 1))
        
        ' Parse the value based on what type it should be.
        Dim value As Variant
        value = parseStringValue(StringValue)
        
        ' Finally add the Key and Value to the data dictionary.
        this.data.add key, value
        this.lastCached = LastModified
        
        ' Move to the next line of text content until all
        ' lines have been read and parsed.
Continue:
    Next row
End Sub

' Helper function. The values are by default a string,
' this helps determine what data type it should be and returns
' that value accordingly.
Private Function parseStringValue(ByVal value As String) As Variant
    On Error GoTo catch
    parseStringValue = value
    Select Case True
        Case IsNumeric(value) And InStr(value, ".") > 0
            parseStringValue = CDbl(value)
            
        Case IsNumeric(value)
            parseStringValue = CLng(value)
            
        Case LCase(value) = "true", LCase(value) = "false"
            parseStringValue = CBool(value)
            
        Case Left(value, 1) = """", Left(value, 1) = "'"
            parseStringValue = Mid(value, 2, Len(value) - 2)
            ' For now do not inject with strings, until
            ' this is proper
            'parseStringValue = Replace(parseStringValue, "\t", vbTab)
            'parseStringValue = Replace(parseStringValue, "\n", vbNewLine)
    End Select
catch:
End Function

' Get the last date the file was modified
' this is used for checking if the cached data is still
' accurate.
Public Property Get LastModified() As Date
    On Error Resume Next
    LastModified = this.fso.GetFile(FILEPATH).DateLastModified
End Property

' Print the raw text file into the immediate window.
Public Sub DisplayRawContent()
    Debug.Print this.Source
End Sub

' Print the parsed data into the immediate window.
Public Sub List()
    ParseSourceFile
    Dim index As Long
    For index = 0 To this.data.count - 1
        Debug.Print index, this.data.keys(index), this.data.items(index), TypeName(this.data.items(index))
    Next index
End Sub

' Print parsed values ONLY into the immediate window.
Public Sub ListValues()
    ParseSourceFile
    Dim index As Long
    For index = 0 To this.data.count - 1
        Debug.Print this.data.items(index)
    Next index
End Sub

' Print parsed keys ONLY into the immediate window.
Public Sub ListKeys()
    ParseSourceFile
    Dim index As Long
    For index = 0 To this.data.count - 1
        Debug.Print this.data.keys(index)
    Next index
End Sub
