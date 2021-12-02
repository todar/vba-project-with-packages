Attribute VB_Name = "ArrayLib"
Attribute VB_Description = "A bunch of cool functions for Arrays."
Option Explicit
Option Compare Text
Option Base 0

' ERROR CODES CONSTANTS
Public Const ARRAY_NOT_PASSED_IN As Integer = 5000
Public Const ARRAY_DIMENSION_INCORRECT As Integer = 5001

' PUBLIC FUNCTIONS
' - ArrayAverage
' - ArrayContainsEmpties
' - ArrayDimensionLength
' - ArrayExtractColumn
' - ArrayExtractRow
' - ArrayFilter
' - ArrayFilterTwo
' - ArrayFromRecordset
' - ArrayGetColumnIndex
' - ArrayGetIndexes
' - ArrayIncludes
' - ArrayIndexOf
' - ArrayLength
' - ArrayPluck
' - ArrayPop
' - ArrayPush
' - ArrayPushTwoDim
' - ArrayQuery
' - ArrayRemoveDuplicates
' - ArrayReverse
' - ArrayShift
' - ArraySort
' - ArraySplice
' - ArraySpread
' - ArraySum
' - ArrayToCSVFile
' - ArrayToString
' - ArrayTranspose
' - ArrayUnShift
' - Assign
' - ConvertToArray
' - IsArrayEmpty

' TODO:
' - CLEAN UP CODE! ADD MORE NOTES AND EXAMPLES.
' - NEED TO REALLY TEST ALL OF THESE FUNCTIONS, CHECK FOR ERRORS.
' - ADD MORE CUSTOM ERROR MESSAGES FOR SPECIFIC ERRORS.
'
' - LOOK THROUGH FUNCTIONS DESIGNED FOR SINGLE DIM ARRAYS, SEE IF CAN CONVERT TO WORK
'   WITH 2 DIM AS WELL
'
' - Create ArrayConcat function

'/**
' * EXAMPLES OF VARIOUS FUNCTIONS
' *
' * @author Robert Todar <robert@roberttodar.com>
' */
Private Sub ArrayFunctionExamples()
    Dim a As Variant
    
    ' SINGLE DIM FUNCTIONS TO MANIPULATE
    ArrayPush a, "Banana", "Apple", "Carrot" '--> Banana,Apple,Carrot
    ArrayPop a                               '--> Banana,Apple --> returns Carrot
    ArrayUnShift a, "Mango", "Orange"        '--> Mango,Orange,Banana,Apple
    ArrayShift a                             '--> Orange,Banana,Apple
    ArraySplice a, 2, 0, "Coffee"            '--> Orange,Banana,Coffee,Apple
    ArraySplice a, 0, 1, "Mango", "Coffee"   '--> Mango,Coffee,Banana,Coffee,Apple
    ArrayRemoveDuplicates a                  '--> Mango,Coffee,Banana,Apple
    ArraySort a                              '--> Apple,Banana,Coffee,Mango
    ArrayReverse a                           '--> Mango,Coffee,Banana,Apple
    
    ' ARRAY PROPERTIES
    arrayLength a                            '--> 4
    ArrayIndexOf a, "Coffee"                 '--> 1
    ArrayIncludes a, "Banana"                '--> True
    arrayContains a, Array("Test", "Banana") '--> True
    ArrayContainsEmpties a                   '--> False
    ArrayDimensionLength a                   '--> 1 (single dim array)
    IsArrayEmpty a                           '--> False
    
    ' CAN FLATTEN JAGGED ARRAY WITH SPREAD FORMULA
    ' COULD ALSO SPREAD DICTIONAIRES AND COLLECTIONS AS WELL
    a = Array(1, 2, 3, Array(4, 5, 6, Array(7, 8, 9)))
    a = ArraySpread(a)                       '--> 1,2,3,4,5,6,7,8,9
    
    ' MATH EXAMPLES
    ArraySum a                               '--> 45
    ArrayAverage a                           '--> 5
    
    ' FILTER USE'S REGEX PATTERN
    a = Array("Banana", "Coffee", "Apple", "Carrot", "Canolope")
    a = ArrayFilter(a, "^Ca|^Ap")
    
    ' ARRAY TO STRING WORKS WITH BOTH SINGLE AND DOUBLE DIM ARRAYS!
    Debug.Print ArrayToString(a)
End Sub

'/**
' * TESTER SUB FOR NEW FUNCTIONS
' *
' * @author Robert Todar <robert@roberttodar.com>
' */
Private Sub arrayPlayground()
    Dim arr As Variant
End Sub

'/**
' * Gets the count of rows in a two dim array
' *
' * @autor Robert Todar <robert@roberttodar.com>
' * @param {any[2 dim]} sourceArray
' */
Public Function ArrayRowCount(ByVal sourceArray As Variant) As Long
    On Error GoTo catch
    Dim index As Long
    For index = LBound(sourceArray) To UBound(sourceArray)
        ArrayRowCount = ArrayRowCount + 1
    Next index
catch:
End Function

'/**
' * Removes all non numeric
' *
' * @author Robert Todar <robert@roberttodar.com>
' */
Public Function ArrayRemoveNonNumeric(ByVal sourceArray As Variant) As Variant
    Dim index As Long
    For index = LBound(sourceArray) To UBound(sourceArray)
        If IsNumeric(sourceArray(index)) Then
            ArrayPush ArrayRemoveNonNumeric, sourceArray(index)
        End If
    Next index
End Function

'/**
' * Remove all empty slots
' *
' * @author Robert Todar <robert@roberttodar.com>
' */
Public Function ArrayRemoveEmpties(ByVal sourceArray As Variant) As Variant
    Select Case ArrayDimensionLength(sourceArray)
        Case 1
            Dim index As Long
            For index = LBound(sourceArray) To UBound(sourceArray)
                If Not IsEmpty(sourceArray(index)) Then
                    ArrayPush ArrayRemoveEmpties, sourceArray(index)
                End If
            Next index
        Case 2
            Dim rowIndex As Long
            For rowIndex = LBound(sourceArray, 1) To UBound(sourceArray)
                Dim RowData As Variant
                RowData = ArrayExtractRow(sourceArray, rowIndex)
                
                If Not ArrayIsAllEmpties(RowData) Then
                    ArrayPushTwoDim ArrayRemoveEmpties, RowData
                End If
            Next rowIndex
    End Select
End Function

'/**
' * Returns true if all slots are empty
' *
' * @author Robert Todar <robert@roberttodar.com>
' */
Private Function ArrayIsAllEmpties(ByVal sourceArray As Variant) As Boolean
    Dim index As Long
    For index = LBound(sourceArray) To UBound(sourceArray)
        If Not IsEmpty(sourceArray(index)) Then
            Exit Function
        End If
    Next index
    ArrayIsAllEmpties = True
End Function

'/**
' * FILTER SINGLE DIM ARRAY ELEMENTS BASED ON REGEX PATTERN
' *
' * @author ROBERT TODAR
' * @dim SINGLE DIM ONLY
' * @ref https://regexr.com/
' * @example ArrayFilter(Array("Banana", "Coffee", "Apple"), "^Ba|^Ap") ->  [Banana,Apple]
' */
Public Function ArrayFilter(ByVal sourceArray As Variant, ByVal RegExPattern As String) As Variant
    If ArrayDimensionLength(sourceArray) <> 1 Then
        Err.Raise ARRAY_DIMENSION_INCORRECT, , "SourceArray must be a single dimensional array."
    End If
    
    Dim regex As Object
    Set regex = CreateObject("vbscript.regexp")
    With regex
        .Global = False
        .MultiLine = True
        .IgnoreCase = True
        .pattern = RegExPattern 'SET THE PATTERN THAT WAS PASSED IN
    End With
    
    Dim index As Long
    For index = LBound(sourceArray) To UBound(sourceArray)
        If regex.test(sourceArray(index)) Then
            ArrayPush ArrayFilter, sourceArray(index)
        End If
    Next index
End Function

'/**
' * FILTERS MULTIDIMENSIONAL ARRAY. ARGS ARE PAIR BASED: (HEADING TITLE, REGEX) https://regexr.com/ for help
' *
' * @author ROBERT TODAR
' * @dim TWO DIM ONLY
' * @dependinces: IsValidConditions, ArrayGetConditions, RegExTest
' * @example ArrayFilterTwo(TwoDimArray, "Name", "^R","ID", "\d{6}", ...) can add as many pair args as you'd like
' */
Public Function ArrayFilterTwo(ByVal sourceArray As Variant, ParamArray headerRegexPairs() As Variant) As Variant
    ' THIS FUNCTION IS FOR TWO DIMS ONLY
    If ArrayDimensionLength(sourceArray) <> 2 Then
        Err.Raise ARRAY_DIMENSION_INCORRECT, , "SourceArray must be a two dimensional array."
    End If
    
    ' TODO: SHOULD I ALWAYS RETURN HEADING?? THIS ALSO ASSUMES THERE IS A HEADING...
    ArrayPushTwoDim ArrayFilterTwo, ArrayExtractRow(sourceArray, LBound(sourceArray))
    
    ' GET CONDITIONS JAGGED ARRAY. (HEADING INDEX, AND REGEX CONDITION)
    Dim Conditions As Variant
    Conditions = ArrayGetConditions(sourceArray, headerRegexPairs)
    
    ' CHECK CONDITIONS ON EACH ROW AFTER HEADER
    Dim rowIndex As Integer
    For rowIndex = LBound(sourceArray) + 1 To UBound(sourceArray)
        
        ' some other comment
        If IsValidConditions(sourceArray, Conditions, rowIndex) Then
            ArrayPushTwoDim ArrayFilterTwo, ArrayExtractRow(sourceArray, rowIndex)
        End If
    Next rowIndex
End Function

'/**
' * SUM A SINGLE DIM ARRAY
' *
' * @author ROBERT TODAR
' * @dim SINGLE DIM ONLY
' * @example ArraySum (Array(5, 6, 4, 3, 2)) -> 20
' */
Public Function ArraySum(ByVal sourceArray As Variant) As Double
    'SINGLE DIM ARRAYS ONLY
    If ArrayDimensionLength(sourceArray) <> 1 Then
        Err.Raise ARRAY_DIMENSION_INCORRECT, , "SourceArray must be a 1 dimensional array."
    End If
    
    Dim index As Integer
    For index = LBound(sourceArray, 1) To UBound(sourceArray, 1)
        If Not IsNumeric(sourceArray(index)) Then
            Err.Raise 55, "ArrayFunctions: ArraySum", sourceArray(index) & vbNewLine & "^ Element in Array is not numeric"
        End If
        
        ArraySum = ArraySum + sourceArray(index)
    Next index
End Function

'/**
' * GET AVERAGE OF SINGLE DIM ARRAY
' *
' * @author Robert Todar <robert@roberttodar.com>
' */
Public Function ArrayAverage(ByVal sourceArray As Variant) As Double
    ' SINGLE DIM ARRAYS ONLY
    If ArrayDimensionLength(sourceArray) <> 1 Then
        Err.Raise ARRAY_DIMENSION_INCORRECT, , "SourceArray must be a single dimensional array."
    End If
    
    ArrayAverage = ArraySum(sourceArray) / arrayLength(sourceArray)
End Function

'/**
' * GET LENGTH OF SINGLE DIM ARRAY, REGAURDLESS OF OPTION BASE
' *
' * @author Robert Todar <robert@roberttodar.com>
' */
Public Function arrayLength(ByVal sourceArray As Variant) As Integer
    On Error Resume Next 'empty means 0 lenght
    arrayLength = (UBound(sourceArray, 1) - LBound(sourceArray, 1)) + 1
End Function

'/**
' * SPREADS OUT AN ARRAY INTO A SINGLE ARRAY. EXAMPLE: JAGGED ARRAYS, dictionaries, collections.
' *
' * @author Robert Todar <robert@roberttodar.com>
' */
Public Function ArraySpread(ByVal sourceArray As Variant, Optional SpreadObjects As Boolean = False) As Variant
    ' THIS FUNCTION IS FOR SINGLE DIMS ONLY
    If ArrayDimensionLength(sourceArray) <> 1 Then
        Err.Raise ARRAY_DIMENSION_INCORRECT, , "SourceArray must be a single dimensional array."
    End If
    
    ' CONVERT ANY DICTIONARY OR COLLECTION INTO AN ARRAY FIRST.
    Dim temp As Variant
    temp = ConvertToArray(sourceArray)
    
    Dim index As Integer
    For index = LBound(temp, 1) To UBound(temp, 1)
        
        ' CHECK IF ELEMENT IS AN ARRAY OR OBJECT, RUN RECURSIVE IF SO ON THAT ELEMENT
        If IsArray(temp(index)) Or (IsObject(temp(index)) And SpreadObjects) Then
            
            ' RECURSIVE CALLS UNTIL AT BASE ELEMENTS
            Dim InnerTemp As Variant
            If SpreadObjects Then
                InnerTemp = ArraySpread(ConvertToArray(temp(index)), True)
            Else
                InnerTemp = ArraySpread(temp(index))
            End If
            
            ' ADD EACH ELEMENT TO THE FUNCTION ARRAY
            Dim InnerIndex As Integer
            For InnerIndex = LBound(InnerTemp, 1) To UBound(InnerTemp, 1)
                ArrayPush ArraySpread, InnerTemp(InnerIndex)
            Next InnerIndex
            
        ' ELEMENT IS SINGLE ITEM, SIMPLY TO FUNCTION ARRAY
        Else
            ArrayPush ArraySpread, temp(index)
        End If
    Next index
End Function

'/**
' * RETURNS A SINGLE DIM ARRAY OF THE INDEXES OF COLUMN HEADERS
' * HEADERS NOT FOUND RETURNS EMPTY IN THAT INDEX
' * EXPERIMENTAL CODE PART OF A BIGGER PLAN....
' *
' * @author Robert Todar <robert@roberttodar.com>
' */
Public Function ArrayGetIndexes(ByVal sourceArray As Variant, ByVal IndexArray As Variant) As Variant
    Dim temp As Variant
    ReDim temp(LBound(IndexArray) To UBound(IndexArray))
    
    Dim index As Integer
    For index = LBound(IndexArray) To UBound(IndexArray)
        temp(index) = ArrayGetColumnIndex(sourceArray, IndexArray(index))
        If temp(index) = -1 Then
            temp(index) = Empty
        End If
    Next index
    
    ArrayGetIndexes = temp
End Function

'/**
' * CHECK TO SEE IF SINGLE DIM ARRAY CONTAINS ANY EMPTY INDEXES
' *
' * @author Robert Todar <robert@roberttodar.com>
' */
Public Function ArrayContainsEmpties(ByVal sourceArray As Variant) As Boolean
    ' THIS FUNCTION IS FOR SINGLE DIMS ONLY
    If ArrayDimensionLength(sourceArray) <> 1 Then
        Err.Raise ARRAY_DIMENSION_INCORRECT, , "SourceArray must be a single dimensional array."
    End If
    
    Dim index As Integer
    For index = LBound(sourceArray, 1) To UBound(sourceArray, 1)
        If IsEmpty(sourceArray(index)) Then
            ArrayContainsEmpties = True
            Exit Function
        End If
    Next index
End Function

'/**
' * CHECKS TO SEE IF VALUE IS IN SINGLE DIM ARRAY. VALUE CAN BE SINGLE VALUE OR ARRAY OF VALUES.
' * NEED NOTES....
' *
' * @author Robert Todar <robert@roberttodar.com>
' */
Public Function arrayContains(ByVal sourceArray As Variant, ByVal value As Variant) As Boolean
    If IsArrayEmpty(sourceArray) Then
        Exit Function
    End If
    
    If IsArray(value) Then
        Dim ValueIndex As Long
        For ValueIndex = LBound(value) To UBound(value)
            If arrayContains(sourceArray, value(ValueIndex)) Then
                arrayContains = True
                Exit Function
            End If
        Next ValueIndex
        Exit Function
    End If
    
    Dim index As Long
    For index = LBound(sourceArray, 1) To UBound(sourceArray, 1)
        If sourceArray(index) = value Then
            arrayContains = True
            Exit Function
        End If
    Next index
End Function

'/**
' * CHECK TO SEE IF TWO DIM ARRAY CONTAINS HEADERS STORED IN HEADERS ARRAY
' *
' * @author Robert Todar <robert@roberttodar.com>
' */
Public Function ArrayContainsHeaders(ByVal sourceArray As Variant, ByVal headers As Variant) As Variant
    If Not IsArray(sourceArray) Or ArrayDimensionLength(sourceArray) <> 2 Then
        Err.Raise ARRAY_DIMENSION_INCORRECT, , "SourceArray must be passed in as an two dimensional array"
    End If
    
    If Not IsArray(headers) Or ArrayDimensionLength(headers) <> 1 Then
        Err.Raise ARRAY_DIMENSION_INCORRECT, , "Headers must be passed in as a 1 dimensional array"
    End If
    
    Dim HeaderArray As Variant
    HeaderArray = ArrayExtractRow(sourceArray, LBound(sourceArray, 1))
    
    Dim HedIndex As Integer
    For HedIndex = LBound(headers, 1) To UBound(headers, 1)
        If ArrayIncludes(HeaderArray, headers(HedIndex)) = False Then
            Exit Function
        End If
    Next HedIndex
    
    ArrayContainsHeaders = True
End Function

'/**
' * RETURNS THE LENGHT OF THE DIMENSION OF AN ARRAY
' *
' * @author Robert Todar <robert@roberttodar.com>
' */
Public Function ArrayDimensionLength(sourceArray As Variant) As Integer
    On Error GoTo catch
    Dim i As Integer
    Dim test As Integer
    Do
        i = i + 1
        'WAIT FOR ERROR
        test = UBound(sourceArray, i)
    Loop
catch:
    ArrayDimensionLength = i - 1
End Function

'/**
' * GET A COLUMN FROM A TWO DIM ARRAY, AND RETURN A SINLGE DIM ARRAY
' *
' * @author Robert Todar <robert@roberttodar.com>
' */
Public Function ArrayExtractColumn(ByVal sourceArray As Variant, ByVal ColumnIndex As Integer) As Variant
    ' SINGLE DIM ARRAYS ONLY
    If ArrayDimensionLength(sourceArray) <> 2 Then
        Err.Raise ARRAY_DIMENSION_INCORRECT, , "SourceArray must be a two dimensional array."
    End If
    
    Dim temp As Variant
    ReDim temp(LBound(sourceArray, 1) To UBound(sourceArray, 1))
    
    Dim rowIndex As Integer
    For rowIndex = LBound(sourceArray, 1) To UBound(sourceArray, 1)
        temp(rowIndex) = sourceArray(rowIndex, ColumnIndex)
    Next rowIndex
    
    ArrayExtractColumn = temp
End Function

'/**
' * GET A ROW FROM A TWO DIM ARRAY, AND RETURN A SINLGE DIM ARRAY
' *
' * @author Robert Todar <robert@roberttodar.com>
' */
Public Function ArrayExtractRow(ByVal sourceArray As Variant, ByVal rowIndex As Long) As Variant
    Dim temp As Variant
    ReDim temp(LBound(sourceArray, 2) To UBound(sourceArray, 2))
    
    Dim colIndex As Integer
    For colIndex = LBound(sourceArray, 2) To UBound(sourceArray, 2)
        temp(colIndex) = sourceArray(rowIndex, colIndex)
    Next colIndex
    
    ArrayExtractRow = temp
End Function

'/**
' * RETURNS A 2D ARRAY FROM A RECORDSET, OPTIONALLY INCLUDING HEADERS, AND IT TRANSPOSES TO KEEP
' * ORIGINAL OPTION BASE. (TRANSPOSE WILL SET IT TO BASE 1 AUTOMATICALLY.)
' *
' * @author Robert Todar <robert@roberttodar.com>
' */
Public Function ArrayFromRecordset(rs As Object, Optional IncludeHeaders As Boolean = True) As Variant
    ' @note: -Int(IncludeHeaders) RETURNS A BOOLEAN TO AN INT (0 OR 1)
    Dim HeadingIncrement As Integer
    HeadingIncrement = -Int(IncludeHeaders)
    
    Dim temp As Variant
    Dim HeaderIndex As Long
    
    ' CHECK TO MAKE SURE THERE ARE RECORDS TO PULL FROM
    If rs.BOF Or rs.EOF Then
        ReDim temp(0 To 0, 0 To rs.fields.count - 1)
        For HeaderIndex = 0 To rs.fields.count - 1
            temp(LBound(temp, 1), HeaderIndex) = rs.fields(HeaderIndex).name
        Next HeaderIndex
        
        ArrayFromRecordset = temp
        Exit Function
    End If
    
    ' STORE RS DATA
    Dim RsData As Variant
    RsData = rs.GetRows
    
    ' REDIM TEMP TO ALLOW FOR HEADINGS AS WELL AS DATA
    ReDim temp(LBound(RsData, 2) To UBound(RsData, 2) + HeadingIncrement, LBound(RsData, 1) To UBound(RsData, 1))
        
    If IncludeHeaders = True Then
        ' GET HEADERS
        For HeaderIndex = 0 To rs.fields.count - 1
            temp(LBound(temp, 1), HeaderIndex) = rs.fields(HeaderIndex).name
        Next HeaderIndex
    End If
    
    ' GET DATA
    Dim rowIndex As Long
    Dim colIndex As Long
    For rowIndex = LBound(temp, 1) + HeadingIncrement To UBound(temp, 1)
        For colIndex = LBound(temp, 2) To UBound(temp, 2)
            temp(rowIndex, colIndex) = RsData(colIndex, rowIndex - HeadingIncrement)
        Next colIndex
    Next rowIndex
    
    ' RETURN
    ArrayFromRecordset = temp
End Function

'/**
' * LOOKS FOR VALUE IN FIRST ROW OF A TWO DIMENSIONAL ARRAY, RETURNS IT'S COLUMN INDEX
' *
' * @author Robert Todar <robert@roberttodar.com>
' */
Public Function ArrayGetColumnIndex(ByVal sourceArray As Variant, ByVal HeadingValue As String) As Integer
    Dim ColumnIndex As Integer
    For ColumnIndex = LBound(sourceArray, 2) To UBound(sourceArray, 2)
        If sourceArray(LBound(sourceArray, 1), ColumnIndex) = HeadingValue Then
            ArrayGetColumnIndex = ColumnIndex
            Exit Function
        End If
    Next ColumnIndex
    
    ' RETURN NEGATIVE IF NOT FOUND
    ArrayGetColumnIndex = -1
End Function

'/**
' * CHECKS TO SEE IF VALUE IS IN SINGLE DIM ARRAY
' *
' * @author Robert Todar <robert@roberttodar.com>
' */
Public Function ArrayIncludes(ByVal sourceArray As Variant, ByVal value As Variant) As Boolean
    If IsArrayEmpty(sourceArray) Then
        Exit Function
    End If
    
    Dim index As Long
    For index = LBound(sourceArray, 1) To UBound(sourceArray, 1)
        If sourceArray(index) = value Then
            ArrayIncludes = True
            Exit For
        End If
    Next index
End Function

'/**
' * RETURNS INDEX OF A SINGLE DIM ARRAY ELEMENT
' *
' * @author Robert Todar <robert@roberttodar.com>
' */
Public Function ArrayIndexOf(ByVal sourceArray As Variant, ByVal SearchElement As Variant) As Integer
    Dim index As Long
    For index = LBound(sourceArray, 1) To UBound(sourceArray, 1)
        If sourceArray(index) = SearchElement Then
            ArrayIndexOf = index
            Exit Function
        End If
    Next index
    index = -1
End Function

'/**
' * EXTRACTS LIST OF GIVEN PROPERTY. MUST BE ARRAY THAT CONTAINS DICTIONRIES AT THIS TIME.
' *
' * @author Robert Todar <robert@roberttodar.com>
' */
Public Function ArrayPluck(ByVal sourceArray As Variant, ByVal key As Variant) As Variant
    Dim temp As Variant
    ReDim temp(LBound(sourceArray, 1) To UBound(sourceArray, 1))
    
    Dim index As Integer
    For index = LBound(sourceArray, 1) To UBound(sourceArray, 1)
        assign temp(index), sourceArray(index)(key)
    Next index

    ArrayPluck = temp
End Function

'/**
' * REMOVES LAST ELEMENT IN ARRAY, RETURNS POPPED ELEMENT
' *
' * @author Robert Todar <robert@roberttodar.com>
' */
Public Function ArrayPop(ByRef sourceArray As Variant) As Variant
    If Not IsArrayEmpty(sourceArray) Then
        Select Case ArrayDimensionLength(sourceArray)
            Case 1:
                ArrayPop = sourceArray(UBound(sourceArray, 1))
                ReDim Preserve sourceArray(LBound(sourceArray, 1) To UBound(sourceArray, 1) - 1)
            
            Case 2:
                Dim temp As Variant
                ReDim temp(LBound(sourceArray, 2) To UBound(sourceArray, 2))
                
                Dim colIndex As Integer
                For colIndex = LBound(sourceArray, 2) To UBound(sourceArray, 2)
                    temp(colIndex) = sourceArray(UBound(sourceArray, 1), colIndex)
                Next colIndex
                ArrayPop = temp
                
                ArrayTranspose sourceArray
                ReDim Preserve sourceArray(LBound(sourceArray, 1) To UBound(sourceArray, 1), LBound(sourceArray, 2) To UBound(sourceArray, 2) - 1)
                ArrayTranspose sourceArray
        End Select
    End If
End Function

'/**
' * ADDS A NEW ELEMENT(S) TO AN ARRAY (AT THE END), RETURNS THE NEW ARRAY LENGTH
' *
' * @author Robert Todar <robert@roberttodar.com>
' * @param: <SourceArray> can be either 1 or 2 dimensional.
' * @param: <Element> are the elements to be added.
' */
Public Function ArrayPush(ByRef sourceArray As Variant, ParamArray element() As Variant) As Long
    Dim index As Long
    Dim FirstEmptyBound As Long
    Dim OptionBase As Integer
    
    OptionBase = 0

    ' THIS IS ONLY FOR SINGLE DIMENSIONS.
    If ArrayDimensionLength(sourceArray) = 2 Then  'Or IsArray(Element(LBound(Element)))
        'THIS SECTION IS EXPERIMENTAL... ArrayPushTwoDim IS NOT YET PROVEN. REMOVE IF DESIRED.
        ArrayPush = ArrayPushTwoDim(sourceArray, CVar(element))
        Exit Function
    End If
    
    ' REDIM IF EMPTY, OR INCREASE ARRAY IF NOT EMPTY
    If IsArrayEmpty(sourceArray) Then
        ReDim sourceArray(OptionBase To UBound(element, 1) + OptionBase)
        FirstEmptyBound = LBound(sourceArray, 1)
    Else
        FirstEmptyBound = UBound(sourceArray, 1) + 1
        ReDim Preserve sourceArray(UBound(sourceArray, 1) + UBound(element, 1) + 1)
    End If
    
    ' LOOP EACH NEW ELEMENT
    For index = LBound(element, 1) To UBound(element, 1)
        ' ADD ELEMENT TO THE END OF THE ARRAY
        assign sourceArray(FirstEmptyBound), element(index)
        
        ' INCREMENT TO THE NEXT firstEmptyBound
        FirstEmptyBound = FirstEmptyBound + 1
    Next index
    
    ' RETURN NEW ARRAY LENGTH
    ArrayPush = UBound(sourceArray, 1) + 1
End Function

'/**
' * ADDS A NEW ELEMENT(S) TO AN ARRAY (AT THE END), RETURNS THE NEW ARRAY LENGTH
' *
' * @author Robert Todar <robert@roberttodar.com>
' */
Public Function ArrayPushTwoDim(ByRef sourceArray As Variant, ParamArray element() As Variant) As Long
    Dim FirstEmptyRow As Long
    Dim OptionBase As Integer
    
    OptionBase = 0

    ' REDIM IF EMPTY, OR INCREASE ARRAY IF NOT EMPTY
    If IsArrayEmpty(sourceArray) Then
        ReDim sourceArray(OptionBase To UBound(element, 1) + OptionBase, OptionBase To arrayLength(element(LBound(element))) + OptionBase - 1)
        FirstEmptyRow = LBound(sourceArray, 1)
    Else
        FirstEmptyRow = UBound(sourceArray, 1) + 1
        sourceArray = ArrayTranspose(sourceArray)
        ReDim Preserve sourceArray(LBound(sourceArray, 1) To UBound(sourceArray, 1), LBound(sourceArray, 2) To UBound(sourceArray, 2) + arrayLength(element))
        sourceArray = ArrayTranspose(sourceArray)
    End If
    
    ' LOOP EACH ARRAY
    Dim index As Long
    For index = LBound(element, 1) To UBound(element, 1)
        Dim CurrentIndex As Long
        CurrentIndex = LBound(element(index))
        
        ' LOOP EACH ELEMENT IN CURRENT ARRAY
        Dim colIndex As Long
        For colIndex = LBound(sourceArray, 2) To UBound(sourceArray, 2)
            ' ADD ELEMENT TO THE END OF THE ARRAY. NOTE IF ERROR CHANCES ARE ARRAY DIM WAS NOT THE SAME
            assign sourceArray(FirstEmptyRow, colIndex), element(index)(CurrentIndex)
            CurrentIndex = CurrentIndex + 1
        Next colIndex
    
        ' INCREMENT TO THE NEXT firstEmptyRow
        FirstEmptyRow = FirstEmptyRow + 1
    Next index
    
    ' RETURN NEW ARRAY LENGTH
    ArrayPushTwoDim = UBound(sourceArray, 1) - LBound(sourceArray, 1) + 1
End Function

'/**
' * CREATES TEMP TEXT FILE AND SAVES ARRAY VALUES IN A CSV FORMAT,
' * THEN QUERIES AND RETURNS ARRAY.
' *
' * @author Robert Todar <robert@roberttodar.com>
' * @uses ArrayToCSVFile
' * @uses ArrayFromRecordset
' * @returns 2D ARRAY || EMPTY (IF NO RECORDS)
' * @param {ARR} MUST BE A TWO DIMENSIONAL ARRAY, SETUP AS IF IT WERE A TABLE.
' * @param {SQL} ADO SQL STATEMENT FOR A TEXT FILE. MUST INCLUDE 'FROM []'
' * @param {IncludeHeaders} BOOLEAN TO RETURN HEADERS WITH DATA OR NOT
' * @example SQL = "SELECT * FROM [] WHERE [FIRSTNAME] = 'ROBERT'"
' */
Public Function ArrayQuery(sourceArray As Variant, sql As String, Optional IncludeHeaders As Boolean = True) As Variant
    ' CREATE TEMP FOLDER AND FILE NAMES
    Const filename As String = "temp.txt"
    Dim FILEPATH As String
    FILEPATH = Environ("temp")
    
    ' UPDATE SQL WITH TEMP FILE NAME
    sql = Replace(sql, "FROM []", "FROM [" & filename & "]")
    
    ' SEND ARRAY TO TEMP TEXTFILE IN CSV FORMAT
    ArrayToCSVFile sourceArray, FILEPATH & "\" & filename
    
    ' CREATE CONNECTION TO TEMP FILE - CONNECTION IS SET TO COMMA SEPERATED FORMAT
    Dim cnn As Object
    Set cnn = CreateObject("ADODB.Connection")
    cnn.Provider = "Microsoft.Jet.OLEDB.4.0"
    cnn.connectionString = "Data Source=" & FILEPATH & ";" & "Extended Properties=""text;HDR=Yes;FMT=Delimited;"""
    cnn.Open
    
    ' CREATE RECORDSET AND QUERY ON PASSED IN SQL (QUERIES THE TEMP TEXT FILE)
    Dim rs As Object
    Set rs = CreateObject("ADODB.RecordSet")
    With rs
        .ActiveConnection = cnn
        .Open sql
        
        ' GET AN ARRAY FROM THE RECORDSET
         ArrayQuery = ArrayFromRecordset(rs, IncludeHeaders)
        .Close
    End With
    
    ' CLOSE CONNECTION AND KILL TEMP FILE
    cnn.Close
    Kill FILEPATH & "\" & filename
End Function

'/**
' * REMOVED DUPLICATES FROM SINGLE DIM ARRAY
' *
' * @author Robert Todar <robert@roberttodar.com>
' */
Public Function ArrayRemoveDuplicates(sourceArray As Variant) As Variant
    If Not IsArray(sourceArray) Then
        sourceArray = ConvertToArray(sourceArray)
    End If
    
    Dim Dic As Object
    Dim key As Variant
    Set Dic = CreateObject("Scripting.Dictionary")
    For Each key In sourceArray
        Dic(key) = 0
    Next
    
    ArrayRemoveDuplicates = Dic.keys
    sourceArray = ArrayRemoveDuplicates
End Function

'/**
' * REVERSE ARRAY (CAN BE USED AFTER SORT TO GET THE DECENDING ORDER)
' *
' * @author Robert Todar <robert@roberttodar.com>
' */
Public Function ArrayReverse(sourceArray As Variant) As Variant
    ' REVERSE LOOP (HALF OF IT, WILL WORK FROM BOTH SIDES ON EACH ITERATION)
    Dim temp As Variant
    Dim index As Long
    For index = LBound(sourceArray, 1) To ((UBound(sourceArray) + LBound(sourceArray)) \ 2)
        ' STORE LAST VALUE MINUS THE ITERATION
        assign temp, sourceArray(UBound(sourceArray) + LBound(sourceArray) - index)
        
        ' SET LAST VALUE TO FIRST VALUE OF THE ARRAY
        assign sourceArray(UBound(sourceArray) + LBound(sourceArray) - index), sourceArray(index)
        
        ' SET FIRST VALUE TO THE STORED LAST VALUE
        assign sourceArray(index), temp
    Next index
    
    ArrayReverse = sourceArray
End Function

'/**
' * REMOVES ELEMENT FROM ARRAY - RETURNS REMOVED ELEMENT **[SINGLE DIMENSION]
' *
' * @author Robert Todar <robert@roberttodar.com>
' */
Public Function ArrayShift(sourceArray As Variant, Optional ElementNumber As Long = 0) As Variant
    If Not IsArrayEmpty(sourceArray) Then
        ArrayShift = sourceArray(ElementNumber)
        
        Dim index As Long
        For index = ElementNumber To UBound(sourceArray) - 1
            assign sourceArray(index), sourceArray(index + 1)
        Next index
        
        ReDim Preserve sourceArray(UBound(sourceArray, 1) - 1)
    End If
End Function

'/**
' * SORT AN ARRAY [SINGLE DIMENSION]
' *
' * @author Robert Todar <robert@roberttodar.com>
' */
Public Function ArraySort(sourceArray As Variant) As Variant
    ' SORT ARRAY A-Z
    Dim OuterIndex As Long
    For OuterIndex = LBound(sourceArray) To UBound(sourceArray) - 1
        
        Dim InnerIndex As Long
        For InnerIndex = OuterIndex + 1 To UBound(sourceArray)
            
            If UCase(sourceArray(OuterIndex)) > UCase(sourceArray(InnerIndex)) Then
                Dim temp As Variant
                temp = sourceArray(InnerIndex)
                sourceArray(InnerIndex) = sourceArray(OuterIndex)
                sourceArray(OuterIndex) = temp
            End If
            
        Next InnerIndex
    Next OuterIndex
    
    ArraySort = sourceArray
End Function

'/**
' * CHANGES THE CONTENTS OF AN ARRAY BY REMOVING OR REPLACING EXISTING ELEMENTS AND/OR ADDING NEW ELEMENTS.
' *
' * @author Robert Todar <robert@roberttodar.com>
' */
Public Function ArraySplice(sourceArray As Variant, Where As Long, HowManyRemoved As Integer, ParamArray element() As Variant) As Variant
    ' CHECK TO SEE THAT INSERT IS NOT GREATER THAN THE Array (REDUCE IF SO)
    If Where > UBound(sourceArray, 1) + 1 Then
        Where = UBound(sourceArray, 1) + 1
    End If
    
    ' CHECK TO MAKE SURE REMOVED IS NOT MORE THAN THE Array (REDUCE IF SO)
    If HowManyRemoved > (UBound(sourceArray, 1) + 1) - Where Then
        HowManyRemoved = (UBound(sourceArray, 1) + 1) - Where
    End If
    
    If UBound(sourceArray, 1) + UBound(element, 1) + 1 - HowManyRemoved < 0 Then
        ArraySplice = Empty
        sourceArray = Empty
        Exit Function
    End If
    
    ' SET BOUNDS TO TEMP Array
    Dim temp As Variant
    ReDim temp(LBound(sourceArray, 1) To UBound(sourceArray, 1) + UBound(element, 1) + 1 - HowManyRemoved)
    
    ' LOOP TEMP Array, ADDING\REMOVING WHERE NEEDED
    Dim index As Long
    For index = LBound(temp, 1) To UBound(temp, 1)
        
        ' INSERT ONCE AT WHERE, AND ONLY VISIT ONCE
        Dim Visited As Boolean
        If index = Where And Visited = False Then
            Visited = True
            ' ADD NEW ELEMENTS
            Dim Index2 As Long
            Dim Index3 As Long
            For Index2 = LBound(element, 1) To UBound(element, 1)
                temp(index) = element(Index2)
                
                ' INCREMENT COUNTERS
                Index3 = Index3 + 1
                index = index + 1
            Next Index2
            
            ' GET REMOVED ELEMENTS TO RETURN
            Dim RemovedArray As Variant
            If HowManyRemoved > 0 Then
                ReDim RemovedArray(0 To HowManyRemoved - 1)
                For Index2 = LBound(RemovedArray, 1) To UBound(RemovedArray, 1)
                    RemovedArray(Index2) = sourceArray(Where + Index2)
                Next Index2
            Else
                RemovedArray = Empty
            End If
            
            ' DECREMENT COUNTERS FOR AFTER LOOP
            index = index - 1
            Index3 = Index3 - HowManyRemoved
        Else
            ' ADD PREVIOUS ELEMENTS (Index3 IS A HELPER)
            temp(index) = sourceArray(index - Index3)
        End If
        
    Next index
    sourceArray = temp
    ArraySplice = RemovedArray
End Function

'/**
' * BASICALY ARRAY TO STRING HOWEVER QUOTING STIRNGS, THEN SAVING TO A TEXTFILE
' *
' * @author Robert Todar <robert@roberttodar.com>
' */
Public Function ArrayToCSVFile(sourceArray As Variant, FILEPATH As String) As String
    Dim temp As String
    Const delimiter = ","
    
    Select Case ArrayDimensionLength(sourceArray)
        ' SINGLE DIMENTIONAL ARRAY
        Case 1
            Dim index As Integer
            For index = LBound(sourceArray, 1) To UBound(sourceArray, 1)
                If IsNumeric(sourceArray(index)) Then
                    temp = temp & sourceArray(index)
                Else
                    temp = temp & """" & sourceArray(index) & """"
                End If
            Next index
            
        ' 2 DIMENSIONAL ARRAY
        Case 2
            Dim rowIndex As Long
            Dim colIndex As Long
            
            ' LOOP EACH ROW IN MULTI ARRAY
            For rowIndex = LBound(sourceArray, 1) To UBound(sourceArray, 1)
                ' LOOP EACH COLUMN ADDING VALUE TO STRING
                For colIndex = LBound(sourceArray, 2) To UBound(sourceArray, 2)
                
                    If IsNumeric(sourceArray(rowIndex, colIndex)) Then
                        temp = temp & sourceArray(rowIndex, colIndex)
                    Else
                        temp = temp & """" & sourceArray(rowIndex, colIndex) & """"
                    End If
                    
                    If colIndex <> UBound(sourceArray, 2) Then temp = temp & delimiter
                Next colIndex
                
                ' ADD NEWLINE FOR THE NEXT ROW (MINUS LAST ROW)
                If rowIndex <> UBound(sourceArray, 1) Then temp = temp & vbNewLine
        
            Next rowIndex
    End Select
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim ts As Object
    Set ts = fso.OpenTextFile(FILEPATH, 2, True) '2=WRITEABLE
    ts.Write temp
    
    Set ts = Nothing
    Set fso = Nothing
    
    ArrayToCSVFile = temp
End Function

'/**
' * RESIZE PASSED IN EXCEL RANGE, AND SET VALUE EQUAL TO THE ARRAY
' *
' * @author Robert Todar <robert@roberttodar.com>
' * @todo: NEED TO TEST! ALSO THIS ASSUMES ROW, GIVE OPTION TO TRANSPOSE TO COLUMN??
' * note: THIS ALWAYS FORMATS THE CELLS TO BE A STRING... REMOVE FORMATING IF NEED BE.
' *      THIS WAS CREATED FOR THE PURPOSE OF MAINTAINING LEADING ZEROS FOR MY ALL DATA...
' */
Public Sub ArrayToRange(ByVal sourceArray As Variant, Optional ByRef target As Excel.Range)
    ' ADD WORKBOOK IF NOT
    Dim wb As Workbook
    If target Is Nothing Then
        Set wb = Workbooks.add
        Set target = wb.Worksheets("Sheet1").Range("A1")
    End If
    
    Select Case ArrayDimensionLength(sourceArray)
        Case 1:
            Set target = target.Resize(UBound(sourceArray) - LBound(sourceArray) + 1, 1)
            target.NumberFormat = "@"
            target.value = Application.Transpose(sourceArray)
        Case 2:
            Set target = target.Resize((UBound(sourceArray, 1) + 1) - LBound(sourceArray, 1), (UBound(sourceArray, 2) + 1 - LBound(sourceArray, 2)))
            target.NumberFormat = "@"
            target.value = sourceArray
            'Target.Resize((UBound(SourceArray, 1) + 1) - LBound(SourceArray, 1), (UBound(SourceArray, 2) + 1 - LBound(SourceArray, 2))).Value = SourceArray
    End Select
    
    ' OPTIONAL, PLEASE REMOVE IF DESIRED...
    Columns.AutoFit
End Sub

'/**
' * RETURNS A STRING FROM A 2 DIM ARRAY, SPERATED BY OPTIONAL DELIMITER AND VBNEWLINE FOR EACH ROW
' *
' * @author Robert Todar <robert@roberttodar.com>
' */
Public Function ArrayToString(sourceArray As Variant, Optional delimiter As String = ",") As String
    Dim temp As String
    
    Select Case ArrayDimensionLength(sourceArray)
        ' SINGLE DIMENTIONAL ARRAY
        Case 1
            temp = Join(sourceArray, delimiter)
        
        ' 2 DIMENSIONAL ARRAY
        Case 2
            Dim rowIndex As Long
            Dim colIndex As Long
            
            ' LOOP EACH ROW IN MULTI ARRAY
            For rowIndex = LBound(sourceArray, 1) To UBound(sourceArray, 1)
                
                ' LOOP EACH COLUMN ADDING VALUE TO STRING
                For colIndex = LBound(sourceArray, 2) To UBound(sourceArray, 2)
                    temp = temp & sourceArray(rowIndex, colIndex)
                    If colIndex <> UBound(sourceArray, 2) Then temp = temp & delimiter
                Next colIndex
                
                ' ADD NEWLINE FOR THE NEXT ROW (MINUS LAST ROW)
                If rowIndex <> UBound(sourceArray, 1) Then temp = temp & vbNewLine
        
            Next rowIndex
    End Select
    
    ArrayToString = temp
End Function

'/**
' * SENDS AN ARRAY TO A TEXTFILE
' *
' * @author Robert Todar <robert@roberttodar.com>
' */
Public Sub ArrayToTextFile(arr As Variant, FILEPATH As String, Optional delimeter As String = ",")
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim ts As Object
    Set ts = fso.OpenTextFile(FILEPATH, 2, True) '2=WRITEABLE
    ts.Write ArrayToString(arr, delimeter)
    
    Set ts = Nothing
    Set fso = Nothing
End Sub

'/**
' * APPLICATION.TRANSPOSE HAS A LIMIT ON THE SIZE OF THE ARRAY, AND IS LIMITED TO THE 1ST DIM
' *
' * @author Robert Todar <robert@roberttodar.com>
' */
Public Function ArrayTranspose(sourceArray As Variant) As Variant
    Dim temp As Variant
    Select Case ArrayDimensionLength(sourceArray)
        Case 2:
            ReDim temp(LBound(sourceArray, 2) To UBound(sourceArray, 2), LBound(sourceArray, 1) To UBound(sourceArray, 1))
            
            Dim i As Long
            Dim j As Long
            For i = LBound(sourceArray, 2) To UBound(sourceArray, 2)
                For j = LBound(sourceArray, 1) To UBound(sourceArray, 1)
                    temp(i, j) = sourceArray(j, i)
                Next
            Next
    End Select
    
    ArrayTranspose = temp
    sourceArray = temp
End Function

'/**
' * ADDS NEW ELEMENT TO THE BEGINING OF THE ARRAY
' *
' * @author Robert Todar <robert@roberttodar.com>
' */
Public Function ArrayUnShift(sourceArray As Variant, ParamArray element() As Variant) As Long
    ' FOR NOW THIS IS ONLY FOR SINGLE DIMENSIONS. @todo: UPDATE TO PUSH TO MULTI DIM ARRAYS
    If ArrayDimensionLength(sourceArray) <> 1 Then
        ArrayUnShift = -1
        Exit Function
    End If
    
    ' RESIZE TEMP ARRAY
    Dim temp As Variant
    If IsArrayEmpty(sourceArray) Then
        ReDim temp(0 To UBound(element, 1))
    Else
        ReDim temp(UBound(sourceArray, 1) + UBound(element, 1) + 1)
    End If
    
    Dim count As Long
    count = LBound(temp, 1)
    Dim index As Long
    
    ' ADD ELEMENTS TO TEMP ARRAY
    For index = LBound(element, 1) To UBound(element, 1)
        assign temp(count), element(index)
        count = count + 1
    Next index
    
    If Not count > UBound(temp, 1) Then
        ' ADD ELEMENTS FROM ORIGINAL ARRAY
        For index = LBound(sourceArray, 1) To UBound(sourceArray, 1)
            assign temp(count), sourceArray(index)
            count = count + 1
        Next index
    End If
    
    ' SET ARRAY TO TEMP ARRAY
    sourceArray = temp
    
    ' RETURN THE NEW LENGTH OF THE ARRAY
    ArrayUnShift = UBound(sourceArray, 1) + 1
End Function

'/**
' * QUICK TOOL TO EITHER SET OR LET DEPENDING ON IF ELEMENT IS AN OBJECT
' *
' * @author Robert Todar <robert@roberttodar.com>
' */
Public Function assign(ByRef Variable As Variant, ByVal value As Variant)
    If IsObject(value) Then
        Set Variable = value
    Else
        Let Variable = value
    End If
End Function

'/**
' * CONVERT OTHER LIST OBJECTS TO AN ARRAY
' *
' * @author Robert Todar <robert@roberttodar.com>
' */
Public Function ConvertToArray(ByRef Val As Variant) As Variant
    Select Case TypeName(Val)
        Case "Collection":
            Dim index As Integer
            For index = 1 To Val.count
                ArrayPush ConvertToArray, Val(index)
            Next index
        
        Case "Dictionary":
            ConvertToArray = Val.items()
        
        Case Else
            If IsArray(Val) Then
                ConvertToArray = Val
            Else
                ArrayPush ConvertToArray, Val
            End If
    End Select
End Function

Public Function IsArrayEmpty(arr As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CPEARSON
' This function tests whether the array is empty (unallocated). Returns TRUE or FALSE.
'
' The VBA IsArray function indicates whether a variable is an array, but it does not
' distinguish between allocated and unallocated arrays. It will return TRUE for both
' allocated and unallocated arrays. This function tests whether the array has actually
' been allocated.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Err.Clear
    On Error Resume Next
    If IsArray(arr) = False Then
        ' we weren't passed an array, return True
        IsArrayEmpty = True
    End If

    ' Attempt to get the UBound of the array. If the array is
    ' unallocated, an error will occur.
    Dim ub As Long
    ub = UBound(arr, 1)
    If (Err.Number <> 0) Then
        IsArrayEmpty = True
    Else
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' On rare occasion, under circumstances I cannot reliably replicate, Err.Number
        ' will be 0 for an unallocated, empty array. On these occasions, LBound is 0 and
        ' UBound is -1. To accommodate the weird behavior, test to see if LB > UB.
        ' If so, the array is not allocated.
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Err.Clear
        Dim LB As Long
        LB = LBound(arr)
        If LB > ub Then
            IsArrayEmpty = True
        Else
            IsArrayEmpty = False
        End If
    End If
End Function

'/**
' * CHECKS CURRENT ROW OF A TWO DIM ARRAY TO SEE IF CONDITIONS ARRAY PASSES
' * HELPER FUNCTION FOR ARRAYFILTERTWO
' * @ref RegExTest
' *
' * @author Robert Todar <robert@roberttodar.com>
' */
Private Function IsValidConditions(ByVal sourceArray As Variant, ByVal Conditions As Variant, ByVal rowIndex As Integer)
    ' CHECK CONDITIONS
    Dim index As Integer
    For index = LBound(Conditions) To UBound(Conditions)
        Dim value As String
        value = sourceArray(rowIndex, Conditions(index)(0))
        
        Dim pattern As String
        pattern = CStr(Conditions(index)(1))
        
        If Not RegExTest(value, pattern) Then
            Exit Function
        End If
    Next index
    
    IsValidConditions = True
End Function

'/**
' * GROUPS HEADING INDEX WITH CONDITIONS. RETURNS JAGGED ARRAY.
' * HELPER FUNCTION FOR ARRAYFILTERTWO
' *
' * @author Robert Todar <robert@roberttodar.com>
' */
Private Function ArrayGetConditions(ByVal sourceArray As Variant, ByVal Arguments As Variant) As Variant
    ' ARGUMENTS ARE PAIRED BY TWOS. (0) = COLUMN HEADING, (1) = REGEX CONDITION
    Dim index As Integer
    For index = LBound(Arguments) To UBound(Arguments) Step 2
        Dim ColumnIndex As Integer
        ColumnIndex = ArrayGetColumnIndex(sourceArray, Arguments(index))
        ArrayPush ArrayGetConditions, Array(ColumnIndex, Arguments(index + 1))
    Next index
End Function

'/**
' * SIMPLE FUNCTION TO TEST REGULAR EXPRESSIONS. FOR HELP SEE:
' *
' * @author Robert Todar <robert@roberttodar.com>
' */
Private Function RegExTest(ByVal value As String, ByVal pattern As String) As Boolean
    Dim regex As Object
    Set regex = CreateObject("vbscript.regexp")
    With regex
        .Global = True 'TRUE MEANS IT WILL LOOK FOR ALL MATCHES, FALSE FINDS FIRST ONLY
        .MultiLine = True
        .IgnoreCase = True
        .pattern = pattern
    End With
    RegExTest = regex.test(value)
End Function



