Attribute VB_Name = "RecordsetLib_TESTS"
'@Folder("RecordsetLib")
'@IgnoreModule ProcedureNotUsed, UnassignedVariableUsage, VariableNotAssigned, UseMeaningfulName
Option Explicit

Public Enum RecordSetData
    UserData
End Enum

Private Property Get sampleData(ByVal dataType As RecordSetData) As ADODB.Recordset
    Set sampleData = New Recordset
    Select Case dataType
        Case RecordSetData.UserData:
            With sampleData
                .fields.Append "First", adVarChar, 255
                .fields.Append "Last", adVarChar, 255
                .fields.Append "Age", adInteger, 255

                .Open
                .AddNew Array("First", "Last", "Age"), Array("Mark", "Carter", 35)
                .AddNew Array("First", "Last", "Age"), Array("Robert", "Todar", 33)
                .AddNew Array("First", "Last", "Age"), Array("Bob", "King", 65)
                .UpdateBatch
            End With
    End Select
End Property

Public Function Test_RecordsetToRange() As Boolean
    ' Get Sample data
    Dim rs As ADODB.Recordset
    Set rs = sampleData(UserData)
    
    
    'Set rs = RangeToRecordset(DataSheet.Range("A1").CurrentRegion)
    
    ' Do tests
    'rs.Filter = "Age > 40"
    
    Dim newRange As Range
    Set newRange = RecordsetToRange(rs, Sheet1.Range("A1"))
    
    ' Cleanup code
    Sheet1.Range("A1").CurrentRegion.ClearContents
    
    Debug.Assert Not newRange Is Nothing
    Test_RecordsetToRange = True
End Function
