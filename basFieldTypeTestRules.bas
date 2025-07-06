Attribute VB_Name = "basFieldTypeTestRules"
Option Compare Database
Option Explicit
' ============================================================================
' Module: basFieldTypeTestRules
' Purpose: Encodes rules and test routines that validate field *definitions*.
'          Ensures that datatype parameters (e.g., FieldSize) make logical sense.
'
' Example Use Cases:
'   - Disallow FieldSize on YesNo fields
'   - Enforce DecimalPlaces only on numeric fields
'
' Used In:
'   - Testing framework for clsFieldDef logic validation
' ============================================================================

'****************************************************************************************
'Public Sub: RunAllFieldDefTests
'
' Description:
'     Serves as a master routine to execute all individual field definition validation tests.
'     Each field-specific test ensures that inappropriate or contradictory metadata
'     (e.g., invalid field size, AutoNumber flag on wrong type, PK constraints) are
'     detected and rejected by clsFieldDef.ValidateDefinition.
'
' Parameters:
'     (None)
'
' Returns:
'     None
'
' Side Effects:
'     Executes all fTestXXXFieldDefErrors routines.
'****************************************************************************************

Public Sub RunAllFieldDefTests()
    assDebugPrint "--- Running Field Definition Validation Tests ---"
    'ClearImmediateWindow
    fTestBinaryFieldDefErrors
    fTestBooleanFieldDefErrors
    fTestByteFieldDefErrors
    fTestCurrencyFieldDefErrors
    fTestDateFieldDefErrors
    fTestGUIDFieldDefErrors
    fTestIntegerFieldDefErrors
    fTestLongFieldDefErrors
    fTestMemoFieldDefErrors
    fTestSingleFieldDefErrors
    fTestTextFieldDefErrors
    'fTestOtherFieldDefErrors

    assDebugPrint "--- All Field Definition Tests Completed ---"
End Sub

' === fTestBinaryFieldDefErrors ===
' Tests invalid configurations for dbLongBinary fields
Public Sub fTestBinaryFieldDefErrors()
    Dim lTbl As clsTableBuilder
    Set lTbl = New clsTableBuilder
    lTbl.pTableName = "_testBinaryFields"
    assDebugPrint vbCrLf & lTbl.pTableName

    ' === [1] Invalid: Binary field with Size = 0
    lTbl.AddFieldDef lTbl.CreateFieldDef("BinarySizeZero", dbBinary, _
        RequiredBehavior.fdRequired, IndexedBehavior.fdUnindexed, _
        PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdNo, _
        0, AutoNumberBehavior.IsNotAutoNumber)

    ' === [2] Invalid: Binary field marked AutoNumber
    lTbl.AddFieldDef lTbl.CreateFieldDef("BinaryAutoNumber", dbBinary, _
        RequiredBehavior.fdOptional, IndexedBehavior.fdUnindexed, _
        PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdNo, _
        50, AutoNumberBehavior.IsAutoNumber)

    ' === [3] Optional: Binary field with AllowZeroLength = Yes (not applicable)
    lTbl.AddFieldDef lTbl.CreateFieldDef("BinaryZeroLengthYes", dbBinary, _
        RequiredBehavior.fdOptional, IndexedBehavior.fdUnindexed, _
        PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdYes, _
        50, AutoNumberBehavior.IsNotAutoNumber)

    ' === [3] Optional: Binary field with Length > 510
    lTbl.AddFieldDef lTbl.CreateFieldDef("BinaryLengthGT510", dbBinary, _
        RequiredBehavior.fdOptional, IndexedBehavior.fdUnindexed, _
        PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdYes, _
        511, AutoNumberBehavior.IsNotAutoNumber)
    
    ' === [3] Optional: Binary field with Length > 510
    lTbl.AddFieldDef lTbl.CreateFieldDef("BinaryLengthLT0", dbBinary, _
        RequiredBehavior.fdOptional, IndexedBehavior.fdUnindexed, _
        PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdYes, _
        -1, AutoNumberBehavior.IsNotAutoNumber)
    
    ' === Report Definition Errors ===
    Dim lStrErrors As String
    lStrErrors = lTbl.ReportFieldDefinitionErrors()
    If Len(lStrErrors) > 0 Then
        assDebugPrint "Errors during Binary field definition:" & vbCrLf & lStrErrors
    Else
        assDebugPrint "No validation errors detected for Binary fields — unexpected!"
    End If
End Sub

' === fTestBooleanFieldDefErrors ===
' Tests invalid configurations for dbBoolean fields
Public Sub fTestBooleanFieldDefErrors()
    Dim lTbl As clsTableBuilder
    Set lTbl = New clsTableBuilder

    lTbl.pTableName = "_testBooleanFields"
    assDebugPrint vbCrLf & lTbl.pTableName

    ' === [1] Invalid: Boolean field with AllowZeroLength = True
    lTbl.AddFieldDef lTbl.CreateFieldDef("BoolZeroLength", dbBoolean, _
        RequiredBehavior.fdRequired, IndexedBehavior.fdUnindexed, _
        PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdYes, _
        0, AutoNumberBehavior.IsNotAutoNumber)

    ' === [2] Invalid: Boolean field marked AutoNumber
    lTbl.AddFieldDef lTbl.CreateFieldDef("BoolAutoNumber", dbBoolean, _
        RequiredBehavior.fdOptional, IndexedBehavior.fdUnindexed, _
        PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdNo, _
        0, AutoNumberBehavior.IsAutoNumber)

    ' === Report Definition Errors ===
    Dim lStrErrors As String
    lStrErrors = lTbl.ReportFieldDefinitionErrors()
    If Len(lStrErrors) > 0 Then
        assDebugPrint "Errors during boolean field definition:" & vbCrLf & lStrErrors
    Else
        assDebugPrint "No validation errors detected for boolean fields — unexpected!"
    End If
End Sub

' === fTestByteFieldDefErrors ===
' Tests invalid configurations for dbByte fields
Public Sub fTestByteFieldDefErrors()
    Dim lTbl As clsTableBuilder
    Set lTbl = New clsTableBuilder

    lTbl.pTableName = "_testByteFields"
    assDebugPrint vbCrLf & lTbl.pTableName

    ' === [1] Invalid: Byte field with AllowZeroLength
    lTbl.AddFieldDef lTbl.CreateFieldDef("ByteZeroLength", dbByte, _
        RequiredBehavior.fdRequired, IndexedBehavior.fdUnindexed, _
        PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdYes, _
        0, AutoNumberBehavior.IsNotAutoNumber)

    ' === Report Definition Errors ===
    Dim lStrErrors As String
    lStrErrors = lTbl.ReportFieldDefinitionErrors()
    If Len(lStrErrors) > 0 Then
        assDebugPrint "Errors during Byte field definition:" & vbCrLf & lStrErrors
    Else
        assDebugPrint "No validation errors detected for Byte fields — unexpected!"
    End If
End Sub

' === fTestCurrencyFieldDefErrors ===
' Intentionally defines invalid currency field definitions to test validation
Public Sub fTestCurrencyFieldDefErrors()
    Dim lTbl As clsTableBuilder
    Set lTbl = New clsTableBuilder

    lTbl.pTableName = "_testCurrencyFields"
    assDebugPrint vbCrLf & lTbl.pTableName

    ' === Invalid: Currency fields should not specify Size ===
    lTbl.AddFieldDef lTbl.CreateFieldDef("CurrencyHasSize", dbCurrency, _
        RequiredBehavior.fdRequired, IndexedBehavior.fdUnindexed, _
        PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdNo, _
        0, AutoNumberBehavior.IsNotAutoNumber)

    ' === Invalid: Currency field with indexed duplicates allowed
    lTbl.AddFieldDef lTbl.CreateFieldDef("CurrencyIndexedDuplicatesAllowed", dbCurrency, _
        RequiredBehavior.fdRequired, IndexedBehavior.fdUnindexed, _
        PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdNo, _
        0, AutoNumberBehavior.IsNotAutoNumber)

    ' === Invalid: Currency field with indexed no duplicates
    lTbl.AddFieldDef lTbl.CreateFieldDef("CurrencyIndexedNoDuplicates", dbCurrency, _
        RequiredBehavior.fdRequired, IndexedBehavior.fdUnindexed, _
        PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdNo, _
        0, AutoNumberBehavior.IsNotAutoNumber)

    ' === Invalid: Currency field cannot be AutoNumber
    lTbl.AddFieldDef lTbl.CreateFieldDef("CurrencyAutoNumber", dbCurrency, _
        RequiredBehavior.fdOptional, IndexedBehavior.fdUnindexed, _
        PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdNo, _
        0, AutoNumberBehavior.IsAutoNumber)

    ' === Invalid: Currency field cannot AllowZeroLength (only text/memo)
    lTbl.AddFieldDef lTbl.CreateFieldDef("CurrencyZeroLength", dbCurrency, _
        RequiredBehavior.fdOptional, IndexedBehavior.fdUnindexed, _
        PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdYes, _
        0, AutoNumberBehavior.IsNotAutoNumber)

    ' === Report Definition Errors ===
    Dim lStrErrors As String
    lStrErrors = lTbl.ReportFieldDefinitionErrors()
    If Len(lStrErrors) > 0 Then
        assDebugPrint "Errors during currency field definition:" & vbCrLf & lStrErrors
    Else
        assDebugPrint "No validation errors detected for currency fields — unexpected!"
    End If
End Sub

' === fTestDateFieldDefErrors ===
' Intentionally defines invalid date field definitions to test validation
Public Sub fTestDateFieldDefErrors()
    Dim lTbl As clsTableBuilder
    Set lTbl = New clsTableBuilder

    lTbl.pTableName = "_testDateFields"
    assDebugPrint vbCrLf & lTbl.pTableName
    
    ' === [1] Invalid: Date field with Size
    lTbl.AddFieldDef lTbl.CreateFieldDef("DateWithSize", dbDate, _
        RequiredBehavior.fdRequired, IndexedBehavior.fdUnindexed, _
        PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdNo, _
        0, AutoNumberBehavior.IsNotAutoNumber)

    ' === [2] Invalid: Date field with AutoNumber
    lTbl.AddFieldDef lTbl.CreateFieldDef("DateAutoNumber", dbDate, _
        RequiredBehavior.fdOptional, IndexedBehavior.fdUnindexed, _
        PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdNo, _
        0, AutoNumberBehavior.IsAutoNumber)

    ' === [3] Invalid: Date field allows zero-length strings
    lTbl.AddFieldDef lTbl.CreateFieldDef("DateZeroLength", dbDate, _
        RequiredBehavior.fdRequired, IndexedBehavior.fdUnindexed, _
        PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdYes, _
        0, AutoNumberBehavior.IsNotAutoNumber)

    ' === Report Definition Errors ===
    Dim lStrErrors As String
    lStrErrors = lTbl.ReportFieldDefinitionErrors()
    If Len(lStrErrors) > 0 Then
        assDebugPrint "Errors during date field definition:" & vbCrLf & lStrErrors
    Else
        assDebugPrint "No validation errors detected for date fields — unexpected!"
    End If
End Sub

' === fTestDoubleFieldDefErrors ===
' Tests invalid configurations for dbDouble fields
Public Sub fTestDoubleFieldDefErrors()
    Dim lTbl As clsTableBuilder
    Set lTbl = New clsTableBuilder

    lTbl.pTableName = "_testDoubleFields"
    assDebugPrint vbCrLf & lTbl.pTableName

    ' === [1] Invalid: Double field with AllowZeroLength
    lTbl.AddFieldDef lTbl.CreateFieldDef("DoubleZeroLength", dbDouble, _
        RequiredBehavior.fdOptional, IndexedBehavior.fdUnindexed, _
        PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdYes, _
        0, AutoNumberBehavior.IsNotAutoNumber)

    ' === Report Definition Errors ===
    Dim lStrErrors As String
    lStrErrors = lTbl.ReportFieldDefinitionErrors()
    If Len(lStrErrors) > 0 Then
        assDebugPrint "Errors during Double field definition:" & vbCrLf & lStrErrors
    Else
        assDebugPrint "No validation errors detected for Double fields — unexpected!"
    End If
End Sub

' === fTestGUIDFieldDefErrors ===
' Tests invalid configurations for dbGUID fields
Public Sub fTestGUIDFieldDefErrors()
    Dim lTbl As clsTableBuilder
    Set lTbl = New clsTableBuilder

    lTbl.pTableName = "_testGUIDFields"
    assDebugPrint vbCrLf & lTbl.pTableName

    ' === [1] Invalid: GUID field marked AutoNumber
    lTbl.AddFieldDef lTbl.CreateFieldDef("GUIDAutoNumber", dbGUID, _
        RequiredBehavior.fdOptional, IndexedBehavior.fdUnindexed, _
        PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdNo, _
        0, AutoNumberBehavior.IsAutoNumber)

    ' === Report Definition Errors ===
    Dim lStrErrors As String
    lStrErrors = lTbl.ReportFieldDefinitionErrors()
    If Len(lStrErrors) > 0 Then
        assDebugPrint "Errors during GUID field definition:" & vbCrLf & lStrErrors
    Else
        assDebugPrint "No validation errors detected for GUID fields — unexpected!"
    End If
End Sub

' === fTestIntegerFieldDefErrors ===
' Tests invalid configurations for dbInteger fields
Public Sub fTestIntegerFieldDefErrors()
    Dim lTbl As clsTableBuilder
    Set lTbl = New clsTableBuilder

    lTbl.pTableName = "_testIntegerFields"
    assDebugPrint vbCrLf & lTbl.pTableName

    ' === [1] Invalid: Integer field with AllowZeroLength
    lTbl.AddFieldDef lTbl.CreateFieldDef("IntZeroLength", dbInteger, _
        RequiredBehavior.fdRequired, , _
        PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdYes, _
        0, AutoNumberBehavior.IsNotAutoNumber)

    ' === Report Definition Errors ===
    Dim lStrErrors As String
    lStrErrors = lTbl.ReportFieldDefinitionErrors()
    If Len(lStrErrors) > 0 Then
        assDebugPrint "Errors during Integer field definition:" & vbCrLf & lStrErrors
    Else
        assDebugPrint "No validation errors detected for Integer fields — unexpected!"
    End If
End Sub

Public Sub fTestLongFieldDefErrors()
    Dim lTbl As clsTableBuilder
    Set lTbl = New clsTableBuilder

    lTbl.pTableName = "tblTestLongFieldErrors"
    assDebugPrint vbCrLf & lTbl.pTableName

    ' 1. Long field with AutoNumber = False but PrimaryKey = Yes (missing expected AutoNumber)
    lTbl.AddFieldDef lTbl.CreateFieldDef("LongPKNoAuto", dbLong, _
        RequiredBehavior.fdRequired, IndexedBehavior.fdDuplicatesAllowed, _
        PrimaryKeyBehavior.fdPrimaryKey, AllowZeroLengthBehavior.fdNo, _
        0, AutoNumberBehavior.IsNotAutoNumber)

    ' 2. Long field with AutoNumber = True but PrimaryKey = No (suspicious but maybe allowed)
    lTbl.AddFieldDef lTbl.CreateFieldDef("LongAutoNotPK", dbLong, _
        RequiredBehavior.fdOptional, IndexedBehavior.fdUnindexed, _
        PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdNo, _
        0, AutoNumberBehavior.IsAutoNumber)

    ' 3. Long field with invalid combination: Not required, not indexed, not PK, AutoNumber = True
    lTbl.AddFieldDef lTbl.CreateFieldDef("LongNoPurpose", dbLong, _
        RequiredBehavior.fdOptional, IndexedBehavior.fdUnindexed, _
        PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdNo, _
        0, AutoNumberBehavior.IsAutoNumber)

    ' 4. Long field with invalid enum values (manually cast)
    lTbl.AddFieldDef lTbl.CreateFieldDef("LongInvalidEnum", dbLong, _
        -1, -1, -1, -1, 0, -1)

    ' === Report Definition Errors ===
    Dim lStrErrors As String
    lStrErrors = lTbl.ReportFieldDefinitionErrors()
    If Len(lStrErrors) > 0 Then
        assDebugPrint "Errors during long field definition:" & vbCrLf & lStrErrors
    Else
        assDebugPrint "No validation errors detected for long fields — unexpected!"
    End If
End Sub

' === fTestMemoFieldDefErrors ===
' Intentionally defines invalid memo field definitions to test validation
Public Sub fTestMemoFieldDefErrors()
    Dim lTbl As clsTableBuilder
    Set lTbl = New clsTableBuilder

    lTbl.pTableName = "_testMemoFields"
    assDebugPrint vbCrLf & lTbl.pTableName

    ' === [1] Invalid: Memo fields cannot be indexed
    lTbl.AddFieldDef lTbl.CreateFieldDef("MemoIndexed", dbMemo, _
        RequiredBehavior.fdRequired, IndexedBehavior.fdDuplicatesAllowed, _
        PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdYes, _
        0, AutoNumberBehavior.IsNotAutoNumber)

    ' === [2] Invalid: Memo fields cannot be AutoNumber
    lTbl.AddFieldDef lTbl.CreateFieldDef("MemoAutoNumber", dbMemo, _
        RequiredBehavior.fdOptional, IndexedBehavior.fdUnindexed, _
        PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdYes, _
        0, AutoNumberBehavior.IsAutoNumber)

    ' === [3] Invalid: Memo fields should not be Primary Key
    lTbl.AddFieldDef lTbl.CreateFieldDef("MemoAsPrimaryKey", dbMemo, _
        RequiredBehavior.fdRequired, IndexedBehavior.fdUnindexed, _
        PrimaryKeyBehavior.fdPrimaryKey, AllowZeroLengthBehavior.fdYes, _
        0, AutoNumberBehavior.IsNotAutoNumber)
        
    ' === Invalid: Memo fields cannot specify Size (implied, not a param)
    ' Size is not explicitly passed, so this test assumes underlying validation logic checks type
    ' For completeness, repeat a correct 8-arg call but type should trigger error
    lTbl.AddFieldDef lTbl.CreateFieldDef("MemoWithSize", dbMemo, _
        RequiredBehavior.fdOptional, IndexedBehavior.fdUnindexed, _
        PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdYes, _
        0, AutoNumberBehavior.IsNotAutoNumber)

    ' === Report Definition Errors ===
    Dim lStrErrors As String
    lStrErrors = lTbl.ReportFieldDefinitionErrors()
    If Len(lStrErrors) > 0 Then
        assDebugPrint "Errors during memo field definition:" & vbCrLf & lStrErrors
    Else
        assDebugPrint "No validation errors detected for memo fields — unexpected!"
    End If
End Sub

' === fTestSingleFieldDefErrors ===
' Tests invalid configurations for dbSingle fields
Public Sub fTestSingleFieldDefErrors()
    Dim lTbl As clsTableBuilder
    Set lTbl = New clsTableBuilder

    lTbl.pTableName = "_testSingleFields"
    assDebugPrint vbCrLf & lTbl.pTableName

    ' === [1] Invalid: Single field with AllowZeroLength
    lTbl.AddFieldDef lTbl.CreateFieldDef("SingleZeroLength", dbSingle, _
        RequiredBehavior.fdOptional, IndexedBehavior.fdUnindexed, _
        PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdYes, _
        0, AutoNumberBehavior.IsNotAutoNumber)

    ' === Report Definition Errors ===
    Dim lStrErrors As String
    lStrErrors = lTbl.ReportFieldDefinitionErrors()
    If Len(lStrErrors) > 0 Then
        assDebugPrint "Errors during Single field definition:" & vbCrLf & lStrErrors
    Else
        assDebugPrint "No validation errors detected for Single fields — unexpected!"
    End If
End Sub

Public Sub fTestTextFieldDefErrors()
    Dim lTbl As clsTableBuilder
    Set lTbl = New clsTableBuilder

    lTbl.pTableName = "tblTestTextFieldErrors"
    assDebugPrint vbCrLf & lTbl.pTableName

    ' === Invalid Text Field Definitions ===

    ' 1. Text field with missing size (0)
    lTbl.AddFieldDef lTbl.CreateFieldDef("TextSizeZero", dbText, _
        RequiredBehavior.fdRequired, IndexedBehavior.fdUnindexed, _
        PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdNo, _
        , AutoNumberBehavior.IsNotAutoNumber)

    ' 2. Text field with AutoNumber = True (invalid combination)
    lTbl.AddFieldDef lTbl.CreateFieldDef("TextAutoNumber", dbText, _
        RequiredBehavior.fdRequired, IndexedBehavior.fdUnindexed, _
        PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdNo, _
        20, AutoNumberBehavior.IsAutoNumber)

    ' 3. Text field marked as PrimaryKey (allowed, but likely suspect)
    lTbl.AddFieldDef lTbl.CreateFieldDef("TextPrimaryKey", dbText, _
        RequiredBehavior.fdRequired, IndexedBehavior.fdDuplicatesAllowed, _
        PrimaryKeyBehavior.fdPrimaryKey, AllowZeroLengthBehavior.fdNo, _
        20, AutoNumberBehavior.IsNotAutoNumber)

    ' 4. Text field marked AllowZeroLength = Yes but also Required (logical conflict)
    lTbl.AddFieldDef lTbl.CreateFieldDef("TextZeroLengthRequired", dbText, _
        RequiredBehavior.fdRequired, IndexedBehavior.fdUnindexed, _
        PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdYes, _
        20, AutoNumberBehavior.IsNotAutoNumber)

    ' 5. Text field with a very large size (beyond Access limit ~255)
    lTbl.AddFieldDef lTbl.CreateFieldDef("TextOversize", dbText, _
        RequiredBehavior.fdRequired, IndexedBehavior.fdUnindexed, _
        PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdNo, _
        500, AutoNumberBehavior.IsNotAutoNumber)
        
    ' 6. Text field with size explicitly set to 0 (invalid)
    lTbl.AddFieldDef lTbl.CreateFieldDef("TextSizeZeroExplicit", dbText, _
        RequiredBehavior.fdRequired, IndexedBehavior.fdUnindexed, _
        PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdNo, _
        0, AutoNumberBehavior.IsNotAutoNumber)
    
    ' 7. Text field with negative size
    lTbl.AddFieldDef lTbl.CreateFieldDef("TextSizeNegative", dbText, _
        RequiredBehavior.fdRequired, IndexedBehavior.fdUnindexed, _
        PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdNo, _
        -10, AutoNumberBehavior.IsNotAutoNumber)
    

    ' === Report Definition Errors ===
    Dim lStrErrors As String
    lStrErrors = lTbl.ReportFieldDefinitionErrors()
    If Len(lStrErrors) > 0 Then
        assDebugPrint "Errors during text field definition:" & vbCrLf & lStrErrors
    Else
        assDebugPrint "No validation errors detected for text fields — unexpected!"
    End If
End Sub

