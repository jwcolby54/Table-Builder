Attribute VB_Name = "basFieldDataTestRules"
Option Compare Database
Option Explicit
'****************************************************************************************
'Public Sub: fTestValueValidationTypeMismatch
'
' Description:
'     Executes a suite of test cases to verify that clsFieldDef properly enforces type
'     constraints on assigned values. Attempts to set intentionally mismatched values
'     (e.g., setting a string into a numeric field), and checks that validation fails.
'
'     This test ensures that clsFieldDef.ValidateValue detects type mismatches and
'     raises the appropriate evDataValidationError events, which clsTableBuilder handles.
'
' Parameters:
'     (None)
'
' Returns:
'     None
'
' Side Effects:
'     May output test results to Debug window or raise runtime events.
'
' Test Scope:
'     - Assigns incorrect data types to various field types
'     - Verifies that validation logic rejects them correctly
'****************************************************************************************
'
Public Sub fTestValueValidationTypeMismatch()
    Dim lTbl As clsTableBuilder
    Set lTbl = New clsTableBuilder

    lTbl.pTableName = "_testValueMismatch"

    ' Define fields with known types
    lTbl.AddFieldDef lTbl.CreateFieldDef("TestDate", dbDate, _
        RequiredBehavior.fdRequired, IndexedBehavior.fdUnindexed, _
        PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdNo, _
        0, AutoNumberBehavior.IsNotAutoNumber)

    lTbl.AddFieldDef lTbl.CreateFieldDef("TestLong", dbLong, _
        RequiredBehavior.fdRequired, IndexedBehavior.fdUnindexed, _
        PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdNo, _
        0, AutoNumberBehavior.IsNotAutoNumber)

    lTbl.AddFieldDef lTbl.CreateFieldDef("TestCurrency", dbCurrency, _
        RequiredBehavior.fdRequired, IndexedBehavior.fdUnindexed, _
        PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdNo, _
        0, AutoNumberBehavior.IsNotAutoNumber)

    lTbl.AddFieldDef lTbl.CreateFieldDef("TestBoolean", dbBoolean, _
        RequiredBehavior.fdRequired, IndexedBehavior.fdUnindexed, _
        PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdNo, _
        0, AutoNumberBehavior.IsNotAutoNumber)

    lTbl.AddFieldDef lTbl.CreateFieldDef("TestText", dbText, _
        RequiredBehavior.fdRequired, IndexedBehavior.fdUnindexed, _
        PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdYes, _
        20, AutoNumberBehavior.IsNotAutoNumber)

    ' Create the table before attempting to seed
    lTbl.CreateTable True

    ' Seed with deliberately invalid values
    Dim colSeedData As Collection
    Set colSeedData = New Collection
    colSeedData.Add Array("not a date", "abc", "wrong currency", "maybe", 12345) ' invalid types

    ' Attempt to seed, triggering validation logic
    Dim lRst As DAO.Recordset
    Dim lLngRow As Long, lLngFieldIndex As Long
    Dim lFldDef As clsFieldDef
    Dim lVntRow As Variant, lVntValue As Variant
    Dim lStrErrors As String
    Dim lBlnHasErrors As Boolean

    Set lRst = CurrentDb.OpenRecordset("_testValueMismatch", dbOpenDynaset)

    For lLngRow = 1 To colSeedData.Count
        lVntRow = colSeedData(lLngRow)
        lRst.AddNew
        lBlnHasErrors = False
        lStrErrors = ""

    lLngFieldIndex = 0
    For Each lFldDef In lTbl.pFieldDefs ' ? Assuming you expose mColFieldDefs via a Property Get
        '
        'Set mClsCurrentField so that we can sink any event from that field
        '
        Set lTbl.pClsCurrentField = lFldDef

        lVntValue = lVntRow(lLngFieldIndex)
    
        If Not lFldDef.ValidateValue(, lVntValue) Then
            lStrErrors = lStrErrors & "- Row " & lLngRow & ", Field '" & lFldDef.pName & "': " & lFldDef.pValidationError() & vbCrLf
            assDebugPrint lStrErrors
            lBlnHasErrors = True
        Else
            assDebugPrint "No error detected for " & lFldDef.pName
            lRst(lFldDef.pName).Value = lVntValue
        End If
    
        lLngFieldIndex = lLngFieldIndex + 1
    Next lFldDef

        If lBlnHasErrors Then
            lRst.CancelUpdate
            assDebugPrint "Validation errors in row " & lLngRow & ":" & vbCrLf & lStrErrors
        Else
            lRst.Update
        End If
    Next lLngRow

    lRst.Close
End Sub


