VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTableBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
' ============================================================================
' Class: clsTableBuilder
' Purpose: Central class responsible for building tables in a metadata-driven
'          fashion using field definitions encapsulated in clsFieldDef objects.
'
' Key Responsibilities:
'   - Store metadata for a table and its fields.
'   - Construct Access tables dynamically using DAO.
'   - Support seeding data, field formatting, and structural logging.
'
' Collaborates With:
'   - clsFieldDef for field metadata
'   - basEnums for constants
'   - basTblBuilder for predefined tables
'
' Usage:
'   Instantiate, add field definitions via AddFieldDef(), then call Build().
' ============================================================================

Private mStrTableName As String
Private mColFields As Collection
Private mColFieldsWithErrors As Collection
Private mBlnLog As Boolean
Private mBlnAutoDeleteExisting As Boolean
Private WithEvents mClsCurrentField As clsFieldDef
Attribute mClsCurrentField.VB_VarHelpID = -1

Private mColValidationErrors As Collection
'
'****************************************************************************************
'Private Sub: mClsCurrentField_evDataValidationError
'
' Description:
'     Event handler that receives detailed validation errors from the current clsFieldDef
'     instance when data validation fails (e.g., incompatible values for field type or rules).
'
'     This event is triggered internally by clsFieldDef.ValidateValue and is automatically
'     wired using the WithEvents reference to mClsCurrentField.
'
' Parameters:
'     FieldName As String
'         - The name of the field that failed validation.
'
'     ErrorMessage As String
'         - A high-level error description summarizing the validation failure.
'
'     ValidationErrors As Collection
'         - A collection of detailed error strings specific to the field value or rules.
'
' Returns:
'     None
'
' Side Effects:
'     Stores the collection of validation errors in mColValidationErrors, keyed by FieldName.
'****************************************************************************************
'
Private Sub mClsCurrentField_evDataValidationError(ByVal FieldName As String, ByVal ErrorMessage As String, ByVal ValidationErrors As Collection)
    ' Store the full error collection keyed by field name
    On Error Resume Next ' In case key already exists
    mColValidationErrors.Add ValidationErrors, FieldName
    On Error GoTo 0
End Sub
'****************************************************************************************
'Private Sub: mClsCurrentField_evFieldValidationError
'
' Description:
'     Handles field definition validation errors raised by the current clsFieldDef instance.
'     This event is triggered by ValidateDefinition when a field is improperly configured.
'
'     It ensures that clsTableBuilder is aware of any invalid definitions prior to
'     attempting to create or seed a table.
'
' Parameters:
'     FieldName As String
'         - Name of the offending field.
'
'     ErrorMessage As String
'         - Description of the configuration issue encountered.
'
' Returns:
'     None
'
' Side Effects:
'     Adds the clsFieldDef instance to mColFieldsWithErrors.
'****************************************************************************************
'
Private Sub mClsCurrentField_evFieldValidationError(ByVal FieldName As String, ByVal ErrorMessage As String, ByVal ValidationErrors As Collection)
    ' Store the full error collection keyed by field name
    On Error Resume Next ' In case key already exists
    mColValidationErrors.Add ValidationErrors, FieldName
    On Error GoTo 0
End Sub

' === Properties ===
Public Property Set pClsCurrentField(lclsFldDef As clsFieldDef)
    Set mClsCurrentField = lclsFldDef
End Property
Public Property Get pTableName() As String
    pTableName = mStrTableName
End Property

Public Property Let pTableName(ByVal v As String)
    mStrTableName = v
End Property

Public Property Get pLog() As Boolean
    pLog = mBlnLog
End Property

Public Property Let pLog(ByVal v As Boolean)
    mBlnLog = v
End Property

Public Property Get pAutoDeleteExisting() As AutoDeleteBehavior
    If mBlnAutoDeleteExisting Then
        pAutoDeleteExisting = fdDeleteIfExists
    Else
        pAutoDeleteExisting = fdPreserveExisting
    End If
End Property

Public Property Let pAutoDeleteExisting(ByVal v As AutoDeleteBehavior)
    mBlnAutoDeleteExisting = (v = fdDeleteIfExists)
End Property
Public Property Get pFieldDefs() As Collection
    Set pFieldDefs = mColFields
End Property
' === Initialization ===
'****************************************************************************************
'Private Sub: Class_Initialize
'
' Description:
'     Initializes a new instance of clsTableBuilder. This constructor sets up the internal
'     collections used to manage field definitions and validation errors, and ensures that
'     WithEvents binding for mClsCurrentField is ready for validation event wiring.
'
' Parameters:
'     (None)
'
' Returns:
'     None
'
' Side Effects:
'     Initializes collections: mColFields, mColFieldsWithErrors, mColValidationErrors.
'****************************************************************************************
'
Private Sub Class_Initialize()
    Set mColFields = New Collection
    Set mColFieldsWithErrors = New Collection
End Sub

'****************************************************************************************
'Public Sub: AddFieldDef
'
' Description:
'     Adds a clsFieldDef instance to the internal collection (mColFields) of field definitions
'     that will be used to build or validate a table. This method is typically called from
'     table-building procedures such as Build_TablePeople in basTblBuilder.
'
'     If the field being added is the current focus, it is also stored in mClsCurrentField,
'     which enables automatic event handling for validation feedback.
'
' Parameters:
'     lFldDef As clsFieldDef
'         - A fully configured field definition ready for validation and use.
'
' Returns:
'     None
'
' Side Effects:
'     Adds the object to mColFields.
'     Sets mClsCurrentField to lFldDef to enable WithEvents validation tracking.
'****************************************************************************************

' === Add FieldDef ===
Public Sub AddFieldDef(ByVal lFldDef As clsFieldDef)
    '
    'Set mClsCurrentField so that we can sink any event from that field
    '
    Set mClsCurrentField = lFldDef
    If lFldDef.ValidateDefinition() Then
        mColFields.Add lFldDef, lFldDef.pName
    Else
        mColFieldsWithErrors.Add lFldDef, lFldDef.pName
    End If
End Sub

' === Validation ===
'****************************************************************************************
'Private Function: ValidateOneField
'
' Description:
'     Validates a single field definition by calling its ValidateDefinition method.
'     If the definition fails, this method will raise the evFieldValidationError event
'     and add the field to mColFieldsWithErrors.
'
'     This is a helper method called internally by ValidateAllFields.
'
' Parameters:
'     lFld As clsFieldDef
'         - The field definition to validate.
'
' Returns:
'     String
'         - Returns an empty string if valid, or an error message describing the problem.
'
' Side Effects:
'     Adds invalid fields to mColFieldsWithErrors.
'****************************************************************************************
'
Private Function ValidateOneField(ByVal lFld As clsFieldDef) As String
    On Error GoTo HandleError
    lFld.ValidateDefinition
    Exit Function
HandleError:
    ValidateOneField = "Field '" & lFld.pName & "': " & Err.Description
    Err.Clear
End Function
'****************************************************************************************
'Public Function: ValidateAllFields
'
' Description:
'     Iterates through all field definitions added to the builder and validates each one
'     using ValidateOneField. If any definitions fail, their errors are accumulated.
'
'     This method ensures that the table schema is logically valid before proceeding
'     with creation or data seeding.
'
' Parameters:
'     (None)
'
' Returns:
'     String
'         - Empty string if all fields are valid; otherwise, a summary of all validation errors.
'
' Side Effects:
'     Populates mColFieldsWithErrors and mColValidationErrors.
'****************************************************************************************
'
Public Function ValidateAllFields() As String
    Dim lFld As clsFieldDef
    Dim lStrResult As String
    Dim lStrMessage As String

    For Each lFld In mColFields
        lStrMessage = ValidateOneField(lFld)
        If Len(lStrMessage) > 0 Then
            lStrResult = lStrResult & lStrMessage & vbCrLf
        End If
    Next

    ValidateAllFields = lStrResult
End Function
'
'****************************************************************************************
'Public Function: ReportFieldDefinitionErrors
'
' Description:
'     Constructs a human-readable summary of validation errors collected from
'     the field definitions in mColFieldsWithErrors and mColValidationErrors.
'
'     This is used to output useful diagnostics if table creation fails due to
'     invalid field configurations.
'
' Parameters:
'     (None)
'
' Returns:
'     String
'         - Multi-line string summarizing all definition and value errors per field.
'****************************************************************************************
'
Public Function ReportFieldDefinitionErrors() As String
    Dim lStrResult As String
    Dim lFldDef As clsFieldDef
    Dim lErr As Variant

    For Each lFldDef In mColFieldsWithErrors
        For Each lErr In lFldDef.pValidationErrors
            lStrResult = lStrResult & "- " & lFldDef.pName & ": " & lErr & vbCrLf
        Next
    Next

    ReportFieldDefinitionErrors = lStrResult
End Function
'
'****************************************************************************************
'Public Sub: CreateTable
'
' Description:
'     Creates a new DAO table using the metadata-driven field definitions added via AddFieldDef.
'     This method optionally deletes an existing table of the same name based on the value
'     of the lDeleteBehavior parameter.
'
'     The method configures a DAO.TableDef, iterates through the field collection,
'     and adds each field. Optionally logs output using Debug.Print or similar.
'
' Parameters:
'     lDeleteBehavior As AutoDeleteBehavior
'         - Enum controlling whether to delete an existing table first.
'
'     lBlnLog As Variant
'         - Optional flag to enable logging of the process. Null means follow internal flag.
'
' Returns:
'     None
'
' Side Effects:
'     May delete and recreate a table. Logs if enabled.
'****************************************************************************************
' === Create Table ===

Public Function CreateTable( _
    Optional ByVal lDeleteBehavior As AutoDeleteBehavior = fdDeleteIfExists, _
    Optional ByVal lBlnLog As Variant = Null) As Boolean

On Error GoTo CreateTable_Error
Dim lBlnDelete As Boolean
Dim lBlnUseLog As Boolean
Dim db As DAO.Database
Dim tdf As DAO.TableDef
Dim fld As DAO.Field
Dim idx As DAO.Index
Dim fldDef As clsFieldDef
Dim strTableName As String
Dim strFieldName As String
Dim lDaoIndex As DAO.Index
Dim lClsPKField As clsFieldDef
Dim lClsFieldDef As clsFieldDef
Dim lIntPKCount As Integer

    lBlnDelete = (lDeleteBehavior = fdDeleteIfExists) Or mBlnAutoDeleteExisting
    lBlnUseLog = IIf(IsNull(lBlnLog), mBlnLog, lBlnLog)

    strTableName = mStrTableName
    Set db = CurrentDb

    If lBlnDelete Then
        On Error Resume Next
        db.TableDefs.Delete strTableName
        Err.Clear
On Error GoTo CreateTable_Error
        If lBlnUseLog Then assDebugPrint "Deleted table: " & strTableName
    End If

    Set tdf = db.CreateTableDef(strTableName)

    For Each fldDef In mColFields
        Set fld = tdf.CreateField(fldDef.pName, fldDef.pType, fldDef.pSize)
        fld.Required = fldDef.pIsRequired
        fld.Attributes = fldDef.pAttributes
        If fld.Type = dbText Or fld.Type = dbMemo Then
            fld.AllowZeroLength = fldDef.pAllowZeroLength
        End If

        If fld.Type = dbText Or fld.Type = dbMemo Then
            fld.AllowZeroLength = fldDef.pAllowZeroLength
        End If
        
        If Not IsNull(fldDef.pDefaultValue) Then
            fld.DefaultValue = fldDef.pDefaultValue
        End If

        If Len(fldDef.pValidationRule & "") > 0 Then
            fld.ValidationRule = fldDef.pValidationRule
        End If

        If Len(fldDef.pValidationText & "") > 0 Then
            fld.ValidationText = fldDef.pValidationText
        End If

        If fldDef.pIsAutoNumber Then
            fld.Attributes = fld.Attributes Or dbAutoIncrField
        End If

        tdf.Fields.Append fld

        If lBlnUseLog Then assDebugPrint "  Field: " & fld.Name & " (" & fld.Type & ")"
    Next

    ' Handle composite primary keys
    Set lDaoIndex = tdf.CreateIndex("PrimaryKey")
    For Each lClsPKField In mColFields
        If lClsPKField.pIsPK Then
            lDaoIndex.Fields.Append lDaoIndex.CreateField(lClsPKField.pName)
            lIntPKCount = lIntPKCount + 1
        End If
    Next lClsPKField

    If lIntPKCount > 0 Then
        lDaoIndex.Primary = True
        tdf.Indexes.Append lDaoIndex
    End If

    ' Add secondary indexes
    For Each lClsFieldDef In mColFields
        If lClsFieldDef.pIsIndexed And Not lClsFieldDef.pIsPK Then
            Set lDaoIndex = tdf.CreateIndex("idx_" & lClsFieldDef.pName)
            lDaoIndex.Fields.Append lDaoIndex.CreateField(lClsFieldDef.pName)
            tdf.Indexes.Append lDaoIndex
        End If
    Next lClsFieldDef

    db.TableDefs.Append tdf
    If lBlnUseLog Then assDebugPrint "Created table: " & strTableName
    If tdf.Fields.Count = mColFields.Count Then
        CreateTable = True
    End If
    
Exit_CreateTable:
    On Error GoTo 0
    Exit Function

CreateTable_Error:
Dim strErrMsg As String
    Select Case Err
    Case 0      'insert Errors you wish to ignore here
        Resume Next
    Case Else   'All other errors will trap
        strErrMsg = "Error " & Err.Number & " (" & Err.Description & ") in procedure SecondDatabase.clsTableBuilder.CreateTable, line " & Erl & "."
        Beep
#If boolELE = 1 Then
        WriteErrorLog strErrMsg
#End If
        assDebugPrint strErrMsg
        Resume Exit_CreateTable
    End Select
    Resume Exit_CreateTable
    Resume 0    'FOR TROUBLESHOOTING
End Function
'
'****************************************************************************************
'Public Sub: SeedRows
'
' Description:
'     Inserts a collection of seed data into the table defined and created by this builder.
'     Each item in lColSeedData must be a Dictionary or similar key-value collection
'     matching the field names in the table.
'
'     Used to populate lookup or test tables after structural creation.
'
' Parameters:
'     lColSeedData As Collection
'         - Collection of row dictionaries with field-value pairs for insertion.
'
' Returns:
'     None
'
' Side Effects:
'     Writes data to the table.
'****************************************************************************************
'
Public Sub SeedRows(ByVal lColSeedData As Collection)
    On Error GoTo SeedRows_Error
    Dim lRst As DAO.Recordset
    Dim lLngRow As Long, lLngFieldIndex As Long, lLngArrayIndex As Long
    Dim lFldDef As clsFieldDef
    Dim lVntRow As Variant, lVntValue As Variant
    Dim lBlnHasErrors As Boolean
    Dim lStrErrors As String

    Set lRst = CurrentDb.OpenRecordset(pTableName, dbOpenDynaset)
    
    For lLngRow = 1 To lColSeedData.Count
        '
        'Arrays are zero based so use an array index
        '
        lLngArrayIndex = 0
        lVntRow = lColSeedData(lLngRow)
        lBlnHasErrors = False
        lStrErrors = ""

        lRst.AddNew
        'The first item in the collection is a PK so
        'skip it since we can't seed that value
        For lLngFieldIndex = 1 To mColFields.Count  ' - 1
            Set lFldDef = mColFields(lLngFieldIndex)
            If lFldDef.pIsAutoNumber Then
                '
            Else
                '
                'The values are passed in as an array
                'Arrays are zero based so we have to
                'use an array specific index
                lVntValue = lVntRow(lLngArrayIndex)
    
                If Not lFldDef.ValidateValue(, lVntValue) Then
                    lBlnHasErrors = True
                    'lStrErrors = lStrErrors & "- Row " & lLngRow & ", Field '" & lFldDef.pName & "': " & lFldDef.GetValidationErrors & vbCrLf
                Else
                    lRst(lFldDef.pName).Value = lVntValue
                End If
                lLngArrayIndex = lLngArrayIndex + 1
            End If
        Next lLngFieldIndex

        If lBlnHasErrors Then
            lRst.CancelUpdate
            assDebugPrint "Validation errors in row " & lLngRow & ":" & vbCrLf & lStrErrors
        Else
            lRst.Update
        End If
    Next lLngRow

    lRst.Close
    
Exit_SeedRows:
    On Error GoTo 0
    Exit Sub

SeedRows_Error:
Dim strErrMsg As String
    Select Case Err
    Case 0      'insert Errors you wish to ignore here
        Resume Next
    Case Else   'All other errors will trap
        strErrMsg = "Error " & Err.Number & " (" & Err.Description & ") in procedure SecondDatabase.clsTableBuilder.SeedRows, line " & Erl & "."
        Beep
#If boolELE = 1 Then
        WriteErrorLog strErrMsg
#End If
        assDebugPrint strErrMsg
        Resume Exit_SeedRows
    End Select
    Resume Exit_SeedRows
    Resume 0    'FOR TROUBLESHOOTING
End Sub
'
'****************************************************************************************
'Private Sub: LogStructure
'
' Description:
'     Logs the structure of the current table to the Immediate window for review.
'     Includes the table name and each field’s name, type, and key attributes.
'
' Parameters:
'     (None)
'
' Returns:
'     None
'
' Side Effects:
'     Prints output to the Debug window.
'****************************************************************************************
'
Private Sub LogStructure()
    Dim lClsFieldDef As clsFieldDef

    For Each lClsFieldDef In mColFields
        assDebugPrint "Field: " & lClsFieldDef.pName
        assDebugPrint "  Type: " & fFieldTypeName(lClsFieldDef.pType)
        If lClsFieldDef.pSize > 0 Then assDebugPrint "  Size: " & lClsFieldDef.pSize
        If lClsFieldDef.pIsRequired Then assDebugPrint "  Required: Yes"
        If Not IsNull(lClsFieldDef.pDefaultValue) Then assDebugPrint "  Default: " & lClsFieldDef.pDefaultValue
        If lClsFieldDef.pIsAutoNumber Then assDebugPrint "  AutoNumber: Yes"
        If lClsFieldDef.pIsPK Then assDebugPrint "  Primary Key: Yes"
        If lClsFieldDef.pIsIndexed Then assDebugPrint "  Indexed: Yes"
        assDebugPrint ""
    Next lClsFieldDef
End Sub
'
'****************************************************************************************
'Private Function: fFieldTypeName
'
' Description:
'     Converts a DAO data type integer constant to a human-readable name, e.g., 10 ? "Text".
'     This is used for logging or reporting purposes.
'
' Parameters:
'     lIntType As Integer
'         - DAO field type constant (e.g., dbText, dbLong).
'
' Returns:
'     String
'         - Descriptive name of the data type.
'****************************************************************************************
'
Private Function fFieldTypeName(ByVal lIntType As Integer) As String
    Select Case lIntType
        Case dbText: fFieldTypeName = "Text"
        Case dbMemo: fFieldTypeName = "Memo"
        Case dbLong: fFieldTypeName = "Long"
        Case dbDate: fFieldTypeName = "Date"
        Case dbBoolean: fFieldTypeName = "Yes/No"
        Case dbCurrency: fFieldTypeName = "Currency"
        Case dbInteger: fFieldTypeName = "Integer"
        Case dbGUID: fFieldTypeName = "GUID"
        Case dbDouble: fFieldTypeName = "Double"
        Case dbSingle: fFieldTypeName = "Single"
        Case dbByte: fFieldTypeName = "Byte"
        Case dbDecimal: fFieldTypeName = "Decimal"
        Case dbBinary: fFieldTypeName = "Binary"
        Case dbVarBinary: fFieldTypeName = "VarBinary"
        Case dbBigInt: fFieldTypeName = "BigInt"
        Case Else: fFieldTypeName = "Unknown"
    End Select
End Function
'
'****************************************************************************************
'Public Function: CreateFieldDef
'
' Description:
'     Creates and returns a configured instance of clsFieldDef using common parameters.
'     This factory-style function simplifies adding field definitions by encapsulating
'     all key setup steps in one place.
'
'     This function is often used in table builders (e.g., basTblBuilder) to reduce
'     repetitive code and improve readability.
'
' Parameters:
'     lStrFieldName As String
'         - Name of the field.
'
'     lIntFieldType As Integer
'         - DAO field type (e.g., dbText, dbLong).
'
'     lRequired As RequiredBehavior
'         - Whether the field must contain a non-null value.
'
'     lIndexed As IndexedBehavior
'         - Whether the field is indexed.
'
'     lPrimaryKey As PrimaryKeyBehavior
'         - Whether the field is the primary key.
'
'     lAllowZeroLength As AllowZeroLengthBehavior
'         - For text fields, whether zero-length strings are allowed.
'
'     lIntFieldSize As Long
'         - Length of text fields. Ignored for non-text fields.
'
'     lIsAutoNumber As AutoNumberBehavior
'         - Whether the field should be AutoNumber.
'
' Returns:
'     clsFieldDef
'         - A fully initialized and configured field definition object.
'
' Side Effects:
'     None directly, but typically followed by AddFieldDef.
'****************************************************************************************
'
Public Function CreateFieldDef( _
    ByVal lStrFieldName As String, _
    ByVal lIntFieldType As Integer, _
    Optional ByVal lRequired As RequiredBehavior = fdRequired, _
    Optional ByVal lIndexed As IndexedBehavior = fdUnindexed, _
    Optional ByVal lPrimaryKey As PrimaryKeyBehavior = fdNonPK, _
    Optional ByVal lAllowZeroLength As AllowZeroLengthBehavior = fdNo, _
    Optional ByVal lIntFieldSize As Long = 50, _
    Optional ByVal lIsAutoNumber As AutoNumberBehavior = IsNotAutoNumber) As clsFieldDef

    Dim lFldDef As New clsFieldDef
    lFldDef.pName = lStrFieldName
    lFldDef.pType = lIntFieldType
    lFldDef.pSize = lIntFieldSize

    lFldDef.pIsPK = (lPrimaryKey = fdPrimaryKey)
    lFldDef.pIsRequired = (lRequired = fdRequired)
    lFldDef.pIsIndexed = ((lIndexed = fdDuplicatesAllowed Or lIndexed = fdNoDuplicates) <> 0)
    lFldDef.pAllowZeroLength = (lAllowZeroLength = fdYes)
    lFldDef.pIsAutoNumber = (lIsAutoNumber = IsAutoNumber)

    Set CreateFieldDef = lFldDef
End Function



