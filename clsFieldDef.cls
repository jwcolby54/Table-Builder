VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFieldDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
' ============================================================================
' Class: clsFieldDef
' Purpose: Encapsulates metadata for a single field in a table. Handles
'          datatype logic, validation constraints, and field-level settings.
'
' Key Responsibilities:
'   - Define field name, type, size, default value, and validation rules.
'   - Enforce logical consistency of field configuration.
'   - Generate field creation syntax and support runtime validation.
'
' Collaborates With:
'   - clsTableBuilder for table assembly
'   - basEnums for datatype and validation constants
'   - basFieldTypeTestRules for config validation
'   - basFieldDataTestRules for value validation
' ============================================================================

' ===============================
' clsFieldDef.cls
' ===============================
' This class holds the definition and metadata for a single DAO field used in dynamic table creation.
' It includes value validation with support for multiple error messages via a collection.
' A ValidationError event is raised if validation fails.

Private mStrName As String
Private mIntType As Integer
Private mIntSize As Integer
Private mBlnIsAutoNumber As Boolean
Private mBlnIsIndexed As Boolean
Private mBlnIsPK As Boolean
Private mBlnIsRequired As Boolean
Private mVarDefaultValue As Variant
Private mVarDataValue As Variant

' Extended Properties
Private mBlnAllowZeroLength As Boolean
Private mStrValidationRule As String
Private mStrValidationText As String
Private mIntOrdinalPosition As Integer
Private mLngAttributes As Long
Private mIntPrecision As Integer
Private mIntScale As Integer
Private mIntCollatingOrder As Integer

' Validation
Private mColValidationErrors As Collection

' === Events ===
Public Event evFieldValidationError(ByVal FieldName As String, ByVal ErrorMessage As String, ByVal ValidationErrors As Collection)
Public Event evDataValidationError(ByVal FieldName As String, ByVal ErrorMessage As String, ByVal ValidationErrors As Collection)
'****************************************************************************************
'Private Sub: Class_Initialize
'
' Description:
'     This is the class constructor, automatically invoked when a new instance of clsFieldDef
'     is created. It is responsible for initializing internal state to default values.
'     This includes setting boolean flags (e.g., Required, Indexed) to False and ensuring
'     that variant properties such as default value and data value are empty.
'
'     The constructor ensures a clean starting state for each field definition so that
'     the consuming clsTableBuilder can explicitly configure field properties before
'     using them to create or validate a table schema.
'
' Parameters:
'     (None)
'
' Returns:
'     None
'
' Side Effects:
'     Initializes all internal member variables to their respective defaults.
'****************************************************************************************
'
Private Sub Class_Initialize()
    Set mColValidationErrors = New Collection
End Sub
'****************************************************************************************
'Private Sub: Class_Terminate
'
' Description:
'     This is the class destructor, automatically called when an instance of clsFieldDef
'     is released from memory. Although not commonly used for releasing external resources
'     in this class, it is provided for completeness and future-proofing.
'
'     It can be used for cleanup or debugging purposes if necessary, such as logging
'     destruction of the object or explicitly clearing collections.
'
' Parameters:
'     (None)
'
' Returns:
'     None
'
' Side Effects:
'     None by default, but may be used for cleanup or instrumentation.
'****************************************************************************************
'
Private Sub Class_Terminate()
    Set mColValidationErrors = Nothing
End Sub

Public Property Get pAllowZeroLength() As Boolean
    pAllowZeroLength = mBlnAllowZeroLength
End Property

Public Property Let pAllowZeroLength(ByVal lBlnValue As Boolean)
    mBlnAllowZeroLength = lBlnValue
End Property

Public Property Get pAttributes() As Long
    pAttributes = mLngAttributes
End Property

Public Property Let pAttributes(ByVal lLngValue As Long)
    mLngAttributes = lLngValue
End Property

Public Property Get pCollatingOrder() As Integer
    pCollatingOrder = mIntCollatingOrder
End Property

Public Property Let pCollatingOrder(ByVal lIntValue As Integer)
    mIntCollatingOrder = lIntValue
End Property

Public Property Get pDataValue() As Variant
    pDataValue = mVarDataValue
End Property

Public Property Let pDataValue(ByVal lVarValue As Variant)
    Dim lStrError As String
    If Not ValidateValue(lStrError, lVarValue) Then
        RaiseEvent evFieldValidationError(mStrName, lStrError, mColValidationErrors)
        Err.Raise vbObjectError + 1001, "clsFieldDef.pDataValue", "Validation failed: " & lStrError
    End If
    mVarDataValue = lVarValue
End Property

Public Property Get pDefaultValue() As Variant
    pDefaultValue = mVarDefaultValue
End Property

Public Property Let pDefaultValue(ByVal lVarValue As Variant)
    mVarDefaultValue = lVarValue
End Property

Public Property Get pHasValidationError() As Boolean
    pHasValidationError = (mColValidationErrors.Count > 0)
End Property

Public Property Get pIsAutoNumber() As Boolean
    pIsAutoNumber = mBlnIsAutoNumber
End Property

Public Property Let pIsAutoNumber(ByVal lBlnValue As Boolean)
    mBlnIsAutoNumber = lBlnValue
End Property

Public Property Get pIsIndexed() As Boolean
    pIsIndexed = mBlnIsIndexed
End Property

Public Property Let pIsIndexed(ByVal lBlnValue As Boolean)
    mBlnIsIndexed = lBlnValue
End Property

Public Property Get pIsPK() As Boolean
    pIsPK = mBlnIsPK
End Property

Public Property Let pIsPK(ByVal lBlnValue As Boolean)
    mBlnIsPK = lBlnValue
End Property

Public Property Get pIsRequired() As Boolean
    pIsRequired = mBlnIsRequired
End Property

Public Property Let pIsRequired(ByVal lBlnValue As Boolean)
    mBlnIsRequired = lBlnValue
End Property

Public Property Get pName() As String
    pName = mStrName
End Property

Public Property Let pName(ByVal lStrValue As String)
    mStrName = lStrValue
End Property

Public Property Get pOrdinalPosition() As Integer
    pOrdinalPosition = mIntOrdinalPosition
End Property

Public Property Let pOrdinalPosition(ByVal lIntValue As Integer)
    mIntOrdinalPosition = lIntValue
End Property

Public Property Get pPrecision() As Integer
    pPrecision = mIntPrecision
End Property

Public Property Let pPrecision(ByVal lIntValue As Integer)
    mIntPrecision = lIntValue
End Property

Public Property Get pScale() As Integer
    pScale = mIntScale
End Property

Public Property Let pScale(ByVal lIntValue As Integer)
    mIntScale = lIntValue
End Property

Public Property Get pSize() As Integer
    pSize = mIntSize
End Property

Public Property Let pSize(ByVal lIntValue As Integer)
    mIntSize = lIntValue
End Property

Public Property Get pType() As Integer
    pType = mIntType
End Property

Public Property Let pType(ByVal lIntValue As Integer)
    mIntType = lIntValue
End Property

Public Property Get pValidationError() As String
    Dim lStr As String
    Dim lVar As Variant
    For Each lVar In mColValidationErrors
        lStr = lStr & lVar & vbCrLf
    Next
    If Len(lStr) > 0 Then lStr = Left(lStr, Len(lStr) - 2)
    pValidationError = lStr
End Property
Public Property Get pValidationErrors() As Collection
    Set pValidationErrors = mColValidationErrors
End Property

Public Property Get pValidationRule() As String
    pValidationRule = mStrValidationRule
End Property

Public Property Let pValidationRule(ByVal lStrValue As String)
    mStrValidationRule = lStrValue
End Property

Public Property Get pValidationText() As String
    pValidationText = mStrValidationText
End Property

Public Property Let pValidationText(ByVal lStrValue As String)
    mStrValidationText = lStrValue
End Property
'
'****************************************************************************************
'Public Function: ValidateDefinition
'
' Description:
'     Validates the field definition, ensuring that the properties such as size, type,
'     and constraints (e.g., Required, ZeroLength) are logical and consistent.
'     It raises validation errors if the field configuration is invalid.
'
' Parameters:
'     Optional ByRef lStrError As String
'         - An optional parameter that will contain an error message if validation fails.
'
' Returns:
'     As Boolean
'         - True if the field definition is valid, False if validation errors occur.
'
' Side Effects:
'     Adds validation errors to the mColValidationErrors collection if any issues are found.
'****************************************************************************************
'
Public Function ValidateDefinition(Optional ByRef lStrError As String) As Boolean
    Dim lStrErrors As String
    Set mColValidationErrors = New Collection

    lStrErrors = ""

    ' Validate type
    If mIntType < 1 Or mIntType > 20 Then
        mColValidationErrors.Add "Invalid field type: " & mIntType
    End If
    Select Case mIntType
    
    Case dbText
        ' Size check for text
        If mIntSize <= 0 Then
            mColValidationErrors.Add "Text field must have a positive Size."
        End If
        
        'mIntSize > 255
        If mIntSize > 255 Then
            mColValidationErrors.Add "Text field cannot exceed 255 characters in length."
        End If
        
        'If mIntSize = 0
        If mIntSize = 0 Then
            mColValidationErrors.Add "Text field has a Size of zero, which is invalid."
        End If
        
        'Required and zero length cannot both be true
        If mBlnIsRequired And mBlnAllowZeroLength Then
            mColValidationErrors.Add "Text field cannot be both Required and AllowZeroLength = Yes."
        End If
        'dbText cannot be autonumber
        If mBlnIsAutoNumber Then
            mColValidationErrors.Add "Text field cannot be AutoNumber."
        End If
        'IsPK not indexed
        If mBlnIsPK And Not mBlnIsIndexed Then
            mColValidationErrors.Add "Text field is a Primary Key but is not indexed."
        End If
        
    Case dbLong Or dbBigInt
        ' AutoNumber fields must be dbLong or dbBigInt — already implied here
        ' If AutoNumber = True, it should be indexed and a primary key
        If mBlnIsAutoNumber Then
            If Not mBlnIsPK Then
                mColValidationErrors.Add "AutoNumber field should also be a Primary Key."
            End If
            If Not mBlnIsIndexed Then
                mColValidationErrors.Add "AutoNumber field should be Indexed."
            End If
        End If
    
        ' If field is PrimaryKey but NOT AutoNumber, raise a warning (not invalid, but questionable)
        If mBlnIsPK And Not mBlnIsAutoNumber Then
            mColValidationErrors.Add "Primary Key field is not AutoNumber. Verify this is intentional."
        End If
        
    Case dbBoolean
    
        ' Booleans should not have size, AutoNumber, or AllowZeroLength
        If mIntSize <> 0 Then
            mColValidationErrors.Add "Boolean field should not have a defined Size."
        End If
        '
        'mBlnAllowZeroLength
        If mBlnAllowZeroLength Then
            mColValidationErrors.Add "Boolean field cannot allow zero-length strings."
        End If
        '
        'mBlnIsAutoNumber
        If mBlnIsAutoNumber Then
            mColValidationErrors.Add "Boolean field cannot be AutoNumber."
        End If
        
    Case dbLongBinary
        ' LongBinary fields should not have Size or AutoNumber or AllowZeroLength
        If mIntSize <> 0 Then
            mColValidationErrors.Add "LongBinary field should not have a defined Size."
        End If
        If mBlnIsAutoNumber Then
            mColValidationErrors.Add "LongBinary field cannot be AutoNumber."
        End If
        If mBlnAllowZeroLength Then
            mColValidationErrors.Add "LongBinary field cannot allow zero-length strings."
        End If
        
    Case dbBinary
        'Binary fields should not have Size or AutoNumber or AllowZeroLength
        If mIntSize <= 0 Or mIntSize > 510 Then
            mColValidationErrors.Add "Binary field is " & mIntSize & " should be > 0 and < 508."
        End If
        If mBlnIsAutoNumber Then
            mColValidationErrors.Add "Binary field cannot be AutoNumber."
        End If
        If mBlnAllowZeroLength Then
            mColValidationErrors.Add "Binary field cannot allow zero length."
        End If
        
    Case dbDate
        ' Date fields should not have size, AutoNumber, or AllowZeroLength
        If mIntSize <> 0 Then
            mColValidationErrors.Add "Date field should not have a defined Size."
        End If
        If mBlnIsAutoNumber Then
            mColValidationErrors.Add "Date field cannot be AutoNumber."
        End If
        If mBlnAllowZeroLength Then
            mColValidationErrors.Add "Date field cannot allow zero-length strings."
        End If
        
    Case dbMemo
        If mBlnIsIndexed Then
            mColValidationErrors.Add "Memo field cannot be indexed."
        End If
        If mIntSize > 0 Then
            mColValidationErrors.Add "Memo field should not have a defined Size."
        End If
        If mBlnIsPK Then
            mColValidationErrors.Add "Memo field should not be used as a primary key."
        End If
        If mBlnIsAutoNumber Then
            mColValidationErrors.Add "Memo field cannot be AutoNumber."
        End If
        
    Case dbCurrency
    
        If mIntSize > 0 Then
            mColValidationErrors.Add "Currency field should not have a defined Size."
        End If
        If mBlnAllowZeroLength Then
            mColValidationErrors.Add "Currency field cannot AllowZeroLength."
        End If
        If mBlnIsAutoNumber Then
            mColValidationErrors.Add "Currency field cannot be AutoNumber."
        End If
        
    Case dbDate
        If mIntSize > 0 Then
            mColValidationErrors.Add "Date field should not have a defined Size."
        End If
        If mBlnIsAutoNumber Then
            mColValidationErrors.Add "Date field cannot be AutoNumber."
        End If
        If mBlnAllowZeroLength Then
            mColValidationErrors.Add "Date field cannot allow zero-length strings."
        End If
        
    Case dbByte
        If mBlnAllowZeroLength Then
            mColValidationErrors.Add "Byte field cannot allow zero-length strings."
        End If
        If mBlnIsAutoNumber Then
            mColValidationErrors.Add "Byte field cannot be AutoNumber."
        End If
        
    Case dbInteger
        If mBlnAllowZeroLength Then
            mColValidationErrors.Add "Integer field cannot allow zero-length strings."
        End If
        If mBlnIsAutoNumber Then
            mColValidationErrors.Add "Integer field cannot be AutoNumber."
        End If
        
    Case dbSingle
        If mBlnAllowZeroLength Then
            mColValidationErrors.Add "Single field cannot allow zero-length strings."
        End If
        If mBlnIsAutoNumber Then
            mColValidationErrors.Add "Single field cannot be AutoNumber."
        End If
    End Select

    ' If field is PrimaryKey but NOT AutoNumber, raise a warning (not invalid, but questionable)
    If mBlnIsPK And Not mBlnIsAutoNumber Then
        mColValidationErrors.Add "Primary Key field is not AutoNumber. Verify this is intentional."
    End If
    
    ' AutoNumber must be Long or BigInt
    If mBlnIsAutoNumber Then
        If Not (mIntType = dbLong Or mIntType = dbBigInt) Then
            mColValidationErrors.Add "AutoNumber field must be dbLong or dbBigInt."
        End If
    End If

    ' Required + no default warning
    If mBlnIsRequired And IsNull(mVarDefaultValue) Then
        mColValidationErrors.Add "Field is required but has no default value."
    End If

    If mColValidationErrors.Count > 0 Then
        RaiseEvent evFieldValidationError(mStrName, lStrError, mColValidationErrors)
        ValidateDefinition = False
    Else
        lStrError = ""
        ValidateDefinition = True
    End If
End Function
''****************************************************************************************
'Public Function: ValidateValue
'
' Description:
'     Validates the data value for the field to ensure that it conforms to the
'     expected type and constraints (e.g., size, required, etc.).
'     This function is typically called after setting the value of a field.
'
' Parameters:
'     Optional ByRef lStrError As String
'         - Optional error message returned if validation fails.
'     Optional ByVal lVarInput As Variant = Empty
'         - The data value to validate. Default is Empty.
'
' Returns:
'     As Boolean
'         - True if the value is valid; False if validation errors occur.
'
' Side Effects:
'     May trigger validation events or modify internal state based on value validation.
'****************************************************************************************
'
Public Function ValidateValue(Optional ByRef lStrError As String, Optional ByVal lVarInput As Variant = Empty) As Boolean
    Dim lVar As Variant
    Set mColValidationErrors = New Collection
    lStrError = ""

    If IsMissing(lVarInput) Then
        lVar = mVarDataValue
    Else
        lVar = lVarInput
    End If

    ' Required check
    If mBlnIsRequired And IsNull(lVar) Then
        mColValidationErrors.Add "Field is required but Null."
    End If

    ' Text checks
    If mIntType = dbText And Not IsNull(lVar) Then
        If Len(lVar) > mIntSize Then
            mColValidationErrors.Add "Field exceeds size " & mIntSize & "."
        End If
        If Len(lVar) = 0 And Not mBlnAllowZeroLength Then
            mColValidationErrors.Add "Field does not allow zero-length."
        End If
    End If

    ' Type check
    Select Case mIntType
        Case dbDate
            If Not (IsDate(lVar) Or IsNull(lVar)) Then mColValidationErrors.Add "Field is not a valid date."
        Case dbBoolean
            If Not (lVar = True Or lVar = False Or IsNull(lVar)) Then mColValidationErrors.Add "Field is not boolean."
        Case dbLong, dbInteger, dbByte, dbBigInt
            If Not (IsNumeric(lVar) Or IsNull(lVar)) Then mColValidationErrors.Add "Field is not numeric."
    End Select

    ' === Currency Field Type Validation Logic ===
    If mIntType = dbCurrency Then
        If IsMissing(lVar) Or IsNull(lVar) Or Trim(CStr(lVar)) = "" Then
            mColValidationErrors.Add "Currency field cannot be empty or missing."
        ElseIf Not IsNumeric(lVar) Then
            mColValidationErrors.Add "Currency field must be numeric."
        End If
    End If
    
    If mColValidationErrors.Count > 0 Then

        Dim msg As Variant
        For Each msg In mColValidationErrors
            lStrError = lStrError & msg & vbCrLf
        Next
        RaiseEvent evDataValidationError(mStrName, lStrError, mColValidationErrors)
        ValidateValue = False
    Else
        ValidateValue = True
    End If
End Function
