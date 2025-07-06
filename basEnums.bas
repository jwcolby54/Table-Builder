Attribute VB_Name = "basEnums"
Option Compare Database
Option Explicit
' ============================================================================
' Module: basEnums
' Purpose: Defines enumerations to standardize constants across the system.
'          Eliminates magic numbers and supports readable metadata definitions.
'
' Key Enums:
'   - FieldTypeEnum: Field data types (Text, Date, YesNo, etc.)
'   - FieldAttributeEnum: Field attributes (Required, Indexed, etc.)
'   - ValidationFlagsEnum: Flags for test rule categories
'
' Used By:
'   - clsFieldDef
'   - clsTableBuilder
'   - Test rule modules
' ============================================================================


Public Enum IndexedBehavior
    fdUnindexed = 0            ' "No"
    fdDuplicatesAllowed = 1    ' "Yes"
    fdNoDuplicates = 2         ' "Yes (No Duplicates)"
End Enum

Public Enum AutoDeleteBehavior
    fdPreserveExisting = 0
    fdDeleteIfExists = -1
End Enum

Public Enum PrimaryKeyBehavior
    fdNonPK = 0
    fdPrimaryKey = 1
End Enum
Public Enum AllowZeroLengthBehavior
    fdNo = 0
    fdYes = 1
End Enum
Public Enum RequiredBehavior
    fdOptional = 0
    fdRequired = 1
End Enum

Public Enum AutoNumberBehavior
    IsNotAutoNumber = 0
    IsAutoNumber = 1
End Enum
