Attribute VB_Name = "basTblBuilder"
Option Compare Database
Option Explicit
' ============================================================================
' Module: basTblBuilder
' Purpose: Defines functions to build application-specific tables using the
'          metadata-driven framework (clsTableBuilder and clsFieldDef).
'
' Tables Defined:
'   - TablePeople
'   - tlkpEyeColor
'   - tlkpHairColor
'   - Additional tables can be defined by following the established pattern.
'
' Usage:
'   Call Build_TableXYZ() to create and optionally seed each table.
' ============================================================================
'name                   kind    scope   params  return_type
'Build_AllTables        Sub     Public  []
'Build_TablePeople      Sub     Public  ['Optional lBlnAutoDelete As Boolean = True']
'Build_tlkpEyeColor     Sub     Public  ['Optional lBlnAutoDelete As Boolean = True']
'Build_tlkpHairColor    Sub     Public  ['Optional lBlnAutoDelete As Boolean = True']
'
'****************************************************************************************
'Public Sub: Build_AllTables
'
' Description:
'     Entry point for building all application tables that are required for use in the
'     current database. Calls each individual Build_XXX routine to create the appropriate
'     tables and seed them if needed.
'
'     This routine ensures all lookup and data tables are defined in a single place
'     for initialization, testing, or reset purposes.
'
' Parameters:
'     (None)
'
' Returns:
'     None
'
' Side Effects:
'     Invokes Build_TablePeople, Build_tlkpEyeColor, Build_tlkpHairColor in order.
'****************************************************************************************
'
Public Sub Build_AllTables()
    Build_tlkpEyeColor
    Build_tlkpHairColor
    Build_TablePeople
    ' Add more as needed
End Sub
'****************************************************************************************
'Public Sub: Build_TablePeople
'
' Description:
'     Builds the main "tblPeople" table, which is used as a working dataset in the application
'     and in examples throughout the book. Defines multiple fields representing a person
'     including names, date of birth, and foreign keys to lookup tables.
'
'     This method instantiates a clsTableBuilder, sets its pTableName property, and uses
'     CreateFieldDef to define each field with the correct type and metadata.
'     Finally, it optionally deletes and recreates the table based on lBlnAutoDelete.
'
' Fields Defined:
'     - PE_ID      (AutoNumber, Primary Key)
'     - PE_FName   (Text, Required)
'     - PE_LName   (Text, Required)
'     - PE_DOB     (Date/Time)
'     - PE_IDEC    (Long Integer, Foreign Key to EyeColor)
'     - PE_IDHC    (Long Integer, Foreign Key to HairColor)
'     - PE_Active  (Yes/No, Default True)
'     - PE_Trash   (Yes/No, Default False)
'
' Parameters:
'     lBlnAutoDelete As Boolean
'         - If True, deletes the table before recreating it.
'
' Returns:
'     None
'
' Side Effects:
'     Deletes and recreates tblPeople if requested.
'****************************************************************************************
'
Public Sub Build_TablePeople(Optional lBlnAutoDelete As Boolean = True)
    Dim lTbl As clsTableBuilder
    Set lTbl = New clsTableBuilder

    With lTbl
        .pTableName = "tblPeople"

        .AddFieldDef .CreateFieldDef("PE_ID", dbLong, _
            RequiredBehavior.fdRequired, IndexedBehavior.fdNoDuplicates, _
            PrimaryKeyBehavior.fdPrimaryKey, AllowZeroLengthBehavior.fdNo, _
            , AutoNumberBehavior.IsAutoNumber)

        .AddFieldDef .CreateFieldDef("PE_FName", dbText, _
            RequiredBehavior.fdOptional, IndexedBehavior.fdUnindexed, _
            PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdYes, _
            , AutoNumberBehavior.IsNotAutoNumber)

        .AddFieldDef .CreateFieldDef("PE_LName", dbText, _
            RequiredBehavior.fdOptional, IndexedBehavior.fdUnindexed, _
            PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdYes, _
            , AutoNumberBehavior.IsNotAutoNumber)

        .AddFieldDef .CreateFieldDef("PE_DOB", dbDate, _
            RequiredBehavior.fdOptional, IndexedBehavior.fdUnindexed, _
            PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdNo, _
            0, AutoNumberBehavior.IsNotAutoNumber)

        .AddFieldDef .CreateFieldDef("PE_IDEC", dbLong, _
            RequiredBehavior.fdOptional, IndexedBehavior.fdUnindexed, _
            PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdYes, _
            , AutoNumberBehavior.IsNotAutoNumber)

        .AddFieldDef .CreateFieldDef("PE_IDHC", dbLong, _
            RequiredBehavior.fdOptional, IndexedBehavior.fdUnindexed, _
            PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdYes, _
            , AutoNumberBehavior.IsNotAutoNumber)

        .AddFieldDef .CreateFieldDef("PE_Active", dbBoolean, _
            RequiredBehavior.fdOptional, IndexedBehavior.fdUnindexed, _
            PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdNo, _
            0, AutoNumberBehavior.IsNotAutoNumber)

        .AddFieldDef .CreateFieldDef("PE_Trash", dbBoolean, _
            RequiredBehavior.fdRequired, IndexedBehavior.fdUnindexed, _
            PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdNo, _
            0, AutoNumberBehavior.IsNotAutoNumber)

        Dim lStrValidation As String
        lStrValidation = .ValidateAllFields()
        If Len(lStrValidation) > 0 Then
            MsgBox "Validation errors (People):" & vbCrLf & lStrValidation, vbExclamation
            Exit Sub
        End If

        '.CreateTable True
    End With
    If lTbl.CreateTable(lBlnAutoDelete) Then

        ' === Seed data if we built a fresh table ===
        If lBlnAutoDelete Then
            Dim colSeedData As Collection
            Set colSeedData = New Collection
    
            colSeedData.Add Array("Daffy", "Duck", #4/17/1937#, 1, 1, True, False)
            colSeedData.Add Array("Donald", "Duck", #6/14/1944#, 2, 1, True, False)
            colSeedData.Add Array("Goofy", "Dawg", #5/25/1932#, 1, 2, False, False)
            colSeedData.Add Array("Micky", "Mouse", #11/18/1928#, 1, 2, False, False)
    
            lTbl.SeedRows colSeedData
        End If
    End If
End Sub


'****************************************************************************************
'Public Sub: Build_tlkpEyeColor
'
' Description:
'     Builds the "tlkpEyeColor" lookup table containing a set of predefined eye color values.
'     This table is referenced by tblPeople via the PE_IDEC foreign key field.
'
'     The function defines the structure using clsTableBuilder, and then seeds the
'     table using a predefined collection of eye colors from basTblBuilder.
'
' Fields:
'     - EC_ID    (AutoNumber, Primary Key)
'     - EC_Desc  (Text, Required, Indexed)
'
' Parameters:
'     lBlnAutoDelete As Boolean
'         - If True, deletes the existing table before creating a new one.
'
' Returns:
'     None
'
' Side Effects:
'     Creates and seeds the tlkpEyeColor table.
'****************************************************************************************
Public Sub Build_tlkpEyeColor(Optional lBlnAutoDelete As Boolean = True)
    Dim lTbl As clsTableBuilder
    Set lTbl = New clsTableBuilder

    lTbl.pTableName = "tlkpEyeColor"

    lTbl.AddFieldDef lTbl.CreateFieldDef("ECID", dbLong, _
        RequiredBehavior.fdRequired, IndexedBehavior.fdNoDuplicates, _
        PrimaryKeyBehavior.fdPrimaryKey, AllowZeroLengthBehavior.fdNo, _
         , AutoNumberBehavior.IsAutoNumber)

    lTbl.AddFieldDef lTbl.CreateFieldDef("ECName", dbText, _
        RequiredBehavior.fdRequired, IndexedBehavior.fdUnindexed, _
        PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdNo, _
         , AutoNumberBehavior.IsNotAutoNumber)

    Dim lStrValidation As String
    lStrValidation = lTbl.ValidateAllFields()
    If Len(lStrValidation) > 0 Then
        MsgBox "Validation errors (EyeColor):" & vbCrLf & lStrValidation, vbExclamation
        Exit Sub
    End If

    If lTbl.CreateTable(lBlnAutoDelete) Then

    ' === Seed data if we built a fresh table ===
        If lBlnAutoDelete Then
            Dim colSeedData As Collection
            Set colSeedData = New Collection
    
            colSeedData.Add Array("Brown")
            colSeedData.Add Array("Blue")
            colSeedData.Add Array("Green")
            colSeedData.Add Array("Violet")
            colSeedData.Add Array("Yellow")
            colSeedData.Add Array("Red")
            colSeedData.Add Array("Black")
    
            lTbl.SeedRows colSeedData
        End If
    End If
End Sub

'****************************************************************************************
'Public Sub: Build_tlkpHairColor
'
' Description:
'     Builds the "tlkpHairColor" lookup table containing a set of predefined hair color values.
'     This table is referenced by tblPeople via the PE_IDHC foreign key field.
'
'     Like the eye color table, it defines the structure and uses a seed collection
'     to insert test values.
'
' Fields:
'     - HC_ID    (AutoNumber, Primary Key)
'     - HC_Desc  (Text, Required, Indexed)
'
' Parameters:
'     lBlnAutoDelete As Boolean
'         - If True, deletes the existing table before creating a new one.
'
' Returns:
'     None
'
' Side Effects:
'     Creates and seeds the tlkpHairColor table.
'****************************************************************************************
'
Public Sub Build_tlkpHairColor(Optional lBlnAutoDelete As Boolean = True)
    Dim lTbl As clsTableBuilder
    Set lTbl = New clsTableBuilder

    lTbl.pTableName = "tlkpHairColor"

    lTbl.AddFieldDef lTbl.CreateFieldDef("HCID", dbLong, _
        RequiredBehavior.fdRequired, IndexedBehavior.fdNoDuplicates, _
        PrimaryKeyBehavior.fdPrimaryKey, AllowZeroLengthBehavior.fdNo, _
        , AutoNumberBehavior.IsAutoNumber)

    lTbl.AddFieldDef lTbl.CreateFieldDef("HCName", dbText, _
        RequiredBehavior.fdRequired, IndexedBehavior.fdUnindexed, _
        PrimaryKeyBehavior.fdNonPK, AllowZeroLengthBehavior.fdNo, _
        , AutoNumberBehavior.IsNotAutoNumber)

    Dim lStrValidation As String
    lStrValidation = lTbl.ValidateAllFields()
    If Len(lStrValidation) > 0 Then
        MsgBox "Validation errors (HairColor):" & vbCrLf & lStrValidation, vbExclamation
        Exit Sub
    End If

    lTbl.CreateTable lBlnAutoDelete

    ' === Seed data if we built a fresh table ===
    If lBlnAutoDelete Then
        Dim colSeedData As Collection
        Set colSeedData = New Collection

        colSeedData.Add Array("Brown")
        colSeedData.Add Array("Black")
        colSeedData.Add Array("Blond")
        colSeedData.Add Array("Grey")
        colSeedData.Add Array("Red")

        lTbl.SeedRows colSeedData
    End If
End Sub

