Attribute VB_Name = "Step1_SetupTables"
Option Compare Database
Option Explicit

' ============================================================
'  STEP 1 SETUP — Run this ONCE after creating a blank .accdb
'  In the VBA editor: click anywhere inside SetupAllTables,
'  then press F5 (or click Run > Run Sub).
' ============================================================

Public Sub SetupAllTables()

    Dim db As DAO.Database
    Set db = CurrentDb

    On Error GoTo ErrHandler

    MsgBox "Starting database setup. This will take a few seconds...", _
           vbInformation, "Setup"

    ' --- Create tables in dependency order ---
    Call CreateTable_LifecycleStatus(db)
    Call CreateTable_QualificationStatus(db)
    Call CreateTable_Manufacturers(db)
    Call CreateTable_Components(db)
    Call CreateTable_Alternatives(db)
    Call CreateTable_Projects(db)
    Call CreateTable_BOM_Headers(db)
    Call CreateTable_BOM_Items(db)
    Call CreateTable_AuditLog(db)
    Call CreateTable_ImportStaging(db)
    Call CreateTable_ColumnProfiles(db)

    ' --- Add foreign key relationships ---
    Call CreateRelationships(db)

    ' --- Seed lookup data ---
    Call SeedLifecycleStatus(db)
    Call SeedQualificationStatus(db)

    MsgBox "SUCCESS! All tables created and data seeded." & vbCrLf & vbCrLf & _
           "Next step: Run Step2_SetupQueries.SetupAllQueries", _
           vbInformation, "Setup Complete"

    Set db = Nothing
    Exit Sub

ErrHandler:
    MsgBox "ERROR during setup:" & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
           "If a table already exists, that error can be ignored. " & _
           "Close this and run again.", vbCritical, "Setup Error"
    Set db = Nothing
End Sub

' ============================================================
'  TABLE CREATION — each in its own Sub for easy debugging
' ============================================================

Private Sub CreateTable_LifecycleStatus(db As DAO.Database)
    If TableExists("Lifecycle_Status") Then Exit Sub
    db.Execute "CREATE TABLE Lifecycle_Status (" & _
        "StatusID    AUTOINCREMENT CONSTRAINT PK_LifecycleStatus PRIMARY KEY, " & _
        "StatusName  TEXT(50) NOT NULL, " & _
        "SortOrder   INTEGER DEFAULT 0" & _
        ")", dbFailOnError
    db.Execute "CREATE UNIQUE INDEX uq_lc_name ON Lifecycle_Status (StatusName)", dbFailOnError
End Sub

Private Sub CreateTable_QualificationStatus(db As DAO.Database)
    If TableExists("Qualification_Status") Then Exit Sub
    db.Execute "CREATE TABLE Qualification_Status (" & _
        "QualID      AUTOINCREMENT CONSTRAINT PK_QualStatus PRIMARY KEY, " & _
        "QualName    TEXT(50) NOT NULL, " & _
        "SortOrder   INTEGER DEFAULT 0" & _
        ")", dbFailOnError
    db.Execute "CREATE UNIQUE INDEX uq_qual_name ON Qualification_Status (QualName)", dbFailOnError
End Sub

Private Sub CreateTable_Manufacturers(db As DAO.Database)
    If TableExists("Manufacturers") Then Exit Sub
    db.Execute "CREATE TABLE Manufacturers (" & _
        "ManufacturerID   AUTOINCREMENT CONSTRAINT PK_Manufacturers PRIMARY KEY, " & _
        "ManufacturerName TEXT(100) NOT NULL, " & _
        "ShortCode        TEXT(20), " & _
        "Website          TEXT(255), " & _
        "Notes            MEMO" & _
        ")", dbFailOnError
    db.Execute "CREATE UNIQUE INDEX uq_mfr_name ON Manufacturers (ManufacturerName)", dbFailOnError
End Sub

Private Sub CreateTable_Components(db As DAO.Database)
    If TableExists("Components") Then Exit Sub
    db.Execute "CREATE TABLE Components (" & _
        "ComponentID      AUTOINCREMENT CONSTRAINT PK_Components PRIMARY KEY, " & _
        "MPN              TEXT(100) NOT NULL, " & _
        "ManufacturerID   LONG NOT NULL, " & _
        "Description      TEXT(255), " & _
        "ComponentType    TEXT(50), " & _
        "Package          TEXT(50), " & _
        "Value            TEXT(50), " & _
        "Voltage_Rating   TEXT(30), " & _
        "Current_Rating   TEXT(30), " & _
        "Tolerance        TEXT(20), " & _
        "Datasheet_URL    TEXT(500), " & _
        "LifecycleID      LONG NOT NULL, " & _
        "QualificationID  LONG NOT NULL, " & _
        "IsPreferred      YESNO DEFAULT No, " & _
        "CreatedDate      DATETIME, " & _
        "ModifiedDate     DATETIME, " & _
        "Notes            MEMO" & _
        ")", dbFailOnError
    db.Execute "CREATE UNIQUE INDEX uq_mpn_mfr ON Components (MPN, ManufacturerID)", dbFailOnError
    db.Execute "CREATE INDEX idx_comp_lc   ON Components (LifecycleID)", dbFailOnError
    db.Execute "CREATE INDEX idx_comp_qual ON Components (QualificationID)", dbFailOnError
    db.Execute "CREATE INDEX idx_comp_mfr  ON Components (ManufacturerID)", dbFailOnError
End Sub

Private Sub CreateTable_Alternatives(db As DAO.Database)
    If TableExists("Alternatives") Then Exit Sub
    db.Execute "CREATE TABLE Alternatives (" & _
        "AltID           AUTOINCREMENT CONSTRAINT PK_Alternatives PRIMARY KEY, " & _
        "PrimaryCompID   LONG NOT NULL, " & _
        "AltCompID       LONG NOT NULL, " & _
        "AltType         TEXT(30), " & _
        "Restrictions    TEXT(255), " & _
        "QualificationID LONG, " & _
        "AddedDate       DATETIME, " & _
        "AddedBy         TEXT(50), " & _
        "Notes           MEMO" & _
        ")", dbFailOnError
    db.Execute "CREATE UNIQUE INDEX uq_alt_pair ON Alternatives (PrimaryCompID, AltCompID)", dbFailOnError
End Sub

Private Sub CreateTable_Projects(db As DAO.Database)
    If TableExists("Projects") Then Exit Sub
    db.Execute "CREATE TABLE Projects (" & _
        "ProjectID    AUTOINCREMENT CONSTRAINT PK_Projects PRIMARY KEY, " & _
        "ProjectCode  TEXT(50) NOT NULL, " & _
        "ProjectName  TEXT(200) NOT NULL, " & _
        "Description  MEMO, " & _
        "PCB_Rev      TEXT(20), " & _
        "Owner        TEXT(100), " & _
        "Status       TEXT(30) DEFAULT 'Active', " & _
        "CreatedDate  DATETIME, " & _
        "Notes        MEMO" & _
        ")", dbFailOnError
    db.Execute "CREATE UNIQUE INDEX uq_proj_code ON Projects (ProjectCode)", dbFailOnError
End Sub

Private Sub CreateTable_BOM_Headers(db As DAO.Database)
    If TableExists("BOM_Headers") Then Exit Sub
    db.Execute "CREATE TABLE BOM_Headers (" & _
        "HeaderID       AUTOINCREMENT CONSTRAINT PK_BOMHeaders PRIMARY KEY, " & _
        "ProjectID      LONG NOT NULL, " & _
        "BOM_Revision   TEXT(20), " & _
        "ImportDate     DATETIME, " & _
        "ImportedBy     TEXT(100), " & _
        "SourceFile     TEXT(500), " & _
        "RowCount       INTEGER DEFAULT 0, " & _
        "Status         TEXT(20) DEFAULT 'Pending', " & _
        "Notes          MEMO" & _
        ")", dbFailOnError
    db.Execute "CREATE INDEX idx_bom_proj ON BOM_Headers (ProjectID)", dbFailOnError
End Sub

Private Sub CreateTable_BOM_Items(db As DAO.Database)
    If TableExists("BOM_Items") Then Exit Sub
    db.Execute "CREATE TABLE BOM_Items (" & _
        "ItemID           AUTOINCREMENT CONSTRAINT PK_BOMItems PRIMARY KEY, " & _
        "HeaderID         LONG NOT NULL, " & _
        "ComponentID      LONG, " & _
        "MPN_Raw          TEXT(100), " & _
        "Manufacturer_Raw TEXT(100), " & _
        "Description_Raw  TEXT(255), " & _
        "Quantity         DOUBLE DEFAULT 1, " & _
        "RefDes           TEXT(500), " & _
        "Notes            TEXT(500), " & _
        "MatchStatus      TEXT(20) DEFAULT 'Unmatched'" & _
        ")", dbFailOnError
    db.Execute "CREATE INDEX idx_item_header ON BOM_Items (HeaderID)", dbFailOnError
    db.Execute "CREATE INDEX idx_item_comp   ON BOM_Items (ComponentID)", dbFailOnError
End Sub

Private Sub CreateTable_AuditLog(db As DAO.Database)
    If TableExists("Audit_Log") Then Exit Sub
    db.Execute "CREATE TABLE Audit_Log (" & _
        "LogID      AUTOINCREMENT CONSTRAINT PK_AuditLog PRIMARY KEY, " & _
        "LogDate    DATETIME, " & _
        "UserName   TEXT(100), " & _
        "Action     TEXT(50), " & _
        "TableName  TEXT(50), " & _
        "RecordID   LONG, " & _
        "FieldName  TEXT(100), " & _
        "OldValue   MEMO, " & _
        "NewValue   MEMO, " & _
        "Notes      MEMO" & _
        ")", dbFailOnError
End Sub

Private Sub CreateTable_ImportStaging(db As DAO.Database)
    If TableExists("tbl_ImportStaging") Then Exit Sub
    db.Execute "CREATE TABLE tbl_ImportStaging (" & _
        "RowNum           AUTOINCREMENT CONSTRAINT PK_Staging PRIMARY KEY, " & _
        "MPN_Raw          TEXT(100), " & _
        "Manufacturer_Raw TEXT(100), " & _
        "Description_Raw  TEXT(255), " & _
        "Quantity         DOUBLE DEFAULT 1, " & _
        "RefDes           TEXT(500), " & _
        "Notes            TEXT(500), " & _
        "MatchStatus      TEXT(20) DEFAULT 'Unmatched', " & _
        "ComponentID      LONG, " & _
        "UserAction       TEXT(20) DEFAULT 'Auto'" & _
        ")", dbFailOnError
End Sub

Private Sub CreateTable_ColumnProfiles(db As DAO.Database)
    If TableExists("Import_ColumnProfiles") Then Exit Sub
    db.Execute "CREATE TABLE Import_ColumnProfiles (" & _
        "ProfileID    AUTOINCREMENT CONSTRAINT PK_ColProfiles PRIMARY KEY, " & _
        "ProfileName  TEXT(100) NOT NULL, " & _
        "ColMPN       INTEGER DEFAULT 1, " & _
        "ColMfr       INTEGER DEFAULT 2, " & _
        "ColDesc      INTEGER DEFAULT 3, " & _
        "ColQty       INTEGER DEFAULT 4, " & _
        "ColRefDes    INTEGER DEFAULT 5, " & _
        "ColNotes     INTEGER DEFAULT 6, " & _
        "HeaderRow    INTEGER DEFAULT 1" & _
        ")", dbFailOnError
    ' Seed a default profile
    db.Execute "INSERT INTO Import_ColumnProfiles " & _
               "(ProfileName, ColMPN, ColMfr, ColDesc, ColQty, ColRefDes, ColNotes, HeaderRow) " & _
               "VALUES ('Default', 1, 2, 3, 4, 5, 6, 1)", dbFailOnError
End Sub

' ============================================================
'  RELATIONSHIPS
' ============================================================

Private Sub CreateRelationships(db As DAO.Database)

    On Error Resume Next  ' Skip if relationship already exists

    Dim rel As DAO.Relation
    Dim fld As DAO.Field

    ' Components -> Manufacturers
    Set rel = db.CreateRelation("FK_Comp_Mfr", "Manufacturers", "Components", dbRelationUpdateCascade)
    rel.Fields.Append rel.CreateField("ManufacturerID")
    rel.Fields("ManufacturerID").ForeignName = "ManufacturerID"
    db.Relations.Append rel

    ' Components -> Lifecycle_Status
    Set rel = db.CreateRelation("FK_Comp_LC", "Lifecycle_Status", "Components", 0)
    rel.Fields.Append rel.CreateField("StatusID")
    rel.Fields("StatusID").ForeignName = "LifecycleID"
    db.Relations.Append rel

    ' Components -> Qualification_Status
    Set rel = db.CreateRelation("FK_Comp_Qual", "Qualification_Status", "Components", 0)
    rel.Fields.Append rel.CreateField("QualID")
    rel.Fields("QualID").ForeignName = "QualificationID"
    db.Relations.Append rel

    ' Alternatives -> Components (primary side)
    Set rel = db.CreateRelation("FK_Alt_Primary", "Components", "Alternatives", 0)
    rel.Fields.Append rel.CreateField("ComponentID")
    rel.Fields("ComponentID").ForeignName = "PrimaryCompID"
    db.Relations.Append rel

    ' BOM_Headers -> Projects
    Set rel = db.CreateRelation("FK_BOM_Proj", "Projects", "BOM_Headers", dbRelationUpdateCascade)
    rel.Fields.Append rel.CreateField("ProjectID")
    rel.Fields("ProjectID").ForeignName = "ProjectID"
    db.Relations.Append rel

    ' BOM_Items -> BOM_Headers
    Set rel = db.CreateRelation("FK_Item_Header", "BOM_Headers", "BOM_Items", dbRelationUpdateCascade + dbRelationDeleteCascade)
    rel.Fields.Append rel.CreateField("HeaderID")
    rel.Fields("HeaderID").ForeignName = "HeaderID"
    db.Relations.Append rel

    On Error GoTo 0
End Sub

' ============================================================
'  SEED DATA
' ============================================================

Private Sub SeedLifecycleStatus(db As DAO.Database)
    If DCount("*", "Lifecycle_Status") > 0 Then Exit Sub
    db.Execute "INSERT INTO Lifecycle_Status (StatusName, SortOrder) VALUES ('Active', 1)", dbFailOnError
    db.Execute "INSERT INTO Lifecycle_Status (StatusName, SortOrder) VALUES ('Preferred', 2)", dbFailOnError
    db.Execute "INSERT INTO Lifecycle_Status (StatusName, SortOrder) VALUES ('NRND', 3)", dbFailOnError
    db.Execute "INSERT INTO Lifecycle_Status (StatusName, SortOrder) VALUES ('Obsolete', 4)", dbFailOnError
    db.Execute "INSERT INTO Lifecycle_Status (StatusName, SortOrder) VALUES ('Discontinued', 5)", dbFailOnError
    db.Execute "INSERT INTO Lifecycle_Status (StatusName, SortOrder) VALUES ('Upcoming', 6)", dbFailOnError
    db.Execute "INSERT INTO Lifecycle_Status (StatusName, SortOrder) VALUES ('Unknown', 7)", dbFailOnError
End Sub

Private Sub SeedQualificationStatus(db As DAO.Database)
    If DCount("*", "Qualification_Status") > 0 Then Exit Sub
    db.Execute "INSERT INTO Qualification_Status (QualName, SortOrder) VALUES ('Approved', 1)", dbFailOnError
    db.Execute "INSERT INTO Qualification_Status (QualName, SortOrder) VALUES ('Proposed', 2)", dbFailOnError
    db.Execute "INSERT INTO Qualification_Status (QualName, SortOrder) VALUES ('Under Testing', 3)", dbFailOnError
    db.Execute "INSERT INTO Qualification_Status (QualName, SortOrder) VALUES ('Rejected', 4)", dbFailOnError
    db.Execute "INSERT INTO Qualification_Status (QualName, SortOrder) VALUES ('On Hold', 5)", dbFailOnError
End Sub

' ============================================================
'  UTILITY
' ============================================================

Public Function TableExists(tableName As String) As Boolean
    Dim tdf As DAO.TableDef
    For Each tdf In CurrentDb.TableDefs
        If tdf.Name = tableName Then
            TableExists = True
            Exit Function
        End If
    Next
    TableExists = False
End Function
