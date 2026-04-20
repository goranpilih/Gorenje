Attribute VB_Name = "Step1_SetupTables"
Option Compare Database
Option Explicit

' ============================================================
'  STEP 1 SETUP — Run this ONCE after creating a blank .accdb
'  Click anywhere inside SetupAllTables, then press F5.
' ============================================================

Public Sub SetupAllTables()

    Dim db As DAO.Database
    Set db = CurrentDb

    MsgBox "Starting database setup...", vbInformation, "Setup"

    ' Each table is created individually with its own error handling.
    ' If one fails you will see exactly which table caused the problem.

    CreateT_LifecycleStatus db
    CreateT_QualificationStatus db
    CreateT_Manufacturers db
    CreateT_Components db
    CreateT_Alternatives db
    CreateT_Projects db
    CreateT_BOM_Headers db
    CreateT_BOM_Items db
    CreateT_AuditLog db
    CreateT_ImportStaging db
    CreateT_ColumnProfiles db

    CreateRelationships db

    SeedLifecycleStatus db
    SeedQualificationStatus db
    SeedDefaultProfile db

    Set db = Nothing

    MsgBox "SUCCESS! All tables created." & vbCrLf & vbCrLf & _
           "Next: open Step2_SetupQueries and press F5.", _
           vbInformation, "Setup Complete"
End Sub

' ============================================================
'  HELPER — wraps Execute and shows table name on failure
' ============================================================
Private Sub RunSQL(db As DAO.Database, tableName As String, sql As String)
    On Error GoTo Fail
    db.Execute sql, dbFailOnError
    Exit Sub
Fail:
    MsgBox "FAILED on table: " & tableName & vbCrLf & vbCrLf & _
           "Error: " & Err.Description & vbCrLf & vbCrLf & _
           "SQL: " & sql, vbCritical, "Setup Error"
    Err.Raise Err.Number   ' Re-raise so SetupAllTables stops too
End Sub

' ============================================================
'  TABLE CREATION
'  Rules used here to avoid Access DDL syntax errors:
'   - No DEFAULT values in CREATE TABLE (set later or via form)
'   - No NOT NULL in CREATE TABLE (enforced by form code)
'   - TEXT max length = 255
'   - Long text fields use MEMO
'   - INTEGER = 16-bit, LONG = 32-bit
' ============================================================

Private Sub CreateT_LifecycleStatus(db As DAO.Database)
    If TableExists("Lifecycle_Status") Then Exit Sub
    RunSQL db, "Lifecycle_Status", _
        "CREATE TABLE Lifecycle_Status (" & _
        "  StatusID   AUTOINCREMENT CONSTRAINT PK_LifecycleStatus PRIMARY KEY," & _
        "  StatusName TEXT(50)," & _
        "  SortOrder  INTEGER" & _
        ")"
    RunSQL db, "Lifecycle_Status (index)", _
        "CREATE UNIQUE INDEX uq_lc_name ON Lifecycle_Status (StatusName)"
End Sub

Private Sub CreateT_QualificationStatus(db As DAO.Database)
    If TableExists("Qualification_Status") Then Exit Sub
    RunSQL db, "Qualification_Status", _
        "CREATE TABLE Qualification_Status (" & _
        "  QualID    AUTOINCREMENT CONSTRAINT PK_QualStatus PRIMARY KEY," & _
        "  QualName  TEXT(50)," & _
        "  SortOrder INTEGER" & _
        ")"
    RunSQL db, "Qualification_Status (index)", _
        "CREATE UNIQUE INDEX uq_qual_name ON Qualification_Status (QualName)"
End Sub

Private Sub CreateT_Manufacturers(db As DAO.Database)
    If TableExists("Manufacturers") Then Exit Sub
    RunSQL db, "Manufacturers", _
        "CREATE TABLE Manufacturers (" & _
        "  ManufacturerID   AUTOINCREMENT CONSTRAINT PK_Manufacturers PRIMARY KEY," & _
        "  ManufacturerName TEXT(100)," & _
        "  ShortCode        TEXT(20)," & _
        "  Website          TEXT(255)," & _
        "  Notes            MEMO" & _
        ")"
    RunSQL db, "Manufacturers (index)", _
        "CREATE UNIQUE INDEX uq_mfr_name ON Manufacturers (ManufacturerName)"
End Sub

Private Sub CreateT_Components(db As DAO.Database)
    If TableExists("Components") Then Exit Sub
    RunSQL db, "Components", _
        "CREATE TABLE Components (" & _
        "  ComponentID     AUTOINCREMENT CONSTRAINT PK_Components PRIMARY KEY," & _
        "  MPN             TEXT(100)," & _
        "  ManufacturerID  LONG," & _
        "  Description     TEXT(255)," & _
        "  ComponentType    TEXT(50)," & _
        "  Package          TEXT(50)," & _
        "  ComponentValue   TEXT(50)," & _
        "  Voltage_Rating  TEXT(30)," & _
        "  Current_Rating  TEXT(30)," & _
        "  Tolerance       TEXT(20)," & _
        "  Datasheet_URL   MEMO," & _
        "  LifecycleID     LONG," & _
        "  QualificationID LONG," & _
        "  IsPreferred     YESNO," & _
        "  CreatedDate     DATETIME," & _
        "  ModifiedDate    DATETIME," & _
        "  Notes           MEMO" & _
        ")"
    RunSQL db, "Components (index mpn)", _
        "CREATE UNIQUE INDEX uq_mpn_mfr ON Components (MPN, ManufacturerID)"
    RunSQL db, "Components (index lc)", _
        "CREATE INDEX idx_comp_lc ON Components (LifecycleID)"
    RunSQL db, "Components (index qual)", _
        "CREATE INDEX idx_comp_qual ON Components (QualificationID)"
    RunSQL db, "Components (index mfr)", _
        "CREATE INDEX idx_comp_mfr ON Components (ManufacturerID)"
End Sub

Private Sub CreateT_Alternatives(db As DAO.Database)
    If TableExists("Alternatives") Then Exit Sub
    RunSQL db, "Alternatives", _
        "CREATE TABLE Alternatives (" & _
        "  AltID           AUTOINCREMENT CONSTRAINT PK_Alternatives PRIMARY KEY," & _
        "  PrimaryCompID   LONG," & _
        "  AltCompID       LONG," & _
        "  AltType         TEXT(30)," & _
        "  Restrictions    TEXT(255)," & _
        "  QualificationID LONG," & _
        "  AddedDate       DATETIME," & _
        "  AddedBy         TEXT(50)," & _
        "  Notes           MEMO" & _
        ")"
    RunSQL db, "Alternatives (index)", _
        "CREATE UNIQUE INDEX uq_alt_pair ON Alternatives (PrimaryCompID, AltCompID)"
End Sub

Private Sub CreateT_Projects(db As DAO.Database)
    If TableExists("Projects") Then Exit Sub
    RunSQL db, "Projects", _
        "CREATE TABLE Projects (" & _
        "  ProjectID   AUTOINCREMENT CONSTRAINT PK_Projects PRIMARY KEY," & _
        "  ProjectCode TEXT(50)," & _
        "  ProjectName TEXT(200)," & _
        "  Description MEMO," & _
        "  PCB_Rev        TEXT(20)," & _
        "  Owner          TEXT(100)," & _
        "  ProjectStatus  TEXT(30)," & _
        "  CreatedDate DATETIME," & _
        "  Notes       MEMO" & _
        ")"
    RunSQL db, "Projects (index)", _
        "CREATE UNIQUE INDEX uq_proj_code ON Projects (ProjectCode)"
End Sub

Private Sub CreateT_BOM_Headers(db As DAO.Database)
    If TableExists("BOM_Headers") Then Exit Sub
    RunSQL db, "BOM_Headers", _
        "CREATE TABLE BOM_Headers (" & _
        "  HeaderID     AUTOINCREMENT CONSTRAINT PK_BOMHeaders PRIMARY KEY," & _
        "  ProjectID    LONG," & _
        "  BOM_Revision TEXT(20)," & _
        "  ImportDate   DATETIME," & _
        "  ImportedBy   TEXT(100)," & _
        "  SourceFile    MEMO," & _
        "  RowCount      INTEGER," & _
        "  ImportStatus  TEXT(20)," & _
        "  Notes        MEMO" & _
        ")"
    RunSQL db, "BOM_Headers (index)", _
        "CREATE INDEX idx_bom_proj ON BOM_Headers (ProjectID)"
End Sub

Private Sub CreateT_BOM_Items(db As DAO.Database)
    If TableExists("BOM_Items") Then Exit Sub
    RunSQL db, "BOM_Items", _
        "CREATE TABLE BOM_Items (" & _
        "  ItemID           AUTOINCREMENT CONSTRAINT PK_BOMItems PRIMARY KEY," & _
        "  HeaderID         LONG," & _
        "  ComponentID      LONG," & _
        "  MPN_Raw          TEXT(100)," & _
        "  Manufacturer_Raw TEXT(100)," & _
        "  Description_Raw  TEXT(255)," & _
        "  Quantity         DOUBLE," & _
        "  RefDes           MEMO," & _
        "  Notes            TEXT(255)," & _
        "  MatchStatus      TEXT(20)" & _
        ")"
    RunSQL db, "BOM_Items (index header)", _
        "CREATE INDEX idx_item_header ON BOM_Items (HeaderID)"
    RunSQL db, "BOM_Items (index comp)", _
        "CREATE INDEX idx_item_comp ON BOM_Items (ComponentID)"
End Sub

Private Sub CreateT_AuditLog(db As DAO.Database)
    If TableExists("Audit_Log") Then Exit Sub
    RunSQL db, "Audit_Log", _
        "CREATE TABLE Audit_Log (" & _
        "  LogID     AUTOINCREMENT CONSTRAINT PK_AuditLog PRIMARY KEY," & _
        "  LogDate   DATETIME," & _
        "  UserName   TEXT(100)," & _
        "  LogAction  TEXT(50)," & _
        "  TableName  TEXT(50)," & _
        "  RecordID  LONG," & _
        "  FieldName TEXT(100)," & _
        "  OldValue  MEMO," & _
        "  NewValue  MEMO," & _
        "  Notes     MEMO" & _
        ")"
End Sub

Private Sub CreateT_ImportStaging(db As DAO.Database)
    If TableExists("tbl_ImportStaging") Then Exit Sub
    RunSQL db, "tbl_ImportStaging", _
        "CREATE TABLE tbl_ImportStaging (" & _
        "  RowNum           AUTOINCREMENT CONSTRAINT PK_Staging PRIMARY KEY," & _
        "  MPN_Raw          TEXT(100)," & _
        "  Manufacturer_Raw TEXT(100)," & _
        "  Description_Raw  TEXT(255)," & _
        "  Quantity         DOUBLE," & _
        "  RefDes           MEMO," & _
        "  Notes            TEXT(255)," & _
        "  MatchStatus      TEXT(20)," & _
        "  ComponentID      LONG," & _
        "  UserAction       TEXT(20)" & _
        ")"
End Sub

Private Sub CreateT_ColumnProfiles(db As DAO.Database)
    If TableExists("Import_ColumnProfiles") Then Exit Sub
    RunSQL db, "Import_ColumnProfiles", _
        "CREATE TABLE Import_ColumnProfiles (" & _
        "  ProfileID   AUTOINCREMENT CONSTRAINT PK_ColProfiles PRIMARY KEY," & _
        "  ProfileName TEXT(100)," & _
        "  ColMPN      INTEGER," & _
        "  ColMfr      INTEGER," & _
        "  ColDesc     INTEGER," & _
        "  ColQty      INTEGER," & _
        "  ColRefDes   INTEGER," & _
        "  ColNotes    INTEGER," & _
        "  HeaderRow   INTEGER" & _
        ")"
End Sub

' ============================================================
'  RELATIONSHIPS
' ============================================================
Private Sub CreateRelationships(db As DAO.Database)
    On Error Resume Next   ' Skip silently if relationship already exists

    Dim rel As DAO.Relation

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

    ' Alternatives -> Components
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
    Set rel = db.CreateRelation("FK_Item_Hdr", "BOM_Headers", "BOM_Items", _
                                 dbRelationUpdateCascade + dbRelationDeleteCascade)
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
    RunSQL db, "seed Lifecycle_Status", "INSERT INTO Lifecycle_Status (StatusName, SortOrder) VALUES ('Active', 1)"
    RunSQL db, "seed Lifecycle_Status", "INSERT INTO Lifecycle_Status (StatusName, SortOrder) VALUES ('Preferred', 2)"
    RunSQL db, "seed Lifecycle_Status", "INSERT INTO Lifecycle_Status (StatusName, SortOrder) VALUES ('NRND', 3)"
    RunSQL db, "seed Lifecycle_Status", "INSERT INTO Lifecycle_Status (StatusName, SortOrder) VALUES ('Obsolete', 4)"
    RunSQL db, "seed Lifecycle_Status", "INSERT INTO Lifecycle_Status (StatusName, SortOrder) VALUES ('Discontinued', 5)"
    RunSQL db, "seed Lifecycle_Status", "INSERT INTO Lifecycle_Status (StatusName, SortOrder) VALUES ('Upcoming', 6)"
    RunSQL db, "seed Lifecycle_Status", "INSERT INTO Lifecycle_Status (StatusName, SortOrder) VALUES ('Unknown', 7)"
End Sub

Private Sub SeedQualificationStatus(db As DAO.Database)
    If DCount("*", "Qualification_Status") > 0 Then Exit Sub
    RunSQL db, "seed Qualification_Status", "INSERT INTO Qualification_Status (QualName, SortOrder) VALUES ('Approved', 1)"
    RunSQL db, "seed Qualification_Status", "INSERT INTO Qualification_Status (QualName, SortOrder) VALUES ('Proposed', 2)"
    RunSQL db, "seed Qualification_Status", "INSERT INTO Qualification_Status (QualName, SortOrder) VALUES ('Under Testing', 3)"
    RunSQL db, "seed Qualification_Status", "INSERT INTO Qualification_Status (QualName, SortOrder) VALUES ('Rejected', 4)"
    RunSQL db, "seed Qualification_Status", "INSERT INTO Qualification_Status (QualName, SortOrder) VALUES ('On Hold', 5)"
End Sub

Private Sub SeedDefaultProfile(db As DAO.Database)
    If DCount("*", "Import_ColumnProfiles") > 0 Then Exit Sub
    RunSQL db, "seed Import_ColumnProfiles", _
        "INSERT INTO Import_ColumnProfiles " & _
        "(ProfileName, ColMPN, ColMfr, ColDesc, ColQty, ColRefDes, ColNotes, HeaderRow) " & _
        "VALUES ('Default', 1, 2, 3, 4, 5, 6, 1)"
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
