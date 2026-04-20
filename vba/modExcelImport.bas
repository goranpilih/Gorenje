Attribute VB_Name = "modExcelImport"
Option Compare Database
Option Explicit

' ============================================================
'  EXCEL BOM IMPORT ENGINE
'  Called by frmBOMImport wizard buttons.
' ============================================================

Private Const STAGING As String = "tbl_ImportStaging"

' ────────────────────────────────────────────────────────────
'  STEP 1 helper: Open file browser and put path in form
' ────────────────────────────────────────────────────────────
Public Sub BrowseForFile(frm As Form)
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    fd.Title = "Select Excel BOM File"
    fd.Filters.Clear
    fd.Filters.Add "Excel Files", "*.xlsx; *.xls; *.xlsm"
    fd.AllowMultiSelect = False

    If fd.Show = True Then
        frm.txtFilePath.Value = fd.SelectedItems(1)
    End If
End Sub

' ────────────────────────────────────────────────────────────
'  STEP 1 → STEP 2: validate inputs then advance tab
' ────────────────────────────────────────────────────────────
Public Sub Step1_ToStep2(frm As Form)
    If IsNull(frm.cboProject.Value) Or frm.cboProject.Value = "" Then
        MsgBox "Please select a project first.", vbExclamation
        Exit Sub
    End If
    If IsNull(frm.txtFilePath.Value) Or frm.txtFilePath.Value = "" Then
        MsgBox "Please select an Excel file first.", vbExclamation
        Exit Sub
    End If
    If Dir(frm.txtFilePath.Value) = "" Then
        MsgBox "File not found: " & frm.txtFilePath.Value, vbCritical
        Exit Sub
    End If
    frm.tabWizard.Value = 1   ' Move to Step 2
End Sub

' ────────────────────────────────────────────────────────────
'  STEP 2 → STEP 3: parse Excel into staging table
' ────────────────────────────────────────────────────────────
Public Sub Step2_ParseFile(frm As Form)

    ' Read column numbers from form
    Dim colMPN    As Integer : colMPN    = Val(Nz(frm.cboColMPN.Value,   "1"))
    Dim colMfr    As Integer : colMfr    = Val(Nz(frm.cboColMfr.Value,   "2"))
    Dim colDesc   As Integer : colDesc   = Val(Nz(frm.cboColDesc.Value,  "3"))
    Dim colQty    As Integer : colQty    = Val(Nz(frm.cboColQty.Value,   "4"))
    Dim colRefDes As Integer : colRefDes = Val(Nz(frm.cboColRefDes.Value,"5"))
    Dim colNotes  As Integer : colNotes  = Val(Nz(frm.cboColNotes.Value, "6"))
    Dim hdrRow    As Integer : hdrRow    = Val(Nz(frm.txtHeaderRow.Value,"1"))

    If colMPN < 1 Then
        MsgBox "MPN column number must be 1 or greater.", vbExclamation
        Exit Sub
    End If

    Dim filePath As String
    filePath = frm.txtFilePath.Value

    ' Parse Excel into staging
    Dim rowCount As Long
    rowCount = ParseExcelToStaging(filePath, colMPN, colMfr, colDesc, _
                                   colQty, colRefDes, colNotes, hdrRow)
    If rowCount < 0 Then Exit Sub  ' Error already shown

    ' Auto-match staged rows to known components
    MatchStagingToComponents

    ' Update summary label on step 3
    Dim matched   As Long
    Dim unmatched As Long
    matched   = DCount("*", STAGING, "MatchStatus LIKE 'Matched*'")
    unmatched = DCount("*", STAGING, "MatchStatus = 'Unmatched'")

    frm.lblSumMatch.Caption = _
        "Parsed " & rowCount & " rows.  " & _
        "Matched: " & matched & "   " & _
        "Unmatched (will be created as new): " & unmatched

    ' Refresh the subform
    frm.sfPreview.Requery

    frm.tabWizard.Value = 2   ' Move to Step 3
End Sub

' ────────────────────────────────────────────────────────────
'  STEP 3 → STEP 4: build confirm summary
' ────────────────────────────────────────────────────────────
Public Sub Step3_ToConfirm(frm As Form)
    Dim total     As Long : total     = DCount("*", STAGING)
    Dim matched   As Long : matched   = DCount("*", STAGING, "MatchStatus LIKE 'Matched*'")
    Dim newOnes   As Long : newOnes   = DCount("*", STAGING, "MatchStatus = 'Unmatched'")
    Dim skipped   As Long : skipped   = DCount("*", STAGING, "MatchStatus = 'Skipped'")

    frm.lblConfProject.Caption = "Project:        " & frm.cboProject.Column(1)
    frm.lblConfFile.Caption    = "Source file:    " & frm.txtFilePath.Value
    frm.lblConfTotal.Caption   = "Total rows:     " & total
    frm.lblConfMatched.Caption = "Matched:        " & matched
    frm.lblConfNew.Caption     = "New components (will be created): " & newOnes & _
                                  "   Skipped: " & skipped

    frm.tabWizard.Value = 3   ' Move to Step 4
End Sub

' ────────────────────────────────────────────────────────────
'  STEP 4: COMMIT — called by Commit button on frmBOMImport
' ────────────────────────────────────────────────────────────
Public Sub CommitImport_FromForm(frm As Form)
    Dim projectID As Long
    projectID = frm.cboProject.Value

    Dim bomRev As String
    bomRev = Nz(frm.txtBOMRev.Value, "")

    Dim srcFile As String
    srcFile = frm.txtFilePath.Value

    If MsgBox("Commit the import? This cannot be undone.", _
              vbYesNo + vbQuestion, "Confirm Import") = vbNo Then Exit Sub

    Dim ok As Boolean
    ok = CommitImport(projectID, bomRev, srcFile)

    If ok Then
        MsgBox "Import committed successfully!", vbInformation, "Done"
        DoCmd.Close acForm, "frmBOMImport"
    End If
End Sub

' ============================================================
'  INTERNAL ENGINE
' ============================================================

Private Function ParseExcelToStaging( _
    filePath  As String, _
    colMPN    As Integer, _
    colMfr    As Integer, _
    colDesc   As Integer, _
    colQty    As Integer, _
    colRefDes As Integer, _
    colNotes  As Integer, _
    hdrRow    As Integer _
) As Long

    Dim xlApp  As Object
    Dim xlWB   As Object
    Dim xlWS   As Object
    Dim db     As DAO.Database
    Dim rs     As DAO.Recordset
    Dim lastRow As Long
    Dim i       As Long
    Dim cnt     As Long

    On Error GoTo ErrHandler

    ' Clear staging
    CurrentDb.Execute "DELETE FROM " & STAGING, dbFailOnError

    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False

    Set xlWB = xlApp.Workbooks.Open(filePath, ReadOnly:=True)
    Set xlWS = xlWB.Sheets(1)

    ' Find last used row in MPN column
    lastRow = xlWS.Cells(xlWS.Rows.Count, colMPN).End(-4162).Row

    Set db = CurrentDb
    Set rs = db.OpenRecordset(STAGING, dbOpenDynaset)

    cnt = 0
    For i = hdrRow + 1 To lastRow
        Dim mpnVal As String
        mpnVal = Trim(CStr(Nz(xlWS.Cells(i, colMPN).Value, "")))

        If Len(mpnVal) = 0 Then GoTo SkipRow  ' Skip blank MPN rows

        rs.AddNew
            rs!MPN_Raw          = Left(mpnVal, 100)
            rs!Manufacturer_Raw = Left(Trim(CStr(Nz(xlWS.Cells(i, colMfr).Value, ""))), 100)
            rs!Description_Raw  = Left(Trim(CStr(Nz(xlWS.Cells(i, colDesc).Value, ""))), 255)
            rs!Quantity         = SafeQty(xlWS.Cells(i, colQty).Value)
            rs!RefDes           = Left(Trim(CStr(Nz(xlWS.Cells(i, colRefDes).Value, ""))), 500)
            rs!Notes            = Left(Trim(CStr(Nz(xlWS.Cells(i, colNotes).Value, ""))), 500)
            rs!MatchStatus      = "Unmatched"
        rs.Update
        cnt = cnt + 1

SkipRow:
    Next i

    rs.Close
    xlWB.Close False
    xlApp.Quit

    Set rs = Nothing
    Set xlWS = Nothing : Set xlWB = Nothing : Set xlApp = Nothing

    ParseExcelToStaging = cnt
    Exit Function

ErrHandler:
    Dim msg As String
    msg = "Excel parse error at row " & i & ": " & Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not xlWB Is Nothing Then xlWB.Close False
    If Not xlApp Is Nothing Then xlApp.Quit
    On Error GoTo 0
    MsgBox msg, vbCritical, "Parse Error"
    ParseExcelToStaging = -1
End Function

' ────────────────────────────────────────────────────────────

Private Sub MatchStagingToComponents()
    Dim db      As DAO.Database
    Dim rsStage As DAO.Recordset
    Dim rsComp  As DAO.Recordset
    Dim sql     As String

    Set db = CurrentDb
    Set rsStage = db.OpenRecordset(STAGING, dbOpenDynaset)

    Do While Not rsStage.EOF
        Dim mpn As String : mpn = Nz(rsStage!MPN_Raw, "")
        Dim mfr As String : mfr = Nz(rsStage!Manufacturer_Raw, "")

        ' Try MPN + manufacturer match
        sql = "SELECT c.ComponentID FROM Components AS c " & _
              "INNER JOIN Manufacturers AS m ON c.ManufacturerID = m.ManufacturerID " & _
              "WHERE c.MPN = '" & Esc(mpn) & "' " & _
              "AND m.ManufacturerName = '" & Esc(mfr) & "'"
        Set rsComp = db.OpenRecordset(sql, dbOpenSnapshot)

        If Not rsComp.EOF Then
            rsStage.Edit
                rsStage!ComponentID = rsComp!ComponentID
                rsStage!MatchStatus = "Matched"
            rsStage.Update
        ElseIf Len(mfr) = 0 Then
            ' Fallback: MPN only
            sql = "SELECT ComponentID FROM Components WHERE MPN = '" & Esc(mpn) & "'"
            Set rsComp = db.OpenRecordset(sql, dbOpenSnapshot)
            If Not rsComp.EOF Then
                rsStage.Edit
                    rsStage!ComponentID = rsComp!ComponentID
                    rsStage!MatchStatus = "Matched_MPNOnly"
                rsStage.Update
            End If
        End If

        rsComp.Close
        rsStage.MoveNext
    Loop
    rsStage.Close
End Sub

' ────────────────────────────────────────────────────────────

Public Function CommitImport( _
    projectID As Long, _
    bomRev    As String, _
    srcFile   As String _
) As Boolean

    Dim db       As DAO.Database
    Dim ws       As DAO.Workspace
    Dim rsStage  As DAO.Recordset
    Dim headerID As Long
    Dim rowCount As Long
    Dim sql      As String

    On Error GoTo ErrHandler

    Set db = CurrentDb
    Set ws = DBEngine.Workspaces(0)

    ws.BeginTrans

    ' 1. Write BOM_Header
    sql = "INSERT INTO BOM_Headers " & _
          "(ProjectID, BOM_Revision, ImportDate, ImportedBy, SourceFile, Status) " & _
          "VALUES (" & projectID & ", '" & Esc(bomRev) & "', Now(), " & _
          "'" & Esc(Environ("USERNAME")) & "', " & _
          "'" & Esc(srcFile) & "', 'Pending')"
    db.Execute sql, dbFailOnError
    headerID = db.OpenRecordset("SELECT @@IDENTITY")(0)

    ' 2. Create new Components for unmatched rows
    Set rsStage = db.OpenRecordset( _
        "SELECT * FROM " & STAGING & " WHERE MatchStatus = 'Unmatched'", _
        dbOpenSnapshot)

    Dim defLC   As Long : defLC   = GetLifecycleID("Active", db)
    Dim defQual As Long : defQual = GetQualID("Proposed", db)

    Do While Not rsStage.EOF
        Dim mfrID As Long
        mfrID = GetOrCreateManufacturer(Nz(rsStage!Manufacturer_Raw, "Unknown"), db)

        sql = "INSERT INTO Components " & _
              "(MPN, ManufacturerID, Description, LifecycleID, QualificationID, CreatedDate, ModifiedDate) " & _
              "VALUES ('" & Esc(Nz(rsStage!MPN_Raw, "")) & "', " & mfrID & ", " & _
              "'" & Esc(Nz(rsStage!Description_Raw, "")) & "', " & _
              defLC & ", " & defQual & ", Now(), Now())"
        db.Execute sql, dbFailOnError

        Dim newCID As Long
        newCID = db.OpenRecordset("SELECT @@IDENTITY")(0)

        ' Tag staging row with new ComponentID
        db.Execute "UPDATE " & STAGING & " SET ComponentID = " & newCID & _
                   ", MatchStatus = 'NewAdded' WHERE RowNum = " & rsStage!RowNum, _
                   dbFailOnError

        WriteAudit db, "INSERT", "Components", newCID, "MPN", "", Nz(rsStage!MPN_Raw, "")
        rsStage.MoveNext
    Loop
    rsStage.Close

    ' 3. Insert BOM_Items for all non-skipped rows
    Set rsStage = db.OpenRecordset( _
        "SELECT * FROM " & STAGING & " WHERE MatchStatus <> 'Skipped'", _
        dbOpenSnapshot)

    Do While Not rsStage.EOF
        Dim cIDstr As String
        If IsNull(rsStage!ComponentID) Then
            cIDstr = "NULL"
        Else
            cIDstr = CStr(rsStage!ComponentID)
        End If

        sql = "INSERT INTO BOM_Items " & _
              "(HeaderID, ComponentID, MPN_Raw, Manufacturer_Raw, Description_Raw, " & _
              "Quantity, RefDes, Notes, MatchStatus) " & _
              "VALUES (" & headerID & ", " & cIDstr & ", " & _
              "'" & Esc(Nz(rsStage!MPN_Raw, "")) & "', " & _
              "'" & Esc(Nz(rsStage!Manufacturer_Raw, "")) & "', " & _
              "'" & Esc(Nz(rsStage!Description_Raw, "")) & "', " & _
              Nz(rsStage!Quantity, 1) & ", " & _
              "'" & Esc(Nz(rsStage!RefDes, "")) & "', " & _
              "'" & Esc(Nz(rsStage!Notes, "")) & "', " & _
              "'" & Nz(rsStage!MatchStatus, "Unmatched") & "')"
        db.Execute sql, dbFailOnError
        rowCount = rowCount + 1
        rsStage.MoveNext
    Loop
    rsStage.Close

    ' 4. Finalise header
    db.Execute "UPDATE BOM_Headers SET Status = 'Committed', RowCount = " & rowCount & _
               ", ImportDate = Now() WHERE HeaderID = " & headerID, dbFailOnError

    WriteAudit db, "IMPORT", "BOM_Headers", headerID, "SourceFile", "", srcFile

    ws.CommitTrans
    CommitImport = True
    Exit Function

ErrHandler:
    ws.Rollback
    MsgBox "IMPORT FAILED — all changes rolled back." & vbCrLf & vbCrLf & _
           Err.Description, vbCritical, "Import Error"
    CommitImport = False
End Function

' ============================================================
'  PRIVATE HELPERS
' ============================================================

Private Function Esc(s As String) As String
    Esc = Replace(s, "'", "''")
End Function

Private Function SafeQty(v As Variant) As Double
    If IsNumeric(v) Then SafeQty = CDbl(v) Else SafeQty = 1
End Function

Private Function GetOrCreateManufacturer(mfrName As String, db As DAO.Database) As Long
    If Len(Trim(mfrName)) = 0 Then mfrName = "Unknown"
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset( _
        "SELECT ManufacturerID FROM Manufacturers WHERE ManufacturerName = '" & Esc(mfrName) & "'", _
        dbOpenSnapshot)
    If Not rs.EOF Then
        GetOrCreateManufacturer = rs!ManufacturerID
    Else
        db.Execute "INSERT INTO Manufacturers (ManufacturerName) VALUES ('" & Esc(mfrName) & "')", dbFailOnError
        GetOrCreateManufacturer = db.OpenRecordset("SELECT @@IDENTITY")(0)
    End If
    rs.Close
End Function

Private Function GetLifecycleID(statusName As String, db As DAO.Database) As Long
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset( _
        "SELECT StatusID FROM Lifecycle_Status WHERE StatusName = '" & Esc(statusName) & "'", _
        dbOpenSnapshot)
    If rs.EOF Then
        Set rs = db.OpenRecordset("SELECT TOP 1 StatusID FROM Lifecycle_Status ORDER BY SortOrder")
    End If
    GetLifecycleID = rs!StatusID
    rs.Close
End Function

Private Function GetQualID(qualName As String, db As DAO.Database) As Long
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset( _
        "SELECT QualID FROM Qualification_Status WHERE QualName = '" & Esc(qualName) & "'", _
        dbOpenSnapshot)
    If rs.EOF Then
        Set rs = db.OpenRecordset("SELECT TOP 1 QualID FROM Qualification_Status ORDER BY SortOrder")
    End If
    GetQualID = rs!QualID
    rs.Close
End Function

Private Sub WriteAudit(db As DAO.Database, action As String, tbl As String, _
                        recID As Long, fld As String, oldV As String, newV As String)
    On Error Resume Next
    db.Execute "INSERT INTO Audit_Log (LogDate, UserName, Action, TableName, RecordID, " & _
               "FieldName, OldValue, NewValue) VALUES (Now(), " & _
               "'" & Esc(Environ("USERNAME")) & "', '" & action & "', '" & tbl & "', " & _
               recID & ", '" & fld & "', '" & Esc(oldV) & "', '" & Esc(newV) & "')", dbFailOnError
    On Error GoTo 0
End Sub
