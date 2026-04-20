Attribute VB_Name = "modFormLogic"
Option Compare Database
Option Explicit

' ============================================================
'  FORM EVENT LOGIC
'  Wire these to form events in the Access Property Sheet.
'  Instructions on which events to wire are in the comments.
' ============================================================

' ────────────────────────────────────────────────────────────
'  COMPONENTS FORM — BeforeUpdate event
'  Wire: frmComponents > Form > BeforeUpdate
' ────────────────────────────────────────────────────────────
Public Sub Components_BeforeUpdate(frm As Form, Cancel As Integer)

    ' 1. MPN is required
    If IsNull(frm!MPN) Or Trim(frm!MPN & "") = "" Then
        MsgBox "MPN (Part Number) cannot be empty.", vbExclamation
        Cancel = True
        frm!MPN.SetFocus
        Exit Sub
    End If

    ' 2. Manufacturer is required
    If IsNull(frm!ManufacturerID) Then
        MsgBox "Please select a manufacturer.", vbExclamation
        Cancel = True
        frm!ManufacturerID.SetFocus
        Exit Sub
    End If

    ' 3. Duplicate check
    Dim excludeID As Long
    excludeID = Nz(frm!ComponentID, 0)

    If IsDuplicateMPN(frm!MPN & "", CLng(frm!ManufacturerID), excludeID) Then
        MsgBox "A component with MPN '" & frm!MPN & "' already exists for this manufacturer." & _
               vbCrLf & "Please check the existing record.", vbExclamation
        Cancel = True
        frm!MPN.SetFocus
        Exit Sub
    End If

    ' 4. Update ModifiedDate
    frm!ModifiedDate = Now()
End Sub

' ────────────────────────────────────────────────────────────
'  COMPONENTS FORM — AfterUpdate on LifecycleID combo
'  Wire: LifecycleID combo > AfterUpdate
' ────────────────────────────────────────────────────────────
Public Sub LifecycleCombo_AfterUpdate(frm As Form)
    Dim lcName As String
    lcName = Nz(DLookup("StatusName", "Lifecycle_Status", _
                "StatusID = " & Nz(frm!LifecycleID, 0)), "")

    If lcName = "Obsolete" Or lcName = "NRND" Or lcName = "Discontinued" Then
        Dim compID As Long
        compID = Nz(frm!ComponentID, 0)

        Dim altCount As Long
        If compID > 0 Then
            altCount = Nz(DCount("*", "Alternatives", _
                "PrimaryCompID = " & compID & " OR AltCompID = " & compID), 0)
        End If

        If altCount = 0 Then
            MsgBox "WARNING: This component is marked as " & lcName & _
                   " but has NO alternatives defined." & vbCrLf & vbCrLf & _
                   "Consider assigning an alternative to reduce risk.", _
                   vbExclamation, "Risk Alert"
        End If
    End If
End Sub

' ────────────────────────────────────────────────────────────
'  COMPONENTS FORM — New record defaults
'  Wire: frmComponents > Form > OnLoad
' ────────────────────────────────────────────────────────────
Public Sub Components_OnLoad(frm As Form)
    ' Set default lifecycle to "Active" and qual to "Proposed" for new records
    If frm.NewRecord Then
        frm!LifecycleID    = Nz(DLookup("StatusID", "Lifecycle_Status", "StatusName='Active'"), 1)
        frm!QualificationID = Nz(DLookup("QualID", "Qualification_Status", "QualName='Proposed'"), 1)
        frm!CreatedDate    = Now()
        frm!ModifiedDate   = Now()
    End If
End Sub

' ────────────────────────────────────────────────────────────
'  PROJECTS FORM — BeforeUpdate
'  Wire: frmProjects > Form > BeforeUpdate
' ────────────────────────────────────────────────────────────
Public Sub Projects_BeforeUpdate(frm As Form, Cancel As Integer)
    If IsNull(frm!ProjectCode) Or Trim(frm!ProjectCode & "") = "" Then
        MsgBox "Project Code cannot be empty.", vbExclamation
        Cancel = True : frm!ProjectCode.SetFocus : Exit Sub
    End If
    If IsNull(frm!ProjectName) Or Trim(frm!ProjectName & "") = "" Then
        MsgBox "Project Name cannot be empty.", vbExclamation
        Cancel = True : frm!ProjectName.SetFocus : Exit Sub
    End If

    ' Duplicate code check
    Dim excludeID As Long
    excludeID = Nz(frm!ProjectID, 0)
    Dim cnt As Long
    cnt = Nz(DCount("*", "Projects", "ProjectCode = '" & _
             Replace(frm!ProjectCode & "", "'", "''") & "'" & _
             IIf(excludeID > 0, " AND ProjectID <> " & excludeID, "")), 0)
    If cnt > 0 Then
        MsgBox "Project Code '" & frm!ProjectCode & "' already exists.", vbExclamation
        Cancel = True : frm!ProjectCode.SetFocus : Exit Sub
    End If

    frm!CreatedDate = IIf(IsNull(frm!CreatedDate), Now(), frm!CreatedDate)
End Sub

' ────────────────────────────────────────────────────────────
'  AUDIT LOG HELPER — call from anywhere to log a change
' ────────────────────────────────────────────────────────────
Public Sub LogChange(tableName As String, recordID As Long, _
                     fieldName As String, oldVal As Variant, newVal As Variant)
    On Error Resume Next
    CurrentDb.Execute "INSERT INTO Audit_Log " & _
        "(LogDate, UserName, Action, TableName, RecordID, FieldName, OldValue, NewValue) " & _
        "VALUES (Now(), '" & Environ("USERNAME") & "', 'UPDATE', '" & tableName & "', " & _
        recordID & ", '" & fieldName & "', '" & _
        Replace(CStr(Nz(oldVal, "")), "'", "''") & "', '" & _
        Replace(CStr(Nz(newVal, "")), "'", "''") & "')", dbFailOnError
    On Error GoTo 0
End Sub

' ────────────────────────────────────────────────────────────
'  SEARCH — used by a search textbox with a Search button
'  Wire the Search button click to call this, passing the form
' ────────────────────────────────────────────────────────────
Public Sub SearchComponents(frm As Form, searchText As String)
    If Trim(searchText) = "" Then
        frm.RecordSource = "qry_ComponentSearch"
        Exit Sub
    End If

    Dim term As String
    term = Replace(Trim(searchText), "'", "''")

    frm.RecordSource = "SELECT c.ComponentID, c.MPN, m.ManufacturerName, " & _
        "c.Description, c.ComponentType, c.Package, c.Value, " & _
        "ls.StatusName AS Lifecycle, qs.QualName AS QualStatus, c.IsPreferred " & _
        "FROM ((Components AS c " & _
        "INNER JOIN Manufacturers AS m ON c.ManufacturerID = m.ManufacturerID) " & _
        "INNER JOIN Lifecycle_Status AS ls ON c.LifecycleID = ls.StatusID) " & _
        "INNER JOIN Qualification_Status AS qs ON c.QualificationID = qs.QualID " & _
        "WHERE c.MPN LIKE '*" & term & "*' " & _
        "OR c.Description LIKE '*" & term & "*' " & _
        "OR m.ManufacturerName LIKE '*" & term & "*' " & _
        "OR c.ComponentType LIKE '*" & term & "*' " & _
        "ORDER BY c.MPN"
End Sub

' ────────────────────────────────────────────────────────────
'  "WHERE USED" — opens a query filtered to one component
'  Call from a button on frmComponents
' ────────────────────────────────────────────────────────────
Public Sub ShowWhereUsed(componentID As Long, mpn As String)
    If componentID <= 0 Then
        MsgBox "Save the component record first.", vbExclamation
        Exit Sub
    End If

    ' Open BOM extract query filtered to this component
    Dim sql As String
    sql = "SELECT p.ProjectCode, p.ProjectName, bh.BOM_Revision, bh.ImportDate, " & _
          "bi.Quantity, bi.RefDes " & _
          "FROM (BOM_Items AS bi " & _
          "INNER JOIN BOM_Headers AS bh ON bi.HeaderID = bh.HeaderID) " & _
          "INNER JOIN Projects AS p ON bh.ProjectID = p.ProjectID " & _
          "WHERE bi.ComponentID = " & componentID & " AND bh.Status = 'Committed' " & _
          "ORDER BY p.ProjectCode"

    ' Save as a temp query and open it
    Dim db  As DAO.Database
    Dim qdf As DAO.QueryDef
    Set db = CurrentDb

    On Error Resume Next
    db.QueryDefs.Delete "qry_WhereUsed_Temp"
    On Error GoTo 0

    Set qdf = db.CreateQueryDef("qry_WhereUsed_Temp", sql)
    qdf.Close

    DoCmd.OpenQuery "qry_WhereUsed_Temp", acViewNormal
End Sub

' ────────────────────────────────────────────────────────────
'  PRIVATE HELPERS
' ────────────────────────────────────────────────────────────

Private Function IsDuplicateMPN(mpn As String, mfrID As Long, excludeID As Long) As Boolean
    Dim cnt As Long
    Dim whereClause As String
    whereClause = "MPN = '" & Replace(mpn, "'", "''") & "' AND ManufacturerID = " & mfrID
    If excludeID > 0 Then whereClause = whereClause & " AND ComponentID <> " & excludeID
    cnt = Nz(DCount("*", "Components", whereClause), 0)
    IsDuplicateMPN = (cnt > 0)
End Function
