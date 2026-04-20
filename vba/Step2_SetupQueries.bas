Attribute VB_Name = "Step2_SetupQueries"
Option Compare Database
Option Explicit

' ============================================================
'  STEP 2 — Run after Step1.
'  Creates all saved queries visible in the Access query list.
' ============================================================

Public Sub SetupAllQueries()

    On Error GoTo ErrHandler

    Call CreateQuery_ComponentsNoAlternatives
    Call CreateQuery_ObsoleteInProjects
    Call CreateQuery_ComponentAlternativeCount
    Call CreateQuery_BOMExtract
    Call CreateQuery_ImpactAnalysis
    Call CreateQuery_ObsolescenceRisk
    Call CreateQuery_ComponentSearch
    Call CreateQuery_ProjectList
    Call CreateQuery_RecentImports

    MsgBox "All queries created successfully." & vbCrLf & vbCrLf & _
           "Next step: Run Step3_CreateForms.SetupAllForms", _
           vbInformation, "Queries Done"
    Exit Sub

ErrHandler:
    MsgBox "Query setup error: " & Err.Description, vbCritical, "Error"
End Sub

' ============================================================
'  QUERY CREATION HELPERS
' ============================================================

Private Sub SaveQuery(queryName As String, sqlText As String)
    Dim db  As DAO.Database
    Dim qdf As DAO.QueryDef
    Set db = CurrentDb

    ' Delete existing version first so we can recreate cleanly
    On Error Resume Next
    db.QueryDefs.Delete queryName
    On Error GoTo 0

    Set qdf = db.CreateQueryDef(queryName, sqlText)
    qdf.Close
    Set qdf = Nothing
End Sub

' ============================================================
'  THE QUERIES
' ============================================================

Private Sub CreateQuery_ComponentsNoAlternatives()
    Call SaveQuery("qry_ComponentsNoAlternatives", _
        "SELECT c.ComponentID, c.MPN, m.ManufacturerName, ls.StatusName AS Lifecycle, " & _
        "c.Description, c.ComponentType " & _
        "FROM (Components AS c " & _
        "INNER JOIN Manufacturers AS m ON c.ManufacturerID = m.ManufacturerID) " & _
        "INNER JOIN Lifecycle_Status AS ls ON c.LifecycleID = ls.StatusID " & _
        "WHERE c.ComponentID NOT IN ( " & _
        "  SELECT PrimaryCompID FROM Alternatives " & _
        "  UNION SELECT AltCompID FROM Alternatives " & _
        ") " & _
        "ORDER BY ls.StatusName, c.MPN;")
End Sub

Private Sub CreateQuery_ObsoleteInProjects()
    Call SaveQuery("qry_ObsoleteInActiveProjects", _
        "SELECT DISTINCT c.MPN, m.ManufacturerName, ls.StatusName AS Lifecycle, " & _
        "p.ProjectCode, p.ProjectName, p.PCB_Rev " & _
        "FROM (((Components AS c " & _
        "INNER JOIN Manufacturers AS m ON c.ManufacturerID = m.ManufacturerID) " & _
        "INNER JOIN Lifecycle_Status AS ls ON c.LifecycleID = ls.StatusID) " & _
        "INNER JOIN BOM_Items AS bi ON bi.ComponentID = c.ComponentID) " & _
        "INNER JOIN BOM_Headers AS bh ON bi.HeaderID = bh.HeaderID " & _
        "INNER JOIN Projects AS p ON bh.ProjectID = p.ProjectID " & _
        "WHERE ls.StatusName IN ('Obsolete','NRND','Discontinued') " & _
        "AND p.Status = 'Active' " & _
        "AND bh.Status = 'Committed' " & _
        "ORDER BY ls.StatusName, c.MPN;")
End Sub

Private Sub CreateQuery_ComponentAlternativeCount()
    Call SaveQuery("qry_ComponentAltCount", _
        "SELECT c.ComponentID, c.MPN, m.ManufacturerName, " & _
        "Count(a.AltID) AS NumberOfAlternatives " & _
        "FROM (Components AS c " & _
        "INNER JOIN Manufacturers AS m ON c.ManufacturerID = m.ManufacturerID) " & _
        "LEFT JOIN Alternatives AS a ON (c.ComponentID = a.PrimaryCompID OR c.ComponentID = a.AltCompID) " & _
        "GROUP BY c.ComponentID, c.MPN, m.ManufacturerName " & _
        "ORDER BY NumberOfAlternatives DESC;")
End Sub

Private Sub CreateQuery_BOMExtract()
    Call SaveQuery("qry_BOMExtract", _
        "SELECT p.ProjectCode, p.ProjectName, bh.BOM_Revision, bh.ImportDate, " & _
        "c.MPN, m.ManufacturerName, c.Description, c.ComponentType, c.Package, c.Value, " & _
        "bi.Quantity, bi.RefDes, bi.MatchStatus, " & _
        "ls.StatusName AS Lifecycle, qs.QualName AS QualStatus " & _
        "FROM ((((BOM_Items AS bi " & _
        "INNER JOIN BOM_Headers AS bh ON bi.HeaderID = bh.HeaderID) " & _
        "INNER JOIN Projects AS p ON bh.ProjectID = p.ProjectID) " & _
        "LEFT JOIN Components AS c ON bi.ComponentID = c.ComponentID) " & _
        "LEFT JOIN Manufacturers AS m ON c.ManufacturerID = m.ManufacturerID) " & _
        "LEFT JOIN Lifecycle_Status AS ls ON c.LifecycleID = ls.StatusID " & _
        "LEFT JOIN Qualification_Status AS qs ON c.QualificationID = qs.QualID " & _
        "WHERE bh.Status = 'Committed' " & _
        "ORDER BY p.ProjectCode, bi.RefDes;")
End Sub

Private Sub CreateQuery_ImpactAnalysis()
    Call SaveQuery("qry_ImpactAnalysis", _
        "SELECT c.ComponentID, c.MPN, m.ManufacturerName, ls.StatusName AS Lifecycle, " & _
        "Count(DISTINCT bh.ProjectID) AS UsedInProjects, " & _
        "Sum(bi.Quantity) AS TotalQuantityAcrossAll " & _
        "FROM ((Components AS c " & _
        "INNER JOIN Manufacturers AS m ON c.ManufacturerID = m.ManufacturerID) " & _
        "INNER JOIN Lifecycle_Status AS ls ON c.LifecycleID = ls.StatusID) " & _
        "INNER JOIN BOM_Items AS bi ON bi.ComponentID = c.ComponentID " & _
        "INNER JOIN BOM_Headers AS bh ON bi.HeaderID = bh.HeaderID " & _
        "WHERE bh.Status = 'Committed' " & _
        "GROUP BY c.ComponentID, c.MPN, m.ManufacturerName, ls.StatusName " & _
        "ORDER BY UsedInProjects DESC;")
End Sub

Private Sub CreateQuery_ObsolescenceRisk()
    Call SaveQuery("qry_ObsolescenceRisk", _
        "SELECT c.ComponentID, c.MPN, m.ManufacturerName, ls.StatusName AS Lifecycle, " & _
        "Count(bi.ItemID) AS UsageCount, " & _
        "IIf(c.ComponentID IN (SELECT PrimaryCompID FROM Alternatives " & _
        "  UNION SELECT AltCompID FROM Alternatives), 'YES','NO') AS HasAlternative " & _
        "FROM ((Components AS c " & _
        "INNER JOIN Manufacturers AS m ON c.ManufacturerID = m.ManufacturerID) " & _
        "INNER JOIN Lifecycle_Status AS ls ON c.LifecycleID = ls.StatusID) " & _
        "LEFT JOIN BOM_Items AS bi ON bi.ComponentID = c.ComponentID " & _
        "WHERE ls.StatusName IN ('Obsolete','NRND','Discontinued') " & _
        "GROUP BY c.ComponentID, c.MPN, m.ManufacturerName, ls.StatusName, c.ComponentID " & _
        "ORDER BY ls.StatusName, UsageCount DESC;")
End Sub

Private Sub CreateQuery_ComponentSearch()
    Call SaveQuery("qry_ComponentSearch", _
        "SELECT c.ComponentID, c.MPN, m.ManufacturerName, c.Description, " & _
        "c.ComponentType, c.Package, c.Value, ls.StatusName AS Lifecycle, " & _
        "qs.QualName AS QualStatus, c.IsPreferred " & _
        "FROM ((Components AS c " & _
        "INNER JOIN Manufacturers AS m ON c.ManufacturerID = m.ManufacturerID) " & _
        "INNER JOIN Lifecycle_Status AS ls ON c.LifecycleID = ls.StatusID) " & _
        "INNER JOIN Qualification_Status AS qs ON c.QualificationID = qs.QualID " & _
        "ORDER BY c.MPN;")
End Sub

Private Sub CreateQuery_ProjectList()
    Call SaveQuery("qry_ProjectList", _
        "SELECT p.ProjectID, p.ProjectCode, p.ProjectName, p.PCB_Rev, p.Owner, " & _
        "p.Status, p.CreatedDate, " & _
        "Count(bh.HeaderID) AS ImportCount " & _
        "FROM Projects AS p " & _
        "LEFT JOIN BOM_Headers AS bh ON p.ProjectID = bh.ProjectID " & _
        "GROUP BY p.ProjectID, p.ProjectCode, p.ProjectName, p.PCB_Rev, " & _
        "p.Owner, p.Status, p.CreatedDate " & _
        "ORDER BY p.ProjectCode;")
End Sub

Private Sub CreateQuery_RecentImports()
    Call SaveQuery("qry_RecentImports", _
        "SELECT TOP 20 bh.HeaderID, p.ProjectCode, p.ProjectName, " & _
        "bh.BOM_Revision, bh.ImportDate, bh.ImportedBy, " & _
        "bh.SourceFile, bh.RowCount, bh.Status " & _
        "FROM BOM_Headers AS bh " & _
        "INNER JOIN Projects AS p ON bh.ProjectID = p.ProjectID " & _
        "WHERE bh.Status = 'Committed' " & _
        "ORDER BY bh.ImportDate DESC;")
End Sub
