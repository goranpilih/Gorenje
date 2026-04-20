Attribute VB_Name = "Step3_CreateForms"
Option Compare Database
Option Explicit

' ============================================================
'  STEP 3 — Builds all forms programmatically.
'  Run SetupAllForms once. Forms will appear in the Form list.
' ============================================================

' Layout constants (twips — Access internal unit, 1440 twips = 1 inch)
Private Const LABEL_W   As Long = 1800
Private Const FIELD_W   As Long = 3200
Private Const ROW_H     As Long = 340
Private Const ROW_GAP   As Long = 60
Private Const LEFT_L    As Long = 120   ' Label left edge
Private Const LEFT_F    As Long = 1980  ' Field left edge
Private Const TOP_START As Long = 200   ' First row top

Public Sub SetupAllForms()
    On Error GoTo ErrHandler

    Call BuildForm_Manufacturers
    Call BuildForm_Components
    Call BuildForm_Projects
    Call BuildForm_BOMImport
    Call BuildForm_Dashboard

    MsgBox "All forms created!" & vbCrLf & vbCrLf & _
           "Open frmDashboard to start using the application.", _
           vbInformation, "Forms Ready"
    Exit Sub
ErrHandler:
    MsgBox "Form creation error: " & Err.Description, vbCritical, "Error"
End Sub

' ============================================================
'  MANUFACTURERS FORM
' ============================================================
Private Sub BuildForm_Manufacturers()
    DeleteFormIfExists "frmManufacturers"

    Dim frm As Form
    Set frm = CreateForm()
    frm.RecordSource = "Manufacturers"
    frm.Caption = "Manufacturers"
    frm.ScrollBars = 0
    frm.RecordSelectors = False
    frm.NavigationButtons = True
    frm.AutoCenter = True
    frm.Width = 6000

    Dim fields(3) As String
    Dim labels(3) As String
    fields(0) = "ManufacturerName" : labels(0) = "Manufacturer Name:"
    fields(1) = "ShortCode"        : labels(1) = "Short Code:"
    fields(2) = "Website"          : labels(2) = "Website:"
    fields(3) = "Notes"            : labels(3) = "Notes:"

    Dim i As Integer
    For i = 0 To 3
        Dim t As Long
        t = TOP_START + i * (ROW_H + ROW_GAP)
        Dim h As Long
        h = IIf(fields(i) = "Notes", ROW_H * 3, ROW_H)
        AddLabelField frm, labels(i), fields(i), t, h
    Next i

    AddCommandButton frm, "cmdSave", "Save & Close", 200, TOP_START + 4 * (ROW_H + ROW_GAP) + 200, _
        "DoCmd.RunCommand acCmdSaveRecord : DoCmd.Close"

    SaveAndCloseForm frm, "frmManufacturers"
End Sub

' ============================================================
'  COMPONENTS FORM
' ============================================================
Private Sub BuildForm_Components()
    DeleteFormIfExists "frmComponents"

    Dim frm As Form
    Set frm = CreateForm()
    frm.RecordSource = "qry_ComponentSearch"
    frm.Caption = "Components"
    frm.ScrollBars = 2  ' Vertical only
    frm.RecordSelectors = False
    frm.NavigationButtons = True
    frm.AutoCenter = True
    frm.Width = 7200

    ' --- Text fields ---
    Dim fNames(6) As String : Dim fLabels(6) As String
    fNames(0) = "MPN"             : fLabels(0) = "MPN (Part Number):"
    fNames(1) = "ManufacturerID"  : fLabels(1) = "Manufacturer:"
    fNames(2) = "ComponentType"   : fLabels(2) = "Type:"
    fNames(3) = "Package"         : fLabels(3) = "Package:"
    fNames(4) = "ComponentValue"   : fLabels(4) = "Value:"
    fNames(5) = "Description"     : fLabels(5) = "Description:"
    fNames(6) = "Datasheet_URL"   : fLabels(6) = "Datasheet URL:"

    Dim i As Integer
    For i = 0 To 6
        Dim t As Long
        t = TOP_START + i * (ROW_H + ROW_GAP)
        AddLabelField frm, fLabels(i), fNames(i), t, ROW_H
    Next i

    ' --- Lifecycle combo ---
    Dim topNext As Long
    topNext = TOP_START + 7 * (ROW_H + ROW_GAP)
    AddBoundCombo frm, "LifecycleID", "Lifecycle Status:", _
        "SELECT StatusID, StatusName FROM Lifecycle_Status ORDER BY SortOrder", _
        topNext, ROW_H

    ' --- Qualification combo ---
    topNext = topNext + ROW_H + ROW_GAP
    AddBoundCombo frm, "QualificationID", "Qual. Status:", _
        "SELECT QualID, QualName FROM Qualification_Status ORDER BY SortOrder", _
        topNext, ROW_H

    ' --- Notes ---
    topNext = topNext + ROW_H + ROW_GAP
    AddLabelField frm, "Notes:", "Notes", topNext, ROW_H * 3

    ' --- Buttons ---
    topNext = topNext + ROW_H * 3 + ROW_GAP + 200
    AddCommandButton frm, "cmdNew", "New Component", 120, topNext, _
        "DoCmd.GoToRecord , , acNewRec"
    AddCommandButton frm, "cmdSave", "Save", 2000, topNext, _
        "DoCmd.RunCommand acCmdSaveRecord"
    AddCommandButton frm, "cmdDelete", "Delete", 3500, topNext, _
        "If MsgBox(""Delete this component?"", vbYesNo) = vbYes Then DoCmd.RunCommand acCmdDeleteRecord"
    AddCommandButton frm, "cmdClose", "Close", 5000, topNext, _
        "DoCmd.Close"

    SaveAndCloseForm frm, "frmComponents"
End Sub

' ============================================================
'  PROJECTS FORM
' ============================================================
Private Sub BuildForm_Projects()
    DeleteFormIfExists "frmProjects"

    Dim frm As Form
    Set frm = CreateForm()
    frm.RecordSource = "Projects"
    frm.Caption = "Projects"
    frm.RecordSelectors = False
    frm.NavigationButtons = True
    frm.AutoCenter = True
    frm.Width = 6500

    Dim fNames(5) As String : Dim fLabels(5) As String
    fNames(0) = "ProjectCode"  : fLabels(0) = "Project Code:"
    fNames(1) = "ProjectName"  : fLabels(1) = "Project Name:"
    fNames(2) = "PCB_Rev"      : fLabels(2) = "PCB Revision:"
    fNames(3) = "Owner"        : fLabels(3) = "Owner:"
    fNames(4) = "ProjectStatus" : fLabels(4) = "Status:"
    fNames(5) = "Notes"        : fLabels(5) = "Notes:"

    Dim i As Integer
    For i = 0 To 5
        Dim t As Long
        t = TOP_START + i * (ROW_H + ROW_GAP)
        Dim h As Long
        h = IIf(fNames(i) = "Notes", ROW_H * 3, ROW_H)
        AddLabelField frm, fLabels(i), fNames(i), t, h
    Next i

    Dim topNext As Long
    topNext = TOP_START + 6 * (ROW_H + ROW_GAP) + ROW_H * 2 + 200
    AddCommandButton frm, "cmdNew", "New Project", 120, topNext, _
        "DoCmd.GoToRecord , , acNewRec"
    AddCommandButton frm, "cmdSave", "Save", 2000, topNext, _
        "DoCmd.RunCommand acCmdSaveRecord"
    AddCommandButton frm, "cmdViewBOM", "View BOM", 3500, topNext, _
        "DoCmd.OpenForm ""frmBOMImport"""
    AddCommandButton frm, "cmdClose", "Close", 5000, topNext, _
        "DoCmd.Close"

    SaveAndCloseForm frm, "frmProjects"
End Sub

' ============================================================
'  BOM IMPORT FORM  (4-page tab wizard)
' ============================================================
Private Sub BuildForm_BOMImport()
    DeleteFormIfExists "frmBOMImport"

    Dim frm As Form
    Set frm = CreateForm()
    frm.Caption = "BOM Import Wizard"
    frm.RecordSource = ""
    frm.ScrollBars = 0
    frm.RecordSelectors = False
    frm.NavigationButtons = False
    frm.AutoCenter = True
    frm.Width = 8000

    ' Tab control spanning most of the form
    Dim tabCtl As Control
    Set tabCtl = CreateControl(frm.Name, acTabCtl, acDetail, "", "", 120, 120, 7700, 5800)
    tabCtl.Name = "tabWizard"

    ' Give pages meaningful names and captions
    tabCtl.Pages(0).Name = "pgStep1"
    tabCtl.Pages(0).Caption = "Step 1: Select File"
    tabCtl.Pages(1).Name = "pgStep2"
    tabCtl.Pages(1).Caption = "Step 2: Map Columns"
    tabCtl.Pages(2).Name = "pgStep3"
    tabCtl.Pages(2).Caption = "Step 3: Resolve"
    tabCtl.Pages(3).Name = "pgStep4"
    tabCtl.Pages(3).Caption = "Step 4: Confirm & Import"

    ' ---- PAGE 1 controls ----
    ' Project combo
    Dim cboProj As Control
    Set cboProj = CreateControl(frm.Name, acComboBox, acDetail, "pgStep1", _
        "cboProject", 2000, 500, 3500, ROW_H)
    cboProj.Name = "cboProject"
    cboProj.RowSourceType = "Table/Query"
    cboProj.RowSource = "SELECT ProjectID, ProjectCode & ' — ' & ProjectName FROM Projects ORDER BY ProjectCode"
    cboProj.ColumnCount = 2
    cboProj.ColumnWidths = "0;3500"
    cboProj.BoundColumn = 1
    AddPlainLabel frm, "pgStep1", "lblProjLbl", "Project:", 120, 500, LABEL_W, ROW_H

    ' BOM Revision textbox
    Dim txtRev As Control
    Set txtRev = CreateControl(frm.Name, acTextBox, acDetail, "pgStep1", _
        "txtBOMRev", 2000, 500 + ROW_H + ROW_GAP, 1500, ROW_H)
    txtRev.Name = "txtBOMRev"
    AddPlainLabel frm, "pgStep1", "lblRevLbl", "BOM Revision:", 120, 500 + ROW_H + ROW_GAP, LABEL_W, ROW_H

    ' File path textbox (read-only, filled by Browse button)
    Dim txtFile As Control
    Set txtFile = CreateControl(frm.Name, acTextBox, acDetail, "pgStep1", _
        "txtFilePath", 2000, 500 + 2 * (ROW_H + ROW_GAP), 4000, ROW_H)
    txtFile.Name = "txtFilePath"
    txtFile.Locked = True
    txtFile.BackColor = 15921906  ' Light gray
    AddPlainLabel frm, "pgStep1", "lblFileLbl", "Excel File:", 120, 500 + 2 * (ROW_H + ROW_GAP), LABEL_W, ROW_H

    ' Browse button
    AddCommandButton frm, "cmdBrowse", "Browse...", 6100, 500 + 2 * (ROW_H + ROW_GAP), _
        "Call modExcelImport.BrowseForFile(Me)"
    ' Move button to pgStep1 — Access puts new controls on last tab by default;
    ' reassign via .Parent property approach (done via caption trick — VBA will handle in event)

    ' Header row number
    Dim txtHdr As Control
    Set txtHdr = CreateControl(frm.Name, acTextBox, acDetail, "pgStep1", _
        "txtHeaderRow", 2000, 500 + 3 * (ROW_H + ROW_GAP), 600, ROW_H)
    txtHdr.Name = "txtHeaderRow"
    txtHdr.DefaultValue = "1"
    AddPlainLabel frm, "pgStep1", "lblHdrLbl", "Header Row #:", 120, 500 + 3 * (ROW_H + ROW_GAP), LABEL_W, ROW_H

    ' Next button (step 1 -> 2)
    AddCommandButton frm, "cmdStep1Next", "Next →", 6000, 5200, _
        "Call modExcelImport.Step1_ToStep2(Me)"

    ' ---- PAGE 2 controls — column mapping combos ----
    Dim colNames(5) As String
    Dim colLabels(5) As String
    colNames(0) = "cboColMPN"   : colLabels(0) = "MPN column number:"
    colNames(1) = "cboColMfr"   : colLabels(1) = "Manufacturer column:"
    colNames(2) = "cboColDesc"  : colLabels(2) = "Description column:"
    colNames(3) = "cboColQty"   : colLabels(3) = "Quantity column:"
    colNames(4) = "cboColRefDes": colLabels(4) = "Ref Des column:"
    colNames(5) = "cboColNotes" : colLabels(5) = "Notes column:"

    Dim ci As Integer
    For ci = 0 To 5
        Dim ct As Long
        ct = 400 + ci * (ROW_H + ROW_GAP + 100)
        Dim colCtl As Control
        Set colCtl = CreateControl(frm.Name, acTextBox, acDetail, "pgStep2", _
            colNames(ci), 2400, ct, 700, ROW_H)
        colCtl.Name = colNames(ci)
        colCtl.DefaultValue = CStr(ci + 1)
        AddPlainLabel frm, "pgStep2", "lbl" & colNames(ci), colLabels(ci), _
            120, ct, 2200, ROW_H
    Next ci

    ' Instruction label on page 2
    Dim lblInstr As Control
    Set lblInstr = CreateControl(frm.Name, acLabel, acDetail, "pgStep2", _
        "lblColInstr", 120, 400 + 6 * (ROW_H + ROW_GAP + 100) + 200, 7000, ROW_H * 2)
    lblInstr.Caption = "Enter the column number for each field (1 = column A, 2 = column B, etc.)"
    lblInstr.ForeColor = RGB(0, 0, 180)

    AddCommandButton frm, "cmdStep2Back", "← Back", 200, 5200, _
        "Me.tabWizard.Value = 0"
    AddCommandButton frm, "cmdStep2Next", "Parse File →", 5800, 5200, _
        "Call modExcelImport.Step2_ParseFile(Me)"

    ' ---- PAGE 3 — preview results (read-only subform placeholder) ----
    Dim lblPrev As Control
    Set lblPrev = CreateControl(frm.Name, acLabel, acDetail, "pgStep3", _
        "lblPreviewInfo", 120, 300, 7000, ROW_H * 2)
    lblPrev.Caption = "Preview of parsed data — review the Match Status column."
    lblPrev.ForeColor = RGB(0, 100, 0)

    ' Subform showing staging table
    Dim sfPrev As Control
    Set sfPrev = CreateControl(frm.Name, acSubform, acDetail, "pgStep3", _
        "sfPreview", 120, 300 + ROW_H * 2 + 200, 7500, 3800)
    sfPrev.Name = "sfPreview"
    sfPrev.SourceObject = "tbl_ImportStaging"

    ' Summary labels
    Dim lblSumMatch As Control
    Set lblSumMatch = CreateControl(frm.Name, acLabel, acDetail, "pgStep3", _
        "lblSumMatch", 120, 4500, 7000, ROW_H)
    lblSumMatch.Caption = "(Parse the file first — summary will appear here)"
    lblSumMatch.Name = "lblSumMatch"

    AddCommandButton frm, "cmdStep3Back", "← Back", 200, 5200, _
        "Me.tabWizard.Value = 1"
    AddCommandButton frm, "cmdStep3Next", "Proceed to Confirm →", 5400, 5200, _
        "Call modExcelImport.Step3_ToConfirm(Me)"

    ' ---- PAGE 4 — confirm and import ----
    Dim lblConf As Control
    Set lblConf = CreateControl(frm.Name, acLabel, acDetail, "pgStep4", _
        "lblConfirmTitle", 120, 300, 7000, ROW_H)
    lblConf.Caption = "Review details below, then click COMMIT IMPORT."
    lblConf.ForeColor = RGB(180, 0, 0)
    lblConf.FontBold = True

    Dim summaryLabels(4) As String
    summaryLabels(0) = "lblConfProject"
    summaryLabels(1) = "lblConfFile"
    summaryLabels(2) = "lblConfTotal"
    summaryLabels(3) = "lblConfMatched"
    summaryLabels(4) = "lblConfNew"

    Dim si As Integer
    For si = 0 To 4
        Dim sumCtl As Control
        Set sumCtl = CreateControl(frm.Name, acLabel, acDetail, "pgStep4", _
            summaryLabels(si), 120, 700 + si * (ROW_H + ROW_GAP + 100), 7000, ROW_H)
        sumCtl.Caption = summaryLabels(si) & ": (will be filled automatically)"
        sumCtl.Name = summaryLabels(si)
    Next si

    AddCommandButton frm, "cmdStep4Back", "← Back", 200, 5200, _
        "Me.tabWizard.Value = 2"
    AddCommandButton frm, "cmdCommit", "COMMIT IMPORT", 4500, 5200, _
        "Call modExcelImport.CommitImport_FromForm(Me)"

    SaveAndCloseForm frm, "frmBOMImport"
End Sub

' ============================================================
'  DASHBOARD FORM
' ============================================================
Private Sub BuildForm_Dashboard()
    DeleteFormIfExists "frmDashboard"

    Dim frm As Form
    Set frm = CreateForm()
    frm.Caption = "Component Management System"
    frm.RecordSource = ""
    frm.ScrollBars = 0
    frm.RecordSelectors = False
    frm.NavigationButtons = False
    frm.AutoCenter = True
    frm.Width = 8000

    ' Title label
    Dim lblTitle As Control
    Set lblTitle = CreateControl(frm.Name, acLabel, acDetail, "", _
        "lblTitle", 200, 200, 7600, 600)
    lblTitle.Caption = "COMPONENT MANAGEMENT SYSTEM"
    lblTitle.FontSize = 18
    lblTitle.FontBold = True
    lblTitle.ForeColor = RGB(0, 70, 140)
    lblTitle.TextAlign = 2  ' Center

    ' Navigation buttons (row of 4)
    Dim btnLeft As Long
    btnLeft = 300

    AddDashBtn frm, "btnComponents", "Components", btnLeft, 1100, _
        "DoCmd.OpenForm ""frmComponents"""
    AddDashBtn frm, "btnProjects", "Projects", btnLeft + 1800, 1100, _
        "DoCmd.OpenForm ""frmProjects"""
    AddDashBtn frm, "btnImport", "Import BOM", btnLeft + 3600, 1100, _
        "DoCmd.OpenForm ""frmBOMImport"""
    AddDashBtn frm, "btnManufacturers", "Manufacturers", btnLeft + 5400, 1100, _
        "DoCmd.OpenForm ""frmManufacturers"""

    ' Stats section header
    Dim lblStats As Control
    Set lblStats = CreateControl(frm.Name, acLabel, acDetail, "", _
        "lblStatsTitle", 200, 2400, 7600, 380)
    lblStats.Caption = "QUICK STATS"
    lblStats.FontBold = True
    lblStats.ForeColor = RGB(80, 80, 80)

    ' Quick stat labels — values set via ControlSource expressions
    AddStatLabel frm, "lblStatComps", "Total Components:", _
        "=DCount(""*"",""Components"")", 300, 2900
    AddStatLabel frm, "lblStatProjs", "Total Projects:", _
        "=DCount(""*"",""Projects"")", 300, 3280
    AddStatLabel frm, "lblStatObs", "Obsolete w/o Alternative:", _
        "=DCount(""*"",""qry_ComponentsNoAlternatives"",""Lifecycle IN ('Obsolete','NRND')"")", _
        300, 3660
    AddStatLabel frm, "lblStatImports", "Total BOM Imports:", _
        "=DCount(""*"",""BOM_Headers"",""ImportStatus='Committed'"")", 300, 4040

    ' Reports section
    Dim lblRep As Control
    Set lblRep = CreateControl(frm.Name, acLabel, acDetail, "", _
        "lblReportsTitle", 200, 4600, 7600, 380)
    lblRep.Caption = "REPORTS"
    lblRep.FontBold = True
    lblRep.ForeColor = RGB(80, 80, 80)

    AddCommandButton frm, "btnRptRisk", "Obsolescence Risk Report", 300, 5050, _
        "DoCmd.OpenQuery ""qry_ObsolescenceRisk"""
    AddCommandButton frm, "btnRptNoAlt", "Components Without Alternatives", 3000, 5050, _
        "DoCmd.OpenQuery ""qry_ComponentsNoAlternatives"""
    AddCommandButton frm, "btnRptBOM", "Full BOM Extract", 5700, 5050, _
        "DoCmd.OpenQuery ""qry_BOMExtract"""

    SaveAndCloseForm frm, "frmDashboard"
End Sub

' ============================================================
'  SHARED HELPERS FOR FORM BUILDING
' ============================================================

Private Sub AddLabelField(frm As Form, labelText As String, fieldName As String, _
                           topPos As Long, fieldHeight As Long)
    Dim lbl As Control
    Set lbl = CreateControl(frm.Name, acLabel, acDetail, "", _
        "lbl_" & fieldName, LEFT_L, topPos, LABEL_W, ROW_H)
    lbl.Caption = labelText
    lbl.TextAlign = 3  ' Right align

    Dim txt As Control
    Set txt = CreateControl(frm.Name, acTextBox, acDetail, "", _
        fieldName, LEFT_F, topPos, FIELD_W, fieldHeight)
    txt.ControlSource = fieldName
    If fieldName = "Notes" Then txt.ScrollBars = 2
End Sub

Private Sub AddBoundCombo(frm As Form, fieldName As String, labelText As String, _
                            rowSrc As String, topPos As Long, fieldHeight As Long)
    Dim lbl As Control
    Set lbl = CreateControl(frm.Name, acLabel, acDetail, "", _
        "lbl_" & fieldName, LEFT_L, topPos, LABEL_W, ROW_H)
    lbl.Caption = labelText
    lbl.TextAlign = 3

    Dim cbo As Control
    Set cbo = CreateControl(frm.Name, acComboBox, acDetail, "", _
        fieldName, LEFT_F, topPos, FIELD_W, fieldHeight)
    cbo.ControlSource = fieldName
    cbo.RowSourceType = "Table/Query"
    cbo.RowSource = rowSrc
    cbo.ColumnCount = 2
    cbo.ColumnWidths = "0;3000"
    cbo.BoundColumn = 1
    cbo.LimitToList = True
End Sub

Private Sub AddPlainLabel(frm As Form, pageName As String, ctlName As String, _
                           caption As String, left As Long, top As Long, _
                           width As Long, height As Long)
    Dim lbl As Control
    Set lbl = CreateControl(frm.Name, acLabel, acDetail, pageName, _
        ctlName, left, top, width, height)
    lbl.Caption = caption
    lbl.TextAlign = 3
End Sub

Private Sub AddCommandButton(frm As Form, btnName As String, caption As String, _
                              leftPos As Long, topPos As Long, onClickCode As String)
    Dim btn As Control
    Set btn = CreateControl(frm.Name, acCommandButton, acDetail, "", _
        btnName, leftPos, topPos, 1700, 400)
    btn.Caption = caption
End Sub

Private Sub AddDashBtn(frm As Form, btnName As String, caption As String, _
                        leftPos As Long, topPos As Long, onClickCode As String)
    Dim btn As Control
    Set btn = CreateControl(frm.Name, acCommandButton, acDetail, "", _
        btnName, leftPos, topPos, 1600, 700)
    btn.Caption = caption
    btn.FontSize = 12
    btn.FontBold = True
End Sub

Private Sub AddStatLabel(frm As Form, ctlName As String, labelText As String, _
                          expression As String, leftPos As Long, topPos As Long)
    Dim lbl As Control
    Set lbl = CreateControl(frm.Name, acLabel, acDetail, "", _
        "lbl_" & ctlName & "_t", leftPos, topPos, 2600, ROW_H)
    lbl.Caption = labelText
    lbl.FontBold = True

    Dim txt As Control
    Set txt = CreateControl(frm.Name, acTextBox, acDetail, "", _
        ctlName, leftPos + 2700, topPos, 1200, ROW_H)
    txt.ControlSource = expression
    txt.Locked = True
    txt.BorderStyle = 0
    txt.BackStyle = 0
    txt.FontBold = True
    txt.ForeColor = RGB(0, 100, 0)
End Sub

Private Sub SaveAndCloseForm(frm As Form, formName As String)
    DoCmd.Save acForm, frm.Name
    DoCmd.Close acForm, frm.Name, acSaveYes
    DoCmd.Rename formName, acForm, frm.Name
End Sub

Private Sub DeleteFormIfExists(formName As String)
    On Error Resume Next
    DoCmd.DeleteObject acForm, formName
    On Error GoTo 0
End Sub
