# Component Management System — Installation Guide

Follow these steps in order. Each step takes 2–5 minutes.
You only need to do this once.

---

## WHAT YOU NEED BEFORE STARTING

- Microsoft Access (already installed on your PC)
- Microsoft Excel (already installed on your PC)
- This folder with the VBA files

---

## PHASE 1 — Create the Database File

**Step 1.** Open Microsoft Access.

**Step 2.** Click **Blank database** (the big button on the welcome screen).

**Step 3.** Choose where to save it.
- Click the yellow folder icon next to the filename box.
- Navigate to a folder you'll remember (e.g. `C:\ComponentDB\`).
- Type the filename: `ComponentDB`
- Click **OK**, then click **Create**.

Access opens with an empty table called "Table1". That's normal — you'll delete it later.

---

## PHASE 2 — Import the VBA Code

**Step 4.** Press `Alt + F11` on your keyboard.
> This opens the VBA code editor (a separate grey window).

**Step 5.** In the VBA editor menu at the top, click **File → Import File...**

**Step 6.** Navigate to the folder where you saved these files, go into the `vba` folder.
Select `Step1_SetupTables.bas` and click **Open**.

**Step 7.** Repeat Step 6 for each of these files (import them one at a time):
- `Step2_SetupQueries.bas`
- `Step3_CreateForms.bas`
- `modExcelImport.bas`
- `modFormLogic.bas`

After importing all 5, you should see them listed on the left side of the VBA editor
under "Modules".

---

## PHASE 3 — Run the Setup Scripts

**Step 8.** In the VBA editor, on the left panel, double-click **Step1_SetupTables**.
The code appears on the right side.

**Step 9.** Click anywhere inside the text that says `Public Sub SetupAllTables()`.
Your cursor just needs to be somewhere in that section.

**Step 10.** Press `F5` on your keyboard (or click **Run → Run Sub/UserForm** in the menu).
> A message box will appear saying "Starting database setup..."
> Wait a few seconds. Then a second message says "SUCCESS! All tables created..."
> Click **OK**.

**Step 11.** Double-click **Step2_SetupQueries** in the left panel.
Click anywhere inside `Public Sub SetupAllQueries()`.
Press `F5`.
> Click OK when done.

**Step 12.** Double-click **Step3_CreateForms** in the left panel.
Click anywhere inside `Public Sub SetupAllForms()`.
Press `F5`.
> This takes 10–20 seconds. Click OK when done.

---

## PHASE 4 — Final Cleanup

**Step 13.** Press `Alt + F11` again to go back to Access (or click the Access icon on the taskbar).

**Step 14.** You'll see "Table1" that Access created automatically at the start.
Right-click on **Table1** in the left panel → click **Delete** → click **Yes**.

**Step 15.** Close and reopen the database file to refresh everything:
- Click **File → Close**
- Open the file again (File → Open Recent → ComponentDB)

---

## PHASE 5 — Start Using It

**Step 16.** In the left panel (called the "Navigation Pane"), you should now see:
- Under **Tables**: 11 tables (Components, Manufacturers, Projects, etc.)
- Under **Queries**: several queries starting with `qry_`
- Under **Forms**: 5 forms (frmDashboard, frmComponents, frmProjects, etc.)

**Step 17.** Double-click **frmDashboard** to open the main screen.
> You're ready to start.

---

## FIRST THINGS TO DO

### Add your first manufacturer:
1. Click **Manufacturers** on the dashboard.
2. Click in the Manufacturer Name field, type the name (e.g. `STMicroelectronics`).
3. Click **Save & Close**.

### Add your first component:
1. Click **Components** on the dashboard.
2. Click **New Component**.
3. Fill in: MPN, Manufacturer (select from dropdown), Description, Type, Package.
4. Select Lifecycle Status (e.g. `Active`) and Qual Status (e.g. `Approved`).
5. Click **Save**.

### Add a project:
1. Click **Projects** on the dashboard.
2. Click **New Project**.
3. Fill in: Project Code (e.g. `PCB-001`), Project Name, Owner.
4. Click **Save**.

### Import a BOM from Excel:
1. Click **Import BOM** on the dashboard.
2. **Step 1**: Select your project, enter BOM revision, click **Browse** to pick your Excel file. Click **Next**.
3. **Step 2**: Enter which column number contains MPN, Manufacturer, etc. (1 = column A, 2 = column B...). Click **Parse File**.
4. **Step 3**: Review results. Green = matched to existing components. Yellow = new (will be auto-created).
5. **Step 4**: Click **COMMIT IMPORT**.

---

## IF SOMETHING GOES WRONG

| Problem | Fix |
|---|---|
| "Error: table already exists" during setup | This is safe to ignore — the table was already created. Click OK and continue. |
| The VBA editor shows red text | There's a typo in the file. Contact support. |
| "File not found" when importing BOM | Make sure the Excel file path has no special characters. Try moving it to `C:\BOMs\` first. |
| A form opens blank | Close it, go to the left panel, right-click the form → Design View. Check the Record Source property. |
| Numbers show as dates or vice versa | In Excel, make sure the Quantity column is formatted as Number, not Text or Date. |

---

## WIRING FORM EVENTS (optional but recommended)

After setup, to get the duplicate-check and warnings working on the Components form:

1. In the left panel, right-click **frmComponents** → **Design View**.
2. Click the grey area outside the fields (to select the form itself, not a control).
3. Press `F4` to open the Properties panel on the right.
4. Click the **Events** tab.
5. Next to **Before Update**, click the `...` button.
6. Choose **Code Builder** → click OK.
7. Between `Private Sub Form_BeforeUpdate(Cancel As Integer)` and `End Sub`, type:
   ```
   Call modFormLogic.Components_BeforeUpdate(Me, Cancel)
   ```
8. Close the VBA window. In the Properties panel, next to **On Load**, click `...` → Code Builder.
9. Type: `Call modFormLogic.Components_OnLoad(Me)`
10. Save the form: `Ctrl + S`. Close Design View.

Repeat steps 1-10 for **frmProjects**, using `Projects_BeforeUpdate` instead.

---

## FILE SUMMARY

```
vba/
  Step1_SetupTables.bas   → Creates all 11 database tables
  Step2_SetupQueries.bas  → Creates all analysis queries
  Step3_CreateForms.bas   → Builds all forms automatically
  modExcelImport.bas      → Excel BOM import engine
  modFormLogic.bas        → Validation, search, warnings
```
