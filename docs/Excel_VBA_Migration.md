# Excel VBA Migration for Machine History & Maintenance

This folder contains a complete VBA port of the Google Apps Script Machine History & Maintenance workflow. It intentionally does **not** recreate the HTML/CSS UI. Instead, worksheets act as database tables and dashboards, while UserForms provide data-entry workflows.

## Delivered VBA files

### Standard modules

Import these files into the VBA editor as standard modules, in this order:

1. `vba/modules/modConstants.bas`
2. `vba/modules/modUtils.bas`
3. `vba/modules/modTables.bas`
4. `vba/modules/modSetup.bas`
5. `vba/modules/modValidation.bas`
6. `vba/modules/modAudit.bas`
7. `vba/modules/modSchedule.bas`
8. `vba/modules/modDashboard.bas`
9. `vba/modules/modProfile.bas`
10. `vba/modules/modMaintenance.bas`
11. `vba/modules/modDefect.bas`

### UserForm code-behind files

Create the UserForms manually in the VBA editor, add the listed controls, then paste the corresponding code-behind file into each form.

1. `frmMaintenanceLog` -> `vba/userforms/frmMaintenanceLog.code.txt`
2. `frmDefectReport` -> `vba/userforms/frmDefectReport.code.txt`
3. `frmMachinePicker` -> `vba/userforms/frmMachinePicker.code.txt`

## Workbook setup instructions

1. Open Excel and create a new macro-enabled workbook (`.xlsm`).
2. Press `ALT + F11` to open the Visual Basic Editor.
3. Import each `.bas` file from `vba/modules` as a standard module.
4. Create the three UserForms and controls listed below.
5. Paste the matching `vba/userforms/*.code.txt` code into each UserForm code window.
6. Run `SetupWorkbook` from `modSetup`.
7. Load or paste machine and spare part master data into the generated Excel tables.
8. Run `RefreshDashboard` whenever you want to recalculate schedule statuses and redraw the dashboard.

## Required sheets and tables

Running `SetupWorkbook` creates these sheets and Excel Tables/ListObjects:

| Sheet | Table | Purpose |
|---|---|---|
| `Dashboard` | none | Search/filter dashboard, KPIs, machine-card grid, and buttons |
| `Machine_Profile` | none | Rendered machine profile report |
| `Machine_List` | `tblMachineList` | Machine master data |
| `Spare_Parts_Master` | `tblSparePartsMaster` | Spare part lifetime master data |
| `Spare_Parts_Schedule` | `tblSparePartsSchedule` | Machine + part due schedule |
| `Maintenance_Log` | `tblMaintenanceLog` | Maintenance, repair, PM, and replacement log |
| `Defect_Log` | `tblDefectLog` | Pending/resolved defect lifecycle |
| `Audit_Log` | `tblAuditLog` | Key action audit trail |
| `Settings` | none | Allowed-value reference lists |

## UserForm controls

### `frmMaintenanceLog`

Add these controls with exactly these names:

- ComboBox: `cboMachine`
- TextBox: `txtMaintenanceDate`
- ComboBox: `cboActionType`
- ComboBox: `cboPendingDefect`
- ComboBox: `cboPart`
- TextBox: `txtDetails`
- TextBox: `txtRootCause`
- TextBox: `txtCorrectiveAction`
- TextBox: `txtOperator`
- CommandButton: `cmdSave`
- CommandButton: `cmdCancel`

Recommended captions/labels: Machine, Maintenance Date, Action Type, Pending Defect, Part, Details, Root Cause, Corrective Action, Operator, Save, Cancel.

### `frmDefectReport`

Add these controls with exactly these names:

- ComboBox: `cboMachine`
- TextBox: `txtDateFound`
- TextBox: `txtSymptom`
- ComboBox: `cboSeverity`
- TextBox: `txtReportedBy`
- CommandButton: `cmdSubmit`
- CommandButton: `cmdCancel`

Recommended captions/labels: Machine, Date Found, Symptom, Severity, Reported By, Submit, Cancel.

### `frmMachinePicker`

Add these controls with exactly these names:

- ComboBox: `cboMachine`
- CommandButton: `cmdOpen`
- CommandButton: `cmdCancel`

Recommended captions/labels: Machine, Open, Cancel.

## Dashboard controls

`SetupWorkbook` creates these shape buttons on the `Dashboard` sheet:

- Refresh Dashboard -> `RefreshDashboard`
- New Maintenance Log -> `ShowMaintenanceLogForm`
- Report Defect -> `ShowDefectReportForm`
- Open Machine Profile -> `ShowMachinePickerForProfile`
- Print Report -> `PrintCurrentMachineProfile`

Use cell `B3` for free-text search across Machine ID, Machine Name, Scope, and Manufacturer. Use cell `E3` for one of the filter names defined in `modConstants`.

## Business logic coverage

The VBA implementation includes:

- Required field validation for machine logs, dates, action types, operators, defect reports, severity, and reporter.
- Allowed values for action type, severity, machine status, and defect status.
- Sequential IDs: `DF-0001`, `MT-0001`, and `AU-0001`.
- Repair entries that can link to a pending defect and resolve it with root cause, corrective action, resolver, resolved timestamp, and linked maintenance ID.
- Part replacement entries that require a valid part and update `tblSparePartsSchedule` by `Machine_ID + Part_ID`.
- Schedule status calculation: `OVERDUE`, `DUE_SOON`, or `OK`.
- Dashboard counts for total, active, inactive, pending defects, critical pending defects, overdue parts, and due-soon parts.
- Machine profile rendering for basic information, active pending defects, spare parts schedule, maintenance history, and full defect history.
- Audit logging for maintenance creation, defect reporting, defect resolution, and spare parts schedule updates.

## Practical notes

- The code uses `ListObject` table APIs rather than hardcoded data ranges for database operations.
- Standard modules and UserForm code all use `Option Explicit`.
- The code avoids `Select` and `Activate`.
- If your Excel security settings block macros, save the file as `.xlsm`, trust the workbook location, and reopen it.
- If you want dropdown validation on worksheet cells, use the values populated on the `Settings` sheet as the source lists.
