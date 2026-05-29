# Machine History & Maintenance

This repository contains the original Google Apps Script web app assets and an Excel VBA migration package for a Machine History & Maintenance system.

## Excel VBA migration

The VBA port is in:

- `vba/modules/` - standard VBA modules (`modConstants`, `modUtils`, `modTables`, setup, validation, dashboard, profile, maintenance, defect, schedule, audit)
- `vba/userforms/` - code-behind for `frmMaintenanceLog`, `frmDefectReport`, and `frmMachinePicker`
- `docs/Excel_VBA_Migration.md` - setup instructions, required sheets/tables, UserForm controls, and workflow notes

Start with `docs/Excel_VBA_Migration.md`, import the modules into a macro-enabled workbook, create the UserForms, and run `SetupWorkbook`.
