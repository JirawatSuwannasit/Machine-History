Attribute VB_Name = "modConstants"
Option Explicit

'Workbook sheets
Public Const SHEET_DASHBOARD As String = "Dashboard"
Public Const SHEET_PROFILE As String = "Machine_Profile"
Public Const SHEET_MACHINE_LIST As String = "Machine_List"
Public Const SHEET_PARTS_MASTER As String = "Spare_Parts_Master"
Public Const SHEET_PARTS_SCHEDULE As String = "Spare_Parts_Schedule"
Public Const SHEET_MAINTENANCE_LOG As String = "Maintenance_Log"
Public Const SHEET_DEFECT_LOG As String = "Defect_Log"
Public Const SHEET_AUDIT_LOG As String = "Audit_Log"
Public Const SHEET_SETTINGS As String = "Settings"

'Excel table names
Public Const TBL_MACHINE_LIST As String = "tblMachineList"
Public Const TBL_PARTS_MASTER As String = "tblSparePartsMaster"
Public Const TBL_PARTS_SCHEDULE As String = "tblSparePartsSchedule"
Public Const TBL_MAINTENANCE_LOG As String = "tblMaintenanceLog"
Public Const TBL_DEFECT_LOG As String = "tblDefectLog"
Public Const TBL_AUDIT_LOG As String = "tblAuditLog"

'Allowed values
Public Const ACTION_REPAIR As String = "Repair"
Public Const ACTION_PART_REPLACEMENT As String = "Part Replacement"
Public Const ACTION_PM As String = "PM"

Public Const SEVERITY_LOW As String = "Low"
Public Const SEVERITY_MEDIUM As String = "Medium"
Public Const SEVERITY_HIGH As String = "High"
Public Const SEVERITY_CRITICAL As String = "Critical"

Public Const MACHINE_STATUS_ACTIVE As String = "Active"
Public Const MACHINE_STATUS_INACTIVE As String = "Inactive"

Public Const DEFECT_STATUS_PENDING As String = "Pending"
Public Const DEFECT_STATUS_RESOLVED As String = "Resolved"

Public Const SCHEDULE_STATUS_OVERDUE As String = "OVERDUE"
Public Const SCHEDULE_STATUS_DUE_SOON As String = "DUE_SOON"
Public Const SCHEDULE_STATUS_OK As String = "OK"
Public Const DUE_SOON_DAYS As Long = 30

'Dashboard filters
Public Const FILTER_ALL As String = "All"
Public Const FILTER_ACTIVE As String = "Active"
Public Const FILTER_INACTIVE As String = "Inactive"
Public Const FILTER_PENDING_DEFECTS As String = "Has Pending Defects"
Public Const FILTER_CRITICAL_DEFECTS As String = "Critical Defects"
Public Const FILTER_OVERDUE_PARTS As String = "Overdue Parts"
Public Const FILTER_DUE_SOON_PARTS As String = "Due Soon Parts"

'ID prefixes
Public Const PREFIX_DEFECT As String = "DF-"
Public Const PREFIX_MAINTENANCE As String = "MT-"
Public Const PREFIX_AUDIT As String = "AU-"
Public Const ID_DIGITS As Long = 4
