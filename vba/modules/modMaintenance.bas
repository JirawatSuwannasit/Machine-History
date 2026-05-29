Attribute VB_Name = "modMaintenance"
Option Explicit

Public Function CreateMaintenanceEntry(ByVal machineId As String, ByVal maintenanceDateText As String, ByVal actionType As String, ByVal partId As String, ByVal details As String, ByVal operatorName As String, ByVal defectId As String, ByVal rootCause As String, ByVal correctiveAction As String) As String
    Dim validationMessage As String
    Dim maintenanceId As String
    Dim maintenanceDate As Date
    Dim lr As ListRow

    machineId = TrimText(machineId)
    actionType = TrimText(actionType)
    partId = TrimText(partId)
    operatorName = TrimText(operatorName)
    defectId = TrimText(defectId)
    validationMessage = ValidateMaintenanceInput(machineId, maintenanceDateText, actionType, operatorName, partId, defectId)
    If Len(validationMessage) > 0 Then Err.Raise vbObjectError + 400, , validationMessage

    maintenanceDate = ParseDateRequired(maintenanceDateText, "Maintenance_Date")
    maintenanceId = NextSequentialID(TBL_MAINTENANCE_LOG, "Maintenance_ID", PREFIX_MAINTENANCE)

    Set lr = AddTableRow(TBL_MAINTENANCE_LOG)
    SetCellValue lr, "Maintenance_ID", maintenanceId
    SetCellValue lr, "Timestamp", NowStamp
    SetCellValue lr, "Machine_ID", machineId
    SetCellValue lr, "Action_Type", actionType
    SetCellValue lr, "Part_ID", partId
    SetCellValue lr, "Details", TrimText(details)
    SetCellValue lr, "Operator", operatorName
    SetCellValue lr, "Maintenance_Date", maintenanceDate
    SetCellValue lr, "Defect_ID", defectId
    SetCellValue lr, "Root_Cause", TrimText(rootCause)
    SetCellValue lr, "Corrective_Action", TrimText(correctiveAction)

    WriteAuditLog "Maintenance entry created", "Maintenance_Log", maintenanceId, "", "Machine_ID=" & machineId & "; Action_Type=" & actionType

    If StrComp(actionType, ACTION_REPAIR, vbTextCompare) = 0 And Len(defectId) > 0 Then
        ResolveDefect defectId, maintenanceId, operatorName, rootCause, correctiveAction
    End If

    If Len(partId) > 0 Then UpsertSparePartsSchedule machineId, partId, maintenanceDate
    RefreshDashboard
    If NormalizeKey(GetSheet(SHEET_PROFILE).Range("B3").Value) = NormalizeKey(machineId) Then RenderMachineProfile machineId
    CreateMaintenanceEntry = maintenanceId
End Function
