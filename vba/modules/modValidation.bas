Attribute VB_Name = "modValidation"
Option Explicit

Public Function ValidateMaintenanceInput(ByVal machineId As String, ByVal maintenanceDateText As String, ByVal actionType As String, ByVal operatorName As String, ByVal partId As String, ByVal defectId As String) As String
    Dim message As String
    machineId = TrimText(machineId)
    actionType = TrimText(actionType)
    operatorName = TrimText(operatorName)
    partId = TrimText(partId)
    defectId = TrimText(defectId)

    If Len(machineId) = 0 Then message = message & "Machine_ID is required." & vbCrLf
    If Len(machineId) > 0 And Not MachineExists(machineId) Then message = message & "Machine_ID does not exist." & vbCrLf
    If Len(maintenanceDateText) = 0 Or Not IsDate(maintenanceDateText) Then message = message & "Maintenance_Date is required and must be a valid date." & vbCrLf
    If Len(actionType) = 0 Then message = message & "Action_Type is required." & vbCrLf
    If Len(actionType) > 0 And Not IsInArray(actionType, Array(ACTION_REPAIR, ACTION_PART_REPLACEMENT, ACTION_PM)) Then message = message & "Action_Type is invalid." & vbCrLf
    If Len(operatorName) = 0 Then message = message & "Operator is required." & vbCrLf
    If StrComp(actionType, ACTION_PART_REPLACEMENT, vbTextCompare) = 0 And Len(partId) = 0 Then message = message & "Part_ID is required for Part Replacement." & vbCrLf
    If Len(partId) > 0 And Not PartExists(partId) Then message = message & "Part_ID does not exist in tblSparePartsMaster." & vbCrLf

    If StrComp(actionType, ACTION_REPAIR, vbTextCompare) = 0 And Len(defectId) > 0 Then
        message = message & ValidateResolvableDefect(machineId, defectId)
    End If

    ValidateMaintenanceInput = message
End Function

Public Function ValidateResolvableDefect(ByVal machineId As String, ByVal defectId As String) As String
    Dim lr As ListRow
    Set lr = FindRowByValue(TBL_DEFECT_LOG, "Defect_ID", defectId)
    If lr Is Nothing Then
        ValidateResolvableDefect = "Defect_ID does not exist." & vbCrLf
    ElseIf NormalizeKey(GetCellValue(lr, "Machine_ID")) <> NormalizeKey(machineId) Then
        ValidateResolvableDefect = "Defect_ID does not belong to the selected Machine_ID." & vbCrLf
    ElseIf StrComp(TrimText(GetCellValue(lr, "Status")), DEFECT_STATUS_PENDING, vbTextCompare) <> 0 Then
        ValidateResolvableDefect = "Defect_ID is not Pending." & vbCrLf
    End If
End Function

Public Function ValidateDefectInput(ByVal machineId As String, ByVal dateFoundText As String, ByVal symptom As String, ByVal severity As String, ByVal reportedBy As String) As String
    Dim message As String
    machineId = TrimText(machineId)
    severity = TrimText(severity)

    If Len(machineId) = 0 Then message = message & "Machine_ID is required." & vbCrLf
    If Len(machineId) > 0 And Not MachineExists(machineId) Then message = message & "Machine_ID does not exist." & vbCrLf
    If Len(TrimText(dateFoundText)) = 0 Or Not IsDate(dateFoundText) Then message = message & "Date_Found is required and must be a valid date." & vbCrLf
    If Len(TrimText(symptom)) = 0 Then message = message & "Symptom is required." & vbCrLf
    If Len(severity) = 0 Then message = message & "Severity is required." & vbCrLf
    If Len(severity) > 0 And Not IsInArray(severity, Array(SEVERITY_LOW, SEVERITY_MEDIUM, SEVERITY_HIGH, SEVERITY_CRITICAL)) Then message = message & "Severity is invalid." & vbCrLf
    If Len(TrimText(reportedBy)) = 0 Then message = message & "Reported_By is required." & vbCrLf

    ValidateDefectInput = message
End Function
