Attribute VB_Name = "modDefect"
Option Explicit

Public Function CreateDefectReport(ByVal machineId As String, ByVal dateFoundText As String, ByVal symptom As String, ByVal severity As String, ByVal reportedBy As String) As String
    Dim validationMessage As String
    Dim defectId As String
    Dim lr As ListRow

    machineId = TrimText(machineId)
    severity = TrimText(severity)
    validationMessage = ValidateDefectInput(machineId, dateFoundText, symptom, severity, reportedBy)
    If Len(validationMessage) > 0 Then Err.Raise vbObjectError + 500, , validationMessage

    defectId = NextSequentialID(TBL_DEFECT_LOG, "Defect_ID", PREFIX_DEFECT)
    Set lr = AddTableRow(TBL_DEFECT_LOG)
    SetCellValue lr, "Defect_ID", defectId
    SetCellValue lr, "Timestamp", NowStamp
    SetCellValue lr, "Machine_ID", machineId
    SetCellValue lr, "Date_Found", ParseDateRequired(dateFoundText, "Date_Found")
    SetCellValue lr, "Symptom", TrimText(symptom)
    SetCellValue lr, "Severity", severity
    SetCellValue lr, "Reported_By", TrimText(reportedBy)
    SetCellValue lr, "Status", DEFECT_STATUS_PENDING
    SetCellValue lr, "Root_Cause", ""
    SetCellValue lr, "Corrective_Action", ""
    SetCellValue lr, "Resolved_By", ""
    SetCellValue lr, "Resolved_At", ""
    SetCellValue lr, "Linked_Maintenance_ID", ""

    WriteAuditLog "Defect reported", "Defect_Log", defectId, "", "Machine_ID=" & machineId & "; Severity=" & severity
    RefreshDashboard
    If NormalizeKey(GetSheet(SHEET_PROFILE).Range("B3").Value) = NormalizeKey(machineId) Then RenderMachineProfile machineId
    CreateDefectReport = defectId
End Function

Public Sub ResolveDefect(ByVal defectId As String, ByVal maintenanceId As String, ByVal resolvedBy As String, ByVal rootCause As String, ByVal correctiveAction As String)
    Dim lr As ListRow
    Dim beforeText As String
    Dim afterText As String
    Set lr = FindRowByValue(TBL_DEFECT_LOG, "Defect_ID", defectId)
    If lr Is Nothing Then Err.Raise vbObjectError + 501, , "Defect_ID not found: " & defectId
    If StrComp(TrimText(GetCellValue(lr, "Status")), DEFECT_STATUS_PENDING, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 502, , "Defect is not pending: " & defectId

    beforeText = "Status=" & TrimText(GetCellValue(lr, "Status"))
    SetCellValue lr, "Status", DEFECT_STATUS_RESOLVED
    SetCellValue lr, "Root_Cause", TrimText(rootCause)
    SetCellValue lr, "Corrective_Action", TrimText(correctiveAction)
    SetCellValue lr, "Resolved_By", TrimText(resolvedBy)
    SetCellValue lr, "Resolved_At", NowStamp
    SetCellValue lr, "Linked_Maintenance_ID", maintenanceId
    afterText = "Status=" & DEFECT_STATUS_RESOLVED & "; Linked_Maintenance_ID=" & maintenanceId
    WriteAuditLog "Defect resolved", "Defect_Log", defectId, beforeText, afterText
End Sub

Public Function CountPendingDefects(ByVal machineId As String) As Long
    CountPendingDefects = CountDefects(machineId, DEFECT_STATUS_PENDING, vbNullString)
End Function

Public Function CountCriticalPendingDefects(ByVal machineId As String) As Long
    CountCriticalPendingDefects = CountDefects(machineId, DEFECT_STATUS_PENDING, SEVERITY_CRITICAL)
End Function

Private Function CountDefects(ByVal machineId As String, ByVal statusFilter As String, ByVal severityFilter As String) As Long
    Dim lo As ListObject
    Dim lr As ListRow
    Set lo = GetTable(TBL_DEFECT_LOG)
    If lo.DataBodyRange Is Nothing Then Exit Function
    For Each lr In lo.ListRows
        If NormalizeKey(GetCellValue(lr, "Machine_ID")) = NormalizeKey(machineId) _
            And (Len(statusFilter) = 0 Or StrComp(TrimText(GetCellValue(lr, "Status")), statusFilter, vbTextCompare) = 0) _
            And (Len(severityFilter) = 0 Or StrComp(TrimText(GetCellValue(lr, "Severity")), severityFilter, vbTextCompare) = 0) Then
            CountDefects = CountDefects + 1
        End If
    Next lr
End Function
