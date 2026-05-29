Attribute VB_Name = "modSchedule"
Option Explicit

Public Function CalculateScheduleStatus(ByVal nextDue As Date) As String
    If nextDue < TodayDate Then
        CalculateScheduleStatus = SCHEDULE_STATUS_OVERDUE
    ElseIf nextDue <= DateAdd("d", DUE_SOON_DAYS, TodayDate) Then
        CalculateScheduleStatus = SCHEDULE_STATUS_DUE_SOON
    Else
        CalculateScheduleStatus = SCHEDULE_STATUS_OK
    End If
End Function

Public Sub RefreshScheduleStatuses()
    Dim lo As ListObject
    Dim lr As ListRow
    Dim nextDue As Variant
    Set lo = GetTable(TBL_PARTS_SCHEDULE)
    If lo.DataBodyRange Is Nothing Then Exit Sub

    For Each lr In lo.ListRows
        nextDue = GetCellValue(lr, "Next_Due")
        If IsDate(nextDue) Then SetCellValue lr, "Status", CalculateScheduleStatus(CDate(nextDue))
    Next lr
End Sub

Public Sub UpsertSparePartsSchedule(ByVal machineId As String, ByVal partId As String, ByVal maintenanceDate As Date)
    Dim lr As ListRow
    Dim lifetimeYears As Double
    Dim nextDue As Date
    Dim beforeText As String
    Dim afterText As String

    If Len(TrimText(partId)) = 0 Then Exit Sub
    lifetimeYears = GetPartLifetimeYears(partId)
    nextDue = AddYearsFlexible(maintenanceDate, lifetimeYears)

    Set lr = FindScheduleRow(machineId, partId)
    If lr Is Nothing Then
        Set lr = AddTableRow(TBL_PARTS_SCHEDULE)
        SetCellValue lr, "Machine_ID", machineId
        SetCellValue lr, "Part_ID", partId
        beforeText = "<new>"
    Else
        beforeText = "Last_Changed=" & SafeDateText(GetCellValue(lr, "Last_Changed")) & "; Next_Due=" & SafeDateText(GetCellValue(lr, "Next_Due")) & "; Status=" & TrimText(GetCellValue(lr, "Status"))
    End If

    SetCellValue lr, "Last_Changed", maintenanceDate
    SetCellValue lr, "Next_Due", nextDue
    SetCellValue lr, "Status", CalculateScheduleStatus(nextDue)
    afterText = "Last_Changed=" & Format$(maintenanceDate, "yyyy-mm-dd") & "; Next_Due=" & Format$(nextDue, "yyyy-mm-dd") & "; Status=" & CalculateScheduleStatus(nextDue)
    WriteAuditLog "Spare parts schedule updated", "SparePartsSchedule", machineId & ":" & partId, beforeText, afterText
End Sub

Public Function CountMachineSchedules(ByVal machineId As String, ByVal wantedStatus As String) As Long
    Dim lo As ListObject
    Dim lr As ListRow
    Set lo = GetTable(TBL_PARTS_SCHEDULE)
    If lo.DataBodyRange Is Nothing Then Exit Function

    For Each lr In lo.ListRows
        If NormalizeKey(GetCellValue(lr, "Machine_ID")) = NormalizeKey(machineId) Then
            If StrComp(TrimText(GetCellValue(lr, "Status")), wantedStatus, vbTextCompare) = 0 Then CountMachineSchedules = CountMachineSchedules + 1
        End If
    Next lr
End Function
