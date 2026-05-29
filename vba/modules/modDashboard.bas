Attribute VB_Name = "modDashboard"
Option Explicit

Public Sub RefreshDashboard()
    Dim ws As Worksheet
    Dim searchText As String
    Dim filterName As String
    Dim stats As Object

    Application.ScreenUpdating = False
    RefreshScheduleStatuses
    Set ws = GetSheet(SHEET_DASHBOARD)
    searchText = TrimText(ws.Range("B3").Value)
    filterName = TrimText(ws.Range("E3").Value)
    If Len(filterName) = 0 Then filterName = FILTER_ALL

    Set stats = DashboardStats
    ws.Range("A6:G6").Value = Array(stats("Total"), stats("Active"), stats("Inactive"), stats("Pending"), stats("Critical"), stats("Overdue"), stats("DueSoon"))
    RenderMachineCards ws, searchText, filterName
    Application.ScreenUpdating = True
End Sub

Public Function DashboardStats() As Object
    Dim dict As Object
    Dim lo As ListObject
    Dim lr As ListRow
    Dim machineId As String
    Set dict = CreateObject("Scripting.Dictionary")
    dict("Total") = 0: dict("Active") = 0: dict("Inactive") = 0: dict("Pending") = 0: dict("Critical") = 0: dict("Overdue") = 0: dict("DueSoon") = 0

    Set lo = GetTable(TBL_MACHINE_LIST)
    If Not lo.DataBodyRange Is Nothing Then
        For Each lr In lo.ListRows
            machineId = TrimText(GetCellValue(lr, "Machine_ID"))
            dict("Total") = dict("Total") + 1
            If StrComp(TrimText(GetCellValue(lr, "Status")), MACHINE_STATUS_ACTIVE, vbTextCompare) = 0 Then dict("Active") = dict("Active") + 1
            If StrComp(TrimText(GetCellValue(lr, "Status")), MACHINE_STATUS_INACTIVE, vbTextCompare) = 0 Then dict("Inactive") = dict("Inactive") + 1
            If CountPendingDefects(machineId) > 0 Then dict("Pending") = dict("Pending") + 1
            If CountCriticalPendingDefects(machineId) > 0 Then dict("Critical") = dict("Critical") + 1
            If CountMachineSchedules(machineId, SCHEDULE_STATUS_OVERDUE) > 0 Then dict("Overdue") = dict("Overdue") + 1
            If CountMachineSchedules(machineId, SCHEDULE_STATUS_DUE_SOON) > 0 Then dict("DueSoon") = dict("DueSoon") + 1
        Next lr
    End If
    Set DashboardStats = dict
End Function

Private Sub RenderMachineCards(ByVal ws As Worksheet, ByVal searchText As String, ByVal filterName As String)
    Dim lo As ListObject
    Dim lr As ListRow
    Dim outputRow As Long
    Dim machineId As String

    ClearRangeContentsAndFormats ws.Range("A9:K1000")
    ws.Range("A9:K9").Value = Array("Machine_ID", "Machine_Name", "SCOPE", "Manufacturer", "Status", "Pending Defects", "Critical Defects", "Overdue Parts", "Due Soon Parts", "Model", "SN")
    ws.Range("A9:K9").Font.Bold = True
    outputRow = 10
    Set lo = GetTable(TBL_MACHINE_LIST)
    If lo.DataBodyRange Is Nothing Then Exit Sub

    For Each lr In lo.ListRows
        machineId = TrimText(GetCellValue(lr, "Machine_ID"))
        If MachineMatchesSearch(lr, searchText) And MachineMatchesFilter(lr, filterName) Then
            ws.Cells(outputRow, 1).Value = machineId
            ws.Cells(outputRow, 2).Value = GetCellValue(lr, "Machine_Name")
            ws.Cells(outputRow, 3).Value = GetCellValue(lr, "SCOPE")
            ws.Cells(outputRow, 4).Value = GetCellValue(lr, "Manufacturer")
            ws.Cells(outputRow, 5).Value = GetCellValue(lr, "Status")
            ws.Cells(outputRow, 6).Value = CountPendingDefects(machineId)
            ws.Cells(outputRow, 7).Value = CountCriticalPendingDefects(machineId)
            ws.Cells(outputRow, 8).Value = CountMachineSchedules(machineId, SCHEDULE_STATUS_OVERDUE)
            ws.Cells(outputRow, 9).Value = CountMachineSchedules(machineId, SCHEDULE_STATUS_DUE_SOON)
            ws.Cells(outputRow, 10).Value = GetCellValue(lr, "Model")
            ws.Cells(outputRow, 11).Value = GetCellValue(lr, "SN")
            ApplyAlertFormats ws.Range(ws.Cells(outputRow, 6), ws.Cells(outputRow, 9))
            outputRow = outputRow + 1
        End If
    Next lr
    ws.Columns("A:K").AutoFit
End Sub

Private Function MachineMatchesSearch(ByVal lr As ListRow, ByVal searchText As String) As Boolean
    Dim haystack As String
    If Len(searchText) = 0 Then MachineMatchesSearch = True: Exit Function
    haystack = Join(Array(GetCellValue(lr, "Machine_ID"), GetCellValue(lr, "Machine_Name"), GetCellValue(lr, "SCOPE"), GetCellValue(lr, "Manufacturer")), " ")
    MachineMatchesSearch = (InStr(1, haystack, searchText, vbTextCompare) > 0)
End Function

Private Function MachineMatchesFilter(ByVal lr As ListRow, ByVal filterName As String) As Boolean
    Dim machineId As String
    machineId = TrimText(GetCellValue(lr, "Machine_ID"))
    Select Case filterName
        Case FILTER_ALL
            MachineMatchesFilter = True
        Case FILTER_ACTIVE, FILTER_INACTIVE
            MachineMatchesFilter = (StrComp(TrimText(GetCellValue(lr, "Status")), filterName, vbTextCompare) = 0)
        Case FILTER_PENDING_DEFECTS
            MachineMatchesFilter = (CountPendingDefects(machineId) > 0)
        Case FILTER_CRITICAL_DEFECTS
            MachineMatchesFilter = (CountCriticalPendingDefects(machineId) > 0)
        Case FILTER_OVERDUE_PARTS
            MachineMatchesFilter = (CountMachineSchedules(machineId, SCHEDULE_STATUS_OVERDUE) > 0)
        Case FILTER_DUE_SOON_PARTS
            MachineMatchesFilter = (CountMachineSchedules(machineId, SCHEDULE_STATUS_DUE_SOON) > 0)
        Case Else
            MachineMatchesFilter = True
    End Select
End Function

Private Sub ApplyAlertFormats(ByVal target As Range)
    Dim cell As Range
    For Each cell In target.Cells
        If Val(cell.Value) > 0 Then
            cell.Font.Bold = True
            cell.Interior.Color = RGB(255, 199, 206)
        End If
    Next cell
End Sub

Public Sub ShowMaintenanceLogForm()
    frmMaintenanceLog.Show
End Sub

Public Sub ShowDefectReportForm()
    frmDefectReport.Show
End Sub

Public Sub ShowMachinePickerForProfile()
    frmMachinePicker.Show
End Sub
