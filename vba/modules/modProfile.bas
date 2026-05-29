Attribute VB_Name = "modProfile"
Option Explicit

Public Sub RenderMachineProfile(ByVal machineId As String)
    Dim ws As Worksheet
    Dim machineRow As ListRow
    machineId = TrimText(machineId)
    If Len(machineId) = 0 Then Err.Raise vbObjectError + 300, , "Machine_ID is required."

    Set machineRow = FindRowByValue(TBL_MACHINE_LIST, "Machine_ID", machineId)
    If machineRow Is Nothing Then Err.Raise vbObjectError + 301, , "Machine_ID not found: " & machineId

    Application.ScreenUpdating = False
    Set ws = GetSheet(SHEET_PROFILE)
    ClearRangeContentsAndFormats ws.Range("A4:K2000")
    ws.Range("B3").Value = machineId
    RenderMachineInfo ws, machineRow, 5
    RenderPendingDefects ws, machineId, 14
    RenderSchedule ws, machineId, 26
    RenderMaintenanceHistory ws, machineId, 40
    RenderDefectHistory ws, machineId, 58
    ws.Columns("A:K").AutoFit
    Application.ScreenUpdating = True
End Sub

Private Sub RenderMachineInfo(ByVal ws As Worksheet, ByVal lr As ListRow, ByVal startRow As Long)
    Dim fields As Variant
    Dim i As Long
    fields = Array("Machine_ID", "SCOPE", "Machine_Name", "Manufacturer", "Model", "SN", "Range", "Operation_Date", "Status")
    ws.Cells(startRow, 1).Value = "Machine Information"
    ws.Cells(startRow, 1).Font.Bold = True
    For i = LBound(fields) To UBound(fields)
        ws.Cells(startRow + i + 1, 1).Value = fields(i)
        ws.Cells(startRow + i + 1, 2).Value = GetCellValue(lr, CStr(fields(i)))
    Next i
End Sub

Private Sub RenderPendingDefects(ByVal ws As Worksheet, ByVal machineId As String, ByVal startRow As Long)
    ws.Cells(startRow, 1).Value = "Active Pending Defects"
    ws.Cells(startRow, 1).Font.Bold = True
    RenderFilteredTable ws, TBL_DEFECT_LOG, Array("Defect_ID", "Date_Found", "Symptom", "Severity", "Reported_By", "Status"), startRow + 1, machineId, DEFECT_STATUS_PENDING
End Sub

Private Sub RenderSchedule(ByVal ws As Worksheet, ByVal machineId As String, ByVal startRow As Long)
    ws.Cells(startRow, 1).Value = "Spare Parts Schedule"
    ws.Cells(startRow, 1).Font.Bold = True
    RenderFilteredTable ws, TBL_PARTS_SCHEDULE, Array("Machine_ID", "Part_ID", "Last_Changed", "Next_Due", "Status"), startRow + 1, machineId, vbNullString
End Sub

Private Sub RenderMaintenanceHistory(ByVal ws As Worksheet, ByVal machineId As String, ByVal startRow As Long)
    ws.Cells(startRow, 1).Value = "Maintenance History"
    ws.Cells(startRow, 1).Font.Bold = True
    RenderFilteredTable ws, TBL_MAINTENANCE_LOG, Array("Maintenance_ID", "Timestamp", "Action_Type", "Part_ID", "Details", "Operator", "Maintenance_Date", "Defect_ID", "Root_Cause", "Corrective_Action"), startRow + 1, machineId, vbNullString
End Sub

Private Sub RenderDefectHistory(ByVal ws As Worksheet, ByVal machineId As String, ByVal startRow As Long)
    ws.Cells(startRow, 1).Value = "Full Defect History"
    ws.Cells(startRow, 1).Font.Bold = True
    RenderFilteredTable ws, TBL_DEFECT_LOG, Array("Defect_ID", "Timestamp", "Date_Found", "Symptom", "Severity", "Reported_By", "Status", "Root_Cause", "Corrective_Action", "Resolved_By", "Resolved_At"), startRow + 1, machineId, vbNullString
End Sub

Private Sub RenderFilteredTable(ByVal ws As Worksheet, ByVal tableName As String, ByVal columns As Variant, ByVal headerRow As Long, ByVal machineId As String, ByVal statusFilter As String)
    Dim lo As ListObject
    Dim lr As ListRow
    Dim i As Long
    Dim outRow As Long
    Dim includeRow As Boolean
    For i = LBound(columns) To UBound(columns)
        ws.Cells(headerRow, i + 1).Value = columns(i)
    Next i
    ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, UBound(columns) + 1)).Font.Bold = True

    outRow = headerRow + 1
    Set lo = GetTable(tableName)
    If lo.DataBodyRange Is Nothing Then Exit Sub
    For Each lr In lo.ListRows
        includeRow = (NormalizeKey(GetCellValue(lr, "Machine_ID")) = NormalizeKey(machineId))
        If includeRow And Len(statusFilter) > 0 Then includeRow = (StrComp(TrimText(GetCellValue(lr, "Status")), statusFilter, vbTextCompare) = 0)
        If includeRow Then
            For i = LBound(columns) To UBound(columns)
                ws.Cells(outRow, i + 1).Value = GetCellValue(lr, CStr(columns(i)))
            Next i
            outRow = outRow + 1
        End If
    Next lr
End Sub

Public Sub PrintCurrentMachineProfile()
    Dim ws As Worksheet
    Set ws = GetSheet(SHEET_PROFILE)
    If Len(TrimText(ws.Range("B3").Value)) = 0 Then
        MsgBox "Open a machine profile before printing.", vbExclamation
        Exit Sub
    End If
    ws.PrintOut
End Sub
