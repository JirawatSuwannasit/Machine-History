Attribute VB_Name = "modSetup"
Option Explicit

Public Sub SetupWorkbook()
    Application.ScreenUpdating = False
    EnsureAllSheets
    EnsureAllTables
    SetupSettings
    BuildDashboardLayout
    BuildProfileLayout
    Application.ScreenUpdating = True
    MsgBox "Machine Maintenance workbook setup is complete.", vbInformation
End Sub

Public Sub EnsureAllSheets()
    EnsureSheet SHEET_DASHBOARD
    EnsureSheet SHEET_PROFILE
    EnsureSheet SHEET_MACHINE_LIST
    EnsureSheet SHEET_PARTS_MASTER
    EnsureSheet SHEET_PARTS_SCHEDULE
    EnsureSheet SHEET_MAINTENANCE_LOG
    EnsureSheet SHEET_DEFECT_LOG
    EnsureSheet SHEET_AUDIT_LOG
    EnsureSheet SHEET_SETTINGS
End Sub

Private Sub EnsureSheet(ByVal sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If
End Sub

Public Sub EnsureAllTables()
    EnsureTable SHEET_MACHINE_LIST, TBL_MACHINE_LIST, Array("Machine_ID", "SCOPE", "Machine_Name", "Manufacturer", "Model", "SN", "Range", "Operation_Date", "Status")
    EnsureTable SHEET_PARTS_MASTER, TBL_PARTS_MASTER, Array("Part_ID", "Part_Name", "Lifetime_Years", "Description")
    EnsureTable SHEET_PARTS_SCHEDULE, TBL_PARTS_SCHEDULE, Array("Machine_ID", "Part_ID", "Last_Changed", "Next_Due", "Status")
    EnsureTable SHEET_MAINTENANCE_LOG, TBL_MAINTENANCE_LOG, Array("Maintenance_ID", "Timestamp", "Machine_ID", "Action_Type", "Part_ID", "Details", "Operator", "Maintenance_Date", "Defect_ID", "Root_Cause", "Corrective_Action")
    EnsureTable SHEET_DEFECT_LOG, TBL_DEFECT_LOG, Array("Defect_ID", "Timestamp", "Machine_ID", "Date_Found", "Symptom", "Severity", "Reported_By", "Status", "Root_Cause", "Corrective_Action", "Resolved_By", "Resolved_At", "Linked_Maintenance_ID")
    EnsureTable SHEET_AUDIT_LOG, TBL_AUDIT_LOG, Array("Audit_ID", "Timestamp", "User_Name", "Action", "Entity", "Entity_ID", "Before_Value", "After_Value")
End Sub

Private Sub EnsureTable(ByVal sheetName As String, ByVal tableName As String, ByVal headers As Variant)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim headerRange As Range
    Dim i As Long

    Set ws = GetSheet(sheetName)
    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    On Error GoTo 0

    If lo Is Nothing Then
        ws.Cells.Clear
        For i = LBound(headers) To UBound(headers)
            ws.Cells(1, i + 1).Value = headers(i)
        Next i
        Set headerRange = ws.Range(ws.Cells(1, 1), ws.Cells(2, UBound(headers) + 1))
        Set lo = ws.ListObjects.Add(xlSrcRange, headerRange, , xlYes)
        lo.Name = tableName
        If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
    Else
        For i = LBound(headers) To UBound(headers)
            EnsureTableColumn lo, CStr(headers(i))
        Next i
    End If
    lo.TableStyle = "TableStyleMedium2"
    ws.Columns.AutoFit
End Sub

Private Sub EnsureTableColumn(ByVal lo As ListObject, ByVal columnName As String)
    Dim col As ListColumn
    On Error Resume Next
    Set col = lo.ListColumns(columnName)
    On Error GoTo 0
    If col Is Nothing Then
        Set col = lo.ListColumns.Add
        col.Name = columnName
    End If
End Sub

Public Sub SetupSettings()
    Dim ws As Worksheet
    Set ws = GetSheet(SHEET_SETTINGS)
    ws.Cells.Clear
    ws.Range("A1").Value = "Action_Type"
    ws.Range("A2").Resize(3, 1).Value = Application.Transpose(Array(ACTION_REPAIR, ACTION_PART_REPLACEMENT, ACTION_PM))
    ws.Range("C1").Value = "Severity"
    ws.Range("C2").Resize(4, 1).Value = Application.Transpose(Array(SEVERITY_LOW, SEVERITY_MEDIUM, SEVERITY_HIGH, SEVERITY_CRITICAL))
    ws.Range("E1").Value = "Machine_Status"
    ws.Range("E2").Resize(2, 1).Value = Application.Transpose(Array(MACHINE_STATUS_ACTIVE, MACHINE_STATUS_INACTIVE))
    ws.Range("G1").Value = "Defect_Status"
    ws.Range("G2").Resize(2, 1).Value = Application.Transpose(Array(DEFECT_STATUS_PENDING, DEFECT_STATUS_RESOLVED))
    ws.Columns.AutoFit
End Sub

Public Sub BuildDashboardLayout()
    Dim ws As Worksheet
    Set ws = GetSheet(SHEET_DASHBOARD)
    DeleteSheetShapes ws
    ws.Cells.Clear
    ws.Range("A1").Value = "Machine History & Maintenance Dashboard"
    ws.Range("A1").Font.Size = 18
    ws.Range("A1").Font.Bold = True
    ws.Range("A3").Value = "Search"
    ws.Range("B3").Value = ""
    ws.Range("D3").Value = "Filter"
    ws.Range("E3").Value = FILTER_ALL
    ws.Range("A5:G5").Value = Array("Total Machines", "Active Machines", "Inactive Machines", "Pending Defects", "Critical Defects", "Overdue Parts", "Due Soon")
    ws.Range("A5:G6").Font.Bold = True
    ws.Range("A8").Value = "Machine Cards"
    ws.Range("A8").Font.Bold = True
    AddDashboardButton ws, "Refresh Dashboard", "RefreshDashboard", 10, 10
    AddDashboardButton ws, "New Maintenance Log", "ShowMaintenanceLogForm", 150, 10
    AddDashboardButton ws, "Report Defect", "ShowDefectReportForm", 310, 10
    AddDashboardButton ws, "Open Machine Profile", "ShowMachinePickerForProfile", 440, 10
    AddDashboardButton ws, "Print Report", "PrintCurrentMachineProfile", 610, 10
    ws.Columns("A:K").ColumnWidth = 18
End Sub

Private Sub DeleteSheetShapes(ByVal ws As Worksheet)
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
End Sub

Private Sub AddDashboardButton(ByVal ws As Worksheet, ByVal caption As String, ByVal macroName As String, ByVal leftPos As Double, ByVal topPos As Double)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, 130, 28)
    shp.TextFrame2.TextRange.Text = caption
    shp.OnAction = macroName
End Sub

Public Sub BuildProfileLayout()
    Dim ws As Worksheet
    Set ws = GetSheet(SHEET_PROFILE)
    ws.Cells.Clear
    ws.Range("A1").Value = "Machine Profile"
    ws.Range("A1").Font.Size = 18
    ws.Range("A1").Font.Bold = True
    ws.Range("A3").Value = "Selected Machine_ID"
    ws.Range("B3").Value = ""
    ws.Columns("A:K").ColumnWidth = 18
End Sub
