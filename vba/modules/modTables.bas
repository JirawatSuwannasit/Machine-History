Attribute VB_Name = "modTables"
Option Explicit

Public Function GetSheet(ByVal sheetName As String) As Worksheet
    Set GetSheet = ThisWorkbook.Worksheets(sheetName)
End Function

Public Function GetTable(ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
                Set GetTable = lo
                Exit Function
            End If
        Next lo
    Next ws
    Err.Raise vbObjectError + 200, , "Table not found: " & tableName
End Function

Public Function FindRowByValue(ByVal tableName As String, ByVal columnName As String, ByVal matchValue As Variant) As ListRow
    Dim lo As ListObject
    Dim lr As ListRow
    Set lo = GetTable(tableName)

    If lo.DataBodyRange Is Nothing Then Exit Function
    For Each lr In lo.ListRows
        If StrComp(TrimText(GetCellValue(lr, columnName)), TrimText(matchValue), vbTextCompare) = 0 Then
            Set FindRowByValue = lr
            Exit Function
        End If
    Next lr
End Function

Public Function FindScheduleRow(ByVal machineId As String, ByVal partId As String) As ListRow
    Dim lo As ListObject
    Dim lr As ListRow
    Set lo = GetTable(TBL_PARTS_SCHEDULE)

    If lo.DataBodyRange Is Nothing Then Exit Function
    For Each lr In lo.ListRows
        If NormalizeKey(GetCellValue(lr, "Machine_ID")) = NormalizeKey(machineId) _
            And NormalizeKey(GetCellValue(lr, "Part_ID")) = NormalizeKey(partId) Then
            Set FindScheduleRow = lr
            Exit Function
        End If
    Next lr
End Function

Public Function AddTableRow(ByVal tableName As String) As ListRow
    Set AddTableRow = GetTable(tableName).ListRows.Add
End Function

Public Function MachineExists(ByVal machineId As String) As Boolean
    MachineExists = Not FindRowByValue(TBL_MACHINE_LIST, "Machine_ID", machineId) Is Nothing
End Function

Public Function PartExists(ByVal partId As String) As Boolean
    PartExists = Not FindRowByValue(TBL_PARTS_MASTER, "Part_ID", partId) Is Nothing
End Function

Public Function DefectExists(ByVal defectId As String) As Boolean
    DefectExists = Not FindRowByValue(TBL_DEFECT_LOG, "Defect_ID", defectId) Is Nothing
End Function

Public Function GetPartLifetimeYears(ByVal partId As String) As Double
    Dim lr As ListRow
    Set lr = FindRowByValue(TBL_PARTS_MASTER, "Part_ID", partId)
    If lr Is Nothing Then Err.Raise vbObjectError + 201, , "Part_ID does not exist: " & partId
    GetPartLifetimeYears = CDbl(GetCellValue(lr, "Lifetime_Years"))
End Function

Public Function GetMachinesArray(Optional ByVal includeBlankFirst As Boolean = False) As Variant
    GetMachinesArray = GetLookupArray(TBL_MACHINE_LIST, "Machine_ID", "Machine_Name", includeBlankFirst)
End Function

Public Function GetPartsArray(Optional ByVal includeBlankFirst As Boolean = False) As Variant
    GetPartsArray = GetLookupArray(TBL_PARTS_MASTER, "Part_ID", "Part_Name", includeBlankFirst)
End Function

Public Function GetLookupArray(ByVal tableName As String, ByVal idColumn As String, ByVal labelColumn As String, Optional ByVal includeBlankFirst As Boolean = False) As Variant
    Dim lo As ListObject
    Dim lr As ListRow
    Dim result() As String
    Dim index As Long
    Dim itemText As String

    Set lo = GetTable(tableName)
    ReDim result(0 To IIf(includeBlankFirst, lo.ListRows.Count, Application.Max(lo.ListRows.Count - 1, 0)))
    If includeBlankFirst Then
        result(0) = ""
        index = 1
    End If

    If Not lo.DataBodyRange Is Nothing Then
        For Each lr In lo.ListRows
            itemText = TrimText(GetCellValue(lr, idColumn))
            If Len(labelColumn) > 0 Then itemText = itemText & " - " & TrimText(GetCellValue(lr, labelColumn))
            result(index) = itemText
            index = index + 1
        Next lr
    End If

    GetLookupArray = result
End Function

Public Function ExtractIDFromCombo(ByVal displayText As String) As String
    Dim position As Long
    position = InStr(1, displayText, " - ", vbTextCompare)
    If position > 0 Then
        ExtractIDFromCombo = Trim$(Left$(displayText, position - 1))
    Else
        ExtractIDFromCombo = TrimText(displayText)
    End If
End Function

Public Function GetPendingDefectsArray(ByVal machineId As String, Optional ByVal includeBlankFirst As Boolean = True) As Variant
    Dim lo As ListObject
    Dim lr As ListRow
    Dim items As Collection
    Dim arr() As String
    Dim i As Long
    Set lo = GetTable(TBL_DEFECT_LOG)
    Set items = New Collection
    If includeBlankFirst Then items.Add ""

    If Not lo.DataBodyRange Is Nothing Then
        For Each lr In lo.ListRows
            If NormalizeKey(GetCellValue(lr, "Machine_ID")) = NormalizeKey(machineId) _
                And StrComp(TrimText(GetCellValue(lr, "Status")), DEFECT_STATUS_PENDING, vbTextCompare) = 0 Then
                items.Add TrimText(GetCellValue(lr, "Defect_ID")) & " - " & TrimText(GetCellValue(lr, "Symptom"))
            End If
        Next lr
    End If

    ReDim arr(0 To Application.Max(items.Count - 1, 0))
    For i = 1 To items.Count
        arr(i - 1) = CStr(items(i))
    Next i
    GetPendingDefectsArray = arr
End Function
