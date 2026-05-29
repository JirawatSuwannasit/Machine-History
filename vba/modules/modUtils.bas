Attribute VB_Name = "modUtils"
Option Explicit

Public Function Nz(ByVal value As Variant, Optional ByVal fallback As String = "") As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
        Nz = fallback
    Else
        Nz = CStr(value)
    End If
End Function

Public Function TrimText(ByVal value As Variant) As String
    TrimText = Trim$(Nz(value))
End Function

Public Function IsBlank(ByVal value As Variant) As Boolean
    IsBlank = (Len(TrimText(value)) = 0)
End Function

Public Function ParseDateRequired(ByVal value As Variant, ByVal fieldName As String) As Date
    If IsBlank(value) Or Not IsDate(value) Then Err.Raise vbObjectError + 100, , fieldName & " must be a valid date."
    ParseDateRequired = DateValue(CDate(value))
End Function

Public Function SafeDateText(ByVal value As Variant) As String
    If IsDate(value) Then SafeDateText = Format$(CDate(value), "yyyy-mm-dd") Else SafeDateText = ""
End Function

Public Function NowStamp() As Date
    NowStamp = Now
End Function

Public Function TodayDate() As Date
    TodayDate = Date
End Function

Public Function AddYearsFlexible(ByVal startDate As Date, ByVal yearsValue As Double) As Date
    Dim wholeYears As Long
    Dim monthOffset As Long
    wholeYears = Fix(yearsValue)
    monthOffset = Round((yearsValue - wholeYears) * 12, 0)
    AddYearsFlexible = DateAdd("m", monthOffset, DateAdd("yyyy", wholeYears, startDate))
End Function

Public Function IsInArray(ByVal value As String, ByVal allowedValues As Variant) As Boolean
    Dim item As Variant
    For Each item In allowedValues
        If StrComp(value, CStr(item), vbTextCompare) = 0 Then
            IsInArray = True
            Exit Function
        End If
    Next item
End Function

Public Function CurrentUserName() As String
    CurrentUserName = Environ$("USERNAME")
    If Len(CurrentUserName) = 0 Then CurrentUserName = Application.UserName
End Function

Public Function NextSequentialID(ByVal tableName As String, ByVal idColumnName As String, ByVal prefix As String) As String
    Dim lo As ListObject
    Dim lr As ListRow
    Dim maxNumber As Long
    Dim rawId As String
    Dim numericPart As Long

    Set lo = GetTable(tableName)
    If Not lo.DataBodyRange Is Nothing Then
        For Each lr In lo.ListRows
            rawId = TrimText(GetCellValue(lr, idColumnName))
            If Left$(rawId, Len(prefix)) = prefix Then
                numericPart = Val(Mid$(rawId, Len(prefix) + 1))
                If numericPart > maxNumber Then maxNumber = numericPart
            End If
        Next lr
    End If

    NextSequentialID = prefix & Format$(maxNumber + 1, String$(ID_DIGITS, "0"))
End Function

Public Function RowToDictionary(ByVal lr As ListRow) As Object
    Dim dict As Object
    Dim col As ListColumn
    Set dict = CreateObject("Scripting.Dictionary")

    For Each col In lr.Parent.ListColumns
        dict(col.Name) = lr.Range.Cells(1, col.Index).Value
    Next col

    Set RowToDictionary = dict
End Function

Public Function GetCellValue(ByVal lr As ListRow, ByVal columnName As String) As Variant
    GetCellValue = lr.Range.Cells(1, lr.Parent.ListColumns(columnName).Index).Value
End Function

Public Sub SetCellValue(ByVal lr As ListRow, ByVal columnName As String, ByVal value As Variant)
    lr.Range.Cells(1, lr.Parent.ListColumns(columnName).Index).Value = value
End Sub

Public Function NormalizeKey(ByVal value As Variant) As String
    NormalizeKey = UCase$(TrimText(value))
End Function

Public Sub ClearRangeContentsAndFormats(ByVal target As Range)
    target.ClearContents
    target.ClearFormats
End Sub
