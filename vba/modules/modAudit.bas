Attribute VB_Name = "modAudit"
Option Explicit

Public Sub WriteAuditLog(ByVal actionName As String, ByVal entityName As String, ByVal entityId As String, Optional ByVal beforeValue As String = "", Optional ByVal afterValue As String = "")
    Dim lr As ListRow
    Set lr = AddTableRow(TBL_AUDIT_LOG)
    SetCellValue lr, "Audit_ID", NextSequentialID(TBL_AUDIT_LOG, "Audit_ID", PREFIX_AUDIT)
    SetCellValue lr, "Timestamp", NowStamp
    SetCellValue lr, "User_Name", CurrentUserName
    SetCellValue lr, "Action", actionName
    SetCellValue lr, "Entity", entityName
    SetCellValue lr, "Entity_ID", entityId
    SetCellValue lr, "Before_Value", beforeValue
    SetCellValue lr, "After_Value", afterValue
End Sub
