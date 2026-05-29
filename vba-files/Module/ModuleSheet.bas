Attribute VB_Name = "ModuleSheet"
Public Function GetSheetByName(sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheetByName = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

Public Function GetTableByName(ws As Worksheet, tableName As String) As ListObject
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    Set GetTableByName = ws.ListObjects(tableName)
    On Error GoTo 0
End Function

Function GetSheetFromLink(cell As Range) As String
    On Error Resume Next
    GetSheetFromLink = Replace(Split(cell.Hyperlinks(1).SubAddress, "!")(0), "'", "")
    On Error GoTo 0
End Function

Sub DeleteSheetIfExists(sName As String)
    If sName = "" Then Exit Sub
        If SheetExists(sName) Then Sheets(sName).Delete
End Sub

Function SheetExists(sName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function
