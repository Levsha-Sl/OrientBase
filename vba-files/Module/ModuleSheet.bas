Attribute VB_Name = "ModuleSheet"
Public Function GetSheetByName(sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheetByName = ThisWorkbook.Worksheets(sheetName)
    On Error Goto 0
End Function

Public Function GetTableByName(ws As Worksheet, tableName As String) As ListObject
    If ws Is Nothing Then Exit Function
        On Error Resume Next
        Set GetTableByName = ws.ListObjects(tableName)
        On Error Goto 0
End Function

Function GetSheetFromLink(cell As Range) As String
    On Error Resume Next
    GetSheetFromLink = Replace(Split(cell.Hyperlinks(1).SubAddress, "!")(0), "'", "")
    On Error Goto 0
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
    On Error Goto 0
End Function


Public Function CalculateRankExpiry(Byval rankDate As Variant, Byval targetRank As String) As Variant
    If IsEmpty(rankDate) Or rankDate = "" Or targetRank = "" Then
        CalculateRankExpiry = ""
     Exit Function
    End If

    If Not IsDate(rankDate) Then
        CalculateRankExpiry = CVErr(xlErrValue) ' Возвращаем #ЗНАЧ!
     Exit Function
    End If

    On Error Resume Next
    Dim addedYears As Long
    addedYears = RanksData.GetRankValue(targetRank)
    If Err.Number <> 0 Then
        CalculateRankExpiry = CVErr(xlErrValue)
        Err.Clear
     Exit Function
    End If
    On Error Goto 0

        If addedYears <= 0 Then
            CalculateRankExpiry = ""
         Exit Function
        End If

        CalculateRankExpiry = DateAdd("yyyy", addedYears, CDate(rankDate)) - 1
End Function

Public Function CalculateInsuranceExpiry(Byval insDate As Variant, Byval period As Variant) As Variant
    If IsEmpty(insDate) Or insDate = "" Or IsEmpty(period) Or period = "" Then
        CalculateInsuranceExpiry = ""
     Exit Function
    End If

    If Not IsDate(insDate) Or Not IsNumeric(period) Then
        CalculateInsuranceExpiry = CVErr(xlErrValue)
     Exit Function
    End If

    If CLng(period) <= 0 Then
        CalculateInsuranceExpiry = ""
     Exit Function
    End If

    CalculateInsuranceExpiry = CDate(insDate) + CLng(period) - 1
End Function