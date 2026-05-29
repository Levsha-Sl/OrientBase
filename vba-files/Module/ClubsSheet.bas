Attribute VB_Name = "ClubsSheet"
Public wsClubs As Worksheet
Public Sub InitClubsSheet()
    If wsClubs Is Nothing Then
        Set wsClubs = ClubsSheet.GetClubsSheet
    End If
End Sub

Private Function GetClubsSheet() As Worksheet
    Dim ws As Worksheet: Set ws = ModuleSheet.GetSheetByName("Клубы")

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets. _
        Add(Before:=ThisWorkbook.Worksheets(3))
        ws.Name = "Клубы"

        ws.Range("A1:E1").Value = Array("Список участников", "Что предоставить", "Полное название клуба", "Время загрузки", "Управление")
        ws.Range("A1:E1").Font.Bold = True

        ws.Columns("A:E").AutoFit
    End If
    Set GetClubsSheet = ws
End Function

Public Sub AddClub(clubFullName As String, wsClubName As String, wsNeedsName As String, dataTime As Date)
    Dim targetRow As Long: targetRow = wsClubs.Cells(wsClubs.Rows.Count, 3).End(xlUp).Row + 1

    wsClubs.Cells(targetRow, 3).Value = clubFullName
    wsClubs.Cells(targetRow, 4).Value = Format(dataTime, "dd.mm.yyyy hh:mm:ss")
    wsClubs.Hyperlinks.Add Anchor:=wsClubs.Cells(targetRow, 1), Address:="", _
    SubAddress:="'" & wsClubName & "'!A1", TextToDisplay:="Перейти к списку"
    wsClubs.Hyperlinks.Add Anchor:=wsClubs.Cells(targetRow, 2), Address:="", _
    SubAddress:="'" & wsNeedsName & "'!A1", TextToDisplay:="Открыть документы"
    wsClubs.Cells(targetRow, 5).Value = "УДАЛИТЬ КЛУБ"
    wsClubs.Cells(targetRow, 5).Font.Color = vbRed
    wsClubs.Columns("A:D").AutoFit
End Sub

Public Sub RemoveEmptyLinks()
    Dim i As Long, sheetName As String
    Dim lastRow As Long: lastRow = wsClubs.Cells(wsClubs.Rows.Count, 3).End(xlUp).Row

    For i = lastRow To 2 Step -1
        sheetName = ModuleSheet.GetSheetFromLink(wsClubs.Cells(i, 1))

        If sheetName = "" Or Not ModuleSheet.SheetExists(sheetName) Then
            wsClubs.Rows(i).Delete
        End If
    Next i
End Sub

