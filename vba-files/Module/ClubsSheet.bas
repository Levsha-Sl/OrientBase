Attribute VB_Name = "ClubsSheet"
Public wsClubs As Worksheet

Public Const SHEET_NAME As String = "Клубы"
Private Const DELETED As String = "Удалить"

Public Sub Init()
    Set wsClubs = ClubsSheet.GetClubsSheet
End Sub

Private Function GetClubsSheet() As Worksheet
    On Error GoTo ErrorHandler
        Application.EnableEvents = False
        Application.ScreenUpdating = False

        Dim ws As Worksheet: Set ws = ModuleSheet.GetSheetByName(SHEET_NAME)
        If ws Is Nothing Then
            Set ws = ThisWorkbook.Worksheets. _
            Add(Before:=ThisWorkbook.Worksheets(3))
            With ws
                .name = SHEET_NAME
                .Range("A1:D1").value = Array("Список участников", "id", "Название клуба", "Время загрузки")
                .Range("A1:D1").Font.Bold = True

                .Columns("A:D").AutoFit
            End With
        Else
            Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
            If lastRow > 1 Then
                ws.Range("A2:D" & lastRow).ClearContents
            End If
        End If

        Application.DisplayAlerts = False

        Dim sh As Worksheet
        For Each sh In ThisWorkbook.Worksheets
            If sh.Visible = xlSheetVisible And sh.name <> SHEET_NAME And sh.name <> BaseSheet.SHEET_NAME And sh.name <> RanksSheet.SHEET_NAME Then
                sh.Delete
            End If
        Next sh

        Set GetClubsSheet = ws
CleanExit:
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
     Exit Function
ErrorHandler:
        MsgBox "Ошибка при создании листа клубов: " & Err.Description, vbCritical
        Resume CleanExit
End Function

Public Sub AcceptClubs(clubsData As Object)
    On Error GoTo ErrorHandler
        Application.EnableEvents = False
        Application.ScreenUpdating = False

        Dim clubKey As Variant
        Dim clubId As Long
        Dim club As ClubModule
        Dim clSheet As Worksheet

        Dim dataArr()
        Dim sheetNames()

        ReDim dataArr(1 To clubsData.Count, 1 To 4)
        ReDim sheetNames(1 To clubsData.Count)

        For Each clubKey In clubsData.Keys
            clubId = CLng(clubKey)
            Set club = clubsData.Item(clubId)
            Set clSheet = ClubSheet.GetClubSheet(clubId, club)
            sheetNames(clubId) = clSheet.name

            dataArr(clubId, 1) = clubId
            dataArr(clubId, 2) = club.clubName
            dataArr(clubId, 3) = club.timeStamp
            dataArr(clubId, 4) = DELETED
        Next clubKey

        With wsClubs
            .Range("B2").Resize(UBound(dataArr, 1), 4).value = dataArr

            Dim i As Long
            For i = 1 To UBound(sheetNames)
                .Hyperlinks.Add Anchor:=.Cells(i + 1, 1), Address:="", _
                SubAddress:="'" & sheetNames(i) & "'!A1", _
                TextToDisplay:="Перейти к списку"
            Next i

            .Columns(4).NumberFormat = "dd.mm.yyyy hh:mm:ss"
            .Columns(5).Font.Color = vbRed
            .Columns("A:E").AutoFit
            .Activate
        End With
CleanExit:
        Application.EnableEvents = True
        Application.ScreenUpdating = True
     Exit Sub
ErrorHandler:
        MsgBox "Ошибка при записи клубов в листы: " & Err.Description, vbCritical
        Resume CleanExit
End Sub

Public Sub HandleSelectionChange(Target As Range)
    If Target.Column = 5 _
        And Target.Count = 1 _
        And Target.row > 1 _
        And Target.value = DELETED Then

        Dim mainRef As String: mainRef = ModuleSheet.GetSheetFromLink(wsClubs.Cells(Target.row, 1))

        If MsgBox("Удалить лист клуба?", vbYesNo + vbQuestion) = vbYes Then

            Application.DisplayAlerts = False

            ModuleSheet.DeleteSheetIfExists mainRef
            Target.EntireRow.Delete

            Application.DisplayAlerts = True
        End If
    End If
End Sub

Public Sub HandleFollowHyperlink(Target As Hyperlink)
    Dim sheetName As String: sheetName = Split(Target.SubAddress, "!")(0)
    sheetName = Replace(sheetName, "'", "")

    If Not ModuleSheet.SheetExists(sheetName) Then
        MsgBox "Лист [" & sheetName & "] не найден. Лист клуба удален.", vbExclamation
        Target.Range.EntireRow.Delete
    End If
End Sub
