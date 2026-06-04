Attribute VB_Name = "ClubSheet"
Private wsCl As Worksheet
Private tblCl As ListObject

Private Const COL_STATUS As String = "Статус"
Private Const ARR_STATUS As Array = Array("Инф.базы","Обнов","Новый")
Private Const COL_RANK As String = "Разряд"
Private Const COL_RANK_DATE As String = "дата_раз."
Private Const COL_RANK_EXPIRY As String = "окончание"
Private Const COL_INS_DATE As String = "дата_страх."
Private Const COL_PERIOD As String = "период"
Private Const COL_INS_EXPIRY As String = "окончaние"

Public Function GetClubSheet(clubIndex As Long, club As ClubModule) As Worksheet
    On Error Goto ErrorHandler
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        Set wsCl = ThisWorkbook.Worksheets.Add _
    (After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ' Имя листа: L1_2204_1530 (Коротко и уникально)

        Dim dataTime As Date: dataTime = club.TimeStamp
        Dim clubName As String: clubName = club.ClubName

        With wsCl
            .name = "L" & clubIndex & "_" & Format(dataTime, "ddmm_hhmmss")

            Dim headerData(1 To 3, 1 To 1) As Variant
            headerData(1, 1) = "Клуб: " & clubName
            headerData(2, 1) = "Загружено: " & Format(dataTime, "dd.mm.yyyy hh:mm:ss")
            headerData(3, 1) = "Количество участников: 0"
            .Range("A1:A3").Value = headerData

            .Range("A4:J4").Value = Array(COL_STATUS,"id", "ФИО", "День рож.", COL_RANK, COL_RANK_DATE, COL_RANK_EXPIRY, COL_INS_DATE, COL_PERIOD, COL_INS_EXPIRY)
            .Range("A4:J4").Font.Bold = True

            Dim btn As Object: Set btn = .Buttons.Add( _
            Left:=.Cells(1, 5).Left, _
            Top:=.Cells(1, 5).Top, _
            Width:=.Cells(1, 5).Width + .Cells(1, 6).Width, _
        Height:=30)

        With btn
            '    .OnAction = ""
            .Caption = "Выгрузить в базу"
        End With

            .Hyperlinks.Add Anchor:=.Range("B1"), Address:="", SubAddress:=ClubsSheet.SHEET_NAME &"!A" & clubIndex + 1, TextToDisplay:="<< К КЛУБАМ"
            .Columns("A:H").AutoFit
        End With

        ' TODO загрузка данных
        
        Set GetClubSheet = wsCl
 CleanExit:
        Application.EnableEvents = True
    Application.ScreenUpdating = True
     Exit Function
 ErrorHandler:
        MsgBox "Ошибка при создании листа клуба(" & clubName & "): " & Err.Description, vbCritical
        Resume CleanExit
End Function