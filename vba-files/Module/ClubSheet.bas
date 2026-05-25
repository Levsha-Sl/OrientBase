Attribute VB_Name = "ClubSheet"
Private ws As Worksheet

Public Function GetClubSheet(clubIndex As Long, mode As String, dataTime As Date, clubFullName As String) As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add _
    (After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ' Имя листа: L1_2204_1530 (Коротко и уникально)

    ws.Name = IIf(mode = "List", "L", "N") & clubIndex & "_" & Format(dataTime, "ddmm_hhmmss")

    ' ШАПКА ЛИСТА
    ws.Range("A1").Value = "Клуб: " & clubFullName
    ws.Range("A2").Value = "Загружено: " & Format(dataTime, "dd.mm.yyyy hh:mm:ss")
    ws.Range("A3").Value = "Количество участников: 0"

    If mode = "List" Then
        ws.Range("A4:H4").Value = Array("ФИО", "День рож.", "Разряд", "дата раз.", "окончание", "дата страх.", "период", "окончание")

        Dim btn As Object: Set btn = ws.Buttons.Add( _
        Left:=ws.Cells(1, 5).Left, _
        Top:=ws.Cells(1, 5).Top, _
        Width:=ws.Cells(1, 5).Width + ws.Cells(1, 6).Width, _
        Height:=30)

        With btn
            .OnAction = "LoadCurrentListToBase"
            .Caption = "Выгрузить в базу"
        End With

        ws.Range("A4:I4").Font.Bold = True
    Else
        ws.Range("A4:D4").Value = Array("ФИО", "День рож.", "Разряд", "Комментарий")
        ws.Range("A4:D4").Font.Bold = True
    End If

    ' Навигация и закрепление
    ws.Hyperlinks.Add Anchor:=ws.Range("B1"), Address:="", SubAddress:="Клубы!A1", TextToDisplay:="<< К КЛУБАМ"
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A5").Select
    ActiveWindow.FreezePanes = True

    Set GetClubSheet = ws
End Function

Sub LoadCurrentListToBase()
    Set ws = ActiveSheet

    Application.ScreenUpdating = False

    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = 5 To lastRow
        If NeedToUpdate(i) Then
            Dim fio As String, bd As Date
            ' ФИО и дата рождения всегда в первой строке
            fio = Trim(ws.Cells(i, 1).Value)
            bd = ws.Cells(i, 2).Value
            ' Остальне новые данные могут быть только во второй строке, так как
            ' Вперовую строку данные будут подтянуты из базы
            If fio <> "" Then
                i = i + 1 ' Переходим на вторую строку, где могут быть новые данные
                Dim personRow As Long: personRow = BaseSheet.PersonExists(fio, bd)
                If personRow = 0 Then
                    Dim baseLastRow As Long: baseLastRow = wsBase.Cells(wsBase.Rows.Count, 1).End(xlUp).Row + 1

                    wsBase.Cells(baseLastRow, 1).Value = fio
                    wsBase.Cells(baseLastRow, 2).Value = bd

                    ' Загрузаем остальные данные и форматируем строку
                    Call UpdatePersonData(i, baseLastRow)
                Else
                    ' Если человек уже есть - обновляем только данные
                    Call UpdatePersonData(i, personRow)
                End If
            End If
        End If
    Next i

    wsBase.Columns("A:I").AutoFit
    Application.ScreenUpdating = True

    MsgBox "Данные с листа загружены в базу", vbInformation
End Sub

Private  Function NeedToUpdate(row As Long) As Boolean
    NeedToUpdate = ws.Cells(row + 1, 1).Value = "" And _
    ws.Cells(row + 1, 2).Value = "" And _
    ws.Cells(row + 1, 3).Value <> ""
End Function

Private  Function UpdatePersonData(i As Long, personRow As Long)
    ' Обновляем данные по разряду и страховке, но только если формула вернула дату
    If IsDate(ws.Cells(i, 5).Value) Or (RankSheet.GetRankValue(ws.Cells(i,3).Value) = 0) Then
        wsBase.Cells(personRow, 3).Value = ws.Cells(i, 3).Value
        wsBase.Cells(personRow, 4).Value = ws.Cells(i, 4).Value
    End If
    If IsDate(ws.Cells(i, 8).Value) Then
        wsBase.Cells(personRow, 6).Value = ws.Cells(i, 6).Value
        wsBase.Cells(personRow, 7).Value = ws.Cells(i, 7).Value
    End If

    wsBase.Cells(personRow, 9).Value = Now
    Call BaseSheet.ApplyLogicAndFormatting(personRow)
End Function