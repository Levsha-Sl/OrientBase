Attribute VB_Name = "Import"
Private dictClubs As Object

Sub CSV()
    On Error Goto ErrorHandler

        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False

        Dim fd As FileDialog
        Set fd = Application.FileDialog(msoFileDialogFilePicker)

        With fd
            .Title = "Выберите FjwW CSV файл c Orgeo"
            .Filters.Clear
            .Filters.Add "CSV Files", "*.csv" '

            If .Show = True Then
                strFile = .SelectedItems(1)
            Else
                Goto CleanExit
                End If
            End With

            Set dictClubs = CreateObject("Scripting.Dictionary")

            Call ProcessCSVToClubs(fd.SelectedItems(1))

            MsgBox "Загрузка завершена. Клубов: " & dictClubs.Count, vbInformation

 CleanExit:
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
         Exit Sub

 ErrorHandler:
            MsgBox "Ошибка: " & Err.Description, vbCritical
            Resume CleanExit
End Sub

Sub ProcessCSVToClubs(filePath As String)
    Call ClubsSheet.RemoveEmptyLinks

    ' ОПТИМИЗАЦИЯ: Загружаем базу в памяти один раз
    Call BaseSheet.LoadBaseCache

    Dim dataTime As Date: dataTime = Now

    Dim dictWsMain As Object: Set dictWsMain = CreateObject("Scripting.Dictionary")
    Dim dictWsNeeds As Object: Set dictWsNeeds = CreateObject("Scripting.Dictionary")

    Dim objStream As Object
    Dim lineData As String, cols() As String
    Dim fFull As String, fBirth As String, fStat As String, fClub As String
    Dim targetWsClub As Worksheet, targetWsNeeds As Worksheet

    Dim lastRowClub As Long, baseRow As Long

    ' Настройка стрима для UTF-8
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Charset = "utf-8"
    objStream.Open
    objStream.LoadFromFile filePath

    ' Настройка разделителей для формул
    Application.DecimalSeparator = ","
    Application.UseSystemSeparators = True

    ' Пропускаем заголовок
    If Not objStream.EOS Then objStream.ReadText (-2)

        Do Until objStream.EOS
            lineData = objStream.ReadText(-2)
            If Trim(lineData) = "" Then Goto NextLine

                cols = Split(lineData, ";")

                ' 1. РЕГИСТРАЦИЯ КЛУБА И СОЗДАНИЕ ЛИСТОВ
                fClub = cols(1)
                If Not dictWsMain.Exists(fClub) Then
                    dictClubs.Add fClub, 1

                    ' Создаем пару уникальных листов L и N
                    Set targetWsClub = ClubSheet.GetClubSheet(dictClubs.Count, "List", dataTime, fClub)
                    Set targetWsNeeds = ClubSheet.GetClubSheet(dictClubs.Count, "Needs", dataTime, fClub)

                    ' Сохраняем ОБЪЕКТЫ листов в локальные словари
                    Set dictWsMain(fClub) = targetWsClub
                    Set dictWsNeeds(fClub) = targetWsNeeds

                    Call ClubsSheet.AddClub(fClub, targetWsClub.Name, targetWsNeeds.Name, dataTime)
                Else
                    Set targetWsClub = dictWsMain(fClub)
                    Set targetWsNeeds = dictWsNeeds(fClub)
                End If

                ' 2. ДАННЫЕ УЧАСТНИКА
                fFull = Trim(cols(3) & " " & cols(4) & " " & cols(5)) ' NAME + FIRSTNAME + MIDDLENAME
                fBirth = Trim(cols(7))
                fStat = cols(11)

                ' Определяем строку для записи (всегда ниже последнего заполненного Разряда)
                lastRowClub = targetWsClub.Cells(targetWsClub.Rows.Count, "C").End(xlUp).Row + 1

                ' 3. ВЫЗОВ ПРОЦЕДУР ОБРАБОТКИ
                ' TODO
                baseRow = BaseSheet.PersonExists(fFull, CDate(fBirth))
                If baseRow = 0 Then
                    Call ProcessNewPerson(targetWsClub, targetWsNeeds, fFull, fBirth, fStat, lastRowClub)
                Else
                    Call ProcessUpdatePerson(targetWsClub, targetWsNeeds, baseRow, fStat, lastRowClub)
                End If

 NextLine:
                DoEvents ' Обработка очереди событий для стабильности
            Loop
            objStream.Close

            ' 4. ФИНАЛЬНЫЙ РАСЧЕТ УЧАСТНИКОВ
            Dim vItem As Variant, wsToUpdate As Worksheet
            For Each vItem In dictWsMain.Items
                Set wsToUpdate = vItem
                Call UpdateCounterAndFormat(wsToUpdate)
                wsToUpdate.Columns("A:H").AutoFit
            Next vItem
            For Each vItem In dictWsNeeds.Items
                Set wsToUpdate = vItem
                Call UpdateCounterAndFormat(wsToUpdate)
                wsToUpdate.Columns("A:D").AutoFit
            Next vItem

            ' Очищаем кэш базы из памяти
            Call BaseSheet.ClearBaseCache

            ' Восстанавливаем состояние приложения
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True

            wsClubs.Activate
End Sub

' --- ЛОГИКА ДЛЯ НОВОГО ЧЕЛОВЕКА ---
Private Sub ProcessNewPerson(ws As Worksheet, wsN As Worksheet, fio As String, bd As String, st As String, r As Long)
    ws.Cells(r, 1) = fio
    ws.Cells(r, 2) = bd
    ' Пропускаем строку (r), пишем в (r+1)
    Dim dataR As Long: dataR = r + 1
    ws.Cells(dataR, 3) = st

    ' Красим пустые в сиреневый (кроме ФИО, др)
    ws.Range(ws.Cells(dataR, 3), ws.Cells(dataR, 8)).Interior.Color = RGB(221, 160, 221)

    ' Комментарий и формулы
    Call AddComment(wsN, fio, bd, st, "нужно паспорт,страховку и зачетку")
    Call ApplyLogicAndFormatting(ws, dataR)
End Sub

' --- ЛОГИКА ДЛЯ СУЩЕСТВУЮЩЕГО ---
Private Sub ProcessUpdatePerson(ws As Worksheet, wsN As Worksheet, bRow As Long, fStat As String, r As Long)
    Dim bStat As String, bEndStat As Variant, bEndIns As Variant
    Dim newStat As Boolean: newStat = False
    Dim newIns As Boolean: newIns = False
    Dim needRow As Boolean: needRow = False
    Dim comment As String: comment = ""

    bStat = wsBase.Cells(bRow, 3).Value
    bEndStat = wsBase.Cells(bRow, 5).Value
    bEndIns = wsBase.Cells(bRow, 8).Value

    ' Анализ расхождений
    If bStat <> fStat Then
        newStat = True
        needRow = True
        comment = "нужно зачетку"
    Elseif (RanksSheet.GetRankValue(bStat) > 0 And (bEndStat = "" Or bEndStat < Date)) Then
        needRow = True
        comment = "нужно зачетку"
    End If

    If (bEndIns = "" Or bEndIns < Date) Then
        newIns = True
        needRow = True
        comment = IIf(comment = "", "нужно страховку", "нужно страховку и зачетку")
    End If

    ' Копируем данные из Базы в основную строку
    wsBase.Range(wsBase.Cells(bRow, 1), wsBase.Cells(bRow, 8)).Copy Destination:=ws.Cells(r, 1)

    If needRow Then
        Dim nextR As Long: nextR = r + 1
        wsBase.Range(wsBase.Cells(bRow, 3), wsBase.Cells(bRow, 8)).Copy Destination:=ws.Cells(nextR, 3)

        ' Красим пустые в сиреневый (кроме ФИО, др)
        ws.Range(ws.Cells(nextR, 3), ws.Cells(nextR, 8)).Interior.Color = RGB(221, 160, 221)

        If newStat Then
            ws.Cells(nextR, 3).Value = fStat
            ws.Cells(nextR, 4).Value = ""
            ' Перекрашиваем разряд в голубой
            ws.Cells(nextR, 3).Interior.Color = RGB(173, 216, 230)
        End If
        If newIns Then
            ws.Cells(nextR, 6).Value = ""
            ws.Cells(nextR, 7).Value = ""
        End If

        ' В этой строке ФИО и ДР пустые по ТЗ
        Call ApplyLogicAndFormatting(ws, nextR)
    End If

    Call AddComment(wsN, ws.Cells(r, 1).Value, ws.Cells(r, 2).Value, fStat, comment)
    Call ApplyLogicAndFormatting(ws, r)
End Sub

Sub AddComment(wsN As Worksheet, fio As String, bd As String, st As String, txt As String)
    Dim nr As Long: nr = wsN.Cells(wsN.Rows.Count, 1).End(xlUp).Row + 1
    wsN.Cells(nr, 1) = fio: wsN.Cells(nr, 2) = bd: wsN.Cells(nr, 3) = st: wsN.Cells(nr, 4) = txt
End Sub

' ВСПОМОГАТЕЛЬНАЯ ПРОЦЕДУРА ОБНОВЛЕНИЯ СЧЕТЧИКА
Sub UpdateCounterAndFormat(ws As Worksheet)
    Dim countVal As Long: countVal = Application.WorksheetFunction.CountA(ws.Range("A5:A200"))

    ws.Range("A3").Value = "Количество участников: " & countVal
    ws.Columns("A:I").AutoFit
End Sub

' --- ФОРМУЛЫ И УФ (КИРИЛЛИЦА) ---
Private Sub ApplyLogicAndFormatting(ws As Worksheet, r As Long)
    ' Формула окончания разряда (ВПР по листу Разряды)
    ws.Cells(r, 5).FormulaLocal = "=ЕСЛИ(И(D" & r & "<>""""; C" & r & "<>""""); ДАТА(ГОД(D" & r & ")+ВПР(C" & r & ";Разряды!$B$1:$C$15;2;0);МЕСЯЦ(D" & r & ");ДЕНЬ(D" & r & "))-1;"""")"

    ' Формула окончания страховки
    ws.Cells(r, 8).FormulaLocal = "=ЕСЛИ(И(F" & r & "<>""""; G" & r & "<>""""); F" & r & "+G" & r & "-1;"""")"

    Dim rngStat As Range: Set rngStat = ws.Range("D" & r & ":E" & r) ' Диапазон Разряд
    Dim rngIns As Range: Set rngIns = ws.Range("F" & r & ":H" & r) ' Диапазон Страховка

    ' --- УДАЛЯЕМ старые правила ---
    rngStat.FormatConditions.Delete
    rngIns.FormatConditions.Delete

    ' Условное форматирование
    ' --- БЛОК РАЗРЯД ---
    ' Зеленый: Действует
    With rngStat.FormatConditions.Add(Type:=xlExpression, Formula1:="=И($E" & r & "<>""""; $E" & r & ">=СЕГОДНЯ()+10)")
        .Interior.Color = RGB(146, 208, 80)
        .StopIfTrue = True
    End With
    ' Оранжевый: Заканчивается через 10 дней
    With rngStat.FormatConditions.Add(Type:=xlExpression, Formula1:="=И($E" & r & "<>""""; $E" & r & ">СЕГОДНЯ())")
        .Interior.Color = RGB(255, 192, 0)
        .StopIfTrue = True
    End With
    ' Темно-красный: Нет или просрочен
    With rngStat.FormatConditions.Add(Type:=xlExpression, Formula1:="=$E" & r & "<>""""")
        .Interior.Color = RGB(128, 0, 0)
        .Font.Color = vbWhite
        .StopIfTrue = True
    End With

    ' --- БЛОК СТРАХОВКА ---
    ' Зеленый: Действует
    With rngIns.FormatConditions.Add(Type:=xlExpression, Formula1:="=И($H" & r & "<>""""; $H" & r & ">=СЕГОДНЯ()+10)")
        .Interior.Color = RGB(146, 208, 80)
        .StopIfTrue = True
    End With
    ' Желтый: Заканчивается через 10 дней
    With rngIns.FormatConditions.Add(Type:=xlExpression, Formula1:="=И($H" & r & "<>""""; $H" & r & ">СЕГОДНЯ())")
        .Interior.Color = vbYellow
        .StopIfTrue = True
    End With
    ' Ярко-красный: Нет или просрочена
    With rngIns.FormatConditions.Add(Type:=xlExpression, Formula1:="=$H" & r & "<>""""")
        .Interior.Color = vbRed
        .StopIfTrue = True
    End With

    ws.Range("B" & r).NumberFormat = "dd.mm.yyyy"
    ws.Range("D" & r & ":E" & r).NumberFormat = "dd.mm.yyyy"
    ws.Range("F" & r).NumberFormat = "dd.mm.yyyy"
    ws.Range("H" & r).NumberFormat = "dd.mm.yyyy"

    ws.Range("G" & r).NumberFormat = "0"
End Sub



