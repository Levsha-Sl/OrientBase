Attribute VB_Name = "BaseSheet"
Public wsBase As Worksheet
Private dictBaseCache As Object ' Кэш базы в памяти

Sub InitBase()
    If wsBase Is Nothing Then
        Set wsBase = BaseSheet.GetBaseSheet
    End If
End Sub

Private Function GetBaseSheet() As Worksheet
    Dim ws As Worksheet: Set ws = ModuleSheet.GetSheetByName("База")

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets. _
        Add(Before:=ThisWorkbook.Worksheets(2))
        ws.Name = "База"
        ' Создаем заголовки
        ws.Range("A5:I5").Value = _
        Array("ФИО", "День рож.", "Разряд", "дата раз.", _
        "окончание", "дата страх.", _
        "период", "окончание", _
        "изменено")
        ws.Range("A5:I5").Font.Bold = True
        ' Автоподбор ширины колонок
        ws.Columns("A:I").AutoFit

        Dim btnAddPerson As Object: Set btnAddPerson = ws.Buttons.Add( _
        Left:=ws.Cells(4, 1).Left, _
        Top:=ws.Cells(4, 1).Top, _
        Width:=ws.Cells(4, 1).Width, _
        Height:=ws.Cells(4, 1).Height)
        With btnAddPerson
            .OnAction = "AddPerson"
            .Caption = "Добавить человека"
        End With

        Dim btnImportBase As Object: Set btnImportBase = ws.Buttons.Add( _
        Left:=ws.Cells(2, 2).Left, _
        Top:=ws.Cells(2, 2).Top, _
        Width:=ws.Cells(2, 2).Width + ws.Cells(2, 3).Width, _
        Height:=30)
        With btnImportBase
            .OnAction = "ImportCSVToExcel"
            .Caption = "Импорт Базы"
        End With

        Dim btnImport As Object: Set btnImport = ws.Buttons.Add( _
        Left:=ws.Cells(2, 5).Left, _
        Top:=ws.Cells(2, 5).Top, _
        Width:=ws.Cells(2, 5).Width + ws.Cells(2, 6).Width, _
        Height:=30)
        With btnImport
            .OnAction = "Import.CSV"
            .Caption = "Загурзить csv_fjww"
        End With

        Dim btnExportBase As Object: Set btnExportBase = ws.Buttons.Add( _
        Left:=ws.Cells(2, 8).Left, _
        Top:=ws.Cells(2, 8).Top, _
        Width:=ws.Cells(2, 8).Width + ws.Cells(2, 9).Width, _
        Height:=30)
        With btnExportBase
            .OnAction = "ExportToCSV"
            .Caption = "Экспорт Базы"
        End With

        ActiveWindow.FreezePanes = False
        Rows(6).Select
        ActiveWindow.FreezePanes = True
    End If

    Set GetBaseSheet = ws
End Function

Public Sub LoadBaseCache()
    Dim i As Long, lastRow As Long
    Dim key As String

    Set dictBaseCache = CreateObject("Scripting.Dictionary")

    lastRow = wsBase.Cells(wsBase.Rows.Count, 1).End(xlUp).Row

    If lastRow < 5 Then Exit Sub

        For i = 5 To lastRow
            ' Ключ = ФИО + "#" + Дата рождения (уникальный идентификатор)
            key = Trim(wsBase.Cells(i, 1).Value) & "#" & Format(wsBase.Cells(i, 2).Value, "YYYY-MM-DD")
            dictBaseCache(key) = i ' Значение = номер строки в листе
        Next i
End Sub

Public Sub ClearBaseCache()
    Set dictBaseCache = Nothing
End Sub

Public Function PersonExists(fio As String, bd As Date) As Long
    Dim key As String

    ' Если кэш не загружен - загружаем
    If dictBaseCache Is Nothing Then
        Call LoadBaseCache
    End If

    ' Ищем в кэше (максимально быстро - словарь имеет O(1) поиск)
    key = fio & "#" & Format(bd, "YYYY-MM-DD")

    If dictBaseCache.Exists(key) Then
        PersonExists = dictBaseCache(key)
    Else
        PersonExists = 0
    End If
End Function

Public Sub ExportToCSV()
    Dim filePath As Variant
    Dim lastRow As Long, i As Long, j As Long
    Dim lineData As String, objStream As Object

    filePath = Application.GetSaveAsFilename(InitialFileName:="base.csv", _
    FileFilter:="CSV Files (*.csv), *.csv")
    If VarType(filePath) = vbBoolean Then Exit Sub

        Set objStream = CreateObject("ADODB.Stream")
        objStream.Charset = "utf-8"
        objStream.Open

        lastRow = wsBase.Cells(wsBase.Rows.Count, "A").End(xlUp).Row

        ' Сохраняем с 5 строки (заголовки + данные)
        For i = 5 To lastRow
            lineData = ""
            For j = 1 To 4 ' Столбцы A-D
                lineData = lineData & wsBase.Cells(i, j).Value & ";"
            Next j
            ' Столбцы F-G
            lineData = lineData & wsBase.Cells(i, 6).Value & ";" & wsBase.Cells(i, 7).Value

            objStream.WriteText lineData, 1 ' 1 = с переносом строки
        Next i

        objStream.SaveToFile filePath, 2 ' 2 = перезаписать
        objStream.Close
        MsgBox "Экспорт завершен", vbInformation
End Sub

Public Sub ImportCSVToExcel()
    Dim filePath As Variant
    Dim objStream As Object
    Dim lineData As String
    Dim cols() As String
    Dim lastRow As Long
    Dim i As Long

    filePath = Application.GetOpenFilename( _
    FileFilter:="CSV Files (*.csv), *.csv", _
    Title:="Выберите CSV файл для импорта (ADODB)")
    If VarType(filePath) = vbBoolean Then Exit Sub

        Set objStream = CreateObject("ADODB.Stream")
        objStream.Charset = "utf-8"
        objStream.Open
        objStream.LoadFromFile filePath

        Application.ScreenUpdating = False

        If Not objStream.EOS Then objStream.ReadText (-2)
            Do Until objStream.EOS
                lineData = objStream.ReadText(-2)

                ' Проверяем, что строка не пустая
                If Trim(lineData) <> "" Then
                    cols = Split(lineData, ";") ' Жестко заданный разделитель

                    ' Находим место для вставки
                    lastRow = wsBase.Cells(wsBase.Rows.Count, "A").End(xlUp).Row + 1
                    If lastRow < 6 Then lastRow = 6

                        wsBase.Cells(lastRow,1).Value = Trim(cols(0))
                        wsBase.Cells(lastRow,2).Value = Trim(cols(1))
                        wsBase.Cells(lastRow,3).Value = Trim(cols(2))
                        wsBase.Cells(lastRow,4).Value = Trim(cols(3))

                        wsBase.Cells(lastRow,6).Value = Trim(cols(4))
                        wsBase.Cells(lastRow,7).Value = Trim(cols(5))

                        Call ApplyLogicAndFormatting(lastRow)
                    End If
                Loop
                objStream.Close

                wsBase.Columns("A:I").AutoFit
                Application.ScreenUpdating = True

                MsgBox "Данные успешно добавлены!", vbInformation
End Sub

Private Sub AddPerson()
    Dim lastRow As Long: lastRow = wsBase.Cells(wsBase.Rows.Count, "A").End(xlUp).Row + 1
    If lastRow < 6 Then lastRow = 6

        wsBase.Cells(lastRow,1).Value = "Новый человек"

        Call ApplyLogicAndFormatting(lastRow)
End Sub

Public Sub ApplyLogicAndFormatting(r As Long)
    Dim rngStat As Range
    Dim rngIns As Range

    wsBase.Cells(r, 1).HorizontalAlignment = xlLeft
    wsBase.Cells(r, 2).HorizontalAlignment = xlLeft
    wsBase.Cells(r, 3).HorizontalAlignment = xlCenter
    wsBase.Cells(r, 4).HorizontalAlignment = xlLeft
    wsBase.Cells(r, 5).HorizontalAlignment = xlLeft
    wsBase.Cells(r, 6).HorizontalAlignment = xlRight
    wsBase.Cells(r, 7).HorizontalAlignment = xlLeft
    wsBase.Cells(r, 8).HorizontalAlignment = xlRight
    wsBase.Cells(r, 9).HorizontalAlignment = xlLeft

    ' Формула окончания разряда (ВПР по листу Разряды)
    wsBase.Cells(r, 5).FormulaLocal = _
    "=ЕСЛИ(И(D" & r & "<>""""; C" & r & "<>""""); " & _
    "ДАТА(ГОД(D" & r & ")+ВПР(C" & r & ";Разряды!$B$1:$C$15;2;0); " & _
    "МЕСЯЦ(D" & r & ");ДЕНЬ(D" & r & "))-1;"""")"
    ' Формула окончания страховки
    wsBase.Cells(r, 8).FormulaLocal = _
    "=ЕСЛИ(И(F" & r & "<>""""; G" & r & "<>""""); " & _
    "F" & r & "+G" & r & "-1;"""")"

    wsBase.Range("B" & r).NumberFormat = "dd.mm.yyyy"
    wsBase.Range("D" & r & ":E" & r).NumberFormat = "dd.mm.yyyy"
    wsBase.Range("F" & r).NumberFormat = "dd.mm.yyyy"
    wsBase.Range("H" & r).NumberFormat = "dd.mm.yyyy"

    Set rngStat = wsBase.Range("D" & r & ":E" & r)
    Set rngIns = wsBase.Range("F" & r & ":H" & r)

    rngStat.FormatConditions.Delete
    rngIns.FormatConditions.Delete

    ' =========================
    ' Красный ПРОСРОЧЕНО
    ' =========================
    With rngStat.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=И(ВПР($C" & r & ";Разряды!$B$1:$C$15;2;0)<>0;ИЛИ($E" & r & "<=СЕГОДНЯ(); $E" & r & "=""""))")
        .Interior.Color = RGB(192, 0, 0)
        .Font.Color = vbWhite
    End With

    With rngIns.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=ИЛИ($H" & r & "<=СЕГОДНЯ(); $H" & r & "="""")")
        .Interior.Color = vbRed
    End With

    ' =========================
    ' Желтый ЗАКАНЧИВАЕТСЯ < 10 ДНЕЙ
    ' =========================
    With rngStat.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=И($E" & r & ">=СЕГОДНЯ(); $E" & r & "<=СЕГОДНЯ()+10)")
        .Interior.Color = RGB(255, 192, 0)
    End With

    With rngIns.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=И($H" & r & ">=СЕГОДНЯ(); $H" & r & "<=СЕГОДНЯ()+10)")
        .Interior.Color = vbYellow
    End With

    ' =========================
    ' Зеленый НЕДАВНО ОБНОВЛЕНО (ТОЛЬКО ЕСЛИ ВСЁ ОК)
    ' =========================
    With rngStat.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=И($E" & r & ">СЕГОДНЯ()+10; $I" & r & ">=СЕГОДНЯ()-3)")
        .Interior.Color = RGB(146, 208, 80)
    End With

    With rngIns.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=И($H" & r & ">СЕГОДНЯ()+10; $I" & r & ">=СЕГОДНЯ()-3)")
        .Interior.Color = RGB(146, 208, 80)
    End With
End Sub