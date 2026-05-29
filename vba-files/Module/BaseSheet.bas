Attribute VB_Name = "BaseSheet"
Public wsBase As Worksheet
Private tblBase As ListObject

Private Const COL_RANK As String = "Разряд"
Private Const COL_RANK_DATE As String = "дата_раз."
Private Const COL_RANK_EXPIRY As String = "окончание"
Private Const COL_INS_DATE As String = "дата_страх."
Private Const COL_PERIOD As String = "период"
Private Const COL_INS_EXPIRY As String = "окончaние"
Private Const COL_TIME_STEMP As String = "изменено"

Public Sub Init()
    Set wsBase = GetBaseSheet
End Sub

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
    addedYears = Ranks.GetRankValue(targetRank)
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

Public Function ChangesMart(Target As Range) As Boolean
    On Error Goto ErrorHandler
        Dim result As Boolean: result = True
        Dim rowIdx As Long: rowIdx = Target.row
        Dim colIdx As Long: colIdx = Target.Column

        If (colIdx >= 2 And colIdx <> 6 And colIdx <= 8) And rowIdx >= 6 Then
            Application.EnableEvents = False

            Dim baseId As Long, fnm As String, birth As Date
            Dim stat As String, dateStat As Variant, dateIns As Variant, period As Variant, timeStamp As Variant
            Dim hasId As Boolean, hasData As Boolean

            baseId = wsBase.Cells(rowIdx, 1).Value
            fnm = wsBase.Cells(rowIdx, 2).Value
            birth = IIf(IsEmpty(wsBase.Cells(rowIdx, 3).Value) Or wsBase.Cells(rowIdx, 3).Value = "", 0, CDate(wsBase.Cells(rowIdx, 3).Value))

            stat = wsBase.Cells(rowIdx, 4).Value
            dateStat = IIf(IsEmpty(wsBase.Cells(rowIdx, 5).Value) Or wsBase.Cells(rowIdx, 5).Value = "", Empty, CDate(wsBase.Cells(rowIdx, 5).Value))
            dateIns = IIf(IsEmpty(wsBase.Cells(rowIdx, 7).Value) Or wsBase.Cells(rowIdx, 7).Value = "", Empty, CDate(wsBase.Cells(rowIdx, 7).Value))
            period = IIf(IsEmpty(wsBase.Cells(rowIdx, 8).Value) Or wsBase.Cells(rowIdx, 8).Value = "", Empty, CLng(wsBase.Cells(rowIdx, 8).Value))
            timeStamp = Format$(Now, "dd.mm.yyyy")

            hasId = (baseId > 0)
            hasData = (fnm <> "") And (birth <> 0)

            If Not hasId Then
                ' ID пуст: INSERT если есть все данные
                If hasData Then
                    wsBase.Cells(rowIdx, 1).Value = Base.CreateBase(fnm, birth, stat, dateStat, dateIns, period, timeStamp)
                    wsBase.Cells(rowIdx,10).Value = timeStamp
                End If
            Else
                ' ID не пуст: UPDATE если данные есть, DELETE если нет
                If hasData Then
                    Call Base.UpdateBase(fnm, birth, stat, dateStat, dateIns, period, timeStamp, baseId)
                    wsBase.Cells(rowIdx,10).Value = timeStamp
                Else
                    wsBase.Cells(rowIdx, 1).Value = IIf(Base.DeleteBase(baseId), "", baseId)
                End If
            End If
        End If
 CleanExit:
        ChangesMart = result
        Application.EnableEvents = True
     Exit Function
 ErrorHandler:
        Select Case Err.Number
         Case vbObjectError + 1201
            MsgBox Err.Description, vbExclamation, "Ошибка изменения данных"
         Case Else
            MsgBox "Произошла ошибка при сохранении изменений: " & Err.Description, vbCritical, "Ошибка"
        End Select
        result = False
        Resume CleanExit
End Function

Public Sub AcceptBaseData(martArr As Variant)
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    wsBase.Range("A6").Resize(Base.getMaxId, 10).Value = martArr
    wsBase.Columns("A:J").AutoFit

    Set tblBase = GetTableBase

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Function GetBaseSheet() As Worksheet
    On Error Goto ErrorHandler
        Application.EnableEvents = False
        Application.ScreenUpdating = False       
        Dim ws As Worksheet: Set ws = ModuleSheet.GetSheetByName("База")

        If ws Is Nothing Then
            Set ws = ThisWorkbook.Worksheets. Add(Before:=ThisWorkbook.Worksheets(2))

            With ws
                .Name = "База"
                .Range("A5:J5").Value = Array("id","ФИО", "День рож.", COL_RANK, COL_RANK_DATE, COL_RANK_EXPIRY, COL_INS_DATE, COL_PERIOD, COL_INS_EXPIRY, COL_TIME_STEMP)

                With .Range("A5:J5")
                    .Font.Bold = True
                End With

                With .Columns("A:J")
                    .AutoFit
                    .HorizontalAlignment = xlLeft
                End With       

                Dim btnSave As Object: Set btnSave = ws.Buttons.Add( _
                Left:=ws.Cells(1, 5).Left, _
                Top:=ws.Cells(1, 5).Top, _
                Width:=ws.Cells(1, 5).Width + ws.Cells(1, 6).Width, _
                Height:=ws.Cells(1, 1).Height)
                With btnSave
                    .OnAction = "BaseSheet.BtnSave"
                    .Caption = "Сохранить"
                End With

            End With
        End If

        Set GetBaseSheet = ws
 CleanExit:
        Application.EnableEvents = True
        Application.ScreenUpdating = True
     Exit Function
 ErrorHandler:
        MsgBox "Ошибка при создании листа спортсменнов: " & Err.Description, vbCritical
        Resume CleanExit
End Function

Private Sub BtnSave()
    If Not tblBase.DataBodyRange Is Nothing Then 
        tblBase.DataBodyRange.Delete
    End If
    Call Base.SaveChanges
End Sub

Private Function GetTableBase() As ListObject
    Dim tbl As ListObject
    Set tbl = ModuleSheet.GetTableByName(wsBase, "List sportsman")
    If tbl Is Nothing Then
        Set tbl = wsBase.ListObjects.Add(xlSrcRange, wsBase.Range("A5:J" & (5 + Base.getMaxId)), , xlYes)
        With tbl
            .Name = "List sportsman"
            .ShowTableStyleRowStripes = False

            .ListColumns(3).Range.NumberFormat = "dd.mm.yyyy"
            .ListColumns(COL_RANK_DATE).Range.NumberFormat = "dd.mm.yyyy"
            .ListColumns(COL_RANK_EXPIRY).Range.NumberFormat = "dd.mm.yyyy"
            .ListColumns(COL_INS_DATE).Range.NumberFormat = "dd.mm.yyyy"
            .ListColumns(COL_INS_EXPIRY).Range.NumberFormat = "dd.mm.yyyy"
            .ListColumns(COL_TIME_STEMP).Range.NumberFormat = "dd.mm.yyyy"
        End With

        If Not (tbl.DataBodyRange Is Nothing) Then
            Dim rngStat As Range
            Dim rngIns As Range
            Dim firstRow As Long

            Dim colRank As String
            Dim colRankExpiry As String
            Dim colInsExpiry As String
            Dim colTimeStamp As String

            Set rngStat = wsBase.Range( tbl.ListColumns(COL_RANK).DataBodyRange, tbl.ListColumns(COL_RANK_EXPIRY).DataBodyRange)
            Set rngIns = wsBase.Range( tbl.ListColumns(COL_INS_DATE).DataBodyRange, tbl.ListColumns(COL_INS_EXPIRY).DataBodyRange)
            firstRow = tbl.DataBodyRange.Row

            colRank = Split(tbl.ListColumns(COL_RANK).DataBodyRange.Cells(1).Address, "$")(1)
            colRankExpiry = Split(tbl.ListColumns(COL_RANK_EXPIRY).DataBodyRange.Cells(1).Address, "$")(1)
            colInsExpiry = Split(tbl.ListColumns(COL_INS_EXPIRY).DataBodyRange.Cells(1).Address, "$")(1)
            colTimeStamp = Split(tbl.ListColumns(COL_TIME_STEMP).DataBodyRange.Cells(1).Address, "$")(1)

            With rngStat
                Dim ranksVlookup As String: ranksVlookup = "ВПР($" & colRank & firstRow & ";Ranks!$A$1:$B$" & Ranks.getMaxId & ";2;0)"

                .FormatConditions.Delete
                ' Красное (просрочено или пусто)
                With .FormatConditions.Add( Type:=xlExpression, _
                    Formula1:="=И(" & ranksVlookup &"<>0;" & _
                    "ИЛИ($" & colRankExpiry & firstRow & "<=СЕГОДНЯ();" & _
                    "$" & colRankExpiry & firstRow & "=""""))")
                    .Interior.Color = RGB(192, 0, 0)
                    .Font.Color = vbWhite
                End With
                ' Жёлтое (0–10 дней)
                With .FormatConditions.Add( Type:=xlExpression, _
                    Formula1:="=И($" & colRankExpiry & firstRow & ">=СЕГОДНЯ();" & _
                    "$" & colRankExpiry & firstRow & "<=СЕГОДНЯ()+10)")
                    .Interior.Color = RGB(255, 192, 0)
                End With
                ' Зелёное (>10 дней и недавно изменено)
                With .FormatConditions.Add( Type:=xlExpression, _
                    Formula1:="=И($" & colRankExpiry & firstRow & ">СЕГОДНЯ()+10;" & _
                    "$" & colTimeStamp & firstRow & ">=СЕГОДНЯ()-3)")
                    .Interior.Color = RGB(146, 208, 80)
                End With
            End With

            With rngIns
                .FormatConditions.Delete
                ' Красное
                With .FormatConditions.Add( Type:=xlExpression, _
                    Formula1:="=ИЛИ($" & colInsExpiry & firstRow & "<=СЕГОДНЯ();" & _
                    "$" & colInsExpiry & firstRow & "="""")")
                    .Interior.Color = RGB(255, 192, 192)
                End With
                ' Жёлтое
                With .FormatConditions.Add( Type:=xlExpression, _
                    Formula1:="=И($" & colInsExpiry & firstRow & ">=СЕГОДНЯ();" & _
                    "$" & colInsExpiry & firstRow & "<=СЕГОДНЯ()+10)")
                    .Interior.Color = RGB(255, 255, 153)
                End With
                ' Зелёное
                With .FormatConditions.Add( Type:=xlExpression, _
                    Formula1:="=И($" & colInsExpiry & firstRow & ">СЕГОДНЯ()+10;" & _
                    "$" & colTimeStamp & firstRow & ">=СЕГОДНЯ()-3)")
                    .Interior.Color = RGB(146, 208, 80)
                End With
            End With
        End If
    End If

    With tbl
        .ListColumns(COL_RANK_EXPIRY).DataBodyRange.FormulaLocal = _
        "=BaseSheet.CalculateRankExpiry([@[" & COL_RANK_DATE & "]]; [@" & COL_RANK & "])"
        .ListColumns(COL_INS_EXPIRY).DataBodyRange.FormulaLocal = _
        "=BaseSheet.CalculateInsuranceExpiry([@[" & COL_INS_DATE & "]]; [@" & COL_PERIOD & "])"
    End With

    Set GetTableBase = tbl
End Function


' TODO пока не трогать
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

                    End If
                Loop
                objStream.Close

                wsBase.Columns("A:I").AutoFit
                Application.ScreenUpdating = True

                MsgBox "Данные успешно добавлены!", vbInformation
End Sub