Attribute VB_Name = "ClubSheet"
Private Const COL_STATUS As String = "Статус"
Private Const COL_RANK As String = "Разряд"
Private Const COL_RANK_DATE As String = "дата_раз."
Private Const COL_RANK_EXPIRY As String = "окончание"
Private Const COL_INS_DATE As String = "дата_страх."
Private Const COL_PERIOD As String = "период"
Private Const COL_INS_EXPIRY As String = "окончaние"

' TODO
'
'
' под каждый клуб в пямяти его данные, все это не сохраняем, но работаем в кэше
' сохраняем данные,
' Подумать как это подвязать к импорту базы, так чтобы она как клуб грузилась...
' как клуб... при импорте участников кидать в кулб так и делать. Я гений

Public Function GetClubSheet(clubIndex As Long, club As ClubModule) As Worksheet
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Dim wsCl As Worksheet
    Set wsCl = ThisWorkbook.Worksheets.Add _
    (After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ' Имя листа: L1_2204_1530 (Коротко и уникально)

    Dim dataTime As Date: dataTime = club.timeStamp
    Dim clubName As String: clubName = club.clubName

    With wsCl
        .name = "L" & clubIndex & "_" & Format(dataTime, "ddmm_hhmmss")

        Dim headerData(1 To 3, 1 To 1) As Variant
        headerData(1, 1) = "Клуб: " & clubName
        headerData(2, 1) = "Загружено: " & Format(dataTime, "dd.mm.yyyy hh:mm:ss")
        headerData(3, 1) = "Количество участников: " & club.ParticipantCount
        .Range("A1:A3").value = headerData

        .Range("A4:J4").value = Array("ФИО", COL_STATUS, "id", "День рож.", COL_RANK, COL_RANK_DATE, COL_RANK_EXPIRY, COL_INS_DATE, COL_PERIOD, COL_INS_EXPIRY)
        .Range("A4:J4").Font.Bold = True

        Dim btn As Object: Set btn = .Buttons.Add( _
        Left:=.Cells(1, 5).Left, _
        Top:=.Cells(1, 5).Top, _
        Width:=.Cells(1, 5).Width + .Cells(1, 6).Width, _
        Height:=30)

        With btn
            .OnAction = "ClubSheet.BtnExportToBase"
            .Caption = "Выгрузить в базу"
        End With

        .Hyperlinks.Add Anchor:=.Range("B1"), Address:="", SubAddress:=ClubsSheet.SHEET_NAME & "!A" & clubIndex + 1, TextToDisplay:="<< К КЛУБАМ"
        .Columns("A:H").AutoFit
    End With

    ClubsListData.SetClubInCombinedList clubIndex
    Call setData(wsCl, ClubsListData.GetClubForShowcase(clubIndex))

    Set GetClubSheet = wsCl
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Function

Private Sub BtnExportToBase()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim idClub As Long: idClub =  Val(Mid$(ws.name, 2, Len(ws.name) - 11))
    ClubsListData.ExportToBase idClub
End Sub

Public Function ChangesShowcase(ws As Worksheet,Target As Range) As Boolean
    On Error Goto ErrorHandler
        Dim Result As Boolean: Result = True
        Dim rowIdx As Long: rowIdx = Target.row
        Dim colIdx As Long: colIdx = Target.Column

        If (colIdx >= 5 And colIdx <> 7 And colIdx <= 9) And rowIdx >= 5 Then
            Application.EnableEvents = False
            Dim status As String
            status = ws.Cells(rowIdx, 2).value
            If (status <> ClubsListData.arrStatus(0)) Then

                Dim baseId As Long, fnm As String, birth As Date
                Dim stat As String, dateStat As Variant, dateIns As Variant, period As Variant
                Dim hasId As Boolean, hasData As Boolean

                baseId = ws.Cells(rowIdx, 3).value
                fnm = ws.Cells(rowIdx, 1).value
                stat = ws.Cells(rowIdx, 5).value

                Dim cells As Variant
                cells = ws.Cells(rowIdx, 4).value
                birth = IIf(IsEmpty(cells) Or cells = "", 0, CDate(cells))
                cells = ws.Cells(rowIdx, 6).value
                dateStat = IIf(IsEmpty(cells) Or cells = "", Empty, CDate(cells))
                cells = ws.Cells(rowIdx, 8).value
                dateIns = IIf(IsEmpty(cells) Or cells = "", Empty, CDate(cells))
                cells = ws.Cells(rowIdx, 9).value
                period = IIf(IsEmpty(cells) Or cells = "", Empty, CLng(cells))

                hasData = (fnm <> "") And (birth <> 0) And (status <> "")
                Dim keyParticipant As String: keyParticipant = fnm & "#" & birth & "#" & status
                Dim idClub As Long: idClub =  Val(Mid$(ws.name, 2, Len(ws.name) - 11))

                If hasData Then ' если есть все данные
                    ClubsListData.UpdateParticipant idClub, keyParticipant, stat, dateStat, dateIns, period  
                End If
            Else 
                Result = False
            End If
        End If
 CleanExit:
        ChangesShowcase = Result
        Application.EnableEvents = True
     Exit Function
 ErrorHandler:
        Select Case Err.Number
         Case 13
            MsgBox "Ошибка: Неверный формат данных!", vbCritical, "Ошибка 13"
         Case vbObjectError + 1401
            MsgBox Err.Description, vbExclamation, "Ошибка изменения данных"
         Case Else
            MsgBox "Произошла ошибка при сохранении изменений: " & Err.Description, vbCritical, "Ошибка"
        End Select
        Result = False
        Resume CleanExit
End Function

Private Sub setData(ws As Worksheet, martArr As Variant)
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ws.Range("A5").Resize(UBound(martArr, 1), 10).value = martArr
    ws.Columns("A:J").AutoFit

    setLogicTable ws, UBound(martArr, 1)

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Sub setLogicTable(ws As Worksheet, size As Long)
    Dim tbl As ListObject
    Set tbl = ModuleSheet.GetTableByName(ws, ws.name)

    If tbl Is Nothing Then
        Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A4:J" & (4 + size)), , xlYes)

        With tbl
            .name = ws.name
            .ShowTableStyleRowStripes = False

            .ListColumns(4).Range.NumberFormat = "dd.mm.yyyy"
            .ListColumns(COL_RANK_DATE).Range.NumberFormat = "dd.mm.yyyy"
            .ListColumns(COL_RANK_EXPIRY).Range.NumberFormat = "dd.mm.yyyy"
            .ListColumns(COL_INS_DATE).Range.NumberFormat = "dd.mm.yyyy"
            .ListColumns(COL_INS_EXPIRY).Range.NumberFormat = "dd.mm.yyyy"
        End With

        If Not (tbl.DataBodyRange Is Nothing) Then
            Dim rngRank As Range
            Dim rngIns As Range
            Dim rngStatus As Range
            Dim firstRow As Long

            Dim colRank As String
            Dim colRankExpiry As String
            Dim colInsExpiry As String
            Dim colStatus As String

            Set rngRank = ws.Range(tbl.ListColumns(COL_RANK).DataBodyRange, tbl.ListColumns(COL_RANK_EXPIRY).DataBodyRange)
            Set rngIns = ws.Range(tbl.ListColumns(COL_INS_DATE).DataBodyRange, tbl.ListColumns(COL_INS_EXPIRY).DataBodyRange)
            Set rngStatus = ws.Range(tbl.ListColumns(COL_RANK).DataBodyRange, tbl.ListColumns(COL_INS_EXPIRY).DataBodyRange)
            firstRow = tbl.DataBodyRange.row

            colRank = Split(tbl.ListColumns(COL_RANK).DataBodyRange.Cells(1).Address, "$")(1)
            colRankExpiry = Split(tbl.ListColumns(COL_RANK_EXPIRY).DataBodyRange.Cells(1).Address, "$")(1)
            colInsExpiry = Split(tbl.ListColumns(COL_INS_EXPIRY).DataBodyRange.Cells(1).Address, "$")(1)
            colStatus = Split(tbl.ListColumns(COL_STATUS).DataBodyRange.Cells(1).Address, "$")(1)
            
            With rngStatus
                With .FormatConditions.Add(Type:=xlExpression, _
                    Formula1:="=$" & colStatus & firstRow & "=""" & ClubsListData.arrStatus(0) & """")
                    .Interior.Color = RGB(143, 243, 251)
                End With
            End With

            With rngRank
                Dim ranksVlookup As String: ranksVlookup = "ВПР($" & colRank & firstRow & ";" & RanksData.SHEET_NAME &"!$A$1:$B$" & RanksData.getMaxId & ";2;0)"

                ' Красное (просрочено или пусто)
                With .FormatConditions.Add(Type:=xlExpression, _
                    Formula1:="=И(" & ranksVlookup & "<>0;" & _
                    "ИЛИ($" & colRankExpiry & firstRow & "<=СЕГОДНЯ();" & _
                    "$" & colRankExpiry & firstRow & "=""""))")
                    .Interior.Color = RGB(192, 0, 0)
                    .Font.Color = vbWhite
                End With
                ' Жёлтое (0–10 дней)
                With .FormatConditions.Add(Type:=xlExpression, _
                    Formula1:="=И($" & colRankExpiry & firstRow & ">=СЕГОДНЯ();" & _
                    "$" & colRankExpiry & firstRow & "<=СЕГОДНЯ()+10)")
                    .Interior.Color = RGB(255, 192, 0)
                End With
                ' Зелёное (>10 дней и недавно изменено)
                'With .FormatConditions.Add(Type:=xlExpression, _
                '    Formula1:="=И($" & colRankExpiry & firstRow & ">СЕГОДНЯ()+10;" & _
                '    "$" & colTimeStamp & firstRow & ">=СЕГОДНЯ()-3)")
                '    .Interior.Color = RGB(146, 208, 80)
                'End With
            End With

            With rngIns
                ' Красное
                With .FormatConditions.Add(Type:=xlExpression, _
                    Formula1:="=ИЛИ($" & colInsExpiry & firstRow & "<=СЕГОДНЯ();" & _
                    "$" & colInsExpiry & firstRow & "="""")")
                    .Interior.Color = RGB(255, 192, 192)
                End With
                ' Жёлтое
                With .FormatConditions.Add(Type:=xlExpression, _
                    Formula1:="=И($" & colInsExpiry & firstRow & ">=СЕГОДНЯ();" & _
                    "$" & colInsExpiry & firstRow & "<=СЕГОДНЯ()+10)")
                    .Interior.Color = RGB(255, 255, 153)
                End With
                ' Зелёное
                'With .FormatConditions.Add(Type:=xlExpression, _
                '    Formula1:="=И($" & colInsExpiry & firstRow & ">СЕГОДНЯ()+10;" & _
                '    "$" & colTimeStamp & firstRow & ">=СЕГОДНЯ()-3)")
                '    .Interior.Color = RGB(146, 208, 80)
                'End With
            End With
        End If
    End If 

    With tbl
        .ListColumns(COL_RANK_EXPIRY).DataBodyRange.FormulaLocal = _
        "=ModuleSheet.CalculateRankExpiry([@[" & COL_RANK_DATE & "]]; [@" & COL_RANK & "])"
        .ListColumns(COL_INS_EXPIRY).DataBodyRange.FormulaLocal = _
        "=ModuleSheet.CalculateInsuranceExpiry([@[" & COL_INS_DATE & "]]; [@" & COL_PERIOD & "])"
    End With
End Sub 