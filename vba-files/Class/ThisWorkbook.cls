VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Workbook_Open()
    ModuleSheet.InitApp  
End Sub

Private Sub Workbook_SheetSelectionChange(Byval Sh As Object, Byval Target As Range)
    If Sh.name = "Клубы" Then
        If Target.Column <> 5 Or Target.Count > 1 Then Exit Sub
            If Target.Value <> "УДАЛИТЬ КЛУБ" Or Target.row = 1 Then Exit Sub

                Dim mainRef As String, needRef As String

                mainRef = ModuleSheet.GetSheetFromLink(Sh.Cells(Target.row, 1))
                needRef = ModuleSheet.GetSheetFromLink(Sh.Cells(Target.row, 2))

                If MsgBox("Удалить листы этой загрузки?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub

                    Application.DisplayAlerts = False

                    ModuleSheet.DeleteSheetIfExists mainRef
                    ModuleSheet.DeleteSheetIfExists needRef

                    Target.EntireRow.Delete

                    Application.DisplayAlerts = True
                End If
End Sub

Private Sub Workbook_SheetFollowHyperlink(Byval Sh As Object, Byval Target As Hyperlink)
    Dim sheetName As String
    sheetName = Split(Target.SubAddress, "!")(0)
    sheetName = Replace(sheetName, "'", "")

    If Not ModuleSheet.SheetExists(sheetName) Then
        MsgBox "Лист [" & sheetName & "] не найден. Ссылка удалена.", vbExclamation
        Target.Range.ClearContents
    End If
End Sub

Private Sub Workbook_SheetChange(Byval Sh As Object, Byval Target As Range)
    Dim rowIdx As Long: rowIdx = Target.row
    Dim colIdx As Long: colIdx = Target.Column

    If Sh.name = "База" Then
        ' Проверяем, что изменение в пределах A-H и ниже заголовков
        If colIdx <= 8 And rowIdx >= 6 Then

            Dim hasFirstThree As Boolean
            Dim hasAnyOther As Boolean

            ' 1. Проверяем первые три (A, B, C) — должны быть все заполнены
            ' CountA вернет 3, если ячейки A, B и C не пустые
            hasFirstThree = (Application.WorksheetFunction.CountA(Sh.Range(Sh.Cells(rowIdx, 1), Sh.Cells(rowIdx, 3))) = 3)

            ' 2. Проверяем остальные (D, E, F, G, H) — должна быть хотя бы одна
            hasAnyOther = (Application.WorksheetFunction.CountA(Sh.Range(Sh.Cells(rowIdx, 4), Sh.Cells(rowIdx, 8))) > 0)

            Application.EnableEvents = False

            ' Если оба условия выполнены — ставим дату
            If hasFirstThree And hasAnyOther Then
                Sh.Cells(rowIdx, 9).Value = Now
            Else
                ' Если условия перестали соблюдаться (например, стерли ФИО) — очищаем дату
                Sh.Cells(rowIdx, 9).ClearContents
            End If

            Application.EnableEvents = True
        End If
    End If

    If Sh.name = "Разряды" Then 

        ' TODO in RankSheet
        RanksSheet.ChangesMart()
        ' TODO export
        If colIdx >= 1 And colIdx <= 3 And rowIdx >= 2 Then

            Application.EnableEvents = False

            Dim rankId As Variant, rankName As String, periodValue As Variant
            Dim hasRankName As Boolean, hasPeriodValue As Boolean

            rankId = Sh.Cells(rowIdx, 1).Value
            rankName = Sh.Cells(rowIdx, 2).Value
            periodValue = Sh.Cells(rowIdx, 3).Value

            hasRankName = (rankName <> "")
            hasPeriodValue = (periodValue <> "")

            ' TODO update deleted(=Empty Ranks) insert to RanksData
            Dim operation As String

            ' Определяем операцию по состоянию ID и данных
            If rankId = "" Then
                ' ID пуст: может быть INSERT если есть хотя бы данные
                If hasRankName Or hasPeriodValue Then
                    operation = "INSERT"
                Else
                    ' Пустая строка - ничего не делаем
                    Application.EnableEvents = True
                 Exit Sub
                End If
            Else
                ' ID не пуст: UPDATE если данные есть, DELETE если нет
                If hasRankName And hasPeriodValue Then
                    operation = "UPDATE"
                Elseif Not hasRankName And Not hasPeriodValue Then
                    operation = "DELETE"
                Else
                    ' Частичные данные - ничего не делаем
                    Application.EnableEvents = True
                 Exit Sub
                End If
            End If
            
            Application.EnableEvents = True
        
        End If
        ' ELSE Application.Undo
   
    End If
End Sub
